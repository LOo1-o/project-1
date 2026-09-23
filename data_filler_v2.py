# data_filler_v2.py - Новая версия с умным поиском и нормализацией
"""
Версия 2 заполнения данных:
1. Использует find_column_by_year вместо жестких индексов
2. Использует fuzzy matching для поиска строк
3. Применяет нормализацию данных перед вставкой
4. Логирует все действия для отладки
"""

import re
from pathlib import Path
import pandas as pd
from typing import Dict, Set, Tuple, Optional

from config import canonical_okved, find_okved_code, set_paragraph_text_keep_format, get_cleaned_cell_text
from config_v2 import (
    load_column_mapping_v2,
    build_column_mapping_v2_from_excel,
    load_indicator_display_names,
)
from mo import load_mo_map, canonical_mo, find_mo_code
from category_mapping import CATEGORY_PREFIX_KEYWORDS, canonical_category, detect_category_prefix
from smart_loader import (
    find_column_by_year,
    find_row_by_fuzzy_match,
    normalize_text,
    get_cell_value_safely
)
from data_normalizer import clean_excel_value_for_word
from table_manager import TableManager


_YEAR_RE = re.compile(r"\b(20\d{2})\b")
_SOFT_WRAP_HYPHEN_RE = re.compile(r"(?<=\w)-\s*(?=\w)", re.UNICODE)


def _dewrap_hyphen(text: str) -> str:
    """
    Убирает "мягкий" перенос слова, который Excel вставляет в узкой ячейке
    (например, "преды-дущего" вместо "предыдущего"). Без этого маркеры периода
    вроде "предыдущ"/"отчет" не находятся как подстрока и суффикс года
    ошибочно откатывается на дефолт вместо реально определённого года.
    """
    return _SOFT_WRAP_HYPHEN_RE.sub("", text)


def _extract_report_year_from_excel(df: pd.DataFrame) -> Optional[int]:
    """Находит отчетный год в верхних строках Excel-файла."""
    if df.empty:
        return None

    years = []

    # pd.read_excel(path) без header=None использует первую строку листа
    # (например, "Баланс организаций за 2024 год") как df.columns, а не как
    # строку данных — поэтому такой заголовок никогда не встретится в
    # df.iloc[...] и год нужно дополнительно искать среди названий столбцов,
    # иначе report_year всегда будет None.
    for col in df.columns:
        years.extend(int(match.group(1)) for match in _YEAR_RE.finditer(str(col)))

    max_rows = min(10, len(df))
    for row_idx in range(max_rows):
        for value in df.iloc[row_idx].tolist():
            text = str(value)
            years.extend(int(match.group(1)) for match in _YEAR_RE.finditer(text))

    return max(years) if years else None


def _detect_year_suffix_for_column(df: pd.DataFrame, col_idx: Optional[int], default_suffix: str) -> str:
    """
    Определяет фактический двухзначный год для столбца.

    В T24-файлах в шапке обычно указан только отчетный год файла (например, 2024),
    а в строке столбца — "предыдущего года" / "отчетного года". Поэтому нельзя
    жестко сохранять значения как *_22 / *_23: для T24 это должны быть *_23 / *_24.
    """
    if col_idx is None or df.empty or col_idx >= len(df.columns):
        return default_suffix

    report_year = _extract_report_year_from_excel(df)
    previous_year = report_year - 1 if report_year else None

    max_rows = min(12, len(df))
    column_text = " ".join(_dewrap_hyphen(str(df.iloc[row_idx, col_idx]).lower()) for row_idx in range(max_rows))

    # Слова периода разбираем ПЕРЕД явными годами. Заголовок файла
    # («Баланс организаций за 2025 год») в выгрузках Росстата лежит в одном
    # из столбцов с данными, и его год попадал в текст этого столбца. Для
    # столбца «на конец предыдущего года» это давало отчётный год вместо
    # предыдущего: оба года показателя получали один суффикс, схлопывались
    # в один ключ, и данные за первый год терялись. Слово «предыдущего»
    # относится именно к своему столбцу, а год из заголовка — ко всему
    # файлу, поэтому приоритет у слова.
    if report_year:
        if 'предыдущ' in column_text or 'начало' in column_text:
            return str(previous_year)[-2:]
        if 'отчет' in column_text or 'текущ' in column_text or 'конец' in column_text:
            return str(report_year)[-2:]

    explicit_years = [int(match.group(1)) for match in _YEAR_RE.finditer(column_text)]
    if explicit_years:
        return str(explicit_years[-1])[-2:]

    return default_suffix


_DECIMAL_VALUE_RE = re.compile(r'\d[.,]\d')


def _column_has_decimal_values(df: pd.DataFrame, col_idx: Optional[int]) -> bool:
    """Есть ли в столбце Excel хоть одно значение с десятичным разделителем.

    Раньше единообразие "все значения столбца с одним знаком после запятой"
    было включено только для таблиц с "оборачиваемость" в названии — узкий,
    привязанный к конкретной таблице частный случай. Но проблема на самом
    деле про сам СТОЛБЕЦ: если Excel хранит "205.3" в одной строке и просто
    "4" (без ".0") в другой строке того же столбца, оба значения одного
    показателя — и в бюллетене они должны выглядеть единообразно
    ("205,3" и "4,0"), а не как будто это разные по точности величины.
    """
    if col_idx is None or df.empty or col_idx >= len(df.columns):
        return False
    for value in df.iloc[:, col_idx]:
        if pd.isna(value):
            continue
        if _DECIMAL_VALUE_RE.search(str(value)):
            return True
    return False


def pre_load_all_excel_data_v2(excel_dir: Path, table_source_mapping: Dict,
                               okved_codes_set: Set[str], okved_name_to_code: Dict[str, str],
                               column_mapping_path: Path,
                               mo_map_path: Optional[Path] = None,
                               use_fuzzy_match: bool = True, fuzzy_threshold: float = 0.80,
                               errors_report_path: Optional[Path] = None,
                               implausible_report_path: Optional[Path] = None) -> Tuple[Dict, dict]:
    """
    Загружает все данные из Excel с использованием умного поиска.
    
    Args:
        excel_dir: Путь к папке с Excel файлами
        table_source_mapping: Маппинг таблиц → файлы Excel
        okved_codes_set: Набор кодов ОКВЭД для фильтрации
        column_mapping_path: Путь к column_mapping_v2.csv
        use_fuzzy_match: Использовать нечеткое совпадение для поиска строк
        fuzzy_threshold: Порог сходства (0-1) для нечеткого совпадения
    
    Returns:
        (master_data[okved][indicator_year] = value, stats_dict)
    """
    print("\n📊 Шаг 3: Загрузка данных из Excel (УМНЫЙ поиск)")
    
    try:
        _, indicator_to_excel, indicator_to_file, _, indicator_keywords, _ = load_column_mapping_v2(str(column_mapping_path))
    except FileNotFoundError:
        print(f"⚠️ Файл {column_mapping_path} не найден. Генерируем его из Excel...")
        build_column_mapping_v2_from_excel(excel_dir, table_source_mapping, column_mapping_path)
        _, indicator_to_excel, indicator_to_file, _, indicator_keywords, _ = load_column_mapping_v2(str(column_mapping_path))
    except Exception as e:
        print(f"⚠️ Ошибка загрузки маппингов v2: {e}, используем fallback...")
        from config import load_column_mapping
        _, indicator_to_excel, indicator_to_file, _ = load_column_mapping(str(column_mapping_path))
        indicator_keywords = {k: {'2022': [], '2023': []} for k in indicator_to_file.keys()}
    
    # Импортируем функции канонизации
    from config import canonical_okved
    mo_name_to_mo_cleaned = {}
    if mo_map_path is not None:
        _, mo_name_to_mo_cleaned = load_mo_map(mo_map_path)
    
    master_data = {}
    conflicts = []
    implausible_values = []
    # Человекочитаемые названия показателей: по ним видно, что показатель
    # считает штуки, и значит дробное значение у него — дефект выгрузки.
    indicator_display_names = load_indicator_display_names(column_mapping_path)
    stats = {
        'files_processed': 0,
        'errors': [],
        'conflicts': 0,
        'found_by_keyword': 0,
        'found_by_hardcode': 0,
        'found_by_fuzzy': 0,
        'not_found': 0,
    }
    
    excel_files_to_load = set(table_source_mapping.values())

    # Имена файлов год от года (и от территории к территории) меняются —
    # поэтому привязываем формат-специфичные правила не к конкретному имени
    # файла, а к названию таблицы из table_source_data_mapping.csv, которое
    # гораздо стабильнее физического имени Excel-файла.
    table_names_by_file: Dict[str, list] = {}
    for table_name, fname in table_source_mapping.items():
        table_names_by_file.setdefault(fname, []).append(table_name)

    for filename in excel_files_to_load:
        excel_path = excel_dir / filename
        if not excel_path.exists():
            print(f"⚠️ Файл не найден: {filename}")
            stats['errors'].append(f"File not found: {filename}")
            continue
        
        try:
            print(f"\n📥 Загружаем: {filename}")
            df = pd.read_excel(excel_path)
            
            if df.empty:
                print(f"⚠️ Excel файл пуст: {filename}")
                stats['errors'].append(f"Empty file: {filename}")
                continue
            
            is_mo_file = 'mo' in filename.lower()
            header_scan_rows = _detect_header_scan_depth(excel_path)
            category_prefix = detect_category_prefix(excel_path)
            mo_codes_set = set(mo_name_to_mo_cleaned.values()) if mo_name_to_mo_cleaned else None

            if category_prefix:
                # "Категорийный" файл (ОПФ/форма собственности): строки размечены
                # не кодами ОКВЭД/МО, а заголовками категорий, и код "переносится"
                # с заголовка на следующую строку данных — см. _infer_category_entity_keys.
                print(f"ℹ️ {filename}: категорийная классификация (префикс {category_prefix})")
                df['__entity_key__'] = _infer_category_entity_keys(df, category_prefix)
            else:
                code_col_idx = _detect_code_column(df, okved_codes_set, mo_codes_set, is_mo_file)
                if code_col_idx is None:
                    print(f"ℹ️ Не удалось автоматически определить колонку с кодами в файле {filename}. Попробуем использовать первую колонку как имена.")
                    code_col_idx = 0
                else:
                    print(f"ℹ️ Автоопределена колонка с кодами в файле {filename}: {code_col_idx + 1}")

                df['__entity_key__'] = df.apply(
                    lambda row: _infer_entity_key(row, code_col_idx, is_mo_file, okved_codes_set, okved_name_to_code, mo_name_to_mo_cleaned, mo_codes_set),
                    axis=1
                )

            df_filtered = df[df['__entity_key__'].notna()].copy()

            # Если автоопределённый столбец с кодами дал пустой результат, попробуем первую колонку как fallback
            if df_filtered.empty and not category_prefix and code_col_idx is not None and code_col_idx != 0:
                print(f"ℹ️ Автоопределенный столбец с кодами не дал результатов для {filename}. Попробуем первую колонку как имена.")
                df['__entity_key__'] = df.apply(
                    lambda row: _infer_entity_key(row, 0, is_mo_file, okved_codes_set, okved_name_to_code, mo_name_to_mo_cleaned, mo_codes_set),
                    axis=1
                )
                df_filtered = df[df['__entity_key__'].notna()].copy()

            if df_filtered.empty:
                # Не просто "info": если файл целиком не дал ни одной строки
                # (например, справочник МО не смог опознать ни одно
                # название, и колонку с кодами не удалось определить), все
                # показатели этого файла останутся незаполненными без
                # объяснения причины — в unfilled_tags.xlsx будет просто
                # много прочерков без указания, что источник вообще не
                # прочитался. Фиксируем это явно.
                message = f"Файл прочитан, но не дал ни одной строки с кодом: {filename}"
                print(f"ℹ️ Нет данных для нужных кодов ОКВЭД в файле {filename}")
                stats['errors'].append(message)
                continue
            
            stats['files_processed'] += 1
            
            # Обрабатываем каждый показатель для этого файла
            # В таблицах "Оборачиваемость..." (длительность оборота в днях)
            # Excel иногда хранит целое число без десятичной части (например,
            # "295" вместо "295.0"), а в бюллетене все значения этой таблицы
            # должны быть с одним знаком после запятой для единообразия.
            force_decimal = any(
                'оборачиваемост' in normalize_text(t)
                for t in table_names_by_file.get(filename, [])
            )
            for indicator, (col_22_hardcode, col_23_hardcode) in indicator_to_excel.items():
                if indicator_to_file[indicator] != filename:
                    continue  # Пропускаем, если источник не совпадает
                
                keywords_2022 = indicator_keywords.get(indicator, {}).get('2022', [])
                keywords_2023 = indicator_keywords.get(indicator, {}).get('2023', [])
                context_label = f"{filename}: {indicator}"

                # Если ключевые слова обоих годов совпадают, столбец второго
                # года подтвердить нечем и оба года схлопнутся в один — см.
                # _derive_year_keywords. Выводим различающий текст из шапки.
                if keywords_2022 == keywords_2023:
                    derived = _derive_year_keywords(df, col_22_hardcode, col_23_hardcode,
                                                    header_scan_rows)
                    if derived:
                        keywords_2022 = [derived[0]] + keywords_2022
                        keywords_2023 = [derived[1]]
                    elif col_22_hardcode and col_23_hardcode:
                        warning = (f"⚠️ {context_label}: столбцы {col_22_hardcode} и "
                                   f"{col_23_hardcode} ничем не различаются в шапке — "
                                   f"годы развести не удалось, проверьте файл вручную")
                        print(f"   {warning}")

                # === ЭТАП 1: Динамический поиск колонок по году ===
                indicator_name = indicator_display_names.get(indicator, '')
                col_idx_2022 = _find_column_smart(df, col_22_hardcode, "2022", keywords_2022, context_label,
                                                  header_scan_rows, indicator_name, implausible_values)
                col_idx_2023 = _find_column_smart(df, col_23_hardcode, "2023", keywords_2023, context_label,
                                                  header_scan_rows, indicator_name, implausible_values)
                # Одна колонка Excel не может быть двумя годами. Так бывает у
                # показателя с единственной колонкой: номера второго года в
                # маппинге нет, и поиск по названию находит для него ту же
                # колонку, что и для первого. Значит, второго года у
                # показателя нет — как если бы поиск ничего не нашёл.
                if col_idx_2023 is not None and col_idx_2023 == col_idx_2022:
                    col_idx_2023 = None
                suffix_2022 = _detect_year_suffix_for_column(df, col_idx_2022, "22")
                suffix_2023 = _detect_year_suffix_for_column(df, col_idx_2023, "23")
                # force_decimal — по конкретному столбцу (плюс старый признак
                # по названию таблицы, см. комментарий выше), а не по всему
                # файлу целиком: у разных показателей в одном файле разная
                # точность, единообразие нужно внутри одного столбца, а не
                # между всеми столбцами файла.
                force_decimal_2022 = force_decimal or _column_has_decimal_values(df, col_idx_2022)
                force_decimal_2023 = force_decimal or _column_has_decimal_values(df, col_idx_2023)

                # Дробное «количество организаций» — признак дефекта в самой
                # выгрузке (см. is_count_indicator). Заполнение не блокируем:
                # наше дело — показать значение как есть и назвать место, где
                # его стоит перепроверить глазами.
                if is_count_indicator(indicator_display_names.get(indicator, '')):
                    for col_idx in (col_idx_2022, col_idx_2023):
                        if col_idx is None:
                            continue
                        for row_pos in range(len(df)):
                            cell = df.iat[row_pos, col_idx]
                            if pd.isna(cell) or not _is_fractional(cell):
                                continue
                            implausible_values.append({
                                'Файл Excel': filename,
                                'Показатель': indicator_display_names.get(indicator, indicator),
                                'Строка Excel': row_pos + 1,
                                'Название строки': str(df.iat[row_pos, 1])[:80] if df.shape[1] > 1 else '',
                                'Значение': str(cell).strip(),
                                'Что не так': 'дробное значение у показателя, считающего штуки',
                            })
                
                # === ЭТАП 2: Нечеткий поиск строк (экспериментально) ===
                if use_fuzzy_match:
                    # Получаем из column_mapping.csv название показателя
                    # Для примера, используем код индикатора как подсказку
                    pass  # TODO: Реализовать fuzzy match для строк при необходимости
                
                # === ЭТАП 3: Извлечение данных с нормализацией ===
                df_group_keys = df_filtered['__entity_key__']
                for entity_key, group in df_filtered.groupby(df_group_keys):
                    
                    # Проверка на конфликт ключей (только если данные уже есть от другого исходного кода)
                    if entity_key in master_data:
                        # Проверяем, тот же ли это исходный код (просто дубль строки в том же файле)
                        # Если да - это не конфликт, а нормальная ситуация
                        pass  # Данные будут обновлены/дополнены
                    
                    if entity_key not in master_data:
                        master_data[entity_key] = {}
                    
                    # Получаем значения по фактическим годам, определённым для каждой колонки
                    # (suffix_2022/suffix_2023 — это реальный двузначный год, например "23"/"24" для T24-файлов,
                    # а не всегда буквально "22"/"23" — см. _detect_year_suffix_for_column)
                    value_prev = None
                    if col_idx_2022 is not None:
                        for idx, row in group.iterrows():
                            value = get_cell_value_safely(row, col_idx_2022)
                            if value:
                                normalized = clean_excel_value_for_word(value, force_decimal=force_decimal_2022)
                                master_data[entity_key][f"{indicator}_{suffix_2022}"] = normalized
                                value_prev = normalized
                                stats['found_by_keyword' if keywords_2022 else 'found_by_hardcode'] += 1
                                break

                    if col_idx_2023 is not None:
                        for idx, row in group.iterrows():
                            value = get_cell_value_safely(row, col_idx_2023)
                            if value:
                                normalized = clean_excel_value_for_word(value, force_decimal=force_decimal_2023)
                                master_data[entity_key][f"{indicator}_{suffix_2023}"] = normalized
                                stats['found_by_keyword' if keywords_2023 else 'found_by_hardcode'] += 1
                                break
                    elif value_prev is not None:
                        # Если колонка для второго года не указана, но есть данные за первый год, используем их и для второго года
                        master_data[entity_key][f"{indicator}_{suffix_2023}"] = value_prev
                        stats['found_by_keyword' if keywords_2022 else 'found_by_hardcode'] += 1
        
        except Exception as e:
            print(f"❌ Ошибка обработки {filename}: {e}")
            stats['errors'].append(f"Error in {filename}: {str(e)}")
    
    print(f"\n✅ Загружено данных для {len(master_data)} кодов ОКВЭД")
    print(f"📊 Статистика:")
    print(f"   - Файлов обработано: {stats['files_processed']}")
    print(f"   - Найдено по ключевым словам: {stats['found_by_keyword']}")
    print(f"   - Найдено по индексам: {stats['found_by_hardcode']}")
    print(f"   - Конфликтов нормализации: {stats['conflicts']}")
    if conflicts:
        print(f"   ⚠️ Первые конфликты:")
        for c in conflicts[:5]:
            print(f"      - {c}")
    print(f"   - Ошибок: {len(stats['errors'])}")

    # Раньше ошибки загрузки (файл не найден/пуст/битый) были видны только в
    # консоли — если вывод не сохранён или прогон запущен без присмотра, эта
    # информация терялась безвозвратно. Пишем её в файл, как и остальные
    # отчёты (mapping_validation_log.txt, unfilled_tags.xlsx и т.п.).
    if errors_report_path is not None and stats['errors']:
        try:
            with open(errors_report_path, "w", encoding="utf-8") as f:
                f.write("\n".join(stats['errors']))
            print(f"📄 Отчёт об ошибках загрузки Excel сохранён: {errors_report_path}")
        except Exception as exc:
            print(f"⚠️ Не удалось сохранить отчёт об ошибках загрузки: {exc}")

    if implausible_report_path is not None and implausible_values:
        try:
            report_path = Path(implausible_report_path)
            report_path.parent.mkdir(parents=True, exist_ok=True)
            pd.DataFrame(implausible_values).to_excel(report_path, index=False)
            print(f"🔎 Найдено {len(implausible_values)} подозрительных значений "
                  f"(дробные «штуки», подменённые столбцы): {report_path}")
        except Exception as exc:
            print(f"⚠️ Не удалось сохранить отчёт о подозрительных значениях: {exc}")

    return master_data, stats


def _map_year_to_suffix(year: str) -> str:
    """
    Переводит год в двухзначный суффикс.

    Примеры:
        2023 -> 23
        2024 -> 24
        23   -> 23
        24   -> 24
    """
    if not year or not year.isdigit():
        return None
    if len(year) == 4 and year.startswith('20'):
        return year[-2:]
    if len(year) == 2:
        return year
    return None


_HEADER_SCAN_ROWS = 8  # запасной вариант, если определить реальную глубину не удалось


def _detect_header_scan_depth(excel_path: Path, fallback: int = _HEADER_SCAN_ROWS) -> int:
    """Определяет, сколько строк сверху файла нужно просмотреть, чтобы
    гарантированно захватить всю шапку (группа/подзаголовок/буквенный
    маркер "А","Б",1,2...). Глубина преамбулы у разных периодов и разных
    территориальных выгрузок отличается — фиксированное число строк рано
    или поздно "срежет" настоящий заголовок, поэтому там, где возможно,
    вычисляем её по буквенному маркеру (см. category_mapping._find_letter_marker_row
    в config_v2.py — тот же самый надёжный якорь, что используется при
    построении column_mapping_v2.csv)."""
    try:
        from config_v2 import _find_letter_marker_row
        raw_df = pd.read_excel(excel_path, header=None, nrows=40)
    except Exception:
        return fallback
    marker_idx = _find_letter_marker_row(raw_df)
    if marker_idx is None:
        return fallback
    # +2 с запасом: маркер стоит СРАЗУ под шапкой, а после сдвига на строку
    # из-за pandas-автозаголовка (df здесь читается уже БЕЗ header=None) нужен
    # небольшой буфер, чтобы точно не отрезать последнюю строку подзаголовка.
    return max(fallback, marker_idx + 2)


def _column_header_text(df: pd.DataFrame, col_idx: int, max_row: int) -> str:
    """Склеивает текст из первых нескольких строк одного столбца (название
    показателя, подзаголовок с годом, единица измерения и т.п.) — реальная
    шапка росстатовских таблиц занимает НЕСКОЛЬКО строк, а не одну."""
    parts = []
    for row_idx in range(max_row):
        try:
            cell = df.iat[row_idx, col_idx]
        except Exception:
            continue
        if pd.isna(cell):
            continue
        parts.append(str(cell).strip())
    return ' '.join(parts).lower()


def _normalize_header_text(text: str) -> str:
    """Приводит текст шапки или ключевого слова к виду, в котором их можно
    сравнивать: нижний регистр, ё -> е, склеенные переносы, одиночные пробелы.

    Переносы — главное. В узкой ячейке Excel слово разрывается дефисом:
    «управлен-ческие». Ключевое слово в маппинге записано целым словом, и без
    склейки подстрока не находилась: колонка из маппинга считалась
    неподтверждённой, программа решала, что структура файла сдвинулась, и
    брала другую колонку с тем же текстом. В бюллетене №1 так в колонку
    «коммерческие и управленческие расходы» ВСЕХ организаций попали расходы
    только УБЫТОЧНЫХ — 30 ячеек таблицы 17.
    """
    text = str(text).lower().replace('ё', 'е')
    text = _SOFT_WRAP_HYPHEN_RE.sub('', text)
    return re.sub(r'\s+', ' ', text).strip()


def _phrase_in_header(header_text: str, phrase: str) -> bool:
    """Есть ли фраза в тексте шапки — с начала слова, а не где угодно.

    Раньше это была простая подстрока, и «оборотные активы» находились внутри
    «внеоборотные активы». Если колонка «Внеоборотные активы» стоит левее —
    а в балансе она всегда левее, — поиск по ключевому слову возвращал её.
    Конец фразы не ограничиваем: ключевые слова бывают основами слов.
    """
    phrase = _normalize_header_text(phrase)
    if not phrase:
        return False
    return re.search(r'(?<![0-9a-zа-я])' + re.escape(phrase), header_text) is not None


_COUNT_INDICATOR_RE = re.compile(r'количеств|число организац|единиц', re.I)
_NOT_COUNT_INDICATOR_RE = re.compile(r'%|процент|доля|удельн|темп|коэффициент', re.I)


def is_count_indicator(name: str) -> bool:
    """Показатель, значения которого обязаны быть целыми (штуки организаций).

    Программа не проверяет данные на осмысленность, и это уже дало реальный
    случай: в t25OpfVed14.xlsx в столбце «Количество организаций, единиц»
    стоит 2,3 — столько организаций не бывает. В готовый бюллетень это
    попало как есть, и заметил ошибку человек, а не программа.

    Название вида «в % к общему КОЛИЧЕСТВУ организаций» — это доля, она
    дробной быть обязана, поэтому слова процентов/долей снимают признак
    счётчика. На реальных данных правило даёт ровно одно срабатывание —
    тот самый дефект — и ни одного ложного.
    """
    if not name:
        return False
    return bool(_COUNT_INDICATOR_RE.search(name)) and not _NOT_COUNT_INDICATOR_RE.search(name)


def _is_fractional(value: str) -> bool:
    text = str(value).strip().replace(',', '.')
    if not re.fullmatch(r'-?\d+\.\d+', text):
        return False
    try:
        return not float(text).is_integer()
    except ValueError:
        return False


def _looks_like_year(text: str) -> bool:
    return len(text) == 4 and text.isdigit() and 1900 <= int(text) <= 2100


def _derive_year_keywords(df: pd.DataFrame, col_22_idx: str, col_23_idx: str,
                          max_row: Optional[int] = None) -> Optional[Tuple[str, str]]:
    """Ключевые слова, РАЗЛИЧАЮЩИЕ два столбца одного показателя, взятые из
    самой шапки Excel, а не из заранее заданного списка слов.

    Заголовок группы («Внеоборотные активы») в Excel — объединённая ячейка на
    оба года, и pandas кладёт её текст только в ПЕРВЫЙ из двух столбцов. Когда
    в column_mapping_v2.csv ключевым словом для обоих годов стоит это название
    группы, у столбца второго года подтверждать нечем: _find_column_smart
    считает структуру изменившейся, ищет слово заново и возвращается в столбец
    первого года. Оба года схлопываются в один, данные второго года не
    загружаются вообще, а в документе появляются одинаковые цифры за разные
    годы.

    Поэтому берём первую строку шапки, где у двух столбцов РАЗНЫЙ непустой
    текст. Что именно там написано — «на конец предыдущего/отчетного года»,
    «2023»/«2024» или «на 31.12.2023» — не важно: важно, что этот текст лежит
    в самом столбце и различает годы. Так разметка переживает смену
    формулировок в новых периодах и в файлах других регионов.
    """
    try:
        idx_22, idx_23 = int(col_22_idx) - 1, int(col_23_idx) - 1
    except (TypeError, ValueError):
        return None
    if min(idx_22, idx_23) < 0 or max(idx_22, idx_23) >= df.shape[1]:
        return None

    for row_idx in range(min(max_row or 10, len(df))):
        first, second = df.iat[row_idx, idx_22], df.iat[row_idx, idx_23]
        first = '' if pd.isna(first) else str(first).strip()
        second = '' if pd.isna(second) else str(second).strip()
        if not first or not second or first == second:
            continue
        if (first.replace('.', '').isdigit() and second.replace('.', '').isdigit()
                and not (_looks_like_year(first) and _looks_like_year(second))):
            # Строка с номерами столбцов ("1", "2") годы различает, но как
            # ключевое слово бесполезна — короткое число совпадёт где угодно.
            # Пара годов ("2023"/"2024") — наоборот, годный ключ: она
            # длинная и однозначная, и именно так шапка может выглядеть,
            # если Росстат заменит "предыдущий/отчетный" на сами годы.
            continue
        return first, second
    return None


def _find_column_smart(df: pd.DataFrame, hardcode_idx: str, year: str, keywords: list,
                        context_label: str = "", header_scan_rows: Optional[int] = None,
                        indicator_name: str = "", events: Optional[list] = None) -> Optional[int]:
    """
    Умный поиск колонки.

    Номера колонок из column_mapping_v2.csv — это снимок структуры Excel на
    момент составления справочника. Если в новом периоде структура сдвинулась
    (вставили/убрали столбец), старый номер может молча указывать на СОВСЕМ
    ДРУГОЙ показатель, оставаясь при этом в допустимых границах таблицы —
    поэтому жёсткому индексу доверяем, только если заголовок над ним всё ещё
    подтверждается ключевыми словами. Если нет — колонку ищем заново по
    ключевым словам (для этого они и существуют), и только если это тоже не
    удалось — год, и только в самом крайнем случае — всё равно старый номер
    (лучше вернуть что-то, чем молча потерять данные).

    Args:
        df: DataFrame
        hardcode_idx: Жесткий индекс из column_mapping_v2.csv (1-based)
        year: Год для поиска ("2022" или "2023")
        keywords: Список ключевых слов для поиска
        context_label: Для сообщения в консоль, если жёсткий индекс не подтвердился
        header_scan_rows: Сколько строк сверху файла просматривать в поисках
            заголовка/подзаголовка. По умолчанию — _HEADER_SCAN_ROWS, но
            вызывающий код может передать глубину, посчитанную по реальной
            структуре КОНКРЕТНОГО файла (см. _detect_header_scan_depth) —
            преамбула перед шапкой год от года может стать длиннее.
        indicator_name: Название показателя из маппинга. Если оно стоит в шапке
            колонки из маппинга, колонка правильная — что бы ни говорило
            ключевое слово. Иначе одно испорченное ключевое слово уводит
            показатель в чужую колонку: в бюллетене №1 у «дивидендов» оно было
            скопировано со строки «от продажи акций», и 30 ячеек таблицы 20
            заполнились выручкой от продажи акций.
        events: Сюда дописывается каждый случай, когда программа заменила
            колонку из маппинга другой, — для отчёта, а не только для консоли.

    Returns:
        Индекс колонки или None
    """
    max_row = min(header_scan_rows or _HEADER_SCAN_ROWS, len(df))
    header_cache = {}

    def _header(col_idx: int) -> str:
        if col_idx not in header_cache:
            header_cache[col_idx] = _normalize_header_text(_column_header_text(df, col_idx, max_row))
        return header_cache[col_idx]

    def _keyword_match_at(col_idx: int) -> bool:
        if not keywords:
            return False
        return any(_phrase_in_header(_header(col_idx), keyword) for keyword in keywords)

    def _name_match_at(col_idx: int) -> bool:
        return bool(indicator_name) and _phrase_in_header(_header(col_idx), indicator_name)

    def _find_by_keyword() -> Optional[int]:
        for keyword in keywords:
            for col_idx in range(df.shape[1]):
                if _phrase_in_header(_header(col_idx), keyword):
                    return col_idx
        return None

    def _report(what: str, value: str) -> None:
        if events is None:
            return
        source = context_label.split(': ', 1)[0] if ': ' in context_label else ''
        events.append({
            'Файл Excel': source,
            'Показатель': indicator_name or context_label,
            'Строка Excel': '',
            'Название строки': '',
            'Значение': value,
            'Что не так': what,
        })

    hardcode_col = None
    if hardcode_idx and hardcode_idx.strip():
        try:
            idx = int(hardcode_idx) - 1  # CSV использует 1-based индексы
            if 0 <= idx < len(df.columns):
                hardcode_col = idx
        except ValueError:
            pass

    # 1️⃣ Жёсткий индекс — доверяем, если заголовок над ним подтверждён
    #    ключевыми словами или названием самого показателя (или ключевых
    #    слов вообще не задано).
    if hardcode_col is not None and (not keywords or _keyword_match_at(hardcode_col)
                                     or _name_match_at(hardcode_col)):
        return hardcode_col

    # 2️⃣ Структура похожа на изменившуюся — пересчитываем колонку по ключевым словам.
    keyword_col = _find_by_keyword() if keywords else None
    if keyword_col is not None:
        if hardcode_col is not None and keyword_col != hardcode_col:
            print(
                f"   🔄 Структура Excel сдвинулась{(' (' + context_label + ')') if context_label else ''}: "
                f"столбец {hardcode_col + 1} больше не подтверждён ключевыми словами, "
                f"используем столбец {keyword_col + 1}."
            )
            _report(
                f"столбец {hardcode_col + 1} из маппинга не подтвердился по шапке "
                f"(«{_header(hardcode_col)[:80]}»), взят столбец {keyword_col + 1} "
                f"(«{_header(keyword_col)[:80]}») — проверьте, что это тот же показатель",
                f"столбец {hardcode_col + 1} → {keyword_col + 1}",
            )
        return keyword_col

    # 3️⃣ Поиск по году — если ключевые слова не заданы/не нашлись
    if hardcode_idx and hardcode_idx.strip():
        for col_idx in range(df.shape[1]):
            if year in _column_header_text(df, col_idx, max_row):
                return col_idx

    # 4️⃣ Крайний случай: старый номер не подтвердился, но ничего лучше не
    #    нашлось — возвращаем его, чтобы не потерять данные совсем, но это
    #    стоит перепроверить вручную.
    if hardcode_col is not None:
        if keywords:
            print(
                f"   ⚠️ Не удалось подтвердить столбец {hardcode_col + 1}{(' (' + context_label + ')') if context_label else ''} "
                f"по ключевым словам — используем его как есть, проверьте вручную."
            )
            # В отчёт не пишем: это та же колонка, что в маппинге, решение
            # маппинга программа здесь не отменяет. На реальных бюллетенях
            # таких случаев 77–84, и почти все — составные имена показателей,
            # групповая часть которых в Excel стоит только над первой
            # колонкой группы. Отчёт из сотни ложных тревог никто не читает,
            # и настоящая подмена в нём потерялась бы.
        return hardcode_col

    return None


def get_cell_value_safely(row: pd.Series, col_idx: Optional[int]) -> str:
    """Безопасно получает значение ячейки.

    Отрицательный индекс отсекаем явно: pandas трактует его как отсчёт с
    конца, то есть -1 молча вернул бы значение ПОСЛЕДНЕГО столбца вместо
    отказа. Для нас это худший исход — чужая цифра, неотличимая от
    правильной. То же и с None: лучше пустая строка (её видно прочерком и
    отчётом), чем исключение посреди прогона.
    """
    if col_idx is None or col_idx < 0:
        return ""
    try:
        value = row.iloc[col_idx]
        if pd.isna(value):
            return ""
        return str(value).strip()
    except (IndexError, KeyError):
        return ""


def _detect_code_column(df: pd.DataFrame, okved_codes_set: Set[str], mo_codes_set: Optional[Set[str]], is_mo_file: bool) -> Optional[int]:
    """Автоматически находит столбец с кодами ОКВЭД/МО."""
    if df.empty:
        return None

    max_cols = min(df.shape[1], 20)
    max_rows = min(len(df), 40)
    best_score = 0.0
    best_col = None

    for col_idx in range(max_cols):
        matched = 0
        total = 0
        for row_idx in range(5, max_rows):
            try:
                cell = df.iat[row_idx, col_idx]
            except Exception:
                continue
            if pd.isna(cell):
                continue
            cell_text = str(cell).strip()
            if not cell_text or cell_text.lower() == 'nan':
                continue
            total += 1
            if is_mo_file:
                if mo_codes_set and canonical_mo(cell_text) in mo_codes_set:
                    matched += 1
            else:
                if canonical_okved(cell_text) in okved_codes_set:
                    matched += 1
        if total == 0:
            continue
        score = matched / total
        if score > best_score:
            best_score = score
            best_col = col_idx

    if best_col is not None and best_score >= 0.35:
        return best_col
    return None


def _infer_category_entity_keys(df: pd.DataFrame, category_prefix: str, name_col_idx: int = 1) -> pd.Series:
    """
    Строит __entity_key__ для "категорийных" файлов (см. category_mapping.py)
    — там, где заголовок категории (например, "20600 - Ассоциации (союзы)")
    занимает ОТДЕЛЬНУЮ строку без данных, а сами значения — в СЛЕДУЮЩЕЙ строке
    ("101.АГ Всего по обследуемым видам экономической деятельности").

    В отличие от _infer_entity_key (применяется построчно и независимо через
    df.apply), тут нужен ПОСЛЕДОВАТЕЛЬНЫЙ проход: код категории "переносится"
    с строки-заголовка на СЛЕДУЮЩУЮ строку данных, а не определяется по
    содержимому одной и той же строки.
    """
    from category_mapping import CATEGORY_HEADER_RE

    keys = [None] * len(df)
    pending_code = None
    for pos in range(len(df)):
        raw_name = df.iloc[pos, name_col_idx] if name_col_idx < df.shape[1] else None
        name_str = "" if pd.isna(raw_name) else str(raw_name).strip()
        match = CATEGORY_HEADER_RE.match(name_str)
        if match:
            pending_code = f"{category_prefix}_{match.group(1)}"
            continue
        if pending_code is not None:
            keys[pos] = pending_code
            pending_code = None
    return pd.Series(keys, index=df.index)


def _infer_entity_key(row: pd.Series, code_col_idx: int, is_mo_file: bool,
                      okved_codes_set: Set[str], okved_name_to_code: Dict[str, str],
                      mo_name_to_code: Dict[str, str],
                      mo_codes_set: Optional[Set[str]] = None) -> Optional[str]:
    """Определяет ключ сущности по строке из Excel."""
    if code_col_idx is None or code_col_idx >= len(row):
        return None

    raw_value = row.iloc[code_col_idx]
    if pd.isna(raw_value):
        return None
    raw_value = str(raw_value).strip()
    if not raw_value:
        return None

    if is_mo_file:
        candidate = canonical_mo(raw_value)
        # ВАЖНО: mo_name_to_code — это словарь "название -> код", поэтому сам
        # код (например, "30") в нём никогда не окажется как ключ. Проверять
        # нужно членство в множестве кодов (mo_codes_set), иначе прямое
        # совпадение по коду никогда не сработает и код всегда будет уходить
        # в бесполезный поиск по имени через find_mo_code.
        if mo_codes_set is not None and candidate not in mo_codes_set:
            candidate = find_mo_code(raw_value, mo_name_to_code)
        # Префикс "MO_" — по той же причине, по которой категорийные коды
        # (ОПФ/ФС) хранятся как "OPF_10000"/"FS_17", а не голым числом: коды
        # МО кладутся в ОБЩИЙ плоский словарь master_data вместе с кодами
        # ОКВЭД и категорий, и без префикса код МО (например, "30501") может
        # случайно совпасть с каким-то другим кодом в этом же словаре — тег
        # тогда получит чужое значение вместо честного прочерка, вместо того
        # чтобы просто не найтись, если реальный МО-источник не загрузился.
        return f"MO_{candidate}" if candidate else None

    candidate = canonical_okved(raw_value)
    if candidate not in okved_codes_set:
        candidate = find_okved_code(raw_value, okved_name_to_code)
    return candidate if candidate else None


def fill_word_template_by_tags_v2(doc, master_data: Dict, log_path: Optional[Path] = None,
                                  report_path: Optional[Path] = None,
                                  indicator_display_names: Optional[Dict[str, str]] = None) -> list:
    """
    Заполняет теги в Word с использованием нормализованных данных.

    Args:
        doc: Document из python-docx
        master_data: Данные из Excel (с нормализацией)
        log_path: Путь для сохранения логов
        report_path: Путь для сохранения отчета о незаполненных тегах
        indicator_display_names: код показателя -> человекочитаемое название
            (как оно называется в самом бюллетене), для отчёта. Без этого
            параметра отчёт содержит только технический тег вида
            "MO_30501_OtcGoPro_1" — понятный разработчику, но не человеку,
            который просто готовит бюллетень и не знаком с внутренним
            устройством программы.

    Returns:
        Список незаполненных тегов
    """
    print("\n🧩 Шаг 4: Заполнение шаблона по тегам (с нормализацией)")
    
    # Импортируем функцию канонизации
    from config import canonical_okved
    
    unfilled_tags = []
    log = []
    
    tag_regex = re.compile(r"\{\{([^}]+)\}\}")
    # master_data ключи — это коды ОКВЭД, МО (с префиксом "MO_") и категорий
    # (с префиксом "OPF_"/"FS_") вперемешку в общем плоском словаре.
    known_entity_codes = set(master_data.keys())

    def _canonicalize_entity_candidate(entity_raw: str, source_prefix: Optional[str]) -> str:
        if source_prefix == "MO":
            # Префикс обязателен — см. комментарий в _infer_entity_key: без
            # него код МО может случайно совпасть с чужим ключом в общем
            # словаре master_data.
            return f"MO_{canonical_mo(entity_raw)}"
        if source_prefix in CATEGORY_PREFIX_KEYWORDS.values():
            # Категорийные коды (ОПФ/форма собственности) хранятся в master_data
            # с префиксом прямо внутри ключа (например, "OPF_10000") — иначе
            # они пересекались бы с обычными кодами ОКВЭД/МО в общем плоском
            # словаре (например, форма собственности "10" совпадает по цифрам
            # с ОКВЭД "10" — производство пищевых продуктов).
            return f"{source_prefix}_{canonical_category(entity_raw)}"
        if "." in entity_raw or any(c.isalpha() for c in entity_raw):
            return canonical_okved(entity_raw.replace("_", "."))
        return canonical_okved(entity_raw)

    def _split_entity_and_indicator(parts, source_prefix: Optional[str]):
        """
        Разбивает токены тега (после удаления префикса OKVED_/MO_) на код
        сущности, код показателя и (опционально) год.

        И код сущности (например, составной ОКВЭД "101_АГ"), и код показателя
        (например, дедуп-суффиксированный при коллизии имён "ValBal_1") могут
        сами содержать "_", поэтому предположение "перед годом всегда ровно
        один токен показателя" ломается. Единственный надёжный способ найти
        границу — перебрать точки разреза и проверить код сущности по
        известным ключам master_data.
        """
        for split_idx in range(len(parts) - 1, 0, -1):
            entity_candidate = _canonicalize_entity_candidate("_".join(parts[:split_idx]), source_prefix)
            if entity_candidate not in known_entity_codes:
                continue

            indicator_parts = parts[split_idx:]
            if not indicator_parts:
                continue

            year_suffix = None
            last_part = indicator_parts[-1]
            if last_part.isdigit() and len(last_part) in (2, 4):
                candidate_year = _map_year_to_suffix(last_part)
                if candidate_year:
                    year_suffix = candidate_year
                    indicator_parts = indicator_parts[:-1]

            if not indicator_parts:
                continue

            return entity_candidate, "_".join(indicator_parts), year_suffix

        return None, None, None

    table_manager = TableManager(doc)
    for t_idx, table in enumerate(table_manager.iter_tables()):
        for r_idx, row in enumerate(table.rows):
            row_name = get_cleaned_cell_text(row.cells[0]) if row.cells else ""
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = paragraph.text
                    matches = list(tag_regex.finditer(text))
                    if not matches:
                        continue

                    for match in matches:
                        full_tag = match.group(1)
                        raw_tag = full_tag
                        source_prefix = None
                        if raw_tag.startswith("OKVED_"):
                            source_prefix = "OKVED"
                            raw_tag = raw_tag[len("OKVED_"):]
                        elif raw_tag.startswith("MO_"):
                            source_prefix = "MO"
                            raw_tag = raw_tag[len("MO_"):]
                        else:
                            for cat_prefix in CATEGORY_PREFIX_KEYWORDS.values():
                                if raw_tag.startswith(f"{cat_prefix}_"):
                                    source_prefix = cat_prefix
                                    raw_tag = raw_tag[len(cat_prefix) + 1:]
                                    break

                        parts = raw_tag.split("_")
                        if len(parts) < 2:
                            log.append(f"⚠️ Ошибка формата тега: {full_tag}")
                            continue

                        # Разбираем код сущности (ОКВЭД/МО) и код показателя, проверяя
                        # кандидатов по известным ключам master_data — иначе дедуп-суффикс
                        # показателя (например, "ValBal_1") ошибочно принимается за часть
                        # кода сущности или наоборот.
                        entity_code, indicator, year_suffix = _split_entity_and_indicator(parts, source_prefix)

                        if entity_code is None:
                            # Сущности нет в master_data вообще (данные не загрузились ни
                            # из одного файла) — используем старую наивную эвристику только
                            # для того, чтобы тег корректно попал в лог/отчёт как "нет данных".
                            last_part = parts[-1]
                            if last_part.isdigit() and len(last_part) in (2, 4):
                                year_suffix = _map_year_to_suffix(last_part)
                                if year_suffix:
                                    indicator = parts[-2]
                                    okved_parts = parts[:-2]
                                else:
                                    indicator = parts[-1]
                                    okved_parts = parts[:-1]
                            else:
                                indicator = parts[-1]
                                okved_parts = parts[:-1]
                            entity_code = _canonicalize_entity_candidate("_".join(okved_parts), source_prefix)

                        # Обрабатываем год: преобразуем реальный год (202X) в условный код (22/23)
                        lookup_suffix = year_suffix
                        indicator_key = f"{indicator}_{lookup_suffix}" if lookup_suffix else indicator

                        # Ищем значение в master_data по каноническому ключу
                        value = master_data.get(entity_code, {}).get(indicator_key)

                        if value is None and not lookup_suffix:
                            # Если год не указан в теге, пробуем найти с суффиксами года
                            value = master_data.get(entity_code, {}).get(f"{indicator}_22")
                            if value is None:
                                value = master_data.get(entity_code, {}).get(f"{indicator}_23")

                        # ВАЖНО: раньше здесь стояла подстановка соседнего года
                        # (нет данных за _24 — берём _23). Она молча выдавала
                        # данные прошлого года под видом текущего: в документе
                        # оба года выглядели одинаковыми цифрами, в отчёт о
                        # незаполненных тегах такая ячейка не попадала, и
                        # отличить её от честно заполненной было нельзя.
                        # Отсутствие данных должно быть видно прочерком.

                        # Применяем финальную нормализацию
                        # Важно: пустая строка "" - это тоже данные (значит значение есть, но оно пустое/нулевое)
                        # Проверяем именно на None, а не на ложность значения
                        if value is not None:
                            value = clean_excel_value_for_word(value)

                        # Вставляем значение: если нет данных (None), вставляем прочерк
                        if value is not None and value != "":
                            text = text.replace(f"{{{{{full_tag}}}}}", value)
                            log.append(f"✅ Заполнено: {full_tag} → {value}")
                        else:
                            text = text.replace(f"{{{{{full_tag}}}}}", "-")
                            log.append(f"ℹ️ Отсутствующие данные: {full_tag} → [-]")
                            display_indicator = (indicator_display_names or {}).get(indicator, indicator)
                            display_year = f"20{lookup_suffix}" if lookup_suffix else ""
                            unfilled_tags.append({
                                'Таблица': t_idx + 1,
                                'Строка': r_idx + 1,
                                'Название строки': row_name,
                                'Показатель': display_indicator,
                                'Год': display_year,
                                'Тег (для разработчика)': full_tag,
                            })

                    # Сохраняем форматирование (шрифт/размер) исходного
                    # текста ячейки вместо пересоздания run с форматированием
                    # по умолчанию — см. set_paragraph_text_keep_format.
                    # Жирность прочерков относительно соседей по строке
                    # выравнивается отдельным проходом по всему готовому
                    # документу — см. normalize_dash_bold_in_document (main.py).
                    set_paragraph_text_keep_format(paragraph, text)
    
    # Сохраняем логи
    if log_path:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(log))
        print(f"📝 Лог сохранён: {log_path}")

    # Сохраняем отчет по незаполненным тегам. Колонки — не технический тег
    # (его язык понятен только разработчику), а то, что реально видно в
    # самом бюллетене: номер таблицы по счёту в документе, номер и название
    # строки (то, что написано в первой колонке — код ОКВЭД/МО или
    # показатель), человекочитаемое название показателя и год. Технический
    # тег оставлен последней колонкой — на случай, если потребуется
    # эскалация разработчику/Claude.
    if report_path is not None:
        try:
            columns = ['Таблица', 'Строка', 'Название строки', 'Показатель', 'Год', 'Тег (для разработчика)']
            pd.DataFrame(unfilled_tags, columns=columns).to_excel(report_path, index=False)
            print(f"📄 Отчёт по незаполненным тегам сохранён: {report_path}")
        except Exception as exc:
            print(f"⚠️ Не удалось сохранить отчет по незаполненным тегам: {exc}")

    print(f"✅ Заполнено тегов: {len(log) - len(unfilled_tags)}")
    print(f"⚠️ Незаполненных тегов: {len(unfilled_tags)}")
    
    return unfilled_tags
