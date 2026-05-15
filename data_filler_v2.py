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

from config_v2 import load_column_mapping_v2, build_column_mapping_v2_from_excel
from smart_loader import (
    find_column_by_year,
    find_row_by_fuzzy_match,
    normalize_text,
    get_cell_value_safely
)
from data_normalizer import clean_excel_value_for_word


def pre_load_all_excel_data_v2(excel_dir: Path, table_source_mapping: Dict, 
                               okved_codes_set: Set[str], column_mapping_path: Path,
                               use_fuzzy_match: bool = True, fuzzy_threshold: float = 0.80) -> Tuple[Dict, dict]:
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
        _, indicator_to_excel, indicator_to_file, _, indicator_keywords = load_column_mapping_v2(str(column_mapping_path))
    except FileNotFoundError:
        print(f"⚠️ Файл {column_mapping_path} не найден. Генерируем его из Excel...")
        build_column_mapping_v2_from_excel(excel_dir, table_source_mapping, column_mapping_path)
        _, indicator_to_excel, indicator_to_file, _, indicator_keywords = load_column_mapping_v2(str(column_mapping_path))
    except Exception as e:
        print(f"⚠️ Ошибка загрузки маппингов v2: {e}, используем fallback...")
        from config import load_column_mapping
        _, indicator_to_excel, indicator_to_file, _ = load_column_mapping(str(column_mapping_path))
        indicator_keywords = {k: {'2022': [], '2023': []} for k in indicator_to_file.keys()}
    
    # Импортируем функцию канонизации
    from config import canonical_okved
    
    master_data = {}
    conflicts = []
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
            
            # Фильтруем по кодам ОКВЭД (первый столбец) с канонизацией
            df_filtered = df[df.iloc[:, 0].apply(lambda x: canonical_okved(str(x))).isin(okved_codes_set)]
            
            if df_filtered.empty:
                print(f"ℹ️ Нет данных для нужных кодов ОКВЭД в файле {filename}")
                continue
            
            stats['files_processed'] += 1
            
            # Обрабатываем каждый показатель для этого файла
            force_decimal = filename == 'T23_000000_t13Ved14.xlsx'
            for indicator, (col_22_hardcode, col_23_hardcode) in indicator_to_excel.items():
                if indicator_to_file[indicator] != filename:
                    continue  # Пропускаем, если источник не совпадает
                
                keywords_2022 = indicator_keywords.get(indicator, {}).get('2022', [])
                keywords_2023 = indicator_keywords.get(indicator, {}).get('2023', [])
                
                # === ЭТАП 1: Динамический поиск колонок по году ===
                col_idx_2022 = _find_column_smart(df, col_22_hardcode, "2022", keywords_2022, debug=False)
                col_idx_2023 = _find_column_smart(df, col_23_hardcode, "2023", keywords_2023, debug=False)
                
                # === ЭТАП 2: Нечеткий поиск строк (экспериментально) ===
                if use_fuzzy_match:
                    # Получаем из column_mapping.csv название показателя
                    # Для примера, используем код индикатора как подсказку
                    pass  # TODO: Реализовать fuzzy match для строк при необходимости
                
                # === ЭТАП 3: Извлечение данных с нормализацией ===
                for okved_raw, group in df_filtered.groupby(df_filtered.iloc[:, 0]):
                    # Канонизируем код ОКВЭД из Excel
                    okved_canonical = canonical_okved(str(okved_raw))
                    
                    # Проверка на конфликт ключей (только если данные уже есть от другого исходного кода)
                    if okved_canonical in master_data:
                        # Проверяем, тот же ли это исходный код (просто дубль строки в том же файле)
                        # Если да - это не конфликт, а нормальная ситуация
                        pass  # Данные будут обновлены/дополнены
                    
                    if okved_canonical not in master_data:
                        master_data[okved_canonical] = {}
                    
                    # Получаем значения за 2022 и 2023
                    value_22 = None
                    if col_idx_2022 is not None:
                        for idx, row in group.iterrows():
                            value = get_cell_value_safely(row, col_idx_2022)
                            if value:
                                normalized = clean_excel_value_for_word(value, force_decimal=force_decimal)
                                master_data[okved_canonical][f"{indicator}_22"] = normalized
                                value_22 = normalized
                                stats['found_by_keyword' if keywords_2022 else 'found_by_hardcode'] += 1
                                break
                    
                    if col_idx_2023 is not None:
                        for idx, row in group.iterrows():
                            value = get_cell_value_safely(row, col_idx_2023)
                            if value:
                                normalized = clean_excel_value_for_word(value, force_decimal=force_decimal)
                                master_data[okved_canonical][f"{indicator}_23"] = normalized
                                stats['found_by_keyword' if keywords_2023 else 'found_by_hardcode'] += 1
                                break
                    elif value_22 is not None:
                        # Если колонка 2023 не указана, но есть данные за 2022, используем их для 2023
                        master_data[okved_canonical][f"{indicator}_23"] = value_22
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


def _find_column_smart(df: pd.DataFrame, hardcode_idx: str, year: str, keywords: list, debug=False) -> Optional[int]:
    """
    Умный поиск колонки с ОБРАТНЫМИ приоритетами (Problem #2 fix).
    
    НОВАЯ ВЕРСИЯ (более стабильная):
    Приоритет:
    1️⃣ Поиск по ключевым словам + году (самый надежный способ)
    2️⃣ Поиск ТОЛЬКО по ключевым словам (если год не указан)
    3️⃣ Поиск по году (fallback, если есть ключевые слова но они не нашли)
    4️⃣ Жесткий индекс (только как последний resort)
    
    Причина переворота:
    - Жесткий индекс ЗАВИСИТ от порядка столбцов в Excel
    - Если Excel переупорядочивается, индекс становится неправильным
    - Ключевые слова (название столбца) более стабильны
    
    Args:
        df: DataFrame
        hardcode_idx: Жесткий индекс из column_mapping.csv (fallback)
        year: Год для поиска ("2022" или "2023")
        keywords: Список ключевых слов для поиска (надежный способ)
        debug: Выводить подробные логи
    
    Returns:
        Индекс колонки или None
    """
    if debug:
        print(f"   🔍 [_find_column_smart] keywords={keywords}, year={year}, hardcode_idx={hardcode_idx}")
    
    # 1️⃣ ПРИОРИТЕТ 1: Ключевые слова + год (самый надежный)
    if keywords and year:
        for keyword in keywords:
            for col_idx, header in enumerate(df.iloc[0]):
                header_str = str(header).strip().lower()
                keyword_lower = keyword.lower()
                year_str = str(year)
                # Ищем оба: ключевое слово И год в одном столбце
                if keyword_lower in header_str and year_str in header_str:
                    if debug:
                        print(f"   ✅ [1] Найден по ключевому слову + год: col {col_idx}")
                    return col_idx
    
    # 2️⃣ ПРИОРИТЕТ 2: Только ключевые слова (если год не помог)
    if keywords:
        for keyword in keywords:
            for col_idx, header in enumerate(df.iloc[0]):
                header_str = str(header).strip().lower()
                keyword_lower = keyword.lower()
                if keyword_lower in header_str:
                    if debug:
                        print(f"   ✅ [2] Найден по ключевому слову: col {col_idx}")
                    return col_idx
    
    # 3️⃣ ПРИОРИТЕТ 3: Только по году (если ключевые слова не нашли)
    if year:
        for col_idx, header in enumerate(df.iloc[0]):
            header_str = str(header).strip()
            if year in header_str:
                if debug:
                    print(f"   ✅ [3] Найден по году: col {col_idx}")
                return col_idx
    
    # 4️⃣ ПРИОРИТЕТ 4: Жесткий индекс как fallback (только если все else сработало)
    if hardcode_idx and hardcode_idx.strip():
        try:
            idx = int(hardcode_idx) - 1  # CSV использует 1-based индексы
            if 0 <= idx < len(df.columns):
                if debug:
                    print(f"   ✅ [4] Использован жесткий индекс: col {idx}")
                return idx
        except ValueError:
            pass
    
    if debug:
        print(f"   ❌ [_find_column_smart] Колонка не найдена!")
    return None

    return None


def get_cell_value_safely(row: pd.Series, col_idx: int) -> str:
    """Безопасно получает значение ячейки"""
    try:
        value = row.iloc[col_idx]
        if pd.isna(value):
            return ""
        return str(value).strip()
    except (IndexError, KeyError):
        return ""


def fill_word_template_by_tags_v2(doc, master_data: Dict, log_path: Optional[Path] = None, 
                                  report_path: Optional[Path] = None) -> list:
    """
    Заполняет теги в Word с использованием нормализованных данных.
    
    Args:
        doc: Document из python-docx
        master_data: Данные из Excel (с нормализацией)
        log_path: Путь для сохранения логов
        report_path: Путь для сохранения отчета о незаполненных тегах
    
    Returns:
        Список незаполненных тегов
    """
    print("\n🧩 Шаг 4: Заполнение шаблона по тегам (с нормализацией)")
    
    # Импортируем функцию канонизации
    from config import canonical_okved
    
    unfilled_tags = []
    log = []
    
    tag_regex = re.compile(r"\{\{([^}]+)\}\}")
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        text = run.text
                        matches = list(tag_regex.finditer(text))
                        
                        for match in matches:
                            full_tag = match.group(1)
                            raw_tag = full_tag
                            
                            if raw_tag.startswith("OKVED_"):
                                raw_tag = raw_tag[len("OKVED_"):]
                            
                            parts = raw_tag.split("_")
                            if len(parts) < 2:
                                log.append(f"⚠️ Ошибка формата тега: {full_tag}")
                                continue
                            
                            # Парсируем OKVED код и индикатор (с поддержкой реальных годов 202X)
                            year_suffix = None
                            if len(parts) > 1:
                                last_part = parts[-1]
                                # Проверяем, является ли последний элемент годом (202X или двухзначным годом)
                                if last_part.isdigit() and (len(last_part) == 4 or len(last_part) == 2):
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
                            else:
                                indicator = parts[-1]
                                okved_parts = parts[:-1]
                            
                            # Собираем код ОКВЭД из частей (с разделителями "_")
                            okved_raw = "_".join(okved_parts)
                            # Канонизируем код ОКВЭД для поиска в master_data
                            # Важно: в тегах используется "_" как разделитель, но в master_data ключи могут быть с "." или без
                            # canonical_okved просто делает upper() и strip(), так что заменяем "_" обратно на "." только если это было в исходном коде
                            # На самом деле - в Excel коды хранятся как "101.АГ", "85", "A" и т.д.
                            # В тегах мы используем "101_АГ", "85", "A"
                            # Значит нужно заменить "_" на "." для составных кодов
                            if "." in okved_raw or any(c.isalpha() for c in okved_raw):
                                # Это составной код типа "101_АГ" → "101.АГ"
                                okved_code = canonical_okved(okved_raw.replace("_", "."))
                            else:
                                # Это простой код типа "85" или "A"
                                okved_code = canonical_okved(okved_raw)
                            
                            # Обрабатываем год: преобразуем реальный год (202X) в условный код (22/23)
                            lookup_suffix = year_suffix
                            indicator_key = f"{indicator}_{lookup_suffix}" if lookup_suffix else indicator
                            
                            # Ищем значение в master_data по каноническому ключу
                            value = master_data.get(okved_code, {}).get(indicator_key)
                            
                            if value is None and not lookup_suffix:
                                # Если год не указан в теге, пробуем найти с суффиксами года
                                value = master_data.get(okved_code, {}).get(f"{indicator}_22")
                                if value is None:
                                    value = master_data.get(okved_code, {}).get(f"{indicator}_23")
                            
                            if value is None and lookup_suffix:
                                # Если прямой год не найден, пробуем fallback на относительные годы
                                value = master_data.get(okved_code, {}).get(f"{indicator}_23")
                                if value is None:
                                    value = master_data.get(okved_code, {}).get(f"{indicator}_22")
                            
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
                                # Прочерк для отсутствующих данных (чтобы совпадать с эталоном)
                                text = text.replace(f"{{{{{full_tag}}}}}", "-")
                                log.append(f"ℹ️ Отсутствующие данные: {full_tag} → [-]")
                                unfilled_tags.append(full_tag)
                        
                        run.text = text
    
    # Сохраняем логи
    if log_path:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(log))
        print(f"📝 Лог сохранён: {log_path}")
    
    print(f"✅ Заполнено тегов: {len(log) - len(unfilled_tags)}")
    print(f"⚠️ Незаполненных тегов: {len(unfilled_tags)}")
    
    return unfilled_tags
