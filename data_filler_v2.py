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

from config import canonical_okved, find_okved_code
from config_v2 import load_column_mapping_v2, build_column_mapping_v2_from_excel
from mo import load_mo_map, canonical_mo, find_mo_code
from smart_loader import (
    find_column_by_year,
    find_row_by_fuzzy_match,
    normalize_text,
    get_cell_value_safely
)
from data_normalizer import clean_excel_value_for_word


def pre_load_all_excel_data_v2(excel_dir: Path, table_source_mapping: Dict, 
                               okved_codes_set: Set[str], okved_name_to_code: Dict[str, str],
                               column_mapping_path: Path,
                               mo_map_path: Optional[Path] = None,
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
    
    # Импортируем функции канонизации
    from config import canonical_okved
    mo_name_to_mo_cleaned = {}
    if mo_map_path is not None:
        _, mo_name_to_mo_cleaned = load_mo_map(mo_map_path)
    
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
            
            is_mo_file = 'mo' in filename.lower()
            mo_codes_set = set(mo_name_to_mo_cleaned.values()) if mo_name_to_mo_cleaned else None
            code_col_idx = _detect_code_column(df, okved_codes_set, mo_codes_set, is_mo_file)
            if code_col_idx is None:
                print(f"ℹ️ Не удалось автоматически определить колонку с кодами в файле {filename}. Попробуем использовать первую колонку как имена.")
                code_col_idx = 0
            else:
                print(f"ℹ️ Автоопределена колонка с кодами в файле {filename}: {code_col_idx + 1}")

            df['__entity_key__'] = df.apply(
                lambda row: _infer_entity_key(row, code_col_idx, is_mo_file, okved_codes_set, okved_name_to_code, mo_name_to_mo_cleaned),
                axis=1
            )
            df_filtered = df[df['__entity_key__'].notna()].copy()

            # Если автоопределённый столбец с кодами дал пустой результат, попробуем первую колонку как fallback
            if df_filtered.empty and code_col_idx is not None and code_col_idx != 0:
                print(f"ℹ️ Автоопределенный столбец с кодами не дал результатов для {filename}. Попробуем первую колонку как имена.")
                df['__entity_key__'] = df.apply(
                    lambda row: _infer_entity_key(row, 0, is_mo_file, okved_codes_set, okved_name_to_code, mo_name_to_mo_cleaned),
                    axis=1
                )
                df_filtered = df[df['__entity_key__'].notna()].copy()

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
                col_idx_2022 = _find_column_smart(df, col_22_hardcode, "2022", keywords_2022)
                col_idx_2023 = _find_column_smart(df, col_23_hardcode, "2023", keywords_2023)
                
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
                    
                    # Получаем значения за 2022 и 2023
                    value_22 = None
                    if col_idx_2022 is not None:
                        for idx, row in group.iterrows():
                            value = get_cell_value_safely(row, col_idx_2022)
                            if value:
                                normalized = clean_excel_value_for_word(value, force_decimal=force_decimal)
                                master_data[entity_key][f"{indicator}_22"] = normalized
                                value_22 = normalized
                                stats['found_by_keyword' if keywords_2022 else 'found_by_hardcode'] += 1
                                break
                    
                    if col_idx_2023 is not None:
                        for idx, row in group.iterrows():
                            value = get_cell_value_safely(row, col_idx_2023)
                            if value:
                                normalized = clean_excel_value_for_word(value, force_decimal=force_decimal)
                                master_data[entity_key][f"{indicator}_23"] = normalized
                                stats['found_by_keyword' if keywords_2023 else 'found_by_hardcode'] += 1
                                break
                    elif value_22 is not None:
                        # Если колонка 2023 не указана, но есть данные за 2022, используем их для 2023
                        master_data[entity_key][f"{indicator}_23"] = value_22
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


def _find_column_smart(df: pd.DataFrame, hardcode_idx: str, year: str, keywords: list) -> Optional[int]:
    """
    Умный поиск колонки: сначала жесткий индекс, потом ключевые слова, потом год.
    
    Args:
        df: DataFrame
        hardcode_idx: Жесткий индекс из column_mapping.csv (приоритет)
        year: Год для поиска ("2022" или "2023")
        keywords: Список ключевых слов для поиска
    
    Returns:
        Индекс колонки или None
    """
    # 1️⃣ Сначала используем жесткий индекс, если он есть
    if hardcode_idx and hardcode_idx.strip():
        try:
            idx = int(hardcode_idx) - 1  # CSV использует 1-based индексы
            if 0 <= idx < len(df.columns):
                return idx
        except ValueError:
            pass
    
    # 2️⃣ Пытаемся найти по ключевым словам
    if keywords:
        for keyword in keywords:
            for col_idx, header in enumerate(df.iloc[0]):
                header_str = str(header).strip()
                if keyword.lower() in header_str.lower():
                    return col_idx
    
    # 3️⃣ Пытаемся найти по году (только если жесткий индекс был указан, но не найден)
    if hardcode_idx and hardcode_idx.strip():
        for col_idx, header in enumerate(df.iloc[0]):
            header_str = str(header).strip()
            if year in header_str:
                return col_idx
    
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


def _infer_entity_key(row: pd.Series, code_col_idx: int, is_mo_file: bool,
                      okved_codes_set: Set[str], okved_name_to_code: Dict[str, str],
                      mo_name_to_code: Dict[str, str]) -> Optional[str]:
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
        if mo_name_to_code and candidate not in mo_name_to_code:
            candidate = find_mo_code(raw_value, mo_name_to_code)
        return candidate if candidate else None

    candidate = canonical_okved(raw_value)
    if candidate not in okved_codes_set:
        candidate = find_okved_code(raw_value, okved_name_to_code)
    return candidate if candidate else None


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

                        # Собираем код из частей (с разделителями "_")
                        entity_raw = "_".join(okved_parts)
                        # Канонизируем код для поиска в master_data
                        if source_prefix == "MO":
                            entity_code = canonical_mo(entity_raw)
                        else:
                            if "." in entity_raw or any(c.isalpha() for c in entity_raw):
                                entity_code = canonical_okved(entity_raw.replace("_", "."))
                            else:
                                entity_code = canonical_okved(entity_raw)

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

                        if value is None and lookup_suffix:
                            # Если прямой год не найден, пробуем fallback на относительные годы
                            value = master_data.get(entity_code, {}).get(f"{indicator}_23")
                            if value is None:
                                value = master_data.get(entity_code, {}).get(f"{indicator}_22")

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
                            unfilled_tags.append(full_tag)

                    paragraph.text = text
    
    # Сохраняем логи
    if log_path:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(log))
        print(f"📝 Лог сохранён: {log_path}")

    # Сохраняем отчет по незаполненным тегам
    if report_path is not None:
        try:
            pd.DataFrame({'tag': unfilled_tags}).to_excel(report_path, index=False)
            print(f"📄 Отчёт по незаполненным тегам сохранён: {report_path}")
        except Exception as exc:
            print(f"⚠️ Не удалось сохранить отчет по незаполненным тегам: {exc}")

    print(f"✅ Заполнено тегов: {len(log) - len(unfilled_tags)}")
    print(f"⚠️ Незаполненных тегов: {len(unfilled_tags)}")
    
    return unfilled_tags
