# logic.py
from pathlib import Path
import re
from typing import Optional
from config import (
    get_excel_data,
    load_okved_map,
    load_table_source_map,
    load_column_mapping,
    get_cleaned_cell_text,
    _normalize_text,
    find_okved_code,
    get_table_name,
    canonical_okved
)
from config_v2 import load_column_mapping_v2
from docx import Document
from table_registry import TableRegistry

TAG_REGEX = re.compile(r"{{([^}]+?)_([0-9]+)}}")
YEAR_PATTERN = re.compile(r'\b(20\d{2})\b')


def _extract_year(text: str) -> Optional[str]:
    if not text:
        return None
    match = YEAR_PATTERN.search(text)
    return match.group(1) if match else None


def _normalize_match_text(text: str) -> str:
    """
    Нормализация текста для сравнения показателей:
    1. Применяем базовую нормализацию (_normalize_text)
    2. Удаляем специальные символы, оставляя только цифры и буквы
    3. Схлопываем множественные пробелы
    
    Используется для поиска совпадений названий показателей в заголовках.
    """
    if not text:
        return ""
    # Сначала применяем базовую нормализацию
    normalized = _normalize_text(text)
    # Удаляем все символы кроме цифр, букв (Cyrillic/Latin) и пробелов
    normalized = re.sub(r'[^\d\w\s]', ' ', normalized, flags=re.UNICODE)
    # Удаляем лишние пробелы (на случай, если было много спецсимволов подряд)
    normalized = re.sub(r'\s+', ' ', normalized).strip()
    return normalized


def _find_year_header_row(table, min_year_cells=2, max_search_rows=20):
    """Находит последнюю строку заголовка с годами (например, 202х/202х) в таблице."""
    candidates = []
    for row_idx, row in enumerate(table.rows[:max_search_rows]):
        row_years = [get_cleaned_cell_text(cell).strip() for cell in row.cells]
        year_count = sum(1 for year in row_years if _extract_year(year))
        if year_count >= min_year_cells:
            candidates.append((row, row_idx))
    return candidates[-1] if candidates else (None, None)


def _find_header_row_by_indicators(table, source_word_to_indicator, max_search_rows=20, start_row=0):
    """Находит строку заголовка по наилучшему совпадению с названиями показателей."""
    normalized_names = [(_normalize_match_text(name), name) for name in source_word_to_indicator.keys()]
    best_match = None
    best_score = 0
    for row_idx, row in enumerate(table.rows[start_row:max_search_rows], start=start_row):
        row_texts = [_normalize_match_text(get_cleaned_cell_text(cell)) for cell in row.cells]
        score = 0
        for cell_text in row_texts:
            for normalized_name, _ in normalized_names:
                if normalized_name and normalized_name in cell_text:
                    score += 1
        if score > best_score:
            best_score = score
            best_match = (row, row_idx)
    return best_match if best_score > 0 else (None, None)


def _find_header_rows(table, source_word_to_indicator, max_search_rows=80, debug=False):
    """
    Адаптивный поиск строк заголовков таблицы (year или indicator rows).
    
    НОВАЯ ВЕРСИЯ (Problem #4 fix):
    - Автоматически определяет критерии на основе содержания таблицы
    - Поддерживает одногодовые таблицы
    - Использует гибкую иерархию приоритетов
    
    Алгоритм:
    1. Собираем статистику по строкам (года, индикаторы)
    2. Определяем, одно- или многолетние данные
    3. Применяем соответствующие критерии:
       - Многолетние (≥2): (year_count >= 2) OR (indicator_score >= 2)
       - Одногодовые (=1): (year_count >= 1) OR (indicator_score >= 1)
       - Очень слабые: (indicator_score >= 1) (fallback)
    
    Args:
        table: Таблица из Document
        source_word_to_indicator: Маппинг имён показателей на коды
        max_search_rows: Максимально строк для поиска
        debug: Выводить подробные логи
    
    Returns:
        Отсортированный список индексов строк заголовков
    """
    normalized_names = [_normalize_match_text(name) for name in source_word_to_indicator.keys()]
    
    # Собираем статистику по строкам
    row_stats = []
    for row_idx, row in enumerate(table.rows[:max_search_rows]):
        row_years = [get_cleaned_cell_text(cell).strip() for cell in row.cells]
        year_count = sum(1 for year in row_years if _extract_year(year))
        row_texts = [_normalize_match_text(get_cleaned_cell_text(cell)) for cell in row.cells]
        indicator_score = sum(1 for cell_text in row_texts for name in normalized_names if name and name in cell_text)
        row_stats.append({
            'row_idx': row_idx,
            'year_count': year_count,
            'indicator_score': indicator_score,
            'is_potential_header': False,
            'method': None
        })
    
    # Определяем максимальное количество лет в таблице
    max_year_count = max([s['year_count'] for s in row_stats], default=0)
    max_indicator_score = max([s['indicator_score'] for s in row_stats], default=0)
    
    if debug:
        print(f"   🔍 [_find_header_rows DEBUG] max_year_count={max_year_count}, max_indicator_score={max_indicator_score}")
    
    # СТРАТЕГИЯ 1: Многолетние таблицы (есть хотя бы 2 года в одной строке)
    if max_year_count >= 2:
        threshold_years = 2
        threshold_indicators = 2
        if debug:
            print(f"   🔍 [Стратегия 1] Многолетние данные обнаружены (до {max_year_count} лет)")
    # СТРАТЕГИЯ 2: Одногодовые таблицы (1 год максимум)
    elif max_year_count == 1:
        threshold_years = 1
        threshold_indicators = 1
        if debug:
            print(f"   🔍 [Стратегия 2] Одногодовые данные (1 год максимум)")
    # СТРАТЕГИЯ 3: Очень слабые таблицы (нет лет совсем, только индикаторы)
    else:
        threshold_years = 0
        threshold_indicators = 1
        if debug:
            print(f"   🔍 [Стратегия 3] Fallback: нет явных лет, полагаемся на индикаторы (score >= 1)")
    
    # Применяем критерии
    headers = []
    for stat in row_stats:
        # Критерий: (года ИЛИ индикаторы) в зависимости от стратегии
        if stat['year_count'] >= threshold_years or stat['indicator_score'] >= threshold_indicators:
            stat['is_potential_header'] = True
            if stat['year_count'] >= threshold_years:
                stat['method'] = f"year_count={stat['year_count']}"
            else:
                stat['method'] = f"indicator_score={stat['indicator_score']}"
            headers.append(stat['row_idx'])
            if debug:
                print(f"   🔍 [Заголовок {stat['row_idx']}] {stat['method']}")
    
    if debug and not headers:
        print(f"   ⚠️ [_find_header_rows DEBUG] Ни один заголовок не найден! Пытаемся fallback...")
    
    return sorted(set(headers))



def _compute_section_mapping(table, header_idx, source_word_to_indicator):
    """Вычисляет маппинг столбцов для данной секции таблицы."""
    year_row = table.rows[header_idx]
    year_texts = [get_cleaned_cell_text(cell).strip() for cell in year_row.cells]
    has_years = sum(1 for text in year_texts if _extract_year(text)) >= 2

    col_to_indicator_map = {}
    if has_years:
        header_start = max(0, header_idx - 6)
        header_rows_filled = []
        for row in table.rows[header_start:header_idx]:
            row_values = []
            last_non_empty = ""
            for cell in row.cells:
                value = _normalize_text(get_cleaned_cell_text(cell))
                if value:
                    last_non_empty = value
                else:
                    value = last_non_empty
                row_values.append(value)
            header_rows_filled.append(row_values)

        last_indicator = None
        normalized_name_map = {
            _normalize_match_text(name): indicator
            for name, indicator in source_word_to_indicator.items()
        }
        for i, cell in enumerate(year_row.cells):
            raw_year_text = get_cleaned_cell_text(cell).strip()
            year_text = _extract_year(raw_year_text)
            if not year_text:
                continue

            parts = []
            seen = set()
            for header_row in header_rows_filled:
                if i >= len(header_row):
                    continue
                header_part = header_row[i].strip()
                if not header_part or header_part in seen:
                    continue
                seen.add(header_part)
                parts.append(header_part)
            composed_header = " ".join(parts)
            composed_header_norm = _normalize_match_text(composed_header)

            best_match = None
            best_len = 0
            for name_norm, indicator in normalized_name_map.items():
                if name_norm and name_norm in composed_header_norm and len(name_norm) > best_len:
                    best_match = indicator
                    best_len = len(name_norm)

            if best_match:
                last_indicator = best_match
            elif last_indicator and not composed_header_norm:
                best_match = last_indicator

            if not best_match:
                continue

            # Используем сокращенный год из Word: 2023 -> 23, 2024 -> 24
            # Это даёт гибкие теги на будущее без правки кода из-за нового календарного года.
            year_code = year_text[-2:]
            col_to_indicator_map[i] = (best_match, year_code)
            print(f"   🔍 Столбец {i}: '{composed_header[:90]}' -> {best_match}, год {year_text} -> код {year_code}")
    else:
        header_row = year_row
        normalized_name_map = {
            _normalize_match_text(name): indicator
            for name, indicator in source_word_to_indicator.items()
        }
        for i, cell in enumerate(header_row.cells):
            header_text = _normalize_match_text(get_cleaned_cell_text(cell))
            if not header_text:
                continue
            best_match = None
            best_len = 0
            for name_norm, indicator in normalized_name_map.items():
                if name_norm and name_norm in header_text and len(name_norm) > best_len:
                    best_match = indicator
                    best_len = len(name_norm)
            if best_match:
                col_to_indicator_map[i] = (best_match, None)
                print(f"   🔍 Индикаторный заголовок: столбец {i}, текст '{header_text[:50]}' -> {best_match}")

    # Убираем дубликаты
    seen_specs = set()
    filtered_map = {}
    for col_idx in sorted(col_to_indicator_map):
        spec = col_to_indicator_map[col_idx]
        if spec in seen_specs:
            print(f"   ⚠️ Пропускаем дубликат столбца {col_idx} для {spec[0]}_{spec[1] if spec[1] else ''}")
            continue
        seen_specs.add(spec)
        filtered_map[col_idx] = spec

    return filtered_map


# ==========================================================
# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==============================
# ==========================================================
def auto_detect_table_source(table, table_source_mapping, file_word_to_indicator):
    """
    Автоматически определяет источник Excel файла для таблицы 
    по её содержимому (по ключевым показателям из одного файла).
    """
    # Собираем текст из первых 5 строк таблицы
    table_content = ' '.join([
        get_cleaned_cell_text(cell).lower()
        for row in table.rows[:5]
        for cell in row.cells
    ])
    
    # Пробуем найти текстовые совпадения показателей в таблице
    for (excel_file, indicator_name), indicator in file_word_to_indicator.items():
        if indicator_name.lower() in table_content:
            return excel_file
    return None


def get_continuation_table_number(table):
    """Возвращает номер логической таблицы из метки "Продолжение таблицы N"."""
    for row in table.rows[:3]:
        for cell in row.cells:
            text = _normalize_text(get_cleaned_cell_text(cell))
            match = re.search(r'продолжение таблицы\s*(\d+)', text)
            if match:
                try:
                    return int(match.group(1))
                except ValueError:
                    continue
    return None


def determine_table_source(table, registry: TableRegistry, file_word_to_indicator, debug=False):
    """
    Определяет источник таблицы с явной иерархией приоритетов.
    
    НОВЫЙ СПОСОБ (использует TableRegistry вместо позиционной зависимости).
    
    Приоритеты:
    1. Метка "Продолжение таблицы N" (самый надежный способ) → registry.get_source_by_position()
    2. Точное совпадение названия таблицы → registry.get_by_name(threshold=1.0)
    3. Частичное совпадение названия (>=0.90) → registry.get_by_name(threshold=0.90)
    4. Авто-определение по содержимому → auto_detect_table_source()
    5. Не определяем → None
    
    Args:
        table: Таблица из Document
        registry: TableRegistry для безопасного доступа к маппингам
        file_word_to_indicator: Маппинг (файл, название) → индикатор
        debug: Выводить подробные логи
    
    Returns:
        (source_file, method) где source_file = "T24_000000_...", method = "метод определения"
    """
    source = None
    method = None
    
    # 1️⃣ Метка продолжения (самая надежная)
    continuation_number = get_continuation_table_number(table)
    if continuation_number:
        continuation_source = registry.get_source_by_position(continuation_number)
        if continuation_source:
            method = f"continuation_marker(table_{continuation_number})"
            if debug:
                print(f"🔍 [1] Метка продолжения таблицы {continuation_number} → {continuation_source}")
            return continuation_source, method
    
    # 2️⃣ Точное совпадение названия
    table_title = get_table_name(table, registry.name_to_id.keys())
    if table_title:
        table_data = registry.get_by_name(table_title, threshold=1.0)
        if table_data:
            source = table_data['source']
            method = f"exact_title_match({table_data['position']})"
            if debug:
                print(f"🔍 [2] Точное совпадение названия (таблица {table_data['position']}) → {source}")
            return source, method
    
    # 3️⃣ Частичное совпадение названия
    if table_title:
        table_data = registry.get_by_name(table_title, threshold=0.90)
        if table_data:
            source = table_data['source']
            method = f"fuzzy_title_match({table_data['position']}, 0.90)"
            if debug:
                print(f"🔍 [3] Частичное совпадение названия (таблица {table_data['position']}, сходство ~90%) → {source}")
            return source, method
    
    # 4️⃣ Авто-определение по содержимому
    source = auto_detect_table_source(table, registry.registry, file_word_to_indicator)
    if source:
        method = "content_auto_detection"
        if debug:
            print(f"🔍 [4] Авто-определение по содержимому → {source}")
        return source, method
    
    # 5️⃣ Не определяем
    if debug:
        print(f"❌ [5] Не удалось определить источник таблицы")
    return None, "failed"



# ==========================================================
# === ГЕНЕРАЦИЯ ШАБЛОНА ====================================
# ==========================================================
def generate_word_template(input_doc_path, okved_map_path, table_source_mapping_path, column_mapping_path, output_doc_path, debug=False):
    """
    Генерация шаблона Word с тегами {{OKVED_<code>_<indicator>[_22|_23]}}.
    
    НОВАЯ ВЕРСИЯ (использует TableRegistry для стабильности):
    - Защита от сдвига индексов при изменении маппинга
    - Явная иерархия приоритетов определения источника
    - Подробное логирование в DEBUG режиме
    """
    print("\n--- ШАГ 2: Генерация шаблона с умными тегами (v2.1 - со стабилизацией) ---")
    
    # Загружаем справочники
    _, name_to_okved_cleaned = load_okved_map(okved_map_path)
    registry = TableRegistry(table_source_mapping_path)  # ← НОВОЕ: Реестр таблиц
    word_to_indicator, _, indicator_to_file, file_word_to_indicator, _ = load_column_mapping_v2(column_mapping_path)
    
    if debug:
        print("\n🔍 DEBUG РЕЖИМ ВКЛЮЧЕН")
        registry.print_summary()
    
    doc = Document(input_doc_path)
    total_tags = 0
    current_source_file = None
    method_for_current = None  # Способ определения источника
    
    for t_index, table in enumerate(doc.tables):
        print(f"\n📄 Таблица {t_index + 1}")
        
        # НОВЫЙ ПОДХОД: Используем determine_table_source с иерархией приоритетов
        source_file, method = determine_table_source(table, registry, file_word_to_indicator, debug=debug)
        
        if source_file:
            current_source_file = source_file
            method_for_current = method
            print(f"✅ Источник определен: {source_file} ({method})")
        else:
            # Fallback: используем предыдущий источник только если есть хотя бы одна таблица до этого
            if current_source_file:
                print(f"ℹ️ Источник не определен, используем предыдущий: {current_source_file}")
            else:
                print(f"⚠️ Источник не определен и нет предыдущего источника. Пропускаем таблицу.")
                continue
        
        if not current_source_file:
            print("⚠️ Источник не определён. Пропускаем таблицу.")
            continue
        
        # Используем file_word_to_indicator для правильного маппирования по (файл, слово)
        source_word_to_indicator = {
            name: indicator
            for (file, name), indicator in file_word_to_indicator.items()
            if file == current_source_file
        }
        
        if not source_word_to_indicator:
            print(f"⚠️ Нет показателей для источника {current_source_file}. Пропускаем таблицу.")
            continue
        
        # DEBUG для таблицы 17
        if t_index + 1 == 17 and debug:
            print(f"   DEBUG: Таблица 17")
            print(f"   Источник: {current_source_file}")
            print(f"   Показатели ({len(source_word_to_indicator)}):")
            for name, indicator in list(source_word_to_indicator.items())[:5]:
                print(f"      {name} → {indicator}")
        
        header_rows = _find_header_rows(table, source_word_to_indicator, debug=debug)
        section_ranges = []
        for header_idx in header_rows:
            mapping = _compute_section_mapping(table, header_idx, source_word_to_indicator)
            if mapping:
                section_ranges.append((header_idx, mapping))
        
        mapped_sections = []
        col_to_indicator_map = {}
        if not section_ranges:
            # Fallback: старая логика для нестандартных таблиц.
            print(f"   🔄 Fallback: используем старую логику для таблицы {t_index + 1}")
            base_indicator_map = {}
            header_row, header_row_idx = _find_header_row_by_indicators(table, source_word_to_indicator, max_search_rows=20)
            if header_row is not None:
                print(f"   🔍 Fallback: используем строку заголовка {header_row_idx + 1} для поиска показателей")
            else:
                header_row = table.rows[1] if len(table.rows) > 1 else table.rows[0]
            
            for i, cell in enumerate(header_row.cells):
                header_text = _normalize_text(get_cleaned_cell_text(cell))
                if header_text:
                    for name, indicator in source_word_to_indicator.items():
                        if name in header_text:
                            base_indicator_map[i] = indicator
                            print(f"   🔍 Fallback: столбец {i}, текст '{header_text[:50]}' → {indicator}")
                            break
            
            print(f"   📈 base_indicator_map: {base_indicator_map}")
            
            year_row, year_row_idx = _find_year_header_row(table)
            if year_row:
                for i, cell in enumerate(year_row.cells):
                    raw_year_text = get_cleaned_cell_text(cell).strip()
                    indicator = base_indicator_map.get(i)
                    if indicator is None:
                        continue
                    year_text = _extract_year(raw_year_text)
                    if not year_text:
                        continue
                    col_to_indicator_map[i] = (indicator, year_text[-2:])
            else:
                for i, indicator in base_indicator_map.items():
                    col_to_indicator_map[i] = (indicator, None)
            
            print(f"   📊 col_to_indicator_map: {col_to_indicator_map}")
            if not col_to_indicator_map:
                print("⚠️ Заголовки не найдены. Пропускаем таблицу.")
                continue
            
            mapped_sections = [(0, len(table.rows), col_to_indicator_map)]
        else:
            section_ranges.sort(key=lambda x: x[0])
            for idx, (header_idx, mapping) in enumerate(section_ranges):
                start = header_idx + 1
                end = section_ranges[idx + 1][0] if idx + 1 < len(section_ranges) else len(table.rows)
                mapped_sections.append((start, end, mapping))
        
        if not mapped_sections:
            print("⚠️ Заголовки не найдены. Пропускаем таблицу.")
            continue
        
        # 4. Вставка тегов в строки с кодами ОКВЭД
        for row_idx, row in enumerate(table.rows):
            mapping = None
            for start, end, section_map in mapped_sections:
                if start <= row_idx < end:
                    mapping = section_map
                    break
            if mapping is None:
                continue
            
            first_cell_text = get_cleaned_cell_text(row.cells[0])
            okved_code = find_okved_code(first_cell_text, name_to_okved_cleaned)
            if not okved_code:
                continue
            
            # Применяем канонизацию к коду ОКВЭД перед созданием тега
            okved_canonical = canonical_okved(okved_code)
            okved_tag_part = okved_canonical.replace('.', '_')
            
            for col_idx, indicator_spec in mapping.items():
                if col_idx < len(row.cells):
                    indicator, year = indicator_spec
                    if year:
                        tag = f"{{{{OKVED_{okved_tag_part}_{indicator}_{year}}}}}"
                    else:
                        tag = f"{{{{OKVED_{okved_tag_part}_{indicator}}}}}"
                    # Очищаем ячейку перед вставкой тега
                    for p in row.cells[col_idx].paragraphs:
                        p.text = " "
                    if row.cells[col_idx].paragraphs:
                        row.cells[col_idx].paragraphs[0].text = tag
                    total_tags += 1
                    if debug:
                        print(f"   🏷️ Строка {row_idx + 1}: вставлен тег {tag}")
    
    doc.save(output_doc_path)
    print(f"\n✅ Шаблон с тегами сохранён: {output_doc_path}")
    print(f"🔢 Всего вставлено тегов: {total_tags}")
    if debug:
        print(f"ℹ️ DEBUG режим завершен")



# ==========================================================
# === ДИАГНОСТИКА И ПОМОЩНЫЕ ==============================
# ==========================================================
def find_unfilled_tags(doc_path, start_table_number: int = 1, start_table_title: str = None,
                        table_source_mapping: dict = None):
    """Диагностика незаполненных тегов.

    Args:
        doc_path: путь к документу .docx.
        start_table_number: номер таблицы, с которой начинать анализ (1-based).
        start_table_title: точное название таблицы, с которой начать анализ.
        table_source_mapping: маппинг названий таблиц → источников, используемый для определения таблицы по заголовку.
    """
    doc = Document(doc_path)
    unfilled_tags = {}

    def collect_paragraphs(paragraphs):
        for para in paragraphs:
            matches = TAG_REGEX.findall(para.text)
            for tag in matches:
                unfilled_tags.setdefault(f"{{{{{tag}}}}}", 0)
                unfilled_tags[f"{{{{{tag}}}}}"] += 1

    def collect_table(table):
        for row in table.rows:
            for cell in row.cells:
                collect_paragraphs(cell.paragraphs)
                for nested_table in cell.tables:
                    collect_table(nested_table)

    # Всегда анализируем теги из обычных параграфов документа.
    collect_paragraphs(doc.paragraphs)

    if start_table_title and table_source_mapping:
        for idx, table in enumerate(doc.tables, start=1):
            title = get_table_name(table, table_source_mapping.keys())
            if title and title.strip().lower() == start_table_title.strip().lower():
                start_table_number = idx
                break

    start_index = max(1, start_table_number)
    for idx, table in enumerate(doc.tables, start=1):
        if idx < start_index:
            continue
        collect_table(table)

    total_tags = sum(unfilled_tags.values())
    print(f"\n🔍 Незаполненные теги (начиная с таблицы {start_index}): {total_tags}")
    for tag, count in sorted(unfilled_tags.items()):
        print(f" - {tag} ({count} раз)")
    return set(unfilled_tags.keys())


def _extract_okved_code_from_tag(raw_tag: str) -> Optional[str]:
    if not raw_tag or not raw_tag.startswith("OKVED_"):
        return None

    raw_tag = raw_tag[len("OKVED_"):]
    parts = raw_tag.split("_")
    if not parts:
        return None

    if parts[-1] in ("22", "23"):
        parts = parts[:-1]

    if not parts:
        return None

    okved_parts = [parts[0]]
    for part in parts[1:]:
        if part.isdigit() or (part.isalpha() and part.isupper()):
            okved_parts.append(part)
            continue
        # Если встречаем часть, которая выглядит как индикатор, останавливаемся.
        break

    okved_raw = "_".join(okved_parts)
    if not okved_raw:
        return None

    if "." in okved_raw or any(c.isalpha() for c in okved_raw):
        return canonical_okved(okved_raw.replace("_", "."))
    return canonical_okved(okved_raw)


def collect_okved_codes_from_template(doc_path):
    """Собирает коды ОКВЭД из тегов шаблона."""
    doc = Document(doc_path)
    codes = set()

    def add_codes_from_text(text: str):
        for tag, _ in TAG_REGEX.findall(text):
            okved_code = _extract_okved_code_from_tag(tag)
            if okved_code:
                codes.add(okved_code)

    for para in doc.paragraphs:
        add_codes_from_text(para.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    add_codes_from_text(para.text)

    print(f"\n📄 Коды ОКВЭД в шаблоне: {len(codes)}")
    return codes


def compare_okved_sets(template_codes, excel_codes):
    """Сравнение множеств кодов ОКВЭД."""
    missing = template_codes - excel_codes
    extra = excel_codes - template_codes

    print("\n🔍 В шаблоне, но нет в Excel:")
    for code in sorted(missing):
        print(f" - {code}")

    print("\n📁 В Excel, но не используется в шаблоне:")
    for code in sorted(extra):
        print(f" - {code}")
