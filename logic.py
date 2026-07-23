# logic.py
from difflib import SequenceMatcher
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
from mo import load_mo_map, find_mo_code, canonical_mo
from config_v2 import load_column_mapping_v2
from docx import Document

TAG_REGEX = re.compile(r"{{([^}]+?)_([0-9]+)}}")
YEAR_PATTERN = re.compile(r'\b(20\d{2})\b')


def _extract_year(text: str) -> Optional[str]:
    if not text:
        return None
    match = YEAR_PATTERN.search(text)
    return match.group(1) if match else None


def _fuzzy_match_header(target_text: str, normalized_name_map: dict, threshold: float = 0.80) -> str:
    """
    Ищет наилучшее совпадение заголовка с использованием Fuzzy Matching.
    Возвращает код индикатора или None.
    """
    if not target_text:
        return None

    # 1. Точное вхождение (в любую сторону). Среди ВСЕХ кандидатов с вхождением
    # выбираем тот, у кого выше итоговый Ratio, а не первый попавшийся по порядку
    # словаря. Это важно, когда одно название — префикс другого (например,
    # "Себестоимость продаж" является префиксом "Себестоимость продаж с учетом
    # коммерческих и управленческих расходов") — иначе порядок перебора мог бы
    # случайно "перетянуть" короткую колонку на длинный показатель.
    best_exact_match = None
    best_exact_score = -1.0
    for name_norm, indicator in normalized_name_map.items():
        if not name_norm:
            continue
        if name_norm in target_text or target_text in name_norm:
            score = SequenceMatcher(None, target_text, name_norm).ratio()
            if score > best_exact_score:
                best_exact_score = score
                best_exact_match = indicator

    if best_exact_match is not None:
        return best_exact_match

    # 2. Нечеткое сравнение (Fuzzy Match) для случаев без точного вхождения
    best_match = None
    best_score = 0.0
    for name_norm, indicator in normalized_name_map.items():
        if not name_norm:
            continue
        score = SequenceMatcher(None, target_text, name_norm).ratio()
        if score > best_score:
            best_score = score
            best_match = indicator

    if best_score >= threshold:
        return best_match

    # Диагностика: выводим в лог, если показатель почти нашелся, но не дотянул до порога
    if target_text and len(target_text) > 5 and best_score > 0.50:
        print(f"   ⚠️ [Fuzzy Miss] Ожидалось похожее, но Score={best_score:.2f} для '{target_text[:50]}...'")

    return None


def _normalize_match_text(text: str) -> str:
    """
    Нормализация текста для сравнения показателей:
    1. Применяем базовую нормализацию (_normalize_text)
    2. Удаляем скобки, но сохраняем пробелы (важно для многословных названий)
    3. Удаляем все спецсимволы, оставляя только цифры и буквы
    4. Схлопываем множественные пробелы

    Используется для поиска совпадений названий показателей в заголовках.
    """
    if not text:
        return ""
    # Сначала применяем базовую нормализацию
    normalized = _normalize_text(text)
    # Заменяем скобки и их содержимое на пробелы (сохраняем границы слов)
    # Это важно, потому что "текст (в скобках)" должен совпадать с "текст в скобках"
    def _remove_brackets_preserve_words(s: str) -> str:
        # Заменяем скобки на пробелы: (content) -> content
        s = re.sub(r"[\(\[\<]", " ", s)
        s = re.sub(r"[\)\]\>]", " ", s)
        return s
    normalized = _remove_brackets_preserve_words(normalized)
    # Убираем "мягкий" перенос слова Word (дефис без пробелов вокруг, например
    # из-за узкой колонки таблицы: "себестои-мость" -> "себестоимость"). Иначе
    # слово разбивается на два фрагмента и перестаёт совпадать с чистым названием
    # показателя из column_mapping_v2.csv. Важно убрать дефис ДО замены прочих
    # спецсимволов на пробел, иначе перенос уже станет пробелом и информация о
    # том, что дефис не отделял слова, будет потеряна.
    normalized = re.sub(r'(?<=\w)-(?=\w)', '', normalized, flags=re.UNICODE)
    # Удаляем все символы кроме цифр, букв (Cyrillic/Latin) и пробелов
    normalized = re.sub(r'[^\d\w\s]', ' ', normalized, flags=re.UNICODE)
    # Удаляем лишние пробелы
    normalized = re.sub(r'\s+', ' ', normalized).strip()
    return normalized


def _text_matches_name(cell_text: str, name_norm: str, min_len: int = 5) -> bool:
    """
    Проверяет совпадение нормализованного названия показателя с текстом ячейки
    в обе стороны.

    column_mapping_v2.csv иногда хранит "составное" название показателя вместе
    с названием группы столбцов из Excel (например, "Доходы и расходы по обычным
    видам деятельности за отчетный период Выручка"), тогда как в самой ячейке
    Word-таблицы отображается только короткая часть ("Выручка") — групповой
    заголовок в таблице не повторяется. Поэтому helper должен ловить оба случая:
    короткое название внутри длинного текста ячейки и короткий текст ячейки
    внутри длинного составного названия.
    """
    if not name_norm or not cell_text:
        return False
    if name_norm in cell_text:
        return True
    if len(cell_text) >= min_len and cell_text in name_norm:
        return True
    return False


def _find_year_header_row(table, min_year_cells=2, max_search_rows=40):
    """Находит последнюю строку заголовка с годами (например, 2022/2023) в таблице."""
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
                if _text_matches_name(cell_text, normalized_name):
                    score += 1
        if score > best_score:
            best_score = score
            best_match = (row, row_idx)
    return best_match if best_score > 0 else (None, None)


def _build_composed_header_for_column(table, base_row_idx: int, col_idx: int, depth: int = 4) -> str:
    parts = []
    seen = set()
    for offset in range(max(1, depth)):
        row_idx = base_row_idx + offset
        if row_idx >= len(table.rows):
            break
        if col_idx >= len(table.rows[row_idx].cells):
            continue

        # Читаем ячейку, меняем переносы на пробелы, чтобы они не терялись при нормализации
        raw_text = get_cleaned_cell_text(table.rows[row_idx].cells[col_idx]).replace('\n', ' ')
        part = _normalize_match_text(raw_text)

        if not part or part in seen:
            continue
        seen.add(part)
        parts.append(part)

    return " ".join(parts).strip()


def _find_header_rows(table, source_word_to_indicator, max_search_rows=80):
    """Собирает все строки заголовков таблицы (year или indicator rows)."""
    normalized_names = [_normalize_match_text(name) for name in source_word_to_indicator.keys()]
    headers = []
    rows = table.rows[:max_search_rows]
    for row_idx in range(len(rows)):
        row = rows[row_idx]
        row_years = [get_cleaned_cell_text(cell).strip() for cell in row.cells]
        year_count = sum(1 for year in row_years if _extract_year(year))
        row_texts = [_normalize_match_text(get_cleaned_cell_text(cell)) for cell in row.cells]
        indicator_score = sum(1 for cell_text in row_texts for name in normalized_names if _text_matches_name(cell_text, name))

        # Также пробуем объединить текущую строку с следующей — часто заголовок разбит
        combined_score = 0
        if row_idx + 1 < len(rows):
            next_row = rows[row_idx + 1]
            next_texts = [_normalize_match_text(get_cleaned_cell_text(cell)) for cell in next_row.cells]
            combined_texts = [((a + ' ' + b).strip()) for a, b in zip(row_texts, next_texts)]
            combined_score = sum(1 for cell_text in combined_texts for name in normalized_names if _text_matches_name(cell_text, name))

        if year_count >= 2 or indicator_score >= 2 or combined_score >= 2:
            headers.append(row_idx)
            if combined_score >= 2:
                # Помечаем и следующую строку как часть шапки
                headers.append(row_idx + 1)
    return sorted(set(headers))


def _compute_section_mapping(table, header_idx, source_word_to_indicator, header_rows=None):
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

            # Используем Fuzzy-поиск вместо жесткого вхождения
            best_match = _fuzzy_match_header(composed_header_norm, normalized_name_map, threshold=0.80)

            # Fuzzy-fallback для случаев с переносами, дефисами и неявными формулировками
            if not best_match and composed_header_norm:
                fuzzy_score = 0.0
                fuzzy_match = None
                for name_norm, indicator in normalized_name_map.items():
                    if not name_norm:
                        continue
                    score = SequenceMatcher(None, name_norm, composed_header_norm).ratio()
                    if score > fuzzy_score:
                        fuzzy_score = score
                        fuzzy_match = indicator
                if fuzzy_score >= 0.75:
                    best_match = fuzzy_match
                    print(f"   🔍 fuzzy match {fuzzy_score:.2f} для '{composed_header[:80]}' -> {best_match}")

            if best_match:
                last_indicator = best_match
            elif last_indicator and not composed_header_norm:
                best_match = last_indicator

            if not best_match:
                print(f"   ⚠️ Нет match для колонки {i} ('{composed_header[:80]}')")
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
        # Попробуем использовать не только одну строку, но и составной заголовок.
        # ВАЖНО: строку header_idx + 1 подмешиваем только если она сама распознана
        # как заголовочная (входит в header_rows) — иначе, если header_idx оказался
        # последней строкой шапки, next_row_texts протащит текст первой строки ДАННЫХ
        # (например, числа или название ОКВЭД/МО), и это "зашумит" сравнение с
        # названиями показателей, ломая совпадение по некоторым столбцам.
        next_row_texts = []
        next_row_is_header = header_rows is None or (header_idx + 1) in header_rows
        if header_idx + 1 < len(table.rows) and next_row_is_header:
            next_row_texts = [_normalize_match_text(get_cleaned_cell_text(cell)) for cell in table.rows[header_idx + 1].cells]

        for i, cell in enumerate(header_row.cells):
            base_text = _normalize_match_text(get_cleaned_cell_text(cell))
            next_text = next_row_texts[i] if i < len(next_row_texts) else ''
            # Составной заголовок: базовый + подзаголовок
            header_text = (base_text + ' ' + next_text).strip()
            if not header_text:
                continue

            best_match = _fuzzy_match_header(header_text, normalized_name_map, threshold=0.80)

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


def get_table_source_by_number(table_source_mapping, table_number):
    if table_number < 1 or table_number > len(table_source_mapping):
        return None
    return list(table_source_mapping.values())[table_number - 1]


# ==========================================================
# === ОСНОВНАЯ ЛОГИКА ======================================
# ==========================================================
def generate_word_template(input_doc_path, okved_map_path, table_source_mapping_path, column_mapping_path,
                           output_doc_path,
                           mo_map_path: Path = None, validation_log_path: Path = None):
    """
    Генерация шаблона Word с тегами {{OKVED_<code>_<indicator>[_22|_23]}} и {{MO_<code>_<indicator>[_22|_23]}}.
    Расширенный поиск заголовков по первым 5 строкам таблицы.
    """
    print("\n--- ШАГ 2: Генерация шаблона с умными тегами ---")
    _, name_to_okved_cleaned = load_okved_map(okved_map_path)
    _, mo_name_to_mo_cleaned = ({}, {})
    if mo_map_path is not None:
        _, mo_name_to_mo_cleaned = load_mo_map(mo_map_path)

    table_source_mapping = load_table_source_map(table_source_mapping_path)
    mo_source_files = {src for src in table_source_mapping.values() if 'mo' in src.lower()}
    word_to_indicator, _, indicator_to_file, file_word_to_indicator, _ = load_column_mapping_v2(column_mapping_path)
    doc = Document(input_doc_path)
    total_tags = 0
    validation_log = []
    current_source_file = None
    normalized_title_to_src = {k: v for k, v in table_source_mapping.items()}

    for t_index, table in enumerate(doc.tables):
        print(f"\n📄 Таблица {t_index + 1}")

        # 0. Определение источника данных по названию таблицы
        table_title = get_table_name(table, table_source_mapping.keys())
        if table_title:
            table_title_norm = _normalize_text(table_title)
            if table_title_norm in normalized_title_to_src:
                current_source_file = normalized_title_to_src[table_title_norm]
                print(f"🔍 Источник таблицы: {current_source_file}")
            else:
                print(f"⚠️ Не найден источник для заголовка таблицы: '{table_title}'")

        continuation_number = get_continuation_table_number(table)
        if continuation_number:
            continuation_source = get_table_source_by_number(table_source_mapping, continuation_number)
            if continuation_source:
                if current_source_file and current_source_file != continuation_source:
                    print(
                        f"🔁 Источник по метке продолжения таблицы {continuation_number} ({continuation_source}) отличается от источника заголовка ({current_source_file}). Предпочитаем продолжение таблицы.")
                current_source_file = continuation_source
                print(f"🔁 Источник по метке продолжения таблицы {continuation_number}: {current_source_file}")

        if not current_source_file:
            print(f"   → Пытаемся определить по содержимому таблицы...")
            detected = auto_detect_table_source(table, table_source_mapping, file_word_to_indicator)
            if detected:
                current_source_file = detected
                print(f"   ✓ Автоматически определен источник: {current_source_file}")
            else:
                print("ℹ️ Заголовок таблицы не определён, используем предыдущий источник")

        if not current_source_file:
            print("⚠️ Источник не определён. Пропускаем таблицу.")
            continue

        # Используем file_word_to_indicator для правильного маппирования по (файл, слово)
        all_file_entries = {
            name: indicator
            for (file, name), indicator in file_word_to_indicator.items()
            if file == current_source_file
        }

        # Приоритет: если в текущей таблице из Word встречаются названия показателей —
        # используем только их (Word главный источник). В противном случае — все записи из Excel.
        # NOTE: Расширяем диапазон на 20 строк, чтобы захватить все заголовки (включая сложные многострочные)
        table_content = ' '.join([
            get_cleaned_cell_text(cell)
            for row in table.rows[:20]
            for cell in row.cells
        ])
        table_content_norm = _normalize_match_text(table_content)

        source_word_to_indicator = {}
        for name, indicator in all_file_entries.items():
            name_norm = _normalize_match_text(name)
            if name_norm and name_norm in table_content_norm:
                source_word_to_indicator[name] = indicator

        if not source_word_to_indicator:
            # fallback на все показатели из Excel, если в Word ничего не найдено
            source_word_to_indicator = all_file_entries.copy()

        if not source_word_to_indicator:
            print(f"⚠️ Нет показателей для источника {current_source_file}. Пропускаем таблицу.")
            continue

        header_rows = _find_header_rows(table, source_word_to_indicator)
        header_rows_set = set(header_rows)
        section_ranges = []
        for header_idx in header_rows:
            mapping = _compute_section_mapping(table, header_idx, source_word_to_indicator, header_rows=header_rows_set)
            if mapping:
                section_ranges.append((header_idx, mapping))

        mapped_sections = []
        col_to_indicator_map = {}
        if not section_ranges:
            # Fallback: старая логика для нестандартных таблиц.
            print(f"   🔄 Fallback: используем старую логику для таблицы {t_index + 1}")
            base_indicator_map = {}
            # Ищем строку заголовка по явным названиям показателей, если она не на первой позиции.
            header_row, header_row_idx = _find_header_row_by_indicators(table, source_word_to_indicator,
                                                                        max_search_rows=20)
            if header_row is not None:
                print(f"   🔍 Fallback: используем строку заголовка {header_row_idx + 1} для поиска показателей")
            else:
                header_row = table.rows[1] if len(table.rows) > 1 else table.rows[0]
                header_row_idx = 1 if len(table.rows) > 1 else 0

            # 1) Основная стратегия: маппинг только по верхней строке заголовка (row0-first)
            # 2) Если не сработало, fallback: объединяем верхнюю строку с соседней (row0+row1)
            normalized_name_map = {
                _normalize_match_text(name): indicator
                for name, indicator in source_word_to_indicator.items()
            }

            for i, cell in enumerate(header_row.cells):
                header_text = _normalize_match_text(get_cleaned_cell_text(cell))
                if not header_text:
                    continue
                for name_norm, indicator in normalized_name_map.items():
                    if _text_matches_name(header_text, name_norm):
                        base_indicator_map[i] = indicator
                        print(f"   🔍 Fallback row0: столбец {i}, текст '{header_text[:50]}' → {indicator}")
                        break

            if not base_indicator_map and header_row_idx is not None:
                for depth in (2, 3):
                    if base_indicator_map:
                        break
                    for i, cell in enumerate(header_row.cells):
                        composed = _build_composed_header_for_column(table, header_row_idx, i, depth=depth)
                        if not composed:
                            continue
                        for name_norm, indicator in normalized_name_map.items():
                            if _text_matches_name(composed, name_norm):
                                base_indicator_map[i] = indicator
                                print(
                                    f"   🔍 Fallback row0+... (depth={depth}): столбец {i}, "
                                    f"текст '{composed[:70]}' → {indicator}"
                                )
                                break

            print(f"   📈 base_indicator_map: {base_indicator_map}")

            year_row, year_row_idx = _find_year_header_row(table)
            if year_row:
                last_year_code = None
                for i, cell in enumerate(year_row.cells):
                    raw_year_text = get_cleaned_cell_text(cell).strip()
                    indicator = base_indicator_map.get(i)
                    if indicator is None:
                        continue
                    year_text = _extract_year(raw_year_text)
                    if year_text:
                        last_year_code = year_text[-2:]
                        col_to_indicator_map[i] = (indicator, last_year_code)
                    else:
                        # В некоторых шаблонах год в колонке может отсутствовать.
                        # Не пропускаем показатель: используем последний найденный год
                        # в пределах той же шапки, либо сохраняем без года.
                        col_to_indicator_map[i] = (indicator, last_year_code)
                        print(
                            f"   ⚠️ Нет года в колонке {i} fallback-шапки; "
                            f"используем год {last_year_code if last_year_code else 'None'} для {indicator}"
                        )
            else:
                for i, indicator in base_indicator_map.items():
                    col_to_indicator_map[i] = (indicator, None)

            print(f"   📊 col_to_indicator_map: {col_to_indicator_map}")
            if not col_to_indicator_map:
                print("⚠️ Заголовки не найдены. Пропускаем таблицу.")
                expected = sorted(set(source_word_to_indicator.values()))
                if expected:
                    validation_log.append(
                        f"[TABLE {t_index + 1}] source={current_source_file} status=NO_HEADERS expected_indicators={len(expected)} "
                        f"sample={expected[:10]}"
                    )
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

        matched_indicators = set()
        for _, _, section_map in mapped_sections:
            for indicator, _ in section_map.values():
                matched_indicators.add(indicator)
        expected_indicators = set(source_word_to_indicator.values())
        missing_indicators = sorted(expected_indicators - matched_indicators)
        if missing_indicators:
            validation_log.append(
                f"[TABLE {t_index + 1}] source={current_source_file} matched={len(matched_indicators)}/{len(expected_indicators)} "
                f"missing_sample={missing_indicators[:15]}"
            )

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
            mo_code = None
            if not okved_code and current_source_file and current_source_file.lower().endswith(
                    '.xlsx') and current_source_file.lower().find('mo') != -1:
                mo_code = find_mo_code(first_cell_text, mo_name_to_mo_cleaned)

            if not okved_code and not mo_code:
                continue

            if okved_code:
                prefix = "OKVED"
                code_value = canonical_okved(okved_code)
            else:
                prefix = "MO"
                code_value = canonical_mo(mo_code)

            code_tag_part = code_value.replace('.', '_')
            name_cell_tc = row.cells[0]._tc

            for col_idx, indicator_spec in mapping.items():
                if col_idx < len(row.cells):
                    if col_idx != 0 and row.cells[col_idx]._tc is name_cell_tc:
                        # Ячейка данных объединена с ячейкой названия ОКВЭД/МО (col0) —
                        # запись тега сюда стёрла бы название строки. Пропускаем.
                        print(f"   ⚠️ Строка {row_idx + 1}: колонка {col_idx} объединена с колонкой названия, тег пропущен")
                        continue
                    indicator, year = indicator_spec
                    if year:
                        tag = f"{{{{{prefix}_{code_tag_part}_{indicator}_{year}}}}}"
                    else:
                        tag = f"{{{{{prefix}_{code_tag_part}_{indicator}}}}}"
                    # Очищаем ячейку перед вставкой тега
                    for p in row.cells[col_idx].paragraphs:
                        p.text = " "
                    if row.cells[col_idx].paragraphs:
                        row.cells[col_idx].paragraphs[0].text = tag
                    total_tags += 1
                    print(f"   🏷️ Строка {row_idx + 1}: вставлен тег {tag}")

    doc.save(output_doc_path)
    print(f"\n✅ Шаблон с тегами сохранён: {output_doc_path}")
    print(f"🔢 Всего вставлено тегов: {total_tags}")
    if validation_log_path:
        validation_log_path = Path(validation_log_path)
        validation_log_path.parent.mkdir(parents=True, exist_ok=True)
        with open(validation_log_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(validation_log))
        print(f"🧪 Лог валидации маппинга сохранён: {validation_log_path}")


# ==========================================================
# === ДИАГНОСТИКА И ПОМОЩНЫЕ ==============================
# ==========================================================
def find_unfilled_tags(doc_path):
    """Диагностика незаполненных тегов."""
    doc = Document(doc_path)
    unfilled_tags = set()
    for para in doc.paragraphs:
        matches = TAG_REGEX.findall(para.text)
        for code, index in matches:
            unfilled_tags.add(f"{{{{{code}_{index}}}}}")

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    matches = TAG_REGEX.findall(para.text)
                    for code, index in matches:
                        unfilled_tags.add(f"{{{{{code}_{index}}}}}")

    print(f"\n🔍 Незаполненные теги: {len(unfilled_tags)}")
    for tag in sorted(unfilled_tags):
        print(f" - {tag}")
    return unfilled_tags


def _extract_okved_code_from_tag(raw_tag: str) -> Optional[str]:
    if not raw_tag or not raw_tag.startswith("OKVED_"):
        return None

    raw_tag = raw_tag[len("OKVED_"):]
    parts = raw_tag.split("_")
    if not parts:
        return None

    if parts[-1] in ("22", "23"):
        parts = parts[:-1]

    if len(parts) >= 3:
        code_parts = parts[:-1]
        if len(code_parts) >= 2 and code_parts[0].isdigit() and code_parts[1].isdigit():
            return f"{code_parts[0]}.{code_parts[1]}"
        return "_".join(code_parts)

    return None


def save_unfilled_tags_to_excel(tags: set, output_excel_path: str, template_doc_path: str = None):
    """
    Сохраняет незаполненные теги в Excel с дополнительной информацией:
    - tag: сам тег
    - status: status='missing' if there is no data
    - indicator_code: код показателя из тега
    - okved_code: extracted OKVED code
    - okved_name: наименование ОКВЭД (если найдено)
    - table_source: источник таблицы
    """
    import pandas as pd

    rows = []
    for tag in sorted(tags):
        raw = tag.strip('{}')
        okved_code = _extract_okved_code_from_tag(raw)
        rows.append({
            'tag': tag,
            'status': 'missing',
            'okved_code': okved_code,
        })

    df = pd.DataFrame(rows)
    Path(output_excel_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_excel_path, index=False)
    print(f"📄 Отчёт о незаполненных тегах сохранён: {output_excel_path}")


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