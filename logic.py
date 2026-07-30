# logic.py
from difflib import SequenceMatcher
from pathlib import Path
import re
from typing import Optional

import pandas as pd
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


def _squish_text(text: str) -> str:
    """
    Самый толерантный уровень сравнения: убирает вообще все пробелы и
    спецсимволы, оставляя только буквы и цифры "слитно". Ловит случаи, когда
    слово внутри показателя разорвано непредвиденным образом — например,
    "Себе - стоимость продаж" вместо "Себестоимость продаж" (лишний пробел
    вокруг дефиса при переносе строки в Word). Обычная нормализация
    (_normalize_match_text) такое не распознаёт как перенос, потому что перед
    дефисом есть пробел — а значит выглядит как настоящий разделитель вроде
    "деятельности - всего". Склеивание вообще без пробелов снимает разницу:
    "себе стоимость продаж" и "себестоимость продаж" после склейки совпадают.
    """
    if not text:
        return ""
    normalized = _normalize_text(text)
    return re.sub(r'[^0-9a-zA-Zа-яА-ЯёЁ]', '', normalized)


def _one_letter_diff(a: str, b: str) -> bool:
    """
    Проверяет, что две строки отличаются не больше, чем на одну букву
    (расстояние Левенштейна <= 1) — вставка, удаление или замена одного
    символа. Ловит опечатки прямо в самом Word-документе, например
    "уравленческие" вместо "управленческие" (пропущена буква "п"), которые
    ни точное вхождение, ни склейка (_squish_text) не могут распознать —
    ведь там пропущена не пунктуация/пробел, а именно буква слова.

    Специально не считается через полное расстояние Левенштейна (O(n*m)),
    а через линейный проход — этого достаточно, т.к. нас интересует только
    "0 или 1", а не точное значение при большем расхождении.
    """
    if a == b:
        return True
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False
    if la > lb:
        a, b = b, a
        la, lb = lb, la
    i = j = 0
    mismatch_used = False
    while i < la and j < lb:
        if a[i] == b[j]:
            i += 1
            j += 1
            continue
        if mismatch_used:
            return False
        mismatch_used = True
        if la == lb:
            i += 1
            j += 1
        else:
            j += 1
    return True


def _fuzzy_match_header(target_text: str, normalized_name_map: dict, threshold: float = 0.80,
                         near_miss_sink: Optional[list] = None,
                         squish_match_sink: Optional[list] = None) -> str:
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
    best_match_name = None
    best_score = 0.0
    for name_norm, indicator in normalized_name_map.items():
        if not name_norm:
            continue
        score = SequenceMatcher(None, target_text, name_norm).ratio()
        if score > best_score:
            best_score = score
            best_match = indicator
            best_match_name = name_norm

    if best_score >= threshold:
        return best_match

    # 3. Последний, самый толерантный уровень: сравнение "склеенных" (без
    # пробелов и любых спецсимволов) форм. Ловит разрывы слова, которые не
    # распознаются как перенос строки обычной нормализацией (см. _squish_text)
    # — например, "Себе - стоимость продаж" вместо "Себестоимость продаж".
    # Требуем точное совпадение или вхождение (не нечёткий ratio) — этот шаг и
    # так менее строгий, чем предыдущие, добавлять к нему ещё и нечёткость
    # было бы слишком рискованно. Каждое такое совпадение обязательно логируем
    # отдельно (squish_match_sink), т.к. в отличие от точного/нечеткого
    # совпадения текста, здесь два РАЗНЫХ по написанию текста считаются одним
    # и тем же показателем — это стоит проверять человеку, а не доверять молча.
    target_squished = _squish_text(target_text)
    if target_squished and len(target_squished) >= 5:
        best_squish_match = None
        best_squish_name = None
        best_squish_len = -1
        for name_norm, indicator in normalized_name_map.items():
            name_squished = _squish_text(name_norm)
            if not name_squished or len(name_squished) < 5:
                continue
            if name_squished == target_squished or name_squished in target_squished or target_squished in name_squished:
                # При нескольких кандидатах предпочитаем более длинное (более специфичное) совпадение
                if len(name_squished) > best_squish_len:
                    best_squish_len = len(name_squished)
                    best_squish_match = indicator
                    best_squish_name = name_norm
        if best_squish_match is not None:
            print(f"   🧩 [Squish-match] '{target_text[:50]}' -> '{best_squish_name[:50]}' -> {best_squish_match}")
            if squish_match_sink is not None:
                squish_match_sink.append((target_text, best_squish_name, best_squish_match))
            return best_squish_match

    # Диагностика: если показатель почти нашелся, но не дотянул до порога — это
    # тег, который остаётся без значения молча (или, что хуже, ячейка так и
    # хранит нетронутый исходный текст, если это единственный кандидат для всей
    # строки). Раньше это уходило только в консольный лог и никем не читалось;
    # теперь также складываем в near_miss_sink, чтобы в конце прогона можно было
    # выгрузить это в отдельный отчёт для проверки человеком.
    if target_text and len(target_text) > 5 and best_score > 0.50:
        print(f"   ⚠️ [Fuzzy Miss] Ожидалось похожее, но Score={best_score:.2f} для '{target_text[:50]}...'")
        if near_miss_sink is not None:
            near_miss_sink.append((target_text, best_score, best_match_name, best_match))

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
    # Убираем "мягкий" перенос слова Word (дефис без пробела ПЕРЕД ним, например
    # из-за узкой колонки таблицы: "себестои-мость" -> "себестоимость"). После
    # него может остаться один пробел — это коллапс переноса строки ("количе-\n
    # ство" -> "количе- ство" уже на этапе _normalize_text), а не настоящий
    # дефис-разделитель ("деятельности - всего", где пробел есть С ОБЕИХ сторон).
    # Поэтому пробел ДО дефиса — признак настоящего разделителя (не трогаем), а
    # пробел только ПОСЛЕ — признак переноса (убираем вместе с дефисом). Важно
    # убрать дефис ДО замены прочих спецсимволов на пробел, иначе перенос уже
    # станет пробелом и информация о том, что дефис не отделял слова, будет
    # потеряна.
    normalized = re.sub(r'(?<=\w)-\s*(?=\w)', '', normalized, flags=re.UNICODE)
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


def _compute_section_mapping(table, header_idx, source_word_to_indicator, header_rows=None, run_rows=None,
                              near_miss_sink=None, squish_match_sink=None):
    """Вычисляет маппинг столбцов для данной секции таблицы.

    near_miss_sink, если передан, получает по одному элементу
    (column_idx, cell_text, score, closest_name, closest_indicator) на каждую
    почти-но-не-дотянувшую fuzzy-попытку — чтобы такие случаи не терялись
    молча, а могли попасть в отчёт для проверки человеком.

    squish_match_sink, если передан, получает по одному элементу
    (column_idx, cell_text, matched_name, matched_indicator) на каждое
    совпадение, найденное только через самый толерантный уровень сравнения
    (_squish_text) — тоже стоит проверять человеку, см. _fuzzy_match_header.
    """
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
            column_near_misses = [] if near_miss_sink is not None else None
            column_squish_matches = [] if squish_match_sink is not None else None
            best_match = _fuzzy_match_header(composed_header_norm, normalized_name_map, threshold=0.80,
                                              near_miss_sink=column_near_misses,
                                              squish_match_sink=column_squish_matches)
            if column_near_misses:
                near_miss_sink.extend((i, *entry) for entry in column_near_misses)
            if column_squish_matches:
                squish_match_sink.extend((i, *entry) for entry in column_squish_matches)

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
        # header_idx — это ПОСЛЕДНЯЯ строка своего "run"-а подряд идущих
        # заголовочных строк (см. вызывающий код), поэтому строка header_idx+1 —
        # это уже настоящие данные, а не продолжение шапки. Составной заголовок
        # строим по ВСЕМ строкам run'а (обычно это "базовое название показателя"
        # + "суффикс/подзаголовок"), а не только по одной последней строке —
        # иначе несколько колонок с одинаковым суффиксом (например, у двух разных
        # показателей "длительность 1 оборота") неотличимы друг от друга.
        run_rows_sorted = sorted(run_rows) if run_rows else [header_idx]

        for i, cell in enumerate(header_row.cells):
            parts = []
            seen = set()
            for r_idx in run_rows_sorted:
                if r_idx >= len(table.rows):
                    continue
                row_cells = table.rows[r_idx].cells
                if i >= len(row_cells):
                    continue
                part = _normalize_match_text(get_cleaned_cell_text(row_cells[i]))
                if part and part not in seen:
                    seen.add(part)
                    parts.append(part)
            header_text = ' '.join(parts).strip()
            if not header_text:
                continue

            column_near_misses = [] if near_miss_sink is not None else None
            column_squish_matches = [] if squish_match_sink is not None else None
            best_match = _fuzzy_match_header(header_text, normalized_name_map, threshold=0.80,
                                              near_miss_sink=column_near_misses,
                                              squish_match_sink=column_squish_matches)
            if column_near_misses:
                near_miss_sink.extend((i, *entry) for entry in column_near_misses)
            if column_squish_matches:
                squish_match_sink.extend((i, *entry) for entry in column_squish_matches)

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
                           mo_map_path: Path = None, validation_log_path: Path = None,
                           near_miss_report_path: Path = None, squish_match_report_path: Path = None,
                           typo_match_report_path: Path = None):
    """
    Генерация шаблона Word с тегами {{OKVED_<code>_<indicator>[_22|_23]}} и {{MO_<code>_<indicator>[_22|_23]}}.
    Расширенный поиск заголовков по первым 5 строкам таблицы.

    near_miss_report_path: если указан, сюда выгружается таблица всех "почти
    совпавших" заголовков — случаев, когда нечёткое сравнение нашло похожий
    показатель, но недостаточно похожий, чтобы вставить тег автоматически
    (score выше 0.50, но ниже порога принятия 0.80). Раньше такие случаи были
    видны только в консольном логе; теперь их можно проверить и решить,
    исправлять ли column_mapping_v2.csv или сам Word-документ.

    squish_match_report_path: если указан, сюда выгружается таблица всех
    совпадений, найденных только через самый толерантный уровень сравнения
    (склейка без пробелов и спецсимволов, см. _squish_text) — они прошли
    автоматически, но стоит проверить человеком, что совпадение верное.

    typo_match_report_path: если указан, сюда выгружается таблица совпадений,
    найденных только благодаря допуску на опечатку в одну букву (см.
    _one_letter_diff) — случаев, когда сам показатель в Word-документе
    написан с ошибкой (пропущена/заменена одна буква), например
    "уравленческие" вместо "управленческие". Такие совпадения проходят
    автоматически, но их стоит проверить человеком: возможно, стоит
    исправить сам Word-документ, а не мириться с опечаткой.
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
    near_miss_report = []
    squish_match_report = []
    typo_match_report = []
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
        # NOTE: Сканируем ВСЮ таблицу, а не только первые 20 строк — в длинных таблицах
        # (более ~20 строк, например с "Продолжение таблицы N" в середине) второй и
        # последующие блоки заголовков могут начинаться значительно позже. Если их
        # показатели не попадут в source_word_to_indicator, для них не построится
        # section mapping, и правило "конец предыдущей секции = начало следующей"
        # молча растянет предыдущий (неверный) маппинг колонок до конца таблицы.
        #
        # ВАЖНО: составляем текст ПО КОЛОНКАМ (сверху вниз), а не одной строкой на
        # всю таблицу построчно. Название показателя в CSV часто склеено из
        # "группового" заголовка (строка 1) и "листового" подзаголовка (строка 2),
        # которые в Word стоят друг под другом В ОДНОЙ КОЛОНКЕ. При построчной
        # склейке всей таблицы эти две строки становятся смежными только для
        # ПОСЛЕДНЕЙ группы в строке заголовков (её текст просто оказывается
        # впритык к началу следующей строки) — а для остальных групп между ними
        # оказывается текст соседних колонок той же строки, и совпадение не
        # находится. Из-за этого, например, "получившие убыток" находился, а
        # "получившие прибыль" — нет, и оба столбца затем сопоставлялись с одним
        # и тем же (последним найденным) показателем.
        num_cols = max((len(row.cells) for row in table.rows), default=0)
        column_texts_norm = []
        column_texts_squished = []
        # Для допуска на опечатку в одну букву (см. _one_letter_diff) нужны
        # тексты ОТДЕЛЬНЫХ ячеек колонки, а не всего столбца одной строкой —
        # расстояние Левенштейна имеет смысл только между двумя цельными
        # фразами похожей длины, а не при поиске подстроки в длинной "простыне".
        column_cells_squished = []
        for col_idx in range(num_cols):
            cell_texts = [
                get_cleaned_cell_text(row.cells[col_idx])
                for row in table.rows
                if col_idx < len(row.cells)
            ]
            col_text = ' '.join(cell_texts)
            column_texts_norm.append(_normalize_match_text(col_text))
            column_texts_squished.append(_squish_text(col_text))
            column_cells_squished.append([_squish_text(t) for t in cell_texts if t])

        source_word_to_indicator = {}
        for name, indicator in all_file_entries.items():
            name_norm = _normalize_match_text(name)
            if name_norm and any(name_norm in col_text for col_text in column_texts_norm):
                source_word_to_indicator[name] = indicator
                continue
            # Тот же самый показатель может не найтись обычной проверкой, если
            # слово в Word-документе разорвано непредвиденным образом (лишний
            # пробел вокруг дефиса при переносе — см. _squish_text). Пробуем
            # ещё и "склеенное" сравнение, прежде чем считать показатель
            # отсутствующим в этой таблице.
            name_squished = _squish_text(name)
            if name_squished and len(name_squished) >= 5 and any(
                name_squished in col_text for col_text in column_texts_squished
            ):
                source_word_to_indicator[name] = indicator
                continue
            # Последний шанс: опечатка в одну букву в самом Word-документе
            # (например, "уравленческие" вместо "управленческие" — пропущена
            # буква "п"). В отличие от склейки, это не пунктуация/перенос, а
            # искажение самого слова, поэтому сравниваем показатель целиком с
            # каждой ОТДЕЛЬНОЙ ячейкой колонки (не с "простынёй" всего
            # столбца — расстояние Левенштейна не годится для поиска
            # подстроки).
            if name_squished and len(name_squished) >= 8:
                for cells_squished in column_cells_squished:
                    matched_cell = next(
                        (c for c in cells_squished if _one_letter_diff(name_squished, c)),
                        None
                    )
                    if matched_cell is not None:
                        source_word_to_indicator[name] = indicator
                        if typo_match_report_path:
                            typo_match_report.append({
                                'table': t_index + 1,
                                'source_file': current_source_file,
                                'cell_text': matched_cell,
                                'matched_candidate': name_squished,
                                'matched_indicator': indicator,
                            })
                        break

        if not source_word_to_indicator:
            # fallback на все показатели из Excel, если в Word ничего не найдено
            source_word_to_indicator = all_file_entries.copy()

        if not source_word_to_indicator:
            print(f"⚠️ Нет показателей для источника {current_source_file}. Пропускаем таблицу.")
            continue

        header_rows = _find_header_rows(table, source_word_to_indicator)
        header_rows_set = set(header_rows)

        # Многие "шапки" занимают несколько подряд идущих строк (например, строка
        # с базовым названием показателя + строка с суффиксом/годом), и обе строки
        # независимо распознаются _find_header_rows как заголовочные. Раньше для
        # КАЖДОЙ такой строки отдельно вызывался _compute_section_mapping, и обе
        # попадали в section_ranges — а поскольку граница секции считается как
        # "конец = начало следующей", это давало пустой диапазон для первой строки
        # блока (чей маппинг обычно правильный, т.к. содержит различающий текст) и
        # растягивало реальные данные на маппинг последней строки блока (которая
        # часто содержит только повторяющийся суффикс, одинаковый для нескольких
        # колонок, и потому не может их различить). Группируем подряд идущие
        # заголовочные строки в один "run" и используем только его ПОСЛЕДНЮЮ
        # строку как границу секции — а сам маппинг строим по всем строкам run'а.
        header_runs = []
        current_run = []
        for idx in header_rows:
            if current_run and idx != current_run[-1] + 1:
                header_runs.append(current_run)
                current_run = []
            current_run.append(idx)
        if current_run:
            header_runs.append(current_run)

        section_ranges = []
        for run in header_runs:
            header_idx = run[-1]
            table_near_misses = [] if near_miss_report_path else None
            table_squish_matches = [] if squish_match_report_path else None
            mapping = _compute_section_mapping(table, header_idx, source_word_to_indicator,
                                                header_rows=header_rows_set, run_rows=run,
                                                near_miss_sink=table_near_misses,
                                                squish_match_sink=table_squish_matches)
            if table_near_misses:
                near_miss_report.extend(
                    {
                        'table': t_index + 1,
                        'source_file': current_source_file,
                        'column': col_idx,
                        'header_row': header_idx + 1,
                        'cell_text': cell_text,
                        'score': round(score, 2),
                        'closest_candidate': closest_name,
                        'closest_indicator': closest_indicator,
                    }
                    for col_idx, cell_text, score, closest_name, closest_indicator in table_near_misses
                )
            if table_squish_matches:
                squish_match_report.extend(
                    {
                        'table': t_index + 1,
                        'source_file': current_source_file,
                        'column': col_idx,
                        'header_row': header_idx + 1,
                        'cell_text': cell_text,
                        'matched_candidate': matched_name,
                        'matched_indicator': matched_indicator,
                    }
                    for col_idx, cell_text, matched_name, matched_indicator in table_squish_matches
                )
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
    if near_miss_report_path:
        near_miss_report_path = Path(near_miss_report_path)
        near_miss_report_path.parent.mkdir(parents=True, exist_ok=True)
        columns = ['table', 'source_file', 'column', 'header_row', 'cell_text',
                   'score', 'closest_candidate', 'closest_indicator']
        pd.DataFrame(near_miss_report, columns=columns).to_excel(near_miss_report_path, index=False)
        print(f"🔎 Отчёт по «почти совпавшим» показателям сохранён: {near_miss_report_path} "
              f"({len(near_miss_report)} строк)")
    if squish_match_report_path:
        squish_match_report_path = Path(squish_match_report_path)
        squish_match_report_path.parent.mkdir(parents=True, exist_ok=True)
        columns = ['table', 'source_file', 'column', 'header_row', 'cell_text',
                   'matched_candidate', 'matched_indicator']
        pd.DataFrame(squish_match_report, columns=columns).to_excel(squish_match_report_path, index=False)
        print(f"🧩 Отчёт по «склеенным» совпадениям сохранён: {squish_match_report_path} "
              f"({len(squish_match_report)} строк)")
    if typo_match_report_path:
        typo_match_report_path = Path(typo_match_report_path)
        typo_match_report_path.parent.mkdir(parents=True, exist_ok=True)
        columns = ['table', 'source_file', 'cell_text', 'matched_candidate', 'matched_indicator']
        pd.DataFrame(typo_match_report, columns=columns).to_excel(typo_match_report_path, index=False)
        print(f"✏️ Отчёт по опечаткам в одну букву сохранён: {typo_match_report_path} "
              f"({len(typo_match_report)} строк)")


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