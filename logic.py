# logic.py
from difflib import SequenceMatcher
from pathlib import Path
import csv
import re
from typing import Optional

import pandas as pd
from config import (
    get_excel_data,
    load_okved_map,
    load_table_source_map,
    load_manual_table_numbers,
    load_column_mapping,
    get_cleaned_cell_text,
    looks_like_data_value,
    set_paragraph_text_keep_format,
    _normalize_text,
    find_okved_code,
    get_table_name,
    canonical_okved,
    _read_csv_rows_robustly,
    _table_numbers,
)
from mo import load_mo_map, find_mo_code, canonical_mo
from config_v2 import load_column_mapping_v2
from category_mapping import (
    canonical_category,
    extract_category_codes_from_excel,
    detect_category_prefix,
)
from docx import Document

TAG_REGEX = re.compile(r"{{([^}]+?)_([0-9]+)}}")
YEAR_PATTERN = re.compile(r'\b(20\d{2})\b')


def _extract_year(text: str) -> Optional[str]:
    if not text:
        return None
    match = YEAR_PATTERN.search(text)
    return match.group(1) if match else None


def _is_year_header_row(row) -> bool:
    """Строка шапки с годами: название пустое, все непустые ячейки — годы.

    Строка без узнанного названия зачищается как «неизвестная отрасль»:
    числовые ячейки заменяются прочерком. Число от подписи отличается по
    отсутствию букв (looks_like_data_value), и год «2024» под это попадает.
    Когда в одну таблицу Word склеено несколько страниц-продолжений, их шапки
    с годами оказываются среди строк данных — и очистка стирала в них годы.
    В готовых документах вместо «2023 | 2024» стояли прочерки: бюллетень №1,
    таблица 33 — 18 ячеек; бюллетень №2, таблица 29 — 16. В бюллетене №2 без
    годов блок ниже ещё и размечался неверно: теги без года, 6 колонок из 8
    без тегов.

    Проверка по строке, а не по ячейке: число 2024 в ячейке данных вполне
    возможно, а строка без названия, где ВСЕ значения — годы, — нет.
    """
    if get_cleaned_cell_text(row.cells[0]).strip():
        return False
    texts = []
    prev_tc = row.cells[0]._tc
    for cell in row.cells[1:]:
        if cell._tc is prev_tc:
            continue
        prev_tc = cell._tc
        text = get_cleaned_cell_text(cell).strip()
        if text:
            texts.append(text)
    return len(texts) >= 2 and all(re.fullmatch(r'20\d\d', text) for text in texts)


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


# Короткое слово в скобках, которое в Word могли просто не написать: в Excel
# бюллетеня №2 «Темп роста прибыли (убытка) в %», а в Word — «Темп роста
# прибыли, в %». Это не опечатка в букве, а пропущенное целиком слово, поэтому
# _one_letter_diff его не ловит, а нечёткое сравнение выбирало более похожий,
# но чужой показатель «Темп роста суммы прибыли». Скобка — ровно одно слово из
# букв (не «(-)», не «(в % к итогу)»), чтобы правило не съедало смысл.
_BRACKET_WORD_RE = re.compile(r'\(\s*([^\W\d_]{2,12})\s*\)')
# Остаток названия без слова в скобках должен быть не короче трёх слов, иначе
# «Прибыль (убыток)» превратилась бы в «прибыль» и совпала бы с чем угодно.
_BRACKET_VARIANT_MIN_WORDS = 3


def _bracket_word_variants(name: str) -> dict:
    """
    Варианты названия показателя без одного короткого слова в скобках:
    нормализованный вариант -> выброшенное слово. Сравниваются только
    ЦЕЛИКОМ с заголовком колонки Word (не по вхождению).
    """
    variants = {}
    if not name:
        return variants
    for match in _BRACKET_WORD_RE.finditer(name):
        shortened = _normalize_match_text(name[:match.start()] + ' ' + name[match.end():])
        if len(shortened.split()) >= _BRACKET_VARIANT_MIN_WORDS:
            variants[shortened] = match.group(1)
    return variants


# Слова, которые Росстат добавил к названиям показателей в Excel, а в Word их
# нет. В бюллетене №3 (форма отчётности 2025 года): «Прибыль (убыток) от
# продолжающейся деятельности до налогообложения» — в Word по-прежнему
# «Прибыль (убыток) до налогообложения». Это есть в t23, t25 и t27, и без
# правила колонки оставались пустыми во всех этих таблицах — и так будет у
# любого региона. Совпадение, как и у слова в скобках, — только целиком и
# с записью в отчёт «проверьте названия Word и Excel».
_OPTIONAL_PHRASES = ('от продолжающейся деятельности',)


def _name_variants(name: str) -> dict:
    """
    Все допустимые укороченные варианты названия показателя из Excel:
    нормализованный вариант -> что выброшено (для отчёта): «(убытка)» —
    короткое слово в скобках (см. _bracket_word_variants) или фраза из
    _OPTIONAL_PHRASES.
    """
    variants = {variant: f"({word})" for variant, word in _bracket_word_variants(name).items()}
    name_norm = _normalize_match_text(name)
    for phrase in _OPTIONAL_PHRASES:
        phrase_norm = _normalize_match_text(phrase)
        padded = f" {name_norm} "
        if f" {phrase_norm} " not in padded:
            continue
        shortened = ' '.join(padded.replace(f" {phrase_norm} ", ' ', 1).split())
        if len(shortened.split()) >= _BRACKET_VARIANT_MIN_WORDS:
            variants.setdefault(shortened, phrase)
    return variants


def _build_bracket_variant_map(names) -> dict:
    """нормализованный укороченный вариант (см. _name_variants) -> список нормализованных названий."""
    variant_map = {}
    for name in names:
        name_norm = _normalize_match_text(name)
        for variant in _name_variants(name):
            if name_norm not in variant_map.setdefault(variant, []):
                variant_map[variant].append(name_norm)
    return variant_map


# Насколько надёжно колонка сопоставлена с показателем: меньше — надёжнее.
# Нужно, когда две колонки получили один и тот же показатель: его забирает
# колонка с дословным совпадением, а не та, что стоит левее.
_MATCH_RANK = {'exact': 0, 'bracket': 1, 'fuzzy': 2, 'squish': 3}


def _resolve_indicator(normalized_name_map: dict, name_norm: str, consumed: Optional[dict]) -> str:
    """
    Превращает список индикаторов, зарегистрированных под одним и тем же
    нормализованным названием, в ОДИН конкретный код — с учётом того, сколько
    раз это название уже было "занято" ранее В ПРЕДЕЛАХ ТЕКУЩЕЙ секции
    (см. consumed в _compute_section_mapping).

    Нужно для случаев, когда две РАЗНЫЕ колонки таблицы имеют дословно
    ОДИНАКОВЫЙ заголовок (например, "в % к общей задолженности" — один раз
    относительно дебиторской задолженности, другой раз относительно
    кредиторской), но по смыслу это два разных показателя. В CSV для такого
    названия регистрируются НЕСКОЛЬКО строк подряд — и колонки должны разбирать
    их по порядку появления в CSV, а не все получать один и тот же (первый)
    индикатор, откуда предыдущий код молча выкидывал бы все, кроме первого, как
    "дубликат".
    """
    indicators = normalized_name_map[name_norm]
    if consumed is None:
        return indicators[0]
    count = consumed.get(name_norm, 0)
    consumed[name_norm] = count + 1
    idx = min(count, len(indicators) - 1)
    return indicators[idx]


def _fuzzy_match_header(target_text: str, normalized_name_map: dict, threshold: float = 0.80,
                         near_miss_sink: Optional[list] = None,
                         squish_match_sink: Optional[list] = None,
                         consumed: Optional[dict] = None,
                         bracket_variants: Optional[dict] = None,
                         bracket_match_sink: Optional[list] = None,
                         match_kind_sink: Optional[list] = None) -> str:
    """
    Ищет наилучшее совпадение заголовка с использованием Fuzzy Matching.
    Возвращает код индикатора или None.

    normalized_name_map: dict нормализованное_название -> СПИСОК кодов
    индикаторов (обычно из одного элемента; больше одного — когда одно и то же
    название текста в Word соответствует нескольким разным показателям CSV,
    см. _resolve_indicator).

    bracket_variants: вариант названия без короткого слова в скобках ->
    названия (см. _build_bracket_variant_map). Совпадение по нему пишется в
    bracket_match_sink как (заголовок, название, код), чтобы попасть в отчёт.

    match_kind_sink, если передан, получает вид найденного совпадения
    (ключ _MATCH_RANK) — по нему решается, какая из двух колонок с одним и
    тем же показателем его получит.
    """
    if not target_text:
        return None

    def _found(kind, indicator):
        if match_kind_sink is not None:
            match_kind_sink.append(kind)
        return indicator

    # 1. Точное вхождение (в любую сторону). Среди ВСЕХ кандидатов с вхождением
    # выбираем тот, у кого выше итоговый Ratio, а не первый попавшийся по порядку
    # словаря. Это важно, когда одно название — префикс другого (например,
    # "Себестоимость продаж" является префиксом "Себестоимость продаж с учетом
    # коммерческих и управленческих расходов") — иначе порядок перебора мог бы
    # случайно "перетянуть" короткую колонку на длинный показатель.
    best_exact_name = None
    best_exact_score = -1.0
    for name_norm in normalized_name_map:
        if not name_norm:
            continue
        if name_norm in target_text or target_text in name_norm:
            score = SequenceMatcher(None, target_text, name_norm).ratio()
            if score > best_exact_score:
                best_exact_score = score
                best_exact_name = name_norm

    if best_exact_name is not None:
        return _found('exact', _resolve_indicator(normalized_name_map, best_exact_name, consumed))

    # 1а. Название из Excel без одного короткого слова в скобках совпало с
    # заголовком Word ЦЕЛИКОМ: «Темп роста прибыли (убытка) в %» и «Темп
    # роста прибыли, в %». Проверяем раньше нечёткого сравнения — оно здесь
    # выбирало «Темп роста суммы прибыли» (0,87 против 0,85). Если под вариант
    # подходят два разных показателя — не угадываем.
    bracket_names = (bracket_variants or {}).get(target_text, [])
    bracket_names = [name for name in bracket_names if name in normalized_name_map]
    if len(bracket_names) == 1:
        bracket_match = _resolve_indicator(normalized_name_map, bracket_names[0], consumed)
        print(f"   🔗 [Слово в скобках] '{target_text[:50]}' -> '{bracket_names[0][:50]}' -> {bracket_match}")
        if bracket_match_sink is not None:
            bracket_match_sink.append((target_text, bracket_names[0], bracket_match))
        return _found('bracket', bracket_match)

    # 2. Нечеткое сравнение (Fuzzy Match) для случаев без точного вхождения
    best_match_name = None
    best_score = 0.0
    for name_norm in normalized_name_map:
        if not name_norm:
            continue
        score = SequenceMatcher(None, target_text, name_norm).ratio()
        if score > best_score:
            best_score = score
            best_match_name = name_norm

    if best_score >= threshold:
        return _found('fuzzy', _resolve_indicator(normalized_name_map, best_match_name, consumed))

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
        best_squish_name = None
        best_squish_len = -1
        for name_norm in normalized_name_map:
            name_squished = _squish_text(name_norm)
            if not name_squished or len(name_squished) < 5:
                continue
            if name_squished == target_squished or name_squished in target_squished or target_squished in name_squished:
                # При нескольких кандидатах предпочитаем более длинное (более специфичное) совпадение
                if len(name_squished) > best_squish_len:
                    best_squish_len = len(name_squished)
                    best_squish_name = name_norm
        if best_squish_name is not None:
            best_squish_match = _resolve_indicator(normalized_name_map, best_squish_name, consumed)
            print(f"   🧩 [Squish-match] '{target_text[:50]}' -> '{best_squish_name[:50]}' -> {best_squish_match}")
            if squish_match_sink is not None:
                squish_match_sink.append((target_text, best_squish_name, best_squish_match))
            return _found('squish', best_squish_match)

    # Диагностика: если показатель почти нашелся, но не дотянул до порога — это
    # тег, который остаётся без значения молча (или, что хуже, ячейка так и
    # хранит нетронутый исходный текст, если это единственный кандидат для всей
    # строки). Раньше это уходило только в консольный лог и никем не читалось;
    # теперь также складываем в near_miss_sink, чтобы в конце прогона можно было
    # выгрузить это в отдельный отчёт для проверки человеком.
    if target_text and len(target_text) > 5 and best_score > 0.50:
        print(f"   ⚠️ [Fuzzy Miss] Ожидалось похожее, но Score={best_score:.2f} для '{target_text[:50]}...'")
        if near_miss_sink is not None:
            best_match = normalized_name_map[best_match_name][0] if best_match_name else None
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


def _file_wide_prefix(names_norm, min_words=2) -> tuple:
    """
    Начальные слова, общие для ВСЕХ показателей файла, если у каждого после
    них есть ещё слова. В t23-пр бюллетеня №3 все десять показателей
    начинаются с «Предыдущий год (сопоставимый круг)» — это заголовок всего
    файла, а не раздела. Разделов, которые можно перепутать, у такой
    приставки нет, поэтому искать её в таблице Word не нужно (в Word её и
    нет — заголовок таблицы сформулирован иначе).
    """
    lists = [n.split() for n in names_norm if n]
    if len(lists) < 2:
        return ()
    prefix = []
    for words in zip(*lists):
        if any(word != words[0] for word in words):
            break
        prefix.append(words[0])
    if len(prefix) < min_words or any(len(words) <= len(prefix) for words in lists):
        return ()
    return tuple(prefix)


def _compute_shared_group_prefixes(names_norm, min_siblings=2, min_prefix_words=2):
    """Находит "групповые" префиксы, общие для нескольких показателей одного
    Excel-файла — например, "отчетный год" в "отчетный год сальдо прочих
    доходов и расходов" / "отчетный год прибыль (убыток) до налогообложения"
    / "отчетный год прочее" и т.д.

    В исходном Excel такой групповой заголовок объединён (merge) по
    нескольким столбцам (см. _infer_mapping_from_excel в config_v2.py), и
    поэтому "приклеивается" ко всем показателям этой группы при сборке
    column_mapping_v2.csv. В Word же он обычно не повторяется в каждой
    колонке — там просто короткий подзаголовок.

    Важно искать префикс, ОБЩИЙ ХОТЯ БЫ ДЛЯ ДВУХ показателей (а не просто
    "первые N слов одного произвольного названия") — иначе можно случайно
    отрезать значимую часть уникального названия и словить ложное
    совпадение с чужой колонкой. Всегда выбирается МАКСИМАЛЬНО длинный
    общий префикс (в словах) — короткие префиксы более длинных общих
    последовательностей не включаются в результат, чтобы не заслонять
    более точный вариант при последующем поиске.
    """
    from collections import defaultdict

    word_lists = [n.split() for n in names_norm if n]
    max_len = max((len(w) for w in word_lists), default=0)

    prefixes = set()
    covered = set()  # индексы имён, для которых уже найден максимальный префикс
    for k in range(max_len - 1, min_prefix_words - 1, -1):
        groups = defaultdict(list)
        for idx, words in enumerate(word_lists):
            if idx in covered or len(words) <= k:
                continue
            groups[tuple(words[:k])].append(idx)
        for prefix, idxs in groups.items():
            if len(idxs) >= min_siblings:
                prefixes.add(prefix)
                covered.update(idxs)
    return prefixes


# Строки-разметка единиц измерения/периода, которые Word иногда повторяет как
# отдельную строку прямо над шапкой продолжения таблицы ("на конец года,
# тысяч рублей" и т.п.) — это не часть названия показателя, а общая для всей
# группы колонок пометка. Сравниваем ТОЧНО (после нормализации), а не по
# вхождению — иначе можно случайно съесть реальное название, которое просто
# содержит похожие слова.
#
# Список по умолчанию (fallback, если input/mappings/unit_annotations.csv не
# найден) — обычно фактический список загружается из этого CSV через
# load_unit_annotation_texts, чтобы пользователь мог дополнять его без
# правки кода.
_DEFAULT_UNIT_ANNOTATION_TEXTS = {
    'на конец года тысяч рублей',
    'на конец года',
    'тысяч рублей',
    'на начало года тысяч рублей',
    'на начало года',
}


def load_unit_annotation_texts(filepath) -> set:
    """
    Загружает список служебных фраз-разметок единиц измерения/периода из
    CSV (колонка "фраза"; необязательная колонка "комментарий" — только
    для человека, в сравнении не участвует). Каждая фраза нормализуется
    так же, как и сравниваемый текст (см. _normalize_match_text), чтобы
    пользователь мог писать её в любом регистре/с любой пунктуацией.

    Если файл не найден или пуст — возвращает встроенный список по
    умолчанию, чтобы поведение не менялось "из коробки".
    """
    try:
        rows = _read_csv_rows_robustly(filepath, delimiter=';')
    except FileNotFoundError:
        return set(_DEFAULT_UNIT_ANNOTATION_TEXTS)

    if rows:
        rows = rows[1:]  # заголовок

    texts = {
        _normalize_match_text(row[0])
        for row in rows
        if row and row[0].strip()
    }
    return texts or set(_DEFAULT_UNIT_ANNOTATION_TEXTS)


def _is_unit_annotation(part_norm: str, unit_annotation_texts: Optional[set] = None) -> bool:
    texts = unit_annotation_texts if unit_annotation_texts is not None else _DEFAULT_UNIT_ANNOTATION_TEXTS
    return part_norm in texts


# Метка страницы-продолжения ("Продолжение таблицы 13.") стоит над шапкой
# справа, в ячейке, объединённой над последними колонками, и попадает в
# составной заголовок этих колонок. Точное вхождение названия показателя это
# не ломает, а нечёткое сравнение — ломает: у «Прибыль (убыток) до
# налогооблажения» (опечатка в Excel бюллетеня №2) точного вхождения нет,
# а с приставкой «продолжение таблицы 13» сходство падает с 0,97 до 0,72,
# ниже порога 0,80. Колонка оставалась без тегов на всех продолжениях.
_CONTINUATION_LABEL_RE = re.compile(r'продолжение\s+таблицы\s*\d+\s*[.)]?')


def _strip_continuation_label(text: str) -> str:
    """Убирает из нормализованного текста шапки метку "продолжение таблицы N"."""
    if not text or 'продолжение' not in text:
        return text
    return re.sub(r'\s+', ' ', _CONTINUATION_LABEL_RE.sub(' ', text)).strip()


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


def _is_row_label_only(row) -> bool:
    """
    Текст есть только в колонке названий строк, все остальные ячейки пусты:
    «в том числе:», «Городские округа:». Шапка подписывает колонки с
    числами, а такая строка не подписывает ни одной, поэтому шапкой быть
    не может, даже если её текст входит в названия показателей.

    В бюллетене №2 показатели t03Ved14 называются «Внеоборотные активы в
    том числе материальные…» и «…в том числе нематериальные…». Строка «в
    том числе:» под «Всего» в таблице 6 набирала два совпадения и
    считалась шапкой новой секции, где размечена одна колонка названий, —
    и вся первая страница таблицы осталась без тегов.
    """
    cells = row.cells
    if len(cells) < 2 or not get_cleaned_cell_text(cells[0]).strip():
        return False
    first_tc = cells[0]._tc
    others = [cell for cell in cells[1:] if cell._tc is not first_tc]
    return bool(others) and all(not get_cleaned_cell_text(cell).strip() for cell in others)


def _find_header_rows(table, source_word_to_indicator, max_search_rows=None, entity_name_maps=None):
    """Собирает все строки заголовков таблицы (year или indicator rows).

    По умолчанию просматривается вся таблица. Раньше — только первые 80
    строк, но в одну таблицу Word бывает склеено несколько страниц-
    продолжений, каждая со своей шапкой. В бюллетене №2 таблица 29 — это 102
    строки, и шапки последних страниц (строки 78–82 и 91–95) оказывались за
    границей: их строки размечались по раскладке предыдущей страницы, где
    колонки начинаются с другого места, — 28 тегов без года и 6 колонок из 8
    без тегов. В бюллетене №1 то же было в таблице 33 (127 строк), но там у
    всех страниц одинаковая раскладка, и чужая шапка случайно подходила.

    entity_name_maps: необязательный список справочников "название -> код"
    (ОКВЭД/МО/категория — см. category_mapping.py), которые заведомо
    относятся к СТРОКАМ ДАННЫХ, а не заголовков. Если колонка 0 строки-
    кандидата в заголовки однозначно опознаётся как одна из этих сущностей
    (например, "Всего по обследуемым видам экономической деятельности"),
    строка исключается из результата, даже если она набрала нужный score по
    текстовому сходству с показателями — реальный заголовок никогда не
    содержит в колонке 0 название/код сущности, там всегда название метрики.
    Это ЧИСТО ИСКЛЮЧАЮЩЕЕ правило (может только убрать строку из headers,
    никогда не добавить) — специально сделано так, чтобы минимизировать риск
    задеть другие, уже правильно работающие случаи классификации.
    """
    normalized_names = [_normalize_match_text(name) for name in source_word_to_indicator.keys()]
    headers = []
    rows = table.rows[:max_search_rows]
    for row_idx in range(len(rows)):
        row = rows[row_idx]
        if _is_row_label_only(row):
            continue
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
            # Следующая строка — часть шапки, только если склейка с ней нашла
            # БОЛЬШЕ показателей, чем эта строка в одиночку, то есть
            # следующая строка действительно продолжает шапку. Раньше хватало
            # combined_score >= 2, и если строка шапки сама набирала два
            # совпадения, в шапку уходила и строка данных под ней. Строки
            # шапки не очищаются, и в бюллетене №2 в строке «азартные игры»
            # таблицы 21 остались цифры прошлого бюллетеня.
            if (combined_score >= 2 and combined_score > indicator_score
                    and not _is_row_label_only(rows[row_idx + 1])):
                headers.append(row_idx + 1)

    if not entity_name_maps:
        return sorted(set(headers))

    filtered_headers = []
    for row_idx in sorted(set(headers)):
        if row_idx >= len(rows) or not rows[row_idx].cells:
            filtered_headers.append(row_idx)
            continue
        col0_text = get_cleaned_cell_text(rows[row_idx].cells[0]).strip()
        is_known_entity = bool(col0_text) and any(
            find_okved_code(col0_text, name_map) for name_map in entity_name_maps if name_map
        )
        if not is_known_entity:
            filtered_headers.append(row_idx)
    return filtered_headers


def _compute_section_mapping(table, header_idx, source_word_to_indicator, header_rows=None, run_rows=None,
                              near_miss_sink=None, squish_match_sink=None, unit_annotation_sink=None,
                              unit_annotation_texts=None, duplicate_year_sink=None,
                              bracket_match_sink=None, taken_column_sink=None):
    """Вычисляет маппинг столбцов для данной секции таблицы.

    bracket_match_sink, если передан, получает (column_idx, header_text,
    matched_name, matched_indicator) на каждую колонку, узнанную только
    правилом «в Word не написано короткое слово в скобках» (см.
    _bracket_word_variants).

    taken_column_sink, если передан, получает (column_idx, header_text,
    indicator, winner_column_idx) на каждую колонку, оставленную пустой,
    потому что её показатель надёжнее (дословно) совпал с другой колонкой.

    near_miss_sink, если передан, получает по одному элементу
    (column_idx, cell_text, score, closest_name, closest_indicator) на каждую
    почти-но-не-дотянувшую fuzzy-попытку — чтобы такие случаи не терялись
    молча, а могли попасть в отчёт для проверки человеком.

    squish_match_sink, если передан, получает по одному элементу
    (column_idx, cell_text, matched_name, matched_indicator) на каждое
    совпадение, найденное только через самый толерантный уровень сравнения
    (_squish_text) — тоже стоит проверять человеку, см. _fuzzy_match_header.

    unit_annotation_sink, если передан, получает по одному элементу
    (column_idx, row_idx, dropped_text) на каждую строку-разметку единиц
    измерения/периода (см. _is_unit_annotation), выброшенную из составного
    заголовка колонки — чтобы было видно, где именно и что было отфильтровано,
    а не просто молча выброшено.

    duplicate_year_sink, если передан, получает по одному элементу
    (column_idx, indicator, year_code) на каждую колонку, которая была
    выброшена ниже как дубликат (тот же показатель и тот же год, что уже
    занял другую колонку). Реальная причина такого дубля почти всегда —
    опечатка в самой шапке документа (например, «2023» напечатан дважды
    подряд вместо «2023»/«2024») — см. НАЙДЕННЫЕ_ДЕФЕКТЫ_ШАБЛОНА.md, пункт 1.
    Раньше это было видно только как строка в консоли, которую легко
    пропустить в потоке вывода.
    """
    year_row = table.rows[header_idx]
    year_texts = [get_cleaned_cell_text(cell).strip() for cell in year_row.cells]
    has_years = sum(1 for text in year_texts if _extract_year(text)) >= 2

    col_to_indicator_map = {}
    col_rank = {}
    col_header_text = {}
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
        last_rank = None
        normalized_name_map = {
            _normalize_match_text(name): indicators
            for name, indicators in source_word_to_indicator.items()
        }
        bracket_variants = _build_bracket_variant_map(source_word_to_indicator)
        # Отслеживает, сколько раз уже было "занято" каждое название — нужно,
        # когда одно название соответствует НЕСКОЛЬКИМ разным индикаторам
        # (см. _resolve_indicator).
        consumed = {}
        prev_tc = None
        for i, cell in enumerate(year_row.cells):
            # python-docx повторяет один и тот же физический <w:tc> на КАЖДОЙ
            # позиции сетки, которую он перекрывает через gridSpan (например,
            # ячейка "в % к общей задолженности", реально одна, но растянутая
            # на 2 колонки сетки, отдаётся .cells как ДВА одинаковых элемента
            # подряд). Без этой проверки счётчик consumed (см.
            # _resolve_indicator) продвигается по списку индикаторов дважды
            # за одну и ту же ячейку — и настоящая следующая колонка с тем же
            # названием получает уже занятый индикатор, а не свой.
            if cell._tc is prev_tc:
                continue
            prev_tc = cell._tc
            raw_year_text = get_cleaned_cell_text(cell).strip()
            year_text = _extract_year(raw_year_text)

            parts = []
            seen = set()
            for header_row in header_rows_filled:
                if i >= len(header_row):
                    continue
                header_part = _strip_continuation_label(header_row[i].strip())
                if not header_part or header_part in seen:
                    continue
                seen.add(header_part)
                parts.append(header_part)
            # Строка годов бывает смешанной: у одних колонок в ней год, у
            # других — последняя строка их шапки. В бюллетене №2 у колонок
            # «Валюта баланса убыточных организаций, на конец года» в этой
            # строке стоят подписи «итог актива убыточных организаций» и
            # «в % к валюте баланса». Без них от шапки оставалось только
            # групповое название, оно по началу совпадало с «Валютой баланса»,
            # и колонка убыточных организаций размечалась как валюта баланса
            # всех организаций (без года — данных не нашлось, иначе встало бы
            # чужое число), а колонка процентов не размечалась вовсе.
            year_row_label = _normalize_text(raw_year_text) if not year_text else ""
            if year_row_label and year_row_label not in seen:
                parts.append(year_row_label)
            composed_header = " ".join(parts)
            composed_header_norm = _normalize_match_text(composed_header)

            # Используем Fuzzy-поиск вместо жесткого вхождения
            column_near_misses = [] if near_miss_sink is not None else None
            column_squish_matches = [] if squish_match_sink is not None else None
            column_bracket_matches = []
            match_kinds = []
            best_match = _fuzzy_match_header(composed_header_norm, normalized_name_map, threshold=0.80,
                                              near_miss_sink=column_near_misses,
                                              squish_match_sink=column_squish_matches,
                                              consumed=consumed,
                                              bracket_variants=bracket_variants,
                                              bracket_match_sink=column_bracket_matches,
                                              match_kind_sink=match_kinds)
            if column_near_misses:
                near_miss_sink.extend((i, *entry) for entry in column_near_misses)
            if column_squish_matches:
                squish_match_sink.extend((i, *entry) for entry in column_squish_matches)
            if column_bracket_matches and bracket_match_sink is not None:
                bracket_match_sink.extend((i, *entry) for entry in column_bracket_matches)
            match_rank = _MATCH_RANK[match_kinds[-1]] if match_kinds else None

            # Fuzzy-fallback для случаев с переносами, дефисами и неявными формулировками
            if not best_match and composed_header_norm:
                fuzzy_score = 0.0
                fuzzy_match_name = None
                for name_norm in normalized_name_map:
                    if not name_norm:
                        continue
                    score = SequenceMatcher(None, name_norm, composed_header_norm).ratio()
                    if score > fuzzy_score:
                        fuzzy_score = score
                        fuzzy_match_name = name_norm
                if fuzzy_score >= 0.75:
                    best_match = _resolve_indicator(normalized_name_map, fuzzy_match_name, consumed)
                    match_rank = _MATCH_RANK['fuzzy']
                    print(f"   🔍 fuzzy match {fuzzy_score:.2f} для '{composed_header[:80]}' -> {best_match}")

            if not year_text:
                # Секция смешивает колонки, разбитые по годам (2023/2024,
                # например "Уставный капитал"), с колонками БЕЗ такой
                # разбивки — одно значение на весь период (например,
                # "Себестоимость продаж" в той же таблице "Чистые активы").
                # Раньше такая колонка просто пропускалась целиком (год не
                # распознан → continue), хотя её заголовок находится точно
                # так же, как и у остальных — просто без суффикса года.
                if best_match:
                    col_to_indicator_map[i] = (best_match, None)
                    col_rank[i] = match_rank
                    col_header_text[i] = composed_header
                    print(f"   🔍 Столбец {i} (без года): '{composed_header[:80]}' -> {best_match}")
                continue

            if best_match:
                last_indicator = best_match
                last_rank = match_rank
            elif last_indicator and not composed_header_norm:
                best_match = last_indicator
                match_rank = last_rank

            if not best_match:
                print(f"   ⚠️ Нет match для колонки {i} ('{composed_header[:80]}')")
                continue

            # Используем сокращенный год из Word: 2023 -> 23, 2024 -> 24
            # Это даёт гибкие теги на будущее без правки кода из-за нового календарного года.
            year_code = year_text[-2:]
            col_to_indicator_map[i] = (best_match, year_code)
            col_rank[i] = match_rank
            col_header_text[i] = composed_header
            print(f"   🔍 Столбец {i}: '{composed_header[:90]}' -> {best_match}, год {year_text} -> код {year_code}")
    else:
        header_row = year_row
        normalized_name_map = {
            _normalize_match_text(name): indicators
            for name, indicators in source_word_to_indicator.items()
        }
        # Отслеживает, сколько раз уже было "занято" каждое название — нужно,
        # когда одно название соответствует НЕСКОЛЬКИМ разным индикаторам
        # (см. _resolve_indicator).
        consumed = {}
        bracket_variants = _build_bracket_variant_map(source_word_to_indicator)
        # header_idx — это ПОСЛЕДНЯЯ строка своего "run"-а подряд идущих
        # заголовочных строк (см. вызывающий код), поэтому строка header_idx+1 —
        # это уже настоящие данные, а не продолжение шапки. Составной заголовок
        # строим по ВСЕМ строкам run'а (обычно это "базовое название показателя"
        # + "суффикс/подзаголовок"), а не только по одной последней строке —
        # иначе несколько колонок с одинаковым суффиксом (например, у двух разных
        # показателей "длительность 1 оборота") неотличимы друг от друга.
        run_rows_sorted = sorted(run_rows) if run_rows else [header_idx]

        prev_tc = None
        for i, cell in enumerate(header_row.cells):
            # См. аналогичную проверку выше (has_years-ветка): python-docx
            # повторяет один и тот же физический <w:tc> на каждой позиции
            # сетки, которую он занимает через gridSpan — без пропуска такой
            # "фантомной" колонки consumed продвигается по списку индикаторов
            # лишний раз, и настоящая следующая колонка с тем же названием
            # получает уже занятый чужой индикатор.
            if cell._tc is prev_tc:
                continue
            prev_tc = cell._tc
            parts = []
            seen = set()
            for r_idx in run_rows_sorted:
                if r_idx >= len(table.rows):
                    continue
                row_cells = table.rows[r_idx].cells
                if i >= len(row_cells):
                    continue
                part = _strip_continuation_label(_normalize_match_text(get_cleaned_cell_text(row_cells[i])))
                # Строка-разметка единиц измерения ("на конец года, тысяч
                # рублей" и т.п.) иногда попадает в run строки-заголовка (когда
                # секция начинается со страницы "Продолжение таблицы N.") и
                # тогда примешивается к составному названию колонки, сбивая
                # score ниже порога принятия (0.80) — реальное название
                # показателя при этом само по себе совпадает точно.
                if part and _is_unit_annotation(part, unit_annotation_texts):
                    if unit_annotation_sink is not None:
                        unit_annotation_sink.append((i, r_idx, part))
                    continue
                if part and part not in seen:
                    seen.add(part)
                    parts.append(part)
            header_text = ' '.join(parts).strip()
            if not header_text:
                continue

            column_near_misses = [] if near_miss_sink is not None else None
            column_squish_matches = [] if squish_match_sink is not None else None
            column_bracket_matches = []
            match_kinds = []
            best_match = _fuzzy_match_header(header_text, normalized_name_map, threshold=0.80,
                                              near_miss_sink=column_near_misses,
                                              squish_match_sink=column_squish_matches,
                                              consumed=consumed,
                                              bracket_variants=bracket_variants,
                                              bracket_match_sink=column_bracket_matches,
                                              match_kind_sink=match_kinds)
            if column_near_misses:
                near_miss_sink.extend((i, *entry) for entry in column_near_misses)
            if column_squish_matches:
                squish_match_sink.extend((i, *entry) for entry in column_squish_matches)
            if column_bracket_matches and bracket_match_sink is not None:
                bracket_match_sink.extend((i, *entry) for entry in column_bracket_matches)

            if best_match:
                col_to_indicator_map[i] = (best_match, None)
                col_rank[i] = _MATCH_RANK[match_kinds[-1]] if match_kinds else None
                col_header_text[i] = header_text
                print(f"   🔍 Индикаторный заголовок: столбец {i}, текст '{header_text[:50]}' -> {best_match}")

    # Убираем дубликаты. Показатель с годом получает одна колонка. Раньше —
    # всегда самая левая, даже если она узнана лишь приблизительно, а правее
    # стоит колонка с дословно тем же названием: в бюллетене №2 «Темп роста
    # прибыли» (похожесть 0,87) забрала показатель «Темп роста суммы
    # прибыли», получила его числа, а настоящая колонка осталась пустой.
    # Теперь показатель достаётся самому надёжному совпадению, при равной
    # надёжности — по-прежнему самой левой колонке.
    def _rank(col_idx):
        rank = col_rank.get(col_idx)
        return len(_MATCH_RANK) if rank is None else rank

    winner_by_spec = {}
    for col_idx in sorted(col_to_indicator_map):
        spec = col_to_indicator_map[col_idx]
        winner = winner_by_spec.get(spec)
        if winner is None or _rank(col_idx) < _rank(winner):
            winner_by_spec[spec] = col_idx

    filtered_map = {}
    for col_idx in sorted(col_to_indicator_map):
        spec = col_to_indicator_map[col_idx]
        winner = winner_by_spec[spec]
        if col_idx == winner:
            filtered_map[col_idx] = spec
            continue
        if _rank(col_idx) > _rank(winner):
            print(f"   ⚠️ Столбец {col_idx} оставлен пустым: показатель {spec[0]}_{spec[1] if spec[1] else ''} "
                  f"надёжнее совпал со столбцом {winner}")
            if taken_column_sink is not None:
                taken_column_sink.append((col_idx, col_header_text.get(col_idx, ''), spec[0], winner))
            continue
        print(f"   ⚠️ Пропускаем дубликат столбца {col_idx} для {spec[0]}_{spec[1] if spec[1] else ''} "
              f"— похоже на опечатку в шапке (например, один и тот же год напечатан дважды)")
        if duplicate_year_sink is not None and spec[1]:
            duplicate_year_sink.append((col_idx, spec[0], spec[1]))

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


def _clear_manual_table(table) -> int:
    """Стирает числа прошлого периода в таблице, которую заполняют вручную.

    Ячейки остаются ПУСТЫМИ, а не с прочерком: прочерк значит «данных нет»,
    а здесь данные есть — их просто вносит человек. Шапка не трогается:
    в её строках нет названия в первой колонке (годы, «на конец года»).
    Ячейка, объединённая с названием строки, тоже не трогается.
    """
    cleared = 0
    for row in table.rows:
        cells = row.cells
        if not cells or not get_cleaned_cell_text(cells[0]).strip():
            continue
        name_tc = cells[0]._tc
        prev_tc = name_tc
        for cell in cells[1:]:
            if cell._tc is name_tc or cell._tc is prev_tc:
                continue
            prev_tc = cell._tc
            if looks_like_data_value(get_cleaned_cell_text(cell)):
                for p in cell.paragraphs:
                    set_paragraph_text_keep_format(p, "")
                cleared += 1
    return cleared


def get_table_source_by_number(table_source_mapping, table_number):
    """Источник для метки "Продолжение таблицы N".

    Номер берём из самого названия таблицы ("13. ФОРМИРОВАНИЕ ..."), а не из
    позиции строки в маппинге. Позиции доверять нельзя: load_table_source_map
    отбрасывает строки с пустым источником (сводные таблицы, которые ничем не
    заполняются), и порядковый номер перестаёт совпадать с номером таблицы в
    бюллетене. В бюллетене с тремя такими строками в начале сдвиг составлял 3:
    "Продолжение таблицы 4" получало источник таблицы 7, а продолжения таблиц
    21-23 выходили за границу списка и не получали источника вовсе.

    Если ни одно название не пронумеровано, поведение прежнее — по позиции.
    """
    if table_number < 1:
        return None

    numbered = False
    for title, source in table_source_mapping.items():
        match = re.match(r'\s*(\d+)\s*[.)]', str(title))
        if not match:
            continue
        numbered = True
        if int(match.group(1)) == table_number:
            return source

    if numbered:
        # Нумерация есть, но такого номера в маппинге нет — это не повод
        # молча отдать чужой источник по позиции.
        return None

    if table_number > len(table_source_mapping):
        return None
    return list(table_source_mapping.values())[table_number - 1]


# ==========================================================
# === ОСНОВНАЯ ЛОГИКА ======================================
# ==========================================================
def generate_word_template(input_doc_path, okved_map_path, table_source_mapping_path, column_mapping_path,
                           output_doc_path,
                           mo_map_path: Path = None, validation_log_path: Path = None,
                           near_miss_report_path: Path = None, squish_match_report_path: Path = None,
                           typo_match_report_path: Path = None, excel_dir: Path = None,
                           unit_annotation_report_path: Path = None, unit_annotations_path: Path = None,
                           group_prefix_match_report_path: Path = None,
                           unused_indicator_report_path: Path = None,
                           duplicate_year_report_path: Path = None,
                           name_check_report_path: Path = None,
                           clear_only: bool = False):
    """
    Генерация шаблона Word с тегами {{OKVED_<code>_<indicator>[_22|_23]}},
    {{MO_<code>_<indicator>[_22|_23]}} и {{OPF_<code>_<indicator>}}/{{FS_<code>_<indicator>}}
    (см. category_mapping.py — таблицы, где строки размечены не кодом ОКВЭД/МО,
    а названием категории — организационно-правовой формы или формы
    собственности).
    Расширенный поиск заголовков по первым 5 строкам таблицы.

    excel_dir: папка с исходными Excel-файлами — нужна, чтобы построить
    справочник "название категории -> код" для категорийных файлов (см.
    category_mapping.detect_category_prefix). Если не указана, категорийные
    таблицы обрабатываются как раньше (строки останутся без тегов).

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

    unit_annotation_report_path: если указан, сюда выгружается таблица всех
    строк-разметок единиц измерения/периода (см. load_unit_annotation_texts),
    которые были отфильтрованы из составного названия колонки при
    сопоставлении — чтобы это не терялось молча, а было видно, что именно и
    где было выброшено (и можно было проверить, что фильтр не съел что-то
    лишнее).

    unit_annotations_path: путь к CSV со списком фраз-разметок единиц
    измерения/периода (см. load_unit_annotation_texts) — пользователь может
    редактировать этот файл, не трогая код. Если не указан или не найден,
    используется встроенный список по умолчанию.

    group_prefix_match_report_path: если указан, сюда выгружается таблица
    совпадений, найденных только после отбрасывания общего "группового"
    префикса показателя (см. _infer_mapping_from_excel в config_v2.py) —
    случаев, когда в column_mapping_v2.csv показатель записан как
    "<общий заголовок группы колонок> <имя колонки>", а в самом
    Word-документе групповой заголовок не повторяется в каждой колонке.
    Такие совпадения проходят автоматически, но стоит проверить, что
    остаток названия не совпал случайно не с той колонкой — а лучше
    почистить сам column_mapping_v2.csv, укоротив название показателя.

    unused_indicator_report_path: если указан, сюда выгружается таблица
    показателей из column_mapping_v2.csv, которые НИ РАЗУ не попали в пул
    кандидатов ни одной таблицы своего файла-источника — то есть для них
    нигде в документе не нашлось даже отдалённого текстового совпадения.
    Для каждого такого показателя дополнительно считается лучшее нечёткое
    совпадение (SequenceMatcher) среди всех текстов колонок этого файла:
    высокий score (например, ближе к 1.0) — верный признак того, что
    показатель просто по-другому написан в Excel и в Word (например,
    "особо ценного" в Excel против "особо целевого" в самом
    Word-документе — реальный случай, который эта проверка и должна
    ловить), и это нужно поправить руками, сверив оба источника; низкий
    score — показатель, вероятно, в этой редакции документа просто
    отсутствует, и это не обязательно ошибка.

    name_check_report_path: если указан, сюда выгружается отчёт для
    специалиста «проверьте названия Word и Excel»: колонки, узнанные только
    правилом «в Word не написано короткое слово в скобках», и колонки,
    оставленные пустыми, потому что их показатель дословно совпал с другой
    колонкой. По каждой — где она в Word, из какого Excel и какой колонки
    берутся числа, и что поправить, если это неверно.

    clear_only: если True, функция работает как отдельный шаг "очистка
    данных" — находит те же самые ячейки, что и обычно, но вместо
    вставки тега просто ОПУСТОШАЕТ ячейку (старое число из исходного
    бюллетеня удаляется, шапка с годами не трогается). Тег в эту же
    ячейку ставится уже ВТОРЫМ, отдельным проходом (обычный вызов этой
    же функции с clear_only=False, на этот раз — над уже очищенным
    файлом). Это разделение специально сделано так, чтобы в репозитории
    оставался явный промежуточный результат "пустой бланк без чисел и
    без тегов" — специалисты, которые открывают файл между этими двумя
    шагами, не путаются, откуда взялись цифры или что означают {{...}}.
    """
    print("\n--- ШАГ 2: Генерация шаблона с умными тегами ---" if not clear_only
          else "\n--- ШАГ 1.5: Очистка шаблона от данных ---")
    okved_to_name, name_to_okved_cleaned = load_okved_map(okved_map_path)
    mo_to_name, mo_name_to_mo_cleaned = ({}, {})
    if mo_map_path is not None:
        mo_to_name, mo_name_to_mo_cleaned = load_mo_map(mo_map_path)
    unit_annotation_texts = (
        load_unit_annotation_texts(unit_annotations_path)
        if unit_annotations_path is not None
        else set(_DEFAULT_UNIT_ANNOTATION_TEXTS)
    )

    table_source_mapping = load_table_source_map(table_source_mapping_path)
    manual_tables = load_manual_table_numbers(table_source_mapping_path)
    mo_source_files = {src for src in table_source_mapping.values() if 'mo' in src.lower()}
    word_to_indicator, indicator_to_excel, indicator_to_file, file_word_to_indicator, _, file_word_to_indicators = load_column_mapping_v2(column_mapping_path)

    # Справочники "название категории -> код" для категорийных файлов (ОПФ,
    # форма собственности) — строятся прямо из соответствующего Excel-файла,
    # см. category_mapping.py.
    category_name_maps = {}
    category_prefix_by_file = {}
    if excel_dir is not None:
        excel_dir = Path(excel_dir)
        for src_filename in set(table_source_mapping.values()):
            cat_path = excel_dir / src_filename
            if not cat_path.exists():
                continue
            prefix = detect_category_prefix(cat_path)
            if prefix:
                category_prefix_by_file[src_filename] = prefix
                category_name_maps[src_filename] = extract_category_codes_from_excel(cat_path)

    doc = Document(input_doc_path)
    total_tags = 0
    validation_log = []
    near_miss_report = []
    squish_match_report = []
    duplicate_year_report = []
    name_check_report = {}
    typo_match_report = []
    unit_annotation_report = []
    group_prefix_match_report = []
    # Для отчёта "показатель есть в column_mapping_v2.csv, но нигде не
    # нашёл совпадения в Word" (см. unused_indicator_report_path ниже):
    # какие названия показателей реально были включены в пул кандидатов
    # хотя бы одной таблицы каждого файла-источника, и все тексты колонок,
    # которые этот файл когда-либо предъявлял — второе нужно, чтобы потом
    # honestly посчитать, насколько показатель "почти совпал" с чем-то
    # реальным в документе (а не просто отсутствует).
    used_names_by_file = {}
    column_texts_by_file = {}
    current_source_file = None
    normalized_title_to_src = {k: v for k, v in table_source_mapping.items()}
    table_numbers = _table_numbers(doc, doc.tables)

    for t_index, table in enumerate(doc.tables):
        print(f"\n📄 Таблица {t_index + 1}")

        # 0. Определение источника данных по названию таблицы
        table_title, saw_heading_text = get_table_name(table, table_source_mapping.keys(), return_details=True)
        if table_title:
            table_title_norm = _normalize_text(table_title)
            if table_title_norm in normalized_title_to_src:
                current_source_file = normalized_title_to_src[table_title_norm]
                print(f"🔍 Источник таблицы: {current_source_file}")
            else:
                print(f"⚠️ Не найден источник для заголовка таблицы: '{table_title}'")
        elif saw_heading_text:
            # Над таблицей ЕСТЬ содержательный заголовок, но он не совпал ни
            # с одним известным названием таблицы — это не "продолжение без
            # метки" (тогда заголовка не было бы вовсе), а похоже на ДРУГУЮ,
            # незнакомую таблицу. Не наследуем источник предыдущей таблицы
            # вслепую — иначе при случайном текстовом совпадении показателя
            # эта новая таблица получит чужие данные молча.
            current_source_file = None

        continuation_number = get_continuation_table_number(table)
        continuation_source = None
        if continuation_number:
            continuation_source = get_table_source_by_number(table_source_mapping, continuation_number)
            if continuation_source:
                if current_source_file and current_source_file != continuation_source:
                    conflict = (
                        f"🔁 Источник по метке продолжения таблицы {continuation_number} ({continuation_source}) отличается от источника заголовка ({current_source_file}). Предпочитаем продолжение таблицы.")
                    print(conflict)
                    # Оба способа определения источника надёжны, поэтому их
                    # расхождение означает ошибку в одном из них — такое
                    # нельзя оставлять только в консоли, иначе таблица молча
                    # заполнится чужими данными.
                    validation_log.append(f"[TABLE {t_index + 1}] {conflict}")
                current_source_file = continuation_source
                print(f"🔁 Источник по метке продолжения таблицы {continuation_number}: {current_source_file}")

        manual_number = table_numbers[t_index]
        if manual_number in manual_tables and not table_title and not continuation_source:
            # Таблица из списка без Excel-файла (сводная, заполняется
            # вручную). Источник по содержимому не угадываем: раньше таблица
            # «узнавалась» по совпавшему названию показателя (Баланс —
            # как t01Ved14) и очищалась лишь частично, а если не
            # узнавалась — оставалась с числами прошлого периода.
            cleared = _clear_manual_table(table)
            note = (f"✋ Таблица {manual_number} «{manual_tables[manual_number]}» — в списке таблиц нет "
                    f"Excel-файла, заполняется вручную: стёрто чисел {cleared}, ячейки оставлены пустыми.")
            print(note)
            validation_log.append(f"[TABLE {t_index + 1}] {note}")
            current_source_file = None
            continue

        if not current_source_file:
            print(f"   → Пытаемся определить по содержимому таблицы...")
            detected = auto_detect_table_source(table, table_source_mapping, file_word_to_indicator)
            if detected:
                current_source_file = detected
                print(f"   ✓ Автоматически определен источник: {current_source_file}")
            else:
                print("ℹ️ Заголовок таблицы не определён, используем предыдущий источник")

        if not current_source_file:
            warning = "⚠️ Источник не определён. Пропускаем таблицу."
            print(warning)
            validation_log.append(f"[TABLE {t_index + 1}] {warning}")
            continue

        # Используем file_word_to_indicator для правильного маппирования по (файл, слово)
        all_file_entries = {
            name: indicators
            for (file, name), indicators in file_word_to_indicators.items()
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

        if current_source_file:
            column_texts_by_file.setdefault(current_source_file, []).extend(column_texts_norm)

        # Групповые префиксы, общие для нескольких показателей ЭТОГО файла
        # (см. _compute_shared_group_prefixes) — вычисляем один раз на всю
        # таблицу/файл, а не для каждого показателя отдельно, чтобы решение
        # "это групповой заголовок или нет" опиралось на структуру всего
        # набора показателей, а не на длину одного случайно взятого слова.
        all_names_norm = {name: _normalize_match_text(name) for name in all_file_entries}
        shared_group_prefixes = sorted(
            _compute_shared_group_prefixes(all_names_norm.values()),
            key=len, reverse=True
        )
        file_prefix = _file_wide_prefix(all_names_norm.values())

        source_word_to_indicator = {}
        for name, indicators in all_file_entries.items():
            name_norm = all_names_norm[name]
            if name_norm and any(name_norm in col_text for col_text in column_texts_norm):
                source_word_to_indicator[name] = indicators
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
                source_word_to_indicator[name] = indicators
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
                        source_word_to_indicator[name] = indicators
                        if typo_match_report_path:
                            typo_match_report.append({
                                'table': t_index + 1,
                                'source_file': current_source_file,
                                'cell_text': matched_cell,
                                'matched_candidate': name_squished,
                                'matched_indicator': ','.join(indicators),
                            })
                        break

            # Название из Excel без короткого слова в скобках («Темп роста
            # прибыли (убытка) в %» -> «темп роста прибыли в»). Здесь только
            # решается, может ли показатель быть в этой таблице; колонку ему
            # даст лишь дословное совпадение заголовка с этим вариантом (см.
            # _fuzzy_match_header), и каждое такое решение уходит в отчёт.
            if name not in source_word_to_indicator and any(
                variant in col_text
                for variant in _name_variants(name)
                for col_text in column_texts_norm
            ):
                source_word_to_indicator[name] = indicators

            if name in source_word_to_indicator:
                continue
            # В Excel-исходнике заголовок показателя часто склеен из общего
            # "группового" заголовка, объединённого (merge) по нескольким
            # столбцам (например, "Отчетный год" или "Доходы и расходы по
            # обычным видам деятельности за отчетный период"), и собственного
            # подзаголовка колонки ("Прочее", "Выручка") — см.
            # _infer_mapping_from_excel в config_v2.py. В Word же такой
            # групповой заголовок обычно не повторяется в каждом столбце: там
            # просто короткий подзаголовок. Если показатель целиком не
            # нашёлся, пробуем отбросить префикс, который РЕАЛЬНО общий у
            # НЕСКОЛЬКИХ показателей этого файла (shared_group_prefixes) — в
            # отличие от произвольного обрезания по словам, это не хватается
            # за случайное совпадение, а опирается на структуру самого
            # column_mapping_v2.csv.
            name_words = name_norm.split() if name_norm else []
            for prefix in shared_group_prefixes:
                if len(name_words) <= len(prefix) or tuple(name_words[:len(prefix)]) != prefix:
                    continue
                remainder = ' '.join(name_words[len(prefix):])
                # Остаток должен быть содержательной ФРАЗОЙ (минимум 2 слова),
                # а не одним общим словом вроде "единиц" или "всего" — такой
                # остаток совпадает почти где угодно (это же слово случайно
                # оказывается частью и совсем другого, никак не связанного
                # показателя в этой же таблице) и, если зарегистрировать под
                # ним алиас (см. ниже), может увести тег в чужую колонку.
                if len(name_words[len(prefix):]) < 2 or len(remainder) < 10:
                    continue
                # Остаток ищем и в укороченных вариантах (см. _name_variants):
                # в t23 бюллетеня №3 у показателя сразу и групповой префикс
                # «Предыдущий год (сопоставимый круг)», и слова «от
                # продолжающейся деятельности», которых в Word нет.
                remainder_forms = [remainder] + list(_name_variants(remainder))
                matched_col = next(
                    (col_text for col_text in column_texts_norm
                     if any(form in col_text for form in remainder_forms)),
                    None
                )
                if matched_col is None:
                    continue
                # Групповой префикс может быть общим сразу для НЕСКОЛЬКИХ
                # разных секций одного файла (например, "Поступления по
                # текущей деятельности" и "Поступления по инвестиционной
                # деятельности" — обе секции заканчиваются словами "прочие
                # поступления", и Word в обеих пишет ОДИНАКОВЫЙ короткий текст
                # колонки — "в том числе прочие поступления", без вообще
                # какого-либо признака, к какой секции она относится). Просто
                # совпадения remainder недостаточно — так подошёл бы ЛЮБОЙ
                # показатель с таким остатком из ЛЮБОЙ секции файла. Поэтому
                # дополнительно проверяем, что сам префикс (например,
                # "поступления по текущей деятельности") тоже реально
                # присутствует где-то в этой таблице — обычно в соседней
                # колонке "Х - всего" той же секции. Это подтверждает, что мы
                # действительно в "своей" секции, а не просто угадали по
                # общему хвосту фразы.
                prefix_text = ' '.join(prefix)
                # Групповой префикс, унаследованный от объединяющего
                # заголовка вида "в том числе:" (см. _infer_mapping_from_excel
                # в config_v2.py — там он теперь всегда входит в итоговое имя
                # показателя), включает и саму эту связку. Но в Word базовая
                # часть префикса (например, "внеоборотные активы" или
                # "поступления по текущей деятельности") обычно стоит в
                # ИТОГОВОЙ колонке группы, а связка "в том числе" — отдельно,
                # в шапке НАД листовыми колонками. Это две разные ячейки одной
                # таблицы — требовать, чтобы вся фраза целиком нашлась в одной
                # колонке, было бы слишком строго. Поэтому подтверждаем либо
                # полным префиксом, либо его базовой частью без хвостового
                # "в том числе".
                prefix_confirm_candidates = {prefix_text}
                # Часть префикса, общую для всех показателей файла (см.
                # _file_wide_prefix), в таблице не ищем: подтверждаем только
                # то, что после неё, а если после неё ничего — префикс
                # подтверждён сам собой.
                file_prefix_only = False
                if file_prefix and prefix[:len(file_prefix)] == file_prefix:
                    rest_text = ' '.join(prefix[len(file_prefix):])
                    if rest_text:
                        prefix_confirm_candidates.add(rest_text)
                    else:
                        file_prefix_only = True
                for candidate in list(prefix_confirm_candidates):
                    trimmed_prefix = re.sub(r'\s*в том числе:?\s*$', '', candidate).strip()
                    if trimmed_prefix:
                        prefix_confirm_candidates.add(trimmed_prefix)
                if not file_prefix_only and not any(
                    candidate in col_text
                    for candidate in prefix_confirm_candidates
                    for col_text in column_texts_norm
                ):
                    continue
                source_word_to_indicator[name] = indicators
                # Помимо полного (длинного) названия регистрируем показатель
                # ещё и под "коротким" остатком — именно так он выглядит в
                # Word-документе (без группового префикса), а полное название
                # в этой колонке никогда не встретится целиком. Без этого
                # показатель хоть и попадает в пул кандидатов таблицы, но
                # ниже, при постолбцовом сопоставлении (_compute_section_mapping
                # / _fuzzy_match_header), длинное название не находит точного
                # вхождения в короткий текст колонки и не проходит порог
                # нечёткого совпадения — тег так и не проставляется.
                source_word_to_indicator.setdefault(remainder, indicators)
                if group_prefix_match_report_path:
                    group_prefix_match_report.append({
                        'table': t_index + 1,
                        'source_file': current_source_file,
                        'full_name': name,
                        'matched_remainder': remainder,
                        'matched_indicator': ','.join(indicators),
                    })
                break

        if current_source_file:
            used_names_by_file.setdefault(current_source_file, set()).update(source_word_to_indicator.keys())

        if not source_word_to_indicator:
            # fallback на все показатели из Excel, если в Word ничего не найдено
            source_word_to_indicator = all_file_entries.copy()

        if not source_word_to_indicator:
            warning = f"⚠️ Нет показателей для источника {current_source_file}. Пропускаем таблицу."
            print(warning)
            validation_log.append(f"[TABLE {t_index + 1}] {warning}")
            continue

        header_rows = _find_header_rows(
            table, source_word_to_indicator,
            entity_name_maps=[name_to_okved_cleaned, mo_name_to_mo_cleaned, category_name_maps.get(current_source_file, {})]
        )
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
            table_unit_annotations = [] if unit_annotation_report_path else None
            # Всегда собираем (не только если задан report_path) — эта запись
            # ещё и уходит в validation_log ниже, независимо от того, нужен
            # ли отдельный xlsx-отчёт.
            table_duplicate_years = []
            table_bracket_matches = []
            table_taken_columns = []
            mapping = _compute_section_mapping(table, header_idx, source_word_to_indicator,
                                                header_rows=header_rows_set, run_rows=run,
                                                near_miss_sink=table_near_misses,
                                                squish_match_sink=table_squish_matches,
                                                unit_annotation_sink=table_unit_annotations,
                                                unit_annotation_texts=unit_annotation_texts,
                                                duplicate_year_sink=table_duplicate_years,
                                                bracket_match_sink=table_bracket_matches,
                                                taken_column_sink=table_taken_columns)
            if (table_bracket_matches or table_taken_columns) and not clear_only:
                _collect_name_check_entries(
                    name_check_report, validation_log, table, t_index, header_idx, header_rows_set,
                    _describe_word_table(table_title, continuation_number, table_source_mapping, t_index),
                    current_source_file, source_word_to_indicator, indicator_to_excel,
                    table_bracket_matches, table_taken_columns)
            if table_duplicate_years:
                for col_idx, indicator, year_code in table_duplicate_years:
                    warning = (f"⚠️ Таблица {t_index + 1} ({current_source_file}), столбец {col_idx + 1}: "
                               f"год «{year_code}» в шапке повторяется для показателя «{indicator}» — "
                               f"похоже на опечатку в годах (например, «2023» напечатан вместо «2024»). "
                               f"Этот столбец не заполнится, останется исходное значение из документа.")
                    print(f"   {warning}")
                    validation_log.append(f"[TABLE {t_index + 1}] {warning}")
                    duplicate_year_report.append({
                        'table': t_index + 1,
                        'source_file': current_source_file,
                        'column': col_idx + 1,
                        'header_row': header_idx + 1,
                        'indicator': indicator,
                        'year_code': year_code,
                    })
            if table_unit_annotations:
                unit_annotation_report.extend(
                    {
                        'table': t_index + 1,
                        'source_file': current_source_file,
                        'column': col_idx,
                        'row': row_idx + 1,
                        'dropped_text': dropped_text,
                    }
                    for col_idx, row_idx, dropped_text in table_unit_annotations
                )
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
                _normalize_match_text(name): indicators
                for name, indicators in source_word_to_indicator.items()
            }

            for i, cell in enumerate(header_row.cells):
                header_text = _normalize_match_text(get_cleaned_cell_text(cell))
                if not header_text:
                    continue
                for name_norm, indicators in normalized_name_map.items():
                    if _text_matches_name(header_text, name_norm):
                        base_indicator_map[i] = indicators[0]
                        print(f"   🔍 Fallback row0: столбец {i}, текст '{header_text[:50]}' → {indicators[0]}")
                        break

            if not base_indicator_map and header_row_idx is not None:
                for depth in (2, 3):
                    if base_indicator_map:
                        break
                    for i, cell in enumerate(header_row.cells):
                        composed = _build_composed_header_for_column(table, header_row_idx, i, depth=depth)
                        if not composed:
                            continue
                        for name_norm, indicators in normalized_name_map.items():
                            if _text_matches_name(composed, name_norm):
                                base_indicator_map[i] = indicators[0]
                                print(
                                    f"   🔍 Fallback row0+... (depth={depth}): столбец {i}, "
                                    f"текст '{composed[:70]}' → {indicators[0]}"
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
                expected = sorted({ind for indicators in source_word_to_indicator.values() for ind in indicators})
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
            warning = "⚠️ Заголовки не найдены. Пропускаем таблицу."
            print(warning)
            validation_log.append(f"[TABLE {t_index + 1}] source={current_source_file} {warning}")
            continue

        matched_indicators = set()
        for _, _, section_map in mapped_sections:
            for indicator, _ in section_map.values():
                matched_indicators.add(indicator)
        expected_indicators = {ind for indicators in source_word_to_indicator.values() for ind in indicators}
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

            # Шапка с годами у страницы-продолжения, склеенной в эту же
            # таблицу, — не строка данных: ни тегов, ни очистки
            # (см. _is_year_header_row).
            if _is_year_header_row(row):
                continue

            category_prefix = category_prefix_by_file.get(current_source_file)
            if category_prefix:
                # Категорийная таблица (ОПФ / форма собственности) — строки
                # размечены названием категории, а не кодом ОКВЭД/МО.
                category_code = find_okved_code(first_cell_text, category_name_maps.get(current_source_file, {}))
                if not category_code:
                    # Та же опасность, что и для строк ОКВЭД/МО ниже: если
                    # название категории не нашлось в справочнике, строку
                    # нельзя просто пропускать — иначе числа прошлого периода
                    # останутся в ячейках, выглядя как актуальные данные.
                    warning = (f"⚠️ Строка {row_idx + 1} ({current_source_file}): "
                               f"название категории '{first_cell_text}' не найдено — "
                               f"строка НЕ будет заполнена, числовые ячейки очищены")
                    print(f"   {warning}")
                    validation_log.append(f"[TABLE {t_index + 1}] {warning}")
                    if not clear_only and not _is_row_label_only(row):
                        category_map = category_name_maps.get(current_source_file, {})
                        _collect_unknown_row_entry(
                            name_check_report, table_title, continuation_number, table_source_mapping, t_index,
                            first_cell_text, current_source_file, 'категорий (из самого Excel-файла)',
                            {name: (code, name) for name, code in category_map.items()})
                    for cell in row.cells[1:]:
                        if cell._tc is row.cells[0]._tc:
                            continue
                        cell_text = get_cleaned_cell_text(cell)
                        if looks_like_data_value(cell_text):
                            for p in cell.paragraphs:
                                set_paragraph_text_keep_format(p, "-")
                    continue
                prefix = category_prefix
                code_value = canonical_category(category_code)
            else:
                okved_code = find_okved_code(first_cell_text, name_to_okved_cleaned)
                mo_code = None
                if not okved_code and current_source_file and current_source_file.lower().endswith(
                        '.xlsx') and current_source_file.lower().find('mo') != -1:
                    mo_code = find_mo_code(first_cell_text, mo_name_to_mo_cleaned)

                if not okved_code and not mo_code:
                    # ВАЖНО: раньше строка в этом случае просто пропускалась
                    # (continue) — а значит, если название строки ("Алеутский"
                    # и т.п.) не находится в справочнике МО/ОКВЭД (например,
                    # потому что для другого региона подложили чужой mo.csv),
                    # никакая ячейка этой строки не трогалась ВООБЩЕ — ни
                    # тегом, ни очисткой. Числа из предыдущего периода
                    # оставались в готовом бюллетене как будто это актуальные
                    # данные, и это НИГДЕ не отражалось — ни прочерком, ни
                    # записью в unfilled_tags.xlsx (тега там просто никогда не
                    # было). Теперь: громко логируем и на всякий случай чистим
                    # числовые ячейки строки, чтобы устаревшие данные не могли
                    # тихо пережить прогон под видом свежих.
                    warning = (f"⚠️ Строка {row_idx + 1} ({current_source_file}): "
                               f"название '{first_cell_text}' не найдено ни в ОКВЭД, ни в МО "
                               f"справочнике — строка НЕ будет заполнена, числовые ячейки очищены")
                    print(f"   {warning}")
                    validation_log.append(f"[TABLE {t_index + 1}] {warning}")
                    if not clear_only and not _is_row_label_only(row):
                        is_mo_table = bool(current_source_file) and 'mo' in current_source_file.lower()
                        directory = {name: (code, okved_to_name.get(code, name))
                                     for name, code in name_to_okved_cleaned.items()}
                        directory_label = 'ОКВЭД (input/mappings/okved_mapping.csv)'
                        if is_mo_table:
                            directory.update({name: (code, mo_to_name.get(code, name))
                                              for name, code in mo_name_to_mo_cleaned.items()})
                            directory_label = 'ОКВЭД и МО (okved_mapping.csv, mo.csv)'
                        _collect_unknown_row_entry(
                            name_check_report, table_title, continuation_number, table_source_mapping, t_index,
                            first_cell_text, current_source_file, directory_label, directory)
                    for cell in row.cells[1:]:
                        if cell._tc is row.cells[0]._tc:
                            continue
                        cell_text = get_cleaned_cell_text(cell)
                        if looks_like_data_value(cell_text):
                            for p in cell.paragraphs:
                                # "-" — тот же символ "нет данных", что и в
                                # обычном заполнении по тегам (см.
                                # fill_word_template_by_tags_v2), чтобы в
                                # готовом документе не было видимой разницы
                                # между "не нашли тег" и "не распознали строку".
                                set_paragraph_text_keep_format(p, "-")
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
                    if col_idx == 0:
                        # Колонка 0 — это ВСЕГДА название строки (код ОКВЭД/МО/
                        # категории), даже если по какой-то причине алгоритм
                        # сопоставления колонок ошибочно решил, что это колонка
                        # с данными. Ни при каких обстоятельствах не затираем её
                        # тегом — иначе теряется сама подпись строки, и по этой
                        # ячейке уже не восстановить, какая это была строка
                        # (первое, что читает find_okved_code/find_mo_code/
                        # find_category_code чуть выше — именно эта ячейка).
                        print(f"   ⚠️ Строка {row_idx + 1}: колонка 0 — это колонка названия, тег пропущен (индикатор {indicator_spec[0]})")
                        continue
                    if row.cells[col_idx]._tc is name_cell_tc:
                        # Ячейка данных объединена с ячейкой названия ОКВЭД/МО (col0) —
                        # запись тега сюда стёрла бы название строки. Пропускаем.
                        print(f"   ⚠️ Строка {row_idx + 1}: колонка {col_idx} объединена с колонкой названия, тег пропущен")
                        continue
                    indicator, year = indicator_spec
                    if year:
                        tag = f"{{{{{prefix}_{code_tag_part}_{indicator}_{year}}}}}"
                    else:
                        tag = f"{{{{{prefix}_{code_tag_part}_{indicator}}}}}"
                    # Очищаем ячейку перед вставкой тега — сохраняя
                    # форматирование (шрифт/размер) исходного текста, а не
                    # пересоздавая run с форматированием по умолчанию (см.
                    # set_paragraph_text_keep_format).
                    for p in row.cells[col_idx].paragraphs:
                        set_paragraph_text_keep_format(p, "")
                    if clear_only:
                        # Только шаг очистки (см. clear_only) — ячейка с
                        # исходным числом уже опустошена выше, тег на её
                        # место НЕ ставим. Нужно, чтобы получить
                        # промежуточный "очищенный" вариант документа
                        # (только шапки с годами, без старых чисел) — на
                        # него потом отдельным проходом ставятся теги.
                        total_tags += 1
                        print(f"   🧹 Строка {row_idx + 1}: очищена колонка {col_idx} (был бы тег {tag})")
                        continue
                    if row.cells[col_idx].paragraphs:
                        set_paragraph_text_keep_format(row.cells[col_idx].paragraphs[0], tag)
                    total_tags += 1
                    print(f"   🏷️ Строка {row_idx + 1}: вставлен тег {tag}")

            if clear_only:
                # Цикл выше очищает только те колонки, для которых нашёлся
                # показатель (mapping). Но в строке данных могут быть и
                # ДРУГИЕ колонки с числами — например, расчётные/производные
                # (проценты, доли), для которых просто нет отдельного
                # показателя в column_mapping_v2.csv, или колонки, для
                # которых сопоставление не удалось найти. Раз это шаг
                # ПОЛНОЙ очистки от данных прошлого периода — чистим и их
                # тоже, а не только то, что попало в найденный маппинг.
                mapped_cols = set(mapping.keys())
                for col_idx, cell in enumerate(row.cells):
                    if col_idx == 0 or col_idx in mapped_cols:
                        continue
                    if cell._tc is name_cell_tc:
                        continue
                    cell_text = get_cleaned_cell_text(cell)
                    if looks_like_data_value(cell_text):
                        for p in cell.paragraphs:
                            set_paragraph_text_keep_format(p, "")
                        print(f"   🧹 Строка {row_idx + 1}: дополнительно очищена несопоставленная колонка {col_idx} (было {cell_text[:30]!r})")

    doc.save(output_doc_path)
    if clear_only:
        print(f"\n✅ Очищенный от данных бланк сохранён: {output_doc_path}")
        print(f"🧹 Всего очищено ячеек: {total_tags}")
    else:
        print(f"\n✅ Шаблон с тегами сохранён: {output_doc_path}")
        print(f"🔢 Всего вставлено тегов: {total_tags}")
    if name_check_report_path:
        _write_name_check_report(name_check_report, name_check_report_path, column_mapping_path)
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
    if duplicate_year_report_path and duplicate_year_report:
        duplicate_year_report_path = Path(duplicate_year_report_path)
        duplicate_year_report_path.parent.mkdir(parents=True, exist_ok=True)
        columns = ['table', 'source_file', 'column', 'header_row', 'indicator', 'year_code']
        df = pd.DataFrame(duplicate_year_report, columns=columns)
        df.insert(0, 'пояснение', 'Год в шапке повторяется — вероятно, опечатка (например, "2023" вместо "2024"). Столбец не заполнится.')
        df.to_excel(duplicate_year_report_path, index=False)
        print(f"📅 Отчёт по повторяющимся годам в шапке сохранён: {duplicate_year_report_path} "
              f"({len(duplicate_year_report)} строк)")
        print(f"🧩 Отчёт по «склеенным» совпадениям сохранён: {squish_match_report_path} "
              f"({len(squish_match_report)} строк)")
    if typo_match_report_path:
        typo_match_report_path = Path(typo_match_report_path)
        typo_match_report_path.parent.mkdir(parents=True, exist_ok=True)
        columns = ['table', 'source_file', 'cell_text', 'matched_candidate', 'matched_indicator']
        pd.DataFrame(typo_match_report, columns=columns).to_excel(typo_match_report_path, index=False)
        print(f"✏️ Отчёт по опечаткам в одну букву сохранён: {typo_match_report_path} "
              f"({len(typo_match_report)} строк)")
    if group_prefix_match_report_path:
        group_prefix_match_report_path = Path(group_prefix_match_report_path)
        group_prefix_match_report_path.parent.mkdir(parents=True, exist_ok=True)
        columns = ['table', 'source_file', 'full_name', 'matched_remainder', 'matched_indicator']
        pd.DataFrame(group_prefix_match_report, columns=columns).to_excel(group_prefix_match_report_path, index=False)
        print(f"🪓 Отчёт по совпадениям после отбрасывания группового префикса сохранён: "
              f"{group_prefix_match_report_path} ({len(group_prefix_match_report)} строк)")
    if unit_annotation_report_path:
        unit_annotation_report_path = Path(unit_annotation_report_path)
        unit_annotation_report_path.parent.mkdir(parents=True, exist_ok=True)
        columns = ['table', 'source_file', 'column', 'row', 'dropped_text']
        pd.DataFrame(unit_annotation_report, columns=columns).to_excel(unit_annotation_report_path, index=False)
        print(f"📏 Отчёт по отфильтрованным разметкам единиц измерения сохранён: {unit_annotation_report_path} "
              f"({len(unit_annotation_report)} строк)")
    if unused_indicator_report_path:
        # Номер строки в column_mapping_v2.csv для каждого кода показателя —
        # чтобы в отчёте можно было сразу открыть файл и найти нужную строку,
        # а не искать код по всему CSV вручную.
        code_to_csv_line = _mapping_csv_lines(column_mapping_path)

        unused_indicator_report = []
        for (src_file, name), indicators in file_word_to_indicators.items():
            if name in used_names_by_file.get(src_file, set()):
                continue
            name_norm = _normalize_match_text(name)
            candidates = column_texts_by_file.get(src_file, [])
            best_score = 0.0
            best_text = ''
            if name_norm and candidates:
                for col_text in candidates:
                    if not col_text:
                        continue
                    score = SequenceMatcher(None, name_norm, col_text).ratio()
                    if score > best_score:
                        best_score = score
                        best_text = col_text
            for indicator in indicators:
                excel_cols = indicator_to_excel.get(indicator, ('', ''))
                unused_indicator_report.append({
                    'csv_line': code_to_csv_line.get(indicator, ''),
                    'source_file': src_file,
                    'indicator_name': name,
                    'indicator_code': indicator,
                    'excel_column_2022': excel_cols[0],
                    'excel_column_2023': excel_cols[1],
                    'best_match_score': round(best_score, 2),
                    'best_match_text': best_text[:150],
                })

        unused_indicator_report_path = Path(unused_indicator_report_path)
        unused_indicator_report_path.parent.mkdir(parents=True, exist_ok=True)
        columns = ['csv_line', 'source_file', 'indicator_name', 'indicator_code',
                   'excel_column_2022', 'excel_column_2023', 'best_match_score', 'best_match_text']
        unused_indicator_report.sort(key=lambda r: r['best_match_score'], reverse=True)
        pd.DataFrame(unused_indicator_report, columns=columns).to_excel(unused_indicator_report_path, index=False)
        print(f"🕳️ Отчёт по неиспользованным показателям сохранён: {unused_indicator_report_path} "
              f"({len(unused_indicator_report)} строк)")


# ==========================================================
# === ОТЧЁТ «ПРОВЕРЬТЕ НАЗВАНИЯ WORD И EXCEL» ===============
# ==========================================================
_NAME_CHECK_COLUMNS = [
    'Что случилось', 'Таблица в Word', 'Страница начинается со строки', 'Строка в Word', 'Колонка в Word',
    'Excel-источник', 'Показатель в Excel', 'Колонка Excel', 'Строка в column_mapping_v2.csv',
    'Сколько страниц', 'Что сделала программа', 'Что сделать',
]


def _mapping_csv_lines(column_mapping_path, with_names: bool = False) -> dict:
    """Номер строки column_mapping_v2.csv для каждого кода показателя
    (with_names=True: (номер строки, название показателя как в файле))."""
    code_to_csv_line = {}
    try:
        rows = _read_csv_rows_robustly(column_mapping_path, delimiter=';', quotechar='"')
    except FileNotFoundError:
        return code_to_csv_line
    for line_no, row in enumerate(rows, start=1):
        if line_no == 1 or len(row) < 3:
            continue
        code = str(row[2]).strip()
        if code:
            code_to_csv_line[code] = (line_no, str(row[1]).strip()) if with_names else line_no
    return code_to_csv_line


def _excel_column_label(number) -> str:
    """«12» -> «12 (столбец L)»: номер из маппинга и буква, как в самом Excel."""
    try:
        n = int(float(number))
    except (TypeError, ValueError):
        return str(number)
    letters = ''
    k = n
    while k > 0:
        k, rem = divmod(k - 1, 26)
        letters = chr(ord('A') + rem) + letters
    return f"{n} (столбец {letters})" if letters else str(n)


def _describe_word_table(table_title, continuation_number, table_source_mapping, t_index) -> str:
    """Таблица так, как её ищет человек: по номеру и названию из бюллетеня,
    а не по порядковому номеру таблицы внутри файла Word."""
    if continuation_number:
        title = next(
            (str(key) for key in table_source_mapping
             if (m := re.match(r'\s*(\d+)\s*[.)]', str(key))) and int(m.group(1)) == continuation_number),
            '',
        )
        title = re.sub(r'^\s*\d+\s*[.)]\s*', '', title).strip()
        suffix = f" «{title[:1].upper() + title[1:]}»" if title else ''
        return f"Таблица {continuation_number}{suffix}, страница «Продолжение таблицы {continuation_number}»"
    if table_title:
        return f"Таблица «{str(table_title).strip()}»"
    return f"Таблица №{t_index + 1} по порядку в файле Word (заголовок не распознан)"


def _one_line(text) -> str:
    return re.sub(r'\s+', ' ', str(text or '')).strip()


def _raw_column_header(table, header_idx, header_rows_set, col_idx) -> str:
    """Название колонки так, как оно напечатано в Word (все строки шапки)."""
    parts = []
    for row_idx in range(max(0, header_idx - 6), header_idx + 1):
        if row_idx != header_idx and row_idx not in (header_rows_set or ()):
            continue
        cells = table.rows[row_idx].cells
        if col_idx >= len(cells):
            continue
        text = _one_line(get_cleaned_cell_text(cells[col_idx]))
        if not text or _extract_year(text) == text or _strip_continuation_label(text.lower()) == '':
            continue
        if text not in parts:
            parts.append(text)
    return ' / '.join(parts)


def _first_data_row_label(table, header_idx) -> str:
    for row in table.rows[header_idx + 1:]:
        text = _one_line(get_cleaned_cell_text(row.cells[0]))
        if text:
            return text
    return ''


def _collect_name_check_entries(report, validation_log, table, t_index, header_idx, header_rows_set,
                                table_label, source_file, source_word_to_indicator, indicator_to_excel,
                                bracket_matches, taken_columns):
    """Складывает в report (dict, чтобы одинаковые колонки на разных страницах
    стали одной строкой) записи отчёта name_check_report_path."""
    raw_by_norm = {_normalize_match_text(name): name for name in source_word_to_indicator}
    raw_by_indicator = {
        indicator: name
        for name, indicators in source_word_to_indicator.items()
        for indicator in indicators
    }
    first_row = _first_data_row_label(table, header_idx)

    def excel_column(indicator):
        cols = [str(c) for c in indicator_to_excel.get(indicator, ('', '')) if str(c).strip()]
        return ' и '.join(_excel_column_label(c) for c in dict.fromkeys(cols))

    def add(kind, col_idx, indicator, excel_name, did, todo):
        word_header = _raw_column_header(table, header_idx, header_rows_set, col_idx)
        key = (kind, table_label, word_header, source_file, indicator)
        if key in report:
            report[key]['Сколько страниц'] += 1
            return
        report[key] = {
            'Что случилось': kind,
            'Таблица в Word': table_label,
            'Страница начинается со строки': first_row,
            'Колонка в Word': word_header,
            'Excel-источник': source_file,
            'Показатель в Excel': excel_name,
            'Колонка Excel': excel_column(indicator),
            'indicator': indicator,
            'Сколько страниц': 1,
            'Что сделала программа': did.format(word=word_header),
            'Что сделать': todo.format(word=word_header),
        }
        validation_log.append(f"[TABLE {t_index + 1}] ⚠️ {kind}: {table_label}, колонка «{word_header}» "
                              f"— {source_file}, «{excel_name}» ({indicator})")

    for col_idx, header_text, matched_name, indicator in bracket_matches:
        excel_name = raw_by_norm.get(matched_name, matched_name)
        dropped = _name_variants(excel_name).get(header_text, '')
        in_brackets = dropped.startswith('(')
        add(
            'Названия отличаются словом в скобках' if in_brackets else f'В Excel лишние слова «{dropped}»',
            col_idx, indicator, excel_name,
            f"В Word колонка называется «{{word}}», в Excel — «{excel_name}». Без "
            f"{'слова' if in_brackets else 'слов'} «{dropped}» "
            f"названия совпадают, поэтому программа сочла их одним показателем и поставила "
            f"в колонку числа из {source_file}, колонка {excel_column(indicator)}.",
            f"Если числа верные — ничего делать не нужно. Чтобы строка больше не появлялась, "
            f"в input/mappings/column_mapping_v2.csv (строка указана слева) в поле «Название показателя» "
            f"замените «{excel_name}» на «{{word}}».\n"
            f"Если числа неверные (в колонке Word другой показатель) — сделайте так, чтобы нужный показатель "
            f"назывался в Word и в Excel одинаково. Любой из способов: "
            f"1) в Word (input/Бюллетень.docx) переименуйте колонку так, как нужный показатель назван в Excel; "
            f"2) в Excel-источнике {source_file} переименуйте нужный показатель так, как колонка названа в Word; "
            f"3) в input/mappings/column_mapping_v2.csv найдите строку нужного показателя этого файла "
            f"и в поле «Название показателя» впишите «{{word}}».",
        )

    for col_idx, _header_text, indicator, winner_col in taken_columns:
        excel_name = raw_by_indicator.get(indicator, indicator)
        winner_header = _raw_column_header(table, header_idx, header_rows_set, winner_col)
        add(
            'Колонка оставлена пустой', col_idx, indicator, excel_name,
            f"Колонка Word «{{word}}» похожа на показатель Excel «{excel_name}», но этот показатель "
            f"дословно совпал с другой колонкой — «{winner_header}». Чтобы не поставить одни и те же "
            f"числа в две колонки, колонка «{{word}}» оставлена пустой.",
            f"Найдите в Excel-источнике {source_file} показатель, который соответствует колонке «{{word}}», "
            f"и сделайте так, чтобы он назывался в Word и в Excel одинаково. Любой из способов: "
            f"1) в Word (input/Бюллетень.docx) переименуйте колонку так, как показатель назван в Excel; "
            f"2) в Excel-источнике переименуйте показатель так, как колонка названа в Word; "
            f"3) в input/mappings/column_mapping_v2.csv найдите строку этого показателя "
            f"и в поле «Название показателя» впишите «{{word}}».",
        )


# Название строки Word, отличное от справочника не больше, чем на столько,
# считаем вероятной опечаткой и подсказываем правильное название.
_UNKNOWN_ROW_SIMILAR = 0.80


def _table_number(table_title, continuation_number):
    if continuation_number:
        return continuation_number
    match = re.match(r'\s*(\d+)\s*[.)]', str(table_title or ''))
    return int(match.group(1)) if match else None


def _collect_unknown_row_entry(report, table_title, continuation_number, table_source_mapping, t_index,
                               row_name, source_file, directory_label, directory):
    """
    Строка Word с неузнанным названием («торговля оптовая и розничная
    транспортными средствами…» вместо «…автотранспортными…» в бюллетене
    №3) — в отчёт для специалиста, с подсказкой ближайшего названия из
    справочника. Одинаковые названия из разных таблиц — одна строка отчёта.

    directory: нормализованное название -> (код, название как в справочнике).
    """
    row_name = _one_line(row_name)
    number = _table_number(table_title, continuation_number)
    # Не названия строк: пусто, номер, метка «Продолжение таблицы N». И
    # таблицы без номера в маппинге (сводные таблицы в начале бюллетеня) —
    # программа их не заполняет, совет «добавьте в справочник» там неверен.
    if (not row_name or number is None
            or re.fullmatch(r'[\d\s.,()-]+', row_name)
            or not _strip_continuation_label(_normalize_text(row_name))):
        return
    key = ('unknown_row', row_name.lower(), directory_label)
    entry = report.get(key)
    if entry is None:
        text_norm = _normalize_text(row_name)
        best_score, best = 0.0, None
        for name_norm, (code, display) in directory.items():
            score = SequenceMatcher(None, text_norm, name_norm).ratio()
            if score > best_score:
                best_score, best = score, (code, display)
        did = (f"Название строки «{row_name}» не найдено в справочнике {directory_label}. "
               f"Строка не заполнена — в ней прочерки.")
        if best and best_score >= _UNKNOWN_ROW_SIMILAR:
            code, display = best
            todo = (f"Похоже на «{display}» (код {code}) — скорее всего, опечатка в Word. "
                    f"Исправьте название строки в Word (input/Бюллетень.docx) на «{display}» "
                    f"и запустите программу снова.\n"
                    f"Если это действительно другая строка — добавьте её в справочник: "
                    f"строка «код;название» так, как название написано в Word.")
        else:
            todo = (f"Похожей строки в справочнике нет. Если эта строка есть в Excel-источнике "
                    f"({source_file}) — добавьте её в справочник: строка «код;название», код как в "
                    f"первой колонке Excel, название как в Word. Если в Excel её нет — прочерк "
                    f"правильный, ничего делать не нужно.")
        entry = report[key] = {
            'Что случилось': 'Строка не узнана',
            'numbers': set(),
            'Страница начинается со строки': '',
            'Строка в Word': row_name,
            'Excel-источник': source_file,
            'Сколько страниц': 0,
            'Что сделала программа': did,
            'Что сделать': todo,
        }
    entry['Сколько страниц'] += 1
    entry['numbers'].add(number)
    numbers = sorted(entry['numbers'])
    entry['Таблица в Word'] = ('Таблица ' if len(numbers) == 1 else 'Таблицы ') + ', '.join(map(str, numbers))


def _write_name_check_report(report, report_path, column_mapping_path):
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    code_to_csv_line = _mapping_csv_lines(column_mapping_path, with_names=True)
    rows = []
    for entry in report.values():
        row = dict(entry)
        row.pop('numbers', None)
        line_no, csv_name = code_to_csv_line.get(row.pop('indicator', None), ('', ''))
        row['Строка в column_mapping_v2.csv'] = line_no
        # Загрузчик маппинга хранит названия строчными — в отчёте показываем
        # название так, как оно записано в column_mapping_v2.csv.
        loaded_name = row.get('Показатель в Excel', '')
        if csv_name and csv_name.lower() == str(loaded_name).lower():
            row['Показатель в Excel'] = csv_name
            for column in ('Что сделала программа', 'Что сделать'):
                row[column] = row[column].replace(f"«{loaded_name}»", f"«{csv_name}»")
        rows.append(row)
    pd.DataFrame(rows, columns=_NAME_CHECK_COLUMNS).to_excel(report_path, index=False)
    print(f"📝 Отчёт «проверьте названия Word и Excel» сохранён: {report_path} ({len(rows)} строк)")
    for row in rows:
        where = (f"строка «{row['Строка в Word']}»" if row.get('Строка в Word')
                 else f"колонка «{row.get('Колонка в Word', '')}» ← «{row.get('Показатель в Excel', '')}»")
        print(f"   • {row['Что случилось']}: {row['Таблица в Word']}, {where} ({row['Excel-источник']})")


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