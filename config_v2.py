# config_v2.py - Обновленная конфигурация с поддержкой семантического поиска
"""
Расширенная версия config.py с поддержкой:
1. Ключевых слов для поиска колонок (вместо жестких индексов)
2. Нечеткого совпадения при поиске показателей
3. Динамического поиска года в шапке таблицы
"""

import csv
import re
import pandas as pd
from pathlib import Path
from typing import Dict, Tuple, Optional, Set

from config import _read_csv_rows_robustly

# === Константы ===
EMPTY_CELL_MARKER = "—"
OKVED_CODE_COLUMN_INDEX = 0
OKVED_NAME_COLUMN_INDEX = 1

# === Конфигурируемые метрические суффиксы (загружаются из CSV) ===
_METRIC_SUFFIXES: Set[str] = set()

def _load_metric_suffixes(config_path: Optional[Path] = None) -> Set[str]:
    """Загружает конфигурацию метрических суффиксов из CSV.
    
    Файл должен содержать два столбца:
    - Метрический суффикс
    - Тип метрики
    
    Если файл не найден, возвращает пустое множество (система будет работать без метрик).
    """
    global _METRIC_SUFFIXES
    
    if config_path is None:
        config_path = Path(__file__).parent / "input" / "mappings" / "metrics_config.csv"
    
    config_path = Path(config_path)
    _METRIC_SUFFIXES = set()
    
    if not config_path.exists():
        print(f"⚠️ Файл конфигурации метрик не найден: {config_path}")
        print("   Система будет работать без метрических суффиксов.")
        return _METRIC_SUFFIXES
    
    try:
        with open(config_path, encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f, delimiter=';')
            next(reader, None)  # Пропускаем заголовок
            for row in reader:
                if row and row[0].strip():
                    metric = _normalize_text(row[0].strip())
                    _METRIC_SUFFIXES.add(metric)
        
        print(f"✅ Загружены {len(_METRIC_SUFFIXES)} метрических суффиксов из {config_path.name}")
    except Exception as e:
        print(f"⚠️ Ошибка при загрузке метрик: {e}")
    
    return _METRIC_SUFFIXES


# === Утилиты ===
def _normalize_text(text: str) -> str:
    """Нормализует текст для сравнения"""
    return re.sub(r'\s+', ' ', text).strip().lower()


def _normalize_mapping_label(text: str) -> str:
    """Нормализует подписи показателей для column_mapping_v2.

    Убирает артефакты переносов/верстки из Excel-шапок:
    - множественные пробелы;
    - пробелы около скобок;
    - разрывы слов через дефис с пробелами (``дру- гие`` -> ``другие``).
    """
    if not isinstance(text, str):
        return ""
    value = text.replace("\n", " ").replace("\r", " ")
    # Склеиваем разрывы слов вида "дру-гие" / "дру- гие" / "дру -гие", но НЕ
    # трогаем настоящий дефис-разделитель между двумя отдельными словами
    # (например, "деятельности - всего", "текущей деятельности - всего") —
    # такой дефис всегда окружён пробелами С ОБЕИХ сторон, тогда как перенос
    # слова из-за узкой колонки/верстки не имеет пробела хотя бы с одной
    # стороны (дефис прилипает к оборванному фрагменту слова).
    value = re.sub(r'([A-Za-zА-Яа-яЁё])-\s*([A-Za-zА-Яа-яЁё])', r'\1\2', value)
    value = re.sub(r'([A-Za-zА-Яа-яЁё])\s-([A-Za-zА-Яа-яЁё])', r'\1\2', value)
    # Убираем лишние пробелы вокруг скобок
    value = re.sub(r'\(\s+', '(', value)
    value = re.sub(r'\s+\)', ')', value)
    # Схлопываем пробелы
    value = re.sub(r'\s+', ' ', value).strip()
    return value

def _read_csv_robustly(filepath, header_row=0):
    """Читает CSV с несколькими попытками кодировки"""
    encodings = ['utf-8-sig', 'windows-1251', 'cp1251', 'utf-8', 'latin1']
    for encoding in encodings:
        try:
            print(f"Попытка чтения {filepath} с кодировкой {encoding}")
            df = pd.read_csv(filepath, sep=';', encoding=encoding, header=header_row)
            print(f"✅ Успешно прочитано с кодировкой {encoding}")
            if df.shape[1] == 1:
                print(f"⚠️ ВНИМАНИЕ: CSV-файл прочитан как один столбец. Возможно, проблема с разделителями.")
            # Если нет заголовка, явно задаём имена столбцов
            if header_row is None:
                df.columns = [f"col_{i}" for i in range(df.shape[1])]
            return df
        except UnicodeDecodeError:
            pass
        except Exception:
            pass
    raise ValueError("❌ Не удалось прочитать CSV-файл.")


# === Новая функция: Загрузка маппингов с ключевыми словами ===
def _parse_keywords_from_field(value: str) -> list:
    if not value or not isinstance(value, str):
        return []
    value = value.strip()
    if not value:
        return []

    keywords = re.findall(r'"([^"]+)"', value)
    if keywords:
        return [kw.strip() for kw in keywords if kw.strip()]

    # Поле БЕЗ кавычек — это одно название показателя целиком, а не список
    # через запятую. Генератор пишет несколько ключевых слов только в
    # кавычках ("2023","конец","Валюта баланса"), а без кавычек кладёт
    # ровно одно название. Разрезание такого названия по запятым давало
    # первым ключом обрывок вроде "Организации" — слово настолько общее,
    # что поиск по нему находил ЧУЖОЙ столбец (первый попавшийся с таким
    # словом в шапке) и подменял им правильный номер из маппинга. В
    # бюллетене из-за этого семь разных показателей получали одно и то же
    # значение из одного чужого столбца.
    return [value]


def load_indicator_display_names(filepath) -> Dict[str, str]:
    """Код показателя -> человекочитаемое название («как в бюллетене»).

    Отдельная функция, а не ещё один элемент кортежа, который возвращает
    load_column_mapping_v2() — у той функции уже 8 мест вызова с позиционной
    распаковкой фиксированной длины, менять сигнатуру рискованно ради
    единственного отчёта. word_to_indicator (название -> код) там же не
    подходит для обратного поиска "по коду название": одно и то же название
    показателя нередко встречается в разных Excel-файлах с разными кодами
    (например, "ValBal" и "ValBal_1" — это в column_mapping_v2.csv два
    физически разных источника с одинаковым словом-подсказкой), и словарь
    "название -> код" в таком случае молча теряет более ранние коды при
    перезаписи одним и тем же ключом. Строим отдельно и однозначно —
    по коду, а не по названию.
    """
    display_names: Dict[str, str] = {}
    try:
        rows = _read_csv_rows_robustly(filepath, delimiter=';', quotechar='"')
        if rows:
            rows = rows[1:]
        for row in rows:
            if not row or all(not str(cell).strip() for cell in row):
                continue
            word_name = _normalize_mapping_label(str(row[1])) if len(row) > 1 else ''
            indicator = str(row[2]).strip() if len(row) > 2 else ''
            if indicator and word_name:
                display_names[indicator] = word_name
    except FileNotFoundError:
        pass
    return display_names


def load_column_mapping_v2(filepath) -> Tuple[Dict[str, str], Dict[str, Tuple[str, str]],
                                               Dict[str, str], Dict, Dict[str, Dict[str, str]], Dict]:
    """
    Загружает column_mapping_v2.csv с ключевыми словами.

    Returns:
        - word_to_indicator: название показателя → код индикатора
        - indicator_to_excel: код индикатора → (колонка 2022, колонка 2023)
        - indicator_to_file: код индикатора → файл Excel
        - file_word_to_indicator: (файл, название) → код индикатора (если строк
          с одинаковым (файл, название) несколько — тут остаётся ПОСЛЕДНЯЯ)
        - indicator_keywords: код индикатора → {year_2022: [...], year_2023: [...]}
        - file_word_to_indicators: (файл, название) → СПИСОК кодов индикаторов
          в порядке появления в CSV. В отличие от file_word_to_indicator, тут
          НИЧЕГО не перезаписывается — нужно для таблиц, где два РАЗНЫХ
          показателя в Word-документе называются дословно одинаково (например,
          "в % к общей задолженности" — один раз для дебиторской, другой раз
          для кредиторской задолженности), и колонки нужно разбирать по
          порядку следования строк в CSV, а не терять все, кроме одной.
    """
    word_to_indicator = {}
    indicator_to_excel = {}
    indicator_to_file = {}
    file_word_to_indicator = {}
    file_word_to_indicators = {}
    indicator_keywords = {}

    try:
        rows = _read_csv_rows_robustly(filepath, delimiter=';', quotechar='"')

        if rows:
            rows = rows[1:]

        for row in rows:
            if not row or all(not str(cell).strip() for cell in row):
                continue

            excel_file = str(row[0]).strip() if len(row) > 0 else ''
            word_name = _normalize_mapping_label(str(row[1])) if len(row) > 1 else ''
            indicator = str(row[2]).strip() if len(row) > 2 else ''
            col_22 = str(row[3]).strip() if len(row) > 3 else ''
            col_23 = str(row[4]).strip() if len(row) > 4 else ''
            kw_2022 = _normalize_mapping_label(str(row[5])) if len(row) > 5 else ''
            kw_2023 = _normalize_mapping_label(str(row[6])) if len(row) > 6 else ''

            if col_22.lower() == 'nan':
                col_22 = ''
            if col_23.lower() == 'nan':
                col_23 = ''

            if not excel_file or not word_name or not indicator:
                continue

            word_to_indicator[word_name] = indicator
            indicator_to_excel[indicator] = (col_22, col_23)
            indicator_to_file[indicator] = excel_file
            file_word_key = (excel_file, _normalize_text(word_name))
            file_word_to_indicator[file_word_key] = indicator
            file_word_to_indicators.setdefault(file_word_key, []).append(indicator)

            keywords_2022 = _parse_keywords_from_field(kw_2022)
            keywords_2023 = _parse_keywords_from_field(kw_2023)

            indicator_keywords[indicator] = {
                '2022': keywords_2022,
                '2023': keywords_2023
            }

    except FileNotFoundError:
        print(f"⚠️ Файл {filepath} не найден, используем fallback на column_mapping.csv")
        # Fallback на старую версию
        from config import load_column_mapping
        word_to_indicator, indicator_to_excel, indicator_to_file, file_word_to_indicator = load_column_mapping(filepath)
        indicator_keywords = {k: {'2022': [], '2023': []} for k in indicator_to_file.keys()}
        file_word_to_indicators = {k: [v] for k, v in file_word_to_indicator.items()}

    return (word_to_indicator, indicator_to_excel, indicator_to_file, file_word_to_indicator,
            indicator_keywords, file_word_to_indicators)


# === Оригинальные функции (для совместимости) ===
def load_okved_map(filepath):
    print(f"Загрузка справочника ОКВЭД из: {filepath}")
    df = _read_csv_robustly(filepath, header_row=None)
    df = df.iloc[:, :2]
    df.columns = ['Код', 'Наименование']
    df = df.dropna(subset=['Код'])
    df['Код'] = df['Код'].astype(str).str.strip()
    df['Наименование'] = df['Наименование'].astype(str).str.strip()
    df['Наименование_норм'] = df['Наименование'].apply(_normalize_text)
    okved_to_name = dict(zip(df['Код'], df['Наименование']))
    name_to_okved_cleaned = dict(zip(df['Наименование_норм'], df['Код']))
    print(f"📘 Загружено {len(df)} записей ОКВЭД.")
    return okved_to_name, name_to_okved_cleaned


def load_table_source_map(filepath):
    print(f"Загрузка сопоставления таблиц из: {filepath}")
    df = _read_csv_robustly(filepath, header_row=None)
    first_row = [str(x).strip().lower() for x in df.iloc[0].tolist()]
    if "таблица" in first_row and ("файл" in first_row or "источник" in first_row):
        df = _read_csv_robustly(filepath, header_row=0)
    else:
        df.columns = ["Таблица", "Файл"]
    mapping_dict = {}
    file_column = "Файл" if "Файл" in df.columns else "Источник" if "Источник" in df.columns else df.columns[1]
    for _, row in df.iterrows():
        raw_name = str(row["Таблица"]).strip()
        raw_src = str(row[file_column]).strip()
        norm_key = _normalize_text(raw_name)
        mapping_dict[norm_key] = raw_src
    print(f"📘 Загружено сопоставлений: {len(mapping_dict)}")
    return mapping_dict


def validate_table_source_mapping(table_source_mapping: dict, excel_dir: Path):
    """Проверяет, что файлы из маппинга существуют в папке Excel."""
    excel_dir = Path(excel_dir)
    available_files = {p.name for p in excel_dir.glob('*.xlsx')}
    mapped_files = set(table_source_mapping.values())

    missing_files = sorted(src for src in mapped_files if src and src not in available_files)
    unused_files = sorted(name for name in available_files if name not in mapped_files)

    if missing_files:
        print("\n⚠️ ВНИМАНИЕ: найдены источники в table_source_data_mapping.csv, которых нет в input/excel:")
        for src in missing_files:
            print(f"   - {src}")
        print("   Проверьте, не изменилось ли имя файла, и обновите mapping или файл в папке input/excel.")

    if unused_files:
        print("\nℹ️ В папке input/excel найдены файлы, не указанные в маппинге:")
        for name in unused_files[:20]:
            print(f"   - {name}")
        if len(unused_files) > 20:
            print(f"   ...и еще {len(unused_files) - 20} файлов.")

    return missing_files, unused_files


def _detect_year_by_text(text: str) -> str:
    if not isinstance(text, str):
        return None
    s = text.lower()
    if 'преды' in s or '2022' in s:
        return '22'
    if 'отчет' in s or '2023' in s or 'текущ' in s:
        return '23'
    return None


def _find_excel_header_row(df: pd.DataFrame) -> int:
    candidates = []
    for idx in range(min(40, len(df))):
        row = df.iloc[idx, :10].astype(str).fillna('').str.lower().tolist()
        joined = ' '.join(row)
        if 'код' in joined and 'наименование' in joined:
            return idx
        if 'наименование' in joined:
            # Если есть только «Наименование», но нет явного «Код»,
            # это всё ещё может быть строка заголовка.
            if any(cell and 'наименование' not in cell for cell in row):
                candidates.append(idx)

    if candidates:
        return candidates[0]

    # Fallback: некоторые файлы (например, отчёт о движении денежных
    # средств) вообще не подписывают первые два столбца словами
    # «Код»/«Наименование» — там просто буквы-нумераторы «А», «Б» прямо
    # под шапкой (см. типичную вёрстку росстатовских таблиц: строка группы
    # -> строка подзаголовка -> [опционально строка листового уточнения] ->
    # строка "А", "Б", 1, 2, 3...). Эта строка-нумератор — надёжный якорь:
    # она есть во ВСЕХ проверенных файлах, в отличие от текстовых меток.
    # Идём от неё на 2 или 3 строки вверх (в зависимости от того, есть ли
    # третья, "листовая" строка шапки) — выбираем тот вариант, где
    # получившаяся "шапочная" строка действительно что-то содержит.
    letter_row_idx = _find_letter_marker_row(df)
    if letter_row_idx is not None:
        for offset in (2, 3):
            candidate = letter_row_idx - offset
            if candidate < 0:
                continue
            row = df.iloc[candidate, :10].astype(str).fillna('').tolist()
            if _looks_like_real_header_row(row):
                return candidate

    return None


# Строки-"обёртка" над самой шапкой (единица измерения, отметка о том, что
# страница — продолжение предыдущей, название региона и т.п.) — их
# встречаем ДО настоящей строки группового заголовка и не должны спутать с
# ней при поиске "снизу вверх" от строки-нумератора колонок.
_BOILERPLATE_ROW_MARKERS = (
    'тыс.руб', 'тыс. руб', 'продолжение', 'камчатский край',
    'организации без субъектов',
)


def _looks_like_real_header_row(row_values) -> bool:
    non_empty = [c.strip() for c in row_values[2:] if c.strip() and c.strip().lower() != 'nan']
    if not non_empty:
        return False
    return not all(
        any(marker in cell.lower() for marker in _BOILERPLATE_ROW_MARKERS)
        for cell in non_empty
    )


def _find_letter_marker_row(df: pd.DataFrame) -> Optional[int]:
    """Ищет строку-нумератор колонок вида "А", "Б", 1, 2, 3... — она стоит
    сразу под шапкой (группа/подзаголовок[/лист]) практически во всех
    росстатовских Excel-выгрузках, даже когда сама шапка не подписана
    словами «Код»/«Наименование». Первые два непустых значения — это
    буквы "а"/"б" (регистр не важен), дальше идут короткие числа-индексы
    столбцов.
    """
    for idx in range(min(40, len(df))):
        first = str(df.iloc[idx, 0]).strip().lower() if df.shape[1] > 0 else ''
        second = str(df.iloc[idx, 1]).strip().lower() if df.shape[1] > 1 else ''
        if first == 'а' and second == 'б':
            return idx
    return None


def _transliterate_to_latin(text: str) -> str:
    """Транслитерирует русский текст в латинские буквы."""
    cyrillic_to_latin = {
        'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e', 'ж': 'zh',
        'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o',
        'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'h', 'ц': 'ts',
        'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya'
    }
    result = []
    for char in text.lower():
        result.append(cyrillic_to_latin.get(char, char))
    return ''.join(result)


def _create_short_code_from_text(text: str) -> str:
    """Создает короткий код из названия показателя через транслитерацию.
    
    Примеры:
    - "Валюта баланса" → "ValBal"
    - "Внеоборотные активы" → "VneObAk"
    - "основные средства" → "OsnSr"
    """
    normalized = _normalize_text(text)
    words = [w for w in normalized.split() if w]  # Убираем пустые слова
    
    if not words:
        return ''
    
    code_parts = []
    for word in words:
        # Транслитерируем слово
        trans = _transliterate_to_latin(word)
        
        if not trans:
            continue
        
        # Для каждого слова берем нужное количество символов
        if len(trans) <= 2:
            part = trans
        elif len(trans) <= 4:
            part = trans[:2]  # "код" → "ko"
        else:
            # Для длинных слов: "внеоборотные" → "vne+ob" → берем первые слоги
            # Берем первые 3 символа, но стараемся взять целые слоги
            part = trans[:3]
        
        code_parts.append(part.capitalize())
    
    return ''.join(code_parts)


def _slugify_indicator_code(text: str, prefix: str = None) -> str:
    """Создает короткий код показателя через транслитерацию.
    
    Параметр prefix игнорируется для совместимости с более ранним кодом.
    """
    if not text:
        return ''
    return _create_short_code_from_text(text)


def _is_connective_header(text: str) -> bool:
    if not isinstance(text, str):
        return False
    return _normalize_text(text) in {
        'в том числе',
        'в том числе:'
    }


def _is_metric_suffix(text: str) -> bool:
    """Проверяет, является ли текст метрическим суффиксом.
    
    Использует конфигурацию из metrics_config.csv вместо жесткой привязки.
    """
    if not isinstance(text, str):
        return False
    
    # Если конфиг не загружен, загружаем его
    if not _METRIC_SUFFIXES:
        _load_metric_suffixes()
    
    normalized = _normalize_text(text)
    return normalized in _METRIC_SUFFIXES


def _simplify_metric_suffix(base_name: str, suffix: str) -> str:
    """Упрощает метрический суффикс для использования в названии показателя.
    
    Конфигурируемая система: все метрические суффиксы должны быть удалены из имени,
    так как они используются только для кодирования, а не для отображения.
    
    Проверяет конфигурацию из metrics_config.csv.
    """
    if not isinstance(suffix, str):
        return suffix
    
    normalized_suffix = _normalize_text(suffix)
    
    # Если конфиг не загружен, загружаем его
    if not _METRIC_SUFFIXES:
        _load_metric_suffixes()
    
    # Если это известный метрический суффикс, удаляем его из имени
    if normalized_suffix in _METRIC_SUFFIXES:
        return ''
    
    # Иначе, оставляем суффикс как есть
    return suffix


def _is_meaningful_footer_text(text: str) -> bool:
    """Отсеивает "мусорные" тексты третьей строки шапки, которые не
    являются настоящим третьим уровнем детализации:
    - пустые;
    - чисто цифровые (строка-нумератор колонок "А","Б",1,2,3... — она есть
      почти под любой шапкой и не несёт содержательного различия между
      столбцами одной группы).
    """
    text = (text or '').strip()
    if not text:
        return False
    if text.replace('.', '', 1).isdigit():
        return False
    return True


def _extract_keywords_for_year(base_name: str, suffix: str, year: str) -> str:
    """Извлекает ключевые слова из базового названия и суффикса для поиска в Excel.
    
    Эти ключевые слова используются для "умного" поиска нужной колонки в Excel,
    когда жесткий индекс может быть неправильным или устаревшим.
    
    Процесс:
    1. Ищем год в суффиксе через regex (202X) - явный год из Excel
    2. Если год не найден, используем условное преобразование:
       - year='22' → добавляем '2022'
       - year='23' → добавляем '2023'
    3. Добавляем маркер периода:
       - year='22' → 'предыдущий' или 'начало'
       - year='23' → 'конец' или 'отчетный'
    4. Добавляем базовое название показателя как универсальный ключ
    
    Пример:
        base_name = "Валюта баланса"
        suffix = "на конец предыдущего года"
        year = "22"
        
        Результат: '"2022","предыдущий","Валюта баланса"'
    
    Эти ключевые слова затем:
    - Сохраняются в column_mapping_v2.csv
    - Загружаются в memory при предварительной загрузке Excel
    - Используются при поиске колонок в функции _find_column_smart()
    
    Возвращает строку в формате: "слово1","слово2","слово3"
    """
    keywords = []
    suffix_lower = suffix.lower()
    
    # Динамически извлекаем год из суффикса через regex
    # Ищем любой год вида 202X (может быть 2022, 2023, 2024, 2025 и т.д.)
    year_match = re.search(r'(202\d)', suffix)
    if year_match:
        keywords.append(year_match.group(1))
    elif year:
        # Fallback: если год определен как '22' или '23', конвертируем
        if year == '22':
            keywords.append('2022')
        elif year == '23':
            keywords.append('2023')
    
    # Добавляем маркеры периода
    if year == '22':
        # Для первого года (условно 2022/2024)
        if 'начало' in suffix_lower:
            keywords.append('начало')
        elif 'предыдущ' in suffix_lower:
            keywords.append('предыдущий')
        else:
            keywords.append('начало')  # По умолчанию
    elif year == '23':
        # Для второго года (условно 2023/2025)
        if 'конец' in suffix_lower:
            keywords.append('конец')
        elif 'отчетн' in suffix_lower or 'текущ' in suffix_lower:
            keywords.append('конец')
        else:
            keywords.append('конец')  # По умолчанию
    
    # Если нет явного года, просто используем суффикс
    if not keywords and suffix:
        keywords.append(suffix)
    
    # Добавляем базовое название как универсальный ключ поиска
    keywords.append(base_name)
    
    # Возвращаем в формате CSV с кавычками, но БЕЗ двойного экранирования
    # csv.writer сам добавит кавычки вокруг поля, если нужно
    return ','.join(f'"{k}"' for k in keywords)


def _infer_mapping_from_excel(excel_path: Path) -> list:
    """Динамически извлекает маппинг показателей из Excel файла.
    
    Анализирует структуру Excel:
    1. Находит строку заголовка с 'Код' и 'Наименование'
    2. Для каждого столбца определяет:
       - Базовое название (header)
       - Суффикс/подзаголовок, в котором может быть год (202X) или период
       - Год/период на основе текста суффикса
    3. Генерирует ключевые слова для каждого показателя
    
    Результат:
        [(filename, indicator_name, code, col_22, col_23, keywords_2022, keywords_2023), ...]
    
    Где:
    - col_22, col_23 — номера колонок (1-based), использованные для каждого года
    - keywords_2022, keywords_2023 — ключевые слова для "умного" поиска в Excel
    
    Эти ключевые слова позволяют пересчитывать колонки автоматически,
    если структура Excel изменилась, но суть показателей и периодов осталась.
    """
    df = pd.read_excel(excel_path, header=None)
    header_row = _find_excel_header_row(df)
    if header_row is None:
        print(f"⚠️ Не найден заголовок с 'Код' и 'Наименование' в {excel_path.name}")
        return []

    headers = df.iloc[header_row].astype(str).fillna('').tolist()
    subheaders = df.iloc[header_row + 1].astype(str).fillna('').tolist() if header_row + 1 < len(df) else [''] * len(headers)
    footer_row = df.iloc[header_row + 2].astype(str).fillna('').tolist() if header_row + 2 < len(df) else [''] * len(headers)
    max_header_idx = max(
        [i for i, v in enumerate(headers) if str(v).strip()] +
        [i for i, v in enumerate(subheaders) if str(v).strip()] +
        [i for i, v in enumerate(footer_row) if str(v).strip()] +
        [0]
    )

    current_group_label = None
    current_indicator_name = None
    col_entries = []  # список: {col_idx, name, code_source, base, suffix, footer, year}

    first_header = _normalize_text(str(headers[0])) if headers else ''
    second_header = _normalize_text(str(headers[1])) if len(headers) > 1 else ''
    start_col = 2 if 'код' in first_header or (first_header == 'код' and second_header == 'наименование') else 1

    current_base = str(headers[start_col]).strip() if start_col < len(headers) else ''

    # === Проход 1: та же логика, что и раньше (группа + подзаголовок),
    # но результат по каждому столбцу складываем в список, а не сразу лепим
    # в финальные группы — третьей строке шапки (footer) ещё нужно дать
    # шанс подтвердить или переопределить итоговое имя показателя (см.
    # проход 2 ниже).
    for col_idx in range(start_col, max_header_idx + 1):
        raw_header = _normalize_mapping_label(str(headers[col_idx] if col_idx < len(headers) else ''))
        suffix = _normalize_mapping_label(str(subheaders[col_idx]))
        footer_text = _normalize_mapping_label(str(footer_row[col_idx] if col_idx < len(footer_row) else ''))
        year = _detect_year_by_text(suffix)
        code_source = None

        if raw_header:
            if _is_connective_header(raw_header):
                if not current_base or not suffix:
                    continue
                current_group_label = _normalize_mapping_label(f"{current_base} {raw_header.strip(':')}".strip())
                # Имя показателя всегда строим С групповым контекстом (а не
                # голым suffix) — даже если сам групповой заголовок в Word
                # не повторяется в каждой колонке. Это важно, когда суффикс
                # (например, "прочие поступления"/"прочие платежи") —
                # родовая формулировка, которая дословно повторяется в
                # НЕСКОЛЬКИХ независимых секциях одного файла ("текущая",
                # "инвестиционная", "финансовая" деятельность): без
                # группового контекста в имени эти секции схлопнутся в один
                # показатель с чужими колонками как "год 2022"/"год 2023".
                # Несовпадение с коротким текстом Word разбирает runtime
                # (см. shared_group_prefixes в logic.py) — там это уже
                # безопасно, с проверкой, что префикс подтверждён в той же
                # таблице.
                indicator_name = _normalize_mapping_label(f"{current_group_label} {suffix}".strip())
                code_source = indicator_name
            else:
                current_base = raw_header
                current_group_label = None
                if suffix and not year:
                    if _is_metric_suffix(suffix):
                        suffix_part = _simplify_metric_suffix(current_base, suffix)
                        indicator_name = _normalize_mapping_label(f"{current_base} {suffix_part}".strip()) if suffix_part else current_base
                        code_source = _normalize_mapping_label(f"{current_base} {suffix}".strip())
                    else:
                        indicator_name = _normalize_mapping_label(f"{current_base} {suffix}".strip())
                        code_source = indicator_name
                else:
                    indicator_name = current_base
                    code_source = current_base
        elif suffix:
            if not current_base:
                continue
            if current_group_label:
                # См. комментарий выше (в ветке connective-заголовка) — та
                # же причина: сохраняем групповой контекст в имени, а не
                # только голый suffix, иначе одноимённые listовые показатели
                # разных секций ("прочие поступления" у текущей и у
                # инвестиционной деятельности) схлопнутся в один.
                indicator_name = _normalize_mapping_label(f"{current_group_label} {suffix}".strip())
                code_source = indicator_name
            elif year:
                indicator_name = current_indicator_name or current_base
            else:
                indicator_name = _normalize_mapping_label(f"{current_base} {suffix}".strip())
                code_source = indicator_name
        else:
            if not current_indicator_name:
                continue
            indicator_name = current_indicator_name
            code_source = indicator_name

        current_indicator_name = indicator_name
        col_entries.append({
            'col_idx': col_idx,
            'name': indicator_name,
            'code_source': code_source or indicator_name,
            'base': current_base,
            'suffix': suffix,
            'footer': footer_text,
            'year': year,
        })

    # === Проход 2: используем третью строку шапки, если она реально
    # различает столбцы внутри одной и той же (по имени из прохода 1)
    # группы. Без этого два РАЗНЫХ листовых показателя, у которых
    # совпадают группа и подзаголовок (например, "убыточность в % :" —
    # общий подзаголовок и для "к затратам на производство продаж", и для
    # "к коммерческим и управленческим расходам"), схлопнутся в одну запись
    # с двумя столбцами как "год 2022"/"год 2023" — хотя на деле это не
    # два года одного показателя, а два разных показателя за один период.
    #
    # Признак настоящего потерянного уровня — среди столбцов ОДНОЙ группы
    # встречается больше одного РАЗНОГО, содержательного (не число, не
    # признак года) текста в footer. Если footer либо одинаковый у всех,
    # либо это законный маркер года ("на конец отчетного года" и т.п.) —
    # это НЕ баг, а обычная пара "тот же показатель, два периода", трогать
    # не нужно.
    by_name = {}
    for entry in col_entries:
        by_name.setdefault(entry['name'], []).append(entry)

    for entries in by_name.values():
        if len(entries) < 2:
            continue
        distinct_footers = {
            e['footer'] for e in entries
            if _is_meaningful_footer_text(e['footer']) and not _detect_year_by_text(e['footer'])
        }
        if len(distinct_footers) < 2:
            continue
        for e in entries:
            if _is_meaningful_footer_text(e['footer']) and not _detect_year_by_text(e['footer']):
                e['name'] = _normalize_mapping_label(f"{e['name']} {e['footer']}".strip())
                e['code_source'] = e['name']

    # === Проход 3: собираем финальные группы (имя -> код, столбцы по годам) ===
    groups = {}
    groups_order = []
    for entry in col_entries:
        group_key = _normalize_text(entry['name'])
        if group_key not in groups:
            groups_order.append(group_key)
            groups[group_key] = {
                'name': entry['name'],
                'code': _slugify_indicator_code(entry['code_source'], excel_path.stem),
                'cols': {'22': None, '23': None},
                'keywords_2022': [],
                'keywords_2023': []
            }

        year = entry['year']
        if year:
            groups[group_key]['cols'][year] = str(entry['col_idx'] + 1)
            keywords = _extract_keywords_for_year(entry['base'], entry['suffix'], year)
            if year == '22':
                groups[group_key]['keywords_2022'] = keywords
            else:
                groups[group_key]['keywords_2023'] = keywords
        else:
            if groups[group_key]['cols']['22'] is None:
                groups[group_key]['cols']['22'] = str(entry['col_idx'] + 1)
            else:
                groups[group_key]['cols']['23'] = str(entry['col_idx'] + 1)

    rows = []
    for group_key in groups_order:
        group = groups[group_key]
        col_22 = group['cols']['22'] or ''
        col_23 = group['cols']['23'] or ''
        kw_22 = group['keywords_2022'] if group['keywords_2022'] else group['name']
        kw_23 = group['keywords_2023'] if group['keywords_2023'] else group['name']
        rows.append((excel_path.name, group['name'], group['code'], col_22, col_23, kw_22, kw_23))

    return rows


def build_column_mapping_v2_from_excel(excel_dir: Path, table_source_mapping: dict, output_path: Path):
    excel_dir = Path(excel_dir)
    output_path = Path(output_path)
    rows = []
    seen_codes = set()

    unique_files = []
    seen_files = set()
    for filename in table_source_mapping.values():
        if filename not in seen_files:
            seen_files.add(filename)
            unique_files.append(filename)

    failed_files = []
    for filename in unique_files:
        excel_path = excel_dir / filename
        if not excel_path.exists():
            print(f"⚠️ Excel файл не найден: {filename}")
            failed_files.append(f"{filename} (файл не найден)")
            continue
        try:
            inferred = _infer_mapping_from_excel(excel_path)
        except Exception as exc:
            # Битый/нечитаемый Excel-файл (повреждён, не тот формат, защищён
            # паролем и т.д.) не должен обрушивать весь пайплайн — пропускаем
            # этот источник и продолжаем с остальными.
            print(f"❌ Не удалось прочитать Excel файл {filename}: {exc}")
            failed_files.append(f"{filename} (ошибка чтения: {exc})")
            continue
        if not inferred:
            failed_files.append(f"{filename} (не найдена структура заголовков)")
        for excel_file, name, code, col_22, col_23, kw22, kw23 in inferred:
            if code in seen_codes:
                suffix = 1
                base_code = code
                while f"{base_code}_{suffix}" in seen_codes:
                    suffix += 1
                code = f"{base_code}_{suffix}"
            seen_codes.add(code)
            rows.append((excel_file, name, code, col_22, col_23, kw22, kw23))

    if not rows:
        details = "\n".join(f"  - {item}" for item in failed_files) or "  (нет данных)"
        raise ValueError(
            "Не удалось сгенерировать column_mapping_v2 из Excel — ни один "
            f"источник не дал ни одной строки:\n{details}"
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow([
            'Excel файл', 'Название показателя', 'Код показателя',
            'Excel колонка 2022', 'Excel колонка 2023',
            'Ключевое слово 2022', 'Ключевое слово 2023'
        ])
        for row in rows:
            writer.writerow(row)

    print(f"✅ Сгенерирован column_mapping_v2: {output_path} ({len(rows)} строк)")


def ensure_column_mapping_v2(excel_dir: Path, table_source_mapping: dict, output_path: Path):
    """Гарантирует наличие column_mapping_v2.csv.
    Важно: если файл уже существует, он НЕ пересобирается автоматически.
    """
    # Загружаем конфиг метрик
    _load_metric_suffixes()

    output_path = Path(output_path)
    if not output_path.exists():
        print(f"⚠️ Файл {output_path} не найден. Генерируем column_mapping_v2.csv из Excel...")
        build_column_mapping_v2_from_excel(excel_dir, table_source_mapping, output_path)
        return

    existing_sources = set()
    try:
        _, _, _, file_word_to_indicator, _, _ = load_column_mapping_v2(output_path)
        existing_sources = {file for file, _ in file_word_to_indicator.keys()}
    except Exception as exc:
        print(f"⚠️ Не удалось прочитать {output_path} ({exc}). Пересобираем из Excel...")
        build_column_mapping_v2_from_excel(excel_dir, table_source_mapping, output_path)
        return

    expected_sources = set(table_source_mapping.values())
    missing_sources = sorted(expected_sources - existing_sources)
    if missing_sources:
        print("⚠️ Найдены отсутствующие Excel-источники в column_mapping_v2.csv:")
        for src in missing_sources:
            print(f"   - {src}")
        print("    Файл уже существует, автоматическую регенерацию пропускаем.")
    else:
        print(f"✅ column_mapping_v2 уже содержит все источники из {len(expected_sources)} файла(ов).")
