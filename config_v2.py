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
from typing import Dict, Tuple, Optional, List

# === Константы ===
EMPTY_CELL_MARKER = "—"
OKVED_CODE_COLUMN_INDEX = 0
OKVED_NAME_COLUMN_INDEX = 1


# === Утилиты ===
def _normalize_text(text: str) -> str:
    """Нормализует текст для сравнения"""
    return re.sub(r'\s+', ' ', text).strip().lower()


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

    return [kw.strip() for kw in value.split(',') if kw.strip()]


def load_column_mapping_v2(filepath) -> Tuple[Dict[str, str], Dict[str, Tuple[str, str]], 
                                               Dict[str, str], Dict, Dict[str, Dict[str, str]]]:
    """
    Загружает column_mapping_v2.csv с ключевыми словами.
    
    Returns:
        - word_to_indicator: название показателя → код индикатора
        - indicator_to_excel: код индикатора → (колонка 2022, колонка 2023)
        - indicator_to_file: код индикатора → файл Excel
        - file_word_to_indicator: (файл, название) → код индикатора
        - indicator_keywords: код индикатора → {year_2022: [...], year_2023: [...]} 
    """
    word_to_indicator = {}
    indicator_to_excel = {}
    indicator_to_file = {}
    file_word_to_indicator = {}
    indicator_keywords = {}

    try:
        with open(filepath, encoding="utf-8-sig", newline='') as f:
            reader = csv.reader(f, delimiter=';', quotechar='"')
            rows = list(reader)

        if rows:
            rows = rows[1:]

        for row in rows:
            if not row or all(not str(cell).strip() for cell in row):
                continue

            excel_file = str(row[0]).strip() if len(row) > 0 else ''
            word_name = str(row[1]).strip() if len(row) > 1 else ''
            indicator = str(row[2]).strip() if len(row) > 2 else ''
            col_22 = str(row[3]).strip() if len(row) > 3 else ''
            col_23 = str(row[4]).strip() if len(row) > 4 else ''
            kw_2022 = str(row[5]).strip() if len(row) > 5 else ''
            kw_2023 = str(row[6]).strip() if len(row) > 6 else ''

            if col_22.lower() == 'nan':
                col_22 = ''
            if col_23.lower() == 'nan':
                col_23 = ''

            if not excel_file or not word_name or not indicator:
                continue

            word_to_indicator[word_name] = indicator
            indicator_to_excel[indicator] = (col_22, col_23)
            indicator_to_file[indicator] = excel_file
            file_word_to_indicator[(excel_file, _normalize_text(word_name))] = indicator

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
    
    return word_to_indicator, indicator_to_excel, indicator_to_file, file_word_to_indicator, indicator_keywords


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
    for idx in range(min(20, len(df))):
        row = df.iloc[idx, :10].astype(str).fillna('').str.lower().tolist()
        joined = ' '.join(row)
        if 'код' in joined and 'наименование' in joined:
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


def _extract_keywords_for_year(base_name: str, suffix: str, year: str) -> str:
    """Извлекает ключевые слова из базового названия и суффикса для поиска в Excel.
    
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
    df = pd.read_excel(excel_path, header=None)
    header_row = _find_excel_header_row(df)
    if header_row is None:
        print(f"⚠️ Не найден заголовок с 'Код' и 'Наименование' в {excel_path.name}")
        return []

    headers = df.iloc[header_row].astype(str).fillna('').tolist()
    subheaders = df.iloc[header_row + 1].astype(str).fillna('').tolist() if header_row + 1 < len(df) else [''] * len(headers)

    current_base = ''
    current_indicator_name = None
    groups = {}
    col_metadata = {}  # Сохраняем метаданные для каждого столбца (базовое имя, суффикс)
    
    for col_idx in range(2, len(headers)):
        raw_header = str(headers[col_idx]).strip()
        suffix = str(subheaders[col_idx]).strip()
        year = _detect_year_by_text(suffix)

        if raw_header:
            if _is_connective_header(raw_header):
                if not current_base or not suffix:
                    continue
                indicator_name = f"{current_base} {raw_header} {suffix}".strip()
            else:
                current_base = raw_header
                if suffix and not year:
                    indicator_name = f"{current_base} {suffix}".strip()
                else:
                    indicator_name = current_base
        elif suffix:
            if not current_base:
                continue
            if year:
                indicator_name = current_indicator_name or current_base
            else:
                indicator_name = f"{current_base} {suffix}".strip()
        else:
            if not current_indicator_name:
                continue
            indicator_name = current_indicator_name

        current_indicator_name = indicator_name
        group_key = _normalize_text(indicator_name)
        if group_key not in groups:
            groups[group_key] = {
                'name': indicator_name,
                'code': _slugify_indicator_code(indicator_name, excel_path.stem),
                'cols': {'22': None, '23': None},
                'keywords_2022': [],
                'keywords_2023': []
            }
        
        # Сохраняем метаданные столбца
        col_metadata[col_idx] = {'base': current_base, 'suffix': suffix, 'year': year}

        if year:
            groups[group_key]['cols'][year] = str(col_idx + 1)
            # Извлекаем ключевые слова для соответствующего года
            keywords = _extract_keywords_for_year(current_base, suffix, year)
            if year == '22':
                groups[group_key]['keywords_2022'] = keywords
            else:
                groups[group_key]['keywords_2023'] = keywords
        else:
            if groups[group_key]['cols']['22'] is None:
                groups[group_key]['cols']['22'] = str(col_idx + 1)
            else:
                groups[group_key]['cols']['23'] = str(col_idx + 1)

    rows = []
    for group in groups.values():
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

    for filename in sorted(set(table_source_mapping.values())):
        excel_path = excel_dir / filename
        if not excel_path.exists():
            print(f"⚠️ Excel файл не найден: {filename}")
            continue
        inferred = _infer_mapping_from_excel(excel_path)
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
        raise ValueError("Не удалось сгенерировать column_mapping_v2 из Excel.")

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
