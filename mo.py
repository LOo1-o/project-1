import pandas as pd
from config import _normalize_text, _read_csv_robustly

MO_CODE_COLUMN_INDEX = 0
MO_NAME_COLUMN_INDEX = 1


def load_mo_map(filepath):
    """Загружает справочник муниципальных образований из CSV.

    Формат:
        excel;word
    """
    print(f"Загрузка справочника МО из: {filepath}")
    df = _read_csv_robustly(filepath, header_row=0)
    if df.shape[1] < 2:
        raise ValueError(f"Файл {filepath} должен содержать как минимум два столбца: код и наименование МО.")

    df = df.iloc[:, :2]
    df.columns = ['Код', 'Наименование']
    df = df.dropna(subset=['Код', 'Наименование'])

    df['Код'] = df['Код'].astype(str).str.strip()
    df['Наименование'] = df['Наименование'].astype(str).str.strip()
    df['Наименование_норм'] = df['Наименование'].apply(_normalize_text)

    mo_to_name = dict(zip(df['Код'], df['Наименование']))
    name_to_mo_cleaned = dict(zip(df['Наименование_норм'], df['Код']))

    print(f"📘 Загружено записей МО: {len(df)}")
    return mo_to_name, name_to_mo_cleaned


def canonical_mo(code: str) -> str:
    if not code:
        return ""
    code_str = str(code).strip()
    # Если код прочитался из Excel как число с .0, убираем лишнюю дробную часть.
    if code_str.endswith('.0') and code_str.replace('.0', '').isdigit():
        code_str = code_str[:-2]
    return code_str


def find_mo_code(cell_text, name_to_mo_cleaned):
    text_norm = _normalize_text(cell_text)
    if not text_norm:
        return None
    if text_norm in name_to_mo_cleaned:
        return name_to_mo_cleaned[text_norm]
    for name, code in name_to_mo_cleaned.items():
        if len(text_norm) >= 3 and (text_norm in name or name in text_norm):
            print(f"⚠️ Частичное совпадение МО: '{cell_text}' ~ '{name}' → {code}")
            return code
    return None
