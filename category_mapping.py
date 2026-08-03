# category_mapping.py
"""
Сопоставление строк Word-документа с "категорийными" файлами Excel — теми,
где строки размечены не кодами ОКВЭД или МО, а названием категории другой
классификации.

Два таких файла в этом отчёте:
  - T24_000000_t25OpfVed14.xlsx — группировка по ОРГАНИЗАЦИОННО-ПРАВОВЫМ
    ФОРМАМ (например, "Унитарные предприятия", "Фонды").
  - T24_000000_t25FsVed14.xlsx — группировка по ФОРМАМ СОБСТВЕННОСТИ
    (например, "Государственная собственность", "Частная собственность").

Внутри Excel такие таблицы устроены иначе, чем обычные ОКВЭД/МО-таблицы:
вместо строки данных на каждую категорию там СНАЧАЛА идёт строка-заголовок
вида "20600 - Ассоциации (союзы)" (без значений), а сами данные — в СЛЕДУЮЩЕЙ
строке, с кодом "101.АГ" ("Всего по обследуемым видам экономической
деятельности"). В Word-документе при этом строка называется просто
"ассоциации (союзы)" — без числового кода вообще — поэтому сопоставлять её
нужно по названию категории, а не по коду.
"""
import re
from pathlib import Path
from typing import Dict

import openpyxl

from config import _normalize_text

CATEGORY_HEADER_RE = re.compile(r'^\s*(\d+)\s*-\s*(.+?)\s*$')

# Excel-файл -> префикс тега для этой категорийной классификации
# (аналог "OKVED"/"MO", см. generate_word_template и fill_word_template_by_tags_v2).
CATEGORY_FILE_PREFIXES: Dict[str, str] = {
    'T24_000000_t25OpfVed14.xlsx': 'OPF',
    'T24_000000_t25FsVed14.xlsx': 'FS',
}


def canonical_category(code: str) -> str:
    """Коды категорий — просто числа (иногда с ведущими нулями); никакой ОКВЭД-специфичной нормализации не нужно."""
    return str(code).strip()


def extract_category_codes_from_excel(excel_path: Path, name_col_idx: int = 1) -> Dict[str, str]:
    """
    Извлекает {нормализованное_название_категории: код} из строк-заголовков
    вида "20600 - Ассоциации (союзы)" в столбце "Наименование" Excel-файла.

    Не различает "групповые" (например, "10000 - ОРГАНИЗАЦИОННО-ПРАВОВЫЕ
    ФОРМЫ ЮРИДИЧЕСКИХ ЛИЦ, ЯВЛЯЮЩИХСЯ...") и "листовые" (например,
    "70400 - Фонды") заголовки — в Word-документе оба вида встречаются как
    отдельные строки со своими собственными данными.
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    name_to_code: Dict[str, str] = {}
    for row in ws.iter_rows(values_only=True):
        if name_col_idx >= len(row):
            continue
        value = row[name_col_idx]
        if not value:
            continue
        match = CATEGORY_HEADER_RE.match(str(value).strip())
        if not match:
            continue
        code, name = match.group(1), match.group(2)
        name_norm = _normalize_text(name)
        if name_norm:
            name_to_code[name_norm] = code
    return name_to_code
