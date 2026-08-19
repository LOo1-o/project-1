#!/usr/bin/env python3
"""Аудит потерянного третьего уровня шапки Excel-источников.

_infer_mapping_from_excel (config_v2.py) при сборке column_mapping_v2.csv
читает только ДВА уровня шапки — "групповой" заголовок (объединённый по
нескольким столбцам) и подзаголовок колонки. Третья строка шапки (лист,
footer_row) читается, но НИГДЕ не используется при построении имени
показателя.

Из-за этого, если в Excel одна и та же пара (групповой заголовок,
подзаголовок) повторяется для НЕСКОЛЬКИХ столбцов, а различаются они
только третьей строкой (например, "убыточность в %:" -> "к затратам на
производство продаж" / "к коммерческим и управленческим расходам"), это
выглядит для генератора как "тот же самый показатель, просто второй год" —
оба столбца схлопываются в ОДНУ запись CSV с полями "Excel колонка 2022"/
"Excel колонка 2023", хотя на самом деле это два РАЗНЫХ показателя без
всякой связи с годами. При генерации тегов используется только первая из
двух колонок — вторая молча остаётся без тега (см. таблицу 17 бюллетеня,
запись OrgPolUbyOtProUbyV%: до исправления).

Этот скрипт сканирует ВСЕ Excel-файлы, перечисленные в
table_source_data_mapping.csv, повторяет ту же логику группировки, что и
_infer_mapping_from_excel, но дополнительно отслеживает текст третьей
строки шапки для каждого столбца внутри группы — и сообщает, где среди
столбцов одной группы встречаются РАЗНЫЕ непустые тексты третьего уровня
(это и есть признак потерянного различия). Дополнительно сверяет находку с
текущим содержимым column_mapping_v2.csv: если для всех "потерянных"
столбцов там уже есть отдельные записи — считает вопрос закрытым,
предупреждает только если в CSV эти столбцы всё ещё объединены в одну
запись (т.е. баг ещё не исправлен вручную).

Запуск:
    python audit_hidden_header_levels.py
"""
from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd

from config import load_table_source_map
from config_v2 import (
    _detect_year_by_text,
    _find_excel_header_row,
    _is_connective_header,
    _is_metric_suffix,
    _load_metric_suffixes,
    _normalize_mapping_label,
    _normalize_text,
)

BASE_DIR = Path(__file__).parent.resolve()
EXCEL_DIR = BASE_DIR / "input" / "excel"
TABLE_MAPPING_FILE = BASE_DIR / "input" / "mappings" / "table_source_data_mapping.csv"
COLUMN_MAPPING_FILE = BASE_DIR / "input" / "mappings" / "column_mapping_v2.csv"


def _existing_column_spans(column_mapping_path: Path) -> Dict[str, Dict[int, List[str]]]:
    """Excel-файл -> {номер колонки (1-based): [коды показателей, использующих её]}.

    Нужно, чтобы понять, слиты ли в текущем column_mapping_v2.csv столбцы,
    которые скрипт считает потенциально разными показателями, или для них
    уже есть отдельные записи (баг уже исправлен вручную).
    """
    spans: Dict[str, Dict[int, List[str]]] = defaultdict(lambda: defaultdict(list))
    if not column_mapping_path.exists():
        return spans
    with open(column_mapping_path, encoding='utf-8-sig', newline='') as f:
        reader = csv.reader(f, delimiter=';')
        next(reader, None)
        for row in reader:
            if len(row) < 5:
                continue
            file_, _name, code, col22, col23 = row[0], row[1], row[2], row[3], row[4]
            for col in (col22, col23):
                col = (col or '').strip()
                if col.isdigit():
                    spans[file_][int(col)].append(code)
    return spans


def _group_columns_with_footer(excel_path: Path) -> List[dict]:
    """Повторяет группировку _infer_mapping_from_excel, но с трекингом
    текста третьей строки шапки (footer) по каждому столбцу группы.

    Возвращает список групп: {'name', 'cols': [(col_idx_1based, footer_text)]}.
    """
    df = pd.read_excel(excel_path, header=None)
    header_row = _find_excel_header_row(df)
    if header_row is None:
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
    current_base = ''
    groups_order: List[str] = []
    groups: Dict[str, List[Tuple[int, str]]] = defaultdict(list)

    first_header = _normalize_text(str(headers[0])) if headers else ''
    second_header = _normalize_text(str(headers[1])) if len(headers) > 1 else ''
    start_col = 2 if 'код' in first_header or (first_header == 'код' and second_header == 'наименование') else 1
    current_base = str(headers[start_col]).strip() if start_col < len(headers) else ''

    for col_idx in range(start_col, max_header_idx + 1):
        raw_header = _normalize_mapping_label(str(headers[col_idx] if col_idx < len(headers) else ''))
        suffix = _normalize_mapping_label(str(subheaders[col_idx]))
        footer_text = _normalize_mapping_label(str(footer_row[col_idx] if col_idx < len(footer_row) else ''))
        year = _detect_year_by_text(suffix)

        if raw_header:
            if _is_connective_header(raw_header):
                if not current_base or not suffix:
                    continue
                current_group_label = _normalize_mapping_label(f"{current_base} {raw_header.strip(':')}".strip())
                indicator_name = suffix
            else:
                current_base = raw_header
                current_group_label = None
                if suffix and not year:
                    if _is_metric_suffix(suffix):
                        indicator_name = current_base
                    else:
                        indicator_name = _normalize_mapping_label(f"{current_base} {suffix}".strip())
                else:
                    indicator_name = current_base
        elif suffix:
            if not current_base:
                continue
            if current_group_label:
                indicator_name = suffix
            elif year:
                indicator_name = current_indicator_name or current_base
            else:
                indicator_name = _normalize_mapping_label(f"{current_base} {suffix}".strip())
        else:
            if not current_indicator_name:
                continue
            indicator_name = current_indicator_name

        current_indicator_name = indicator_name
        group_key = _normalize_text(indicator_name)
        if group_key not in groups:
            groups_order.append(group_key)
        groups[group_key].append((col_idx + 1, footer_text))  # 1-based Excel column

    return [
        {'name': groups_order_name, 'cols': groups[groups_order_name]}
        for groups_order_name in groups_order
    ]


def _is_meaningful_footer(text: str) -> bool:
    """Отсеивает "мусорные" тексты третьей строки, которые не являются
    настоящим третьим уровнем шапки:
    - пустые;
    - чисто цифровые (строка-нумератор колонок "А","Б",1,2,3... — она есть
      почти под любой шапкой и не несёт содержательного различия).
    """
    text = (text or '').strip()
    if not text:
        return False
    if text.replace('.', '', 1).isdigit():
        return False
    return True


def audit_file(excel_path: Path, existing_spans: Dict[int, List[str]]) -> List[dict]:
    findings = []
    for group in _group_columns_with_footer(excel_path):
        distinct_footers = {
            footer for _col, footer in group['cols'] if _is_meaningful_footer(footer)
        }
        if len(distinct_footers) < 2:
            continue  # либо нет третьего уровня, либо он одинаков (не потерян)
        # Если ВСЕ различающиеся тексты третьей строки распознаются как
        # маркеры года/периода (см. _detect_year_by_text) — это не баг:
        # столбцы законно представляют один и тот же показатель за разные
        # годы, а сохранение их как "Excel колонка 2022"/"Excel колонка
        # 2023" одного показателя семантически верно, даже если сам
        # генератор угадал год не из той строки шапки, из которой "должен
        # был". Реальная проблема — когда различие НЕ сводится к году.
        if all(_detect_year_by_text(footer) for footer in distinct_footers):
            continue
        # Уже исправлено вручную, если для каждого столбца этой группы в CSV
        # есть СВОЙ код (а не один код на несколько столбцов).
        codes_per_col = [existing_spans.get(col, []) for col, _footer in group['cols']]
        already_split = all(len(codes) == 1 for codes in codes_per_col) and \
            len({tuple(codes) for codes in codes_per_col}) == len(codes_per_col)
        findings.append({
            'group_name': group['name'],
            'columns': group['cols'],
            'already_split_in_csv': already_split,
        })
    return findings


def main() -> int:
    _load_metric_suffixes()
    table_mapping = load_table_source_map(TABLE_MAPPING_FILE)
    unique_files = sorted(set(table_mapping.values()))
    existing_spans = _existing_column_spans(COLUMN_MAPPING_FILE)

    print("===== АУДИТ: потерянный третий уровень шапки Excel =====\n")
    total_open = 0
    total_files_with_issue = 0
    for filename in unique_files:
        excel_path = EXCEL_DIR / filename
        if not excel_path.exists():
            print(f"⚠️ Файл не найден, пропуск: {filename}")
            continue
        try:
            findings = audit_file(excel_path, existing_spans.get(filename, {}))
        except Exception as exc:
            print(f"⚠️ Ошибка при разборе {filename}: {exc}")
            continue
        open_findings = [f for f in findings if not f['already_split_in_csv']]
        if not findings:
            continue
        total_files_with_issue += 1 if open_findings else 0
        print(f"### {filename}")
        for f in findings:
            status = "✅ уже разделено в CSV" if f['already_split_in_csv'] else "❌ ТРЕБУЕТ ИСПРАВЛЕНИЯ"
            cols_desc = ", ".join(f"col{c}='{t}'" for c, t in f['columns'])
            print(f"  {status} — группа «{f['group_name']}»: {cols_desc}")
            if not f['already_split_in_csv']:
                total_open += 1
        print()

    print("===== ИТОГ =====")
    print(f"Файлов с потенциальной проблемой: {total_files_with_issue}")
    print(f"Незакрытых находок (нужно поправить column_mapping_v2.csv): {total_open}")
    return 1 if total_open else 0


if __name__ == "__main__":
    raise SystemExit(main())
