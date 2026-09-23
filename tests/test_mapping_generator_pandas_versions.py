"""Генератор column_mapping_v2 обязан давать одинаковый результат на любой
поддерживаемой версии pandas.

Раньше строки шапки превращались в текст так:

    df.iloc[row].astype(str).fillna('')

astype(str) выполняется первым, и что он делает с пустой ячейкой, зависит от
версии pandas. В pandas 3 пустая ячейка остаётся пустой, и fillna('') делает
из неё ''. В pandas 2.x — а это последняя ветка, которая ставится на
Windows 7 с Python 3.8 — пустая ячейка к этому моменту уже стала текстом
'nan', и fillna ей ничего не заменяет.

В росстатовских шапках пустые ячейки есть всегда: название показателя
объединено над двумя колонками лет, и вторая ячейка объединения пустая.
На pandas 2.x генератор принимал её 'nan' за название нового показателя:
настоящий показатель терял вторую колонку года, а в маппинге появлялись
показатели с именами вида "nan nan". Замерено на бюллетене №2: 111 таких
имён из 175 строк, и после заполнения — 223 молча неверных числа, потому
что для показателя без колонки года поиск по ключевому слову находил
«оборотные активы» внутри «внеоборотные активы».
"""
import tempfile
import unittest
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

from config_v2 import _cell_text, _find_excel_header_row, _infer_mapping_from_excel


def _write_rosstat_like_xlsx(path: Path) -> None:
    """Вёрстка как у t01Ved14.xlsx: преамбула, шапка с объединёнными над
    двумя годами названиями, подзаголовок периода, строка-нумератор, данные."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Баланс организаций"])
    ws.append(["Камчатский край"])
    ws.append([None, None, None, None, None, None, None, None, "тыс.руб."])
    ws.append(["Код", "Наименование", "Валюта баланса", None,
               "Внеоборотные активы", None, "Оборотные активы", None])
    ws.append([None, None,
               "на конец предыдущего года", "на конец отчетного года",
               "на конец предыдущего года", "на конец отчетного года",
               "на конец предыдущего года", "на конец отчетного года"])
    ws.append(["А", "Б", 1, 2, 3, 4, 5, 6])
    ws.append(["101.АГ", "Всего", 100, 110, 40, 45, 60, 65])
    ws.append(["A", "Сельское хозяйство", 10, 11, 4, 5, 6, 6])
    for merged in ("C4:D4", "E4:F4", "G4:H4"):
        ws.merge_cells(merged)
    wb.save(path)


class CellTextTest(unittest.TestCase):
    def test_empty_values_become_empty_string(self):
        for value in (np.nan, None, pd.NA, float("nan")):
            with self.subTest(value=repr(value)):
                self.assertEqual(_cell_text(value), "")

    def test_real_values_keep_their_text(self):
        self.assertEqual(_cell_text("Валюта баланса"), "Валюта баланса")
        self.assertEqual(_cell_text(""), "")
        self.assertEqual(_cell_text(3), "3")
        self.assertEqual(_cell_text("nan"), "nan",
                         "настоящий текст 'nan' в ячейке — не пустота, его не трогаем")


class GeneratorIndependentOfPandasVersionTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.xlsx = Path(self._tmp.name) / "t01Ved14.xlsx"
        _write_rosstat_like_xlsx(self.xlsx)

    def tearDown(self):
        self._tmp.cleanup()

    def _rows_by_name(self):
        rows = _infer_mapping_from_excel(self.xlsx)
        return {name: (col_22, col_23) for _, name, _, col_22, col_23, _, _ in rows}

    def test_no_phantom_nan_indicators(self):
        names = list(self._rows_by_name())
        phantom = [n for n in names if "nan" in n.lower().split()]
        self.assertEqual(phantom, [], f"пустые ячейки шапки стали показателями: {phantom}")

    def test_each_indicator_keeps_both_year_columns(self):
        rows = self._rows_by_name()
        self.assertEqual(rows.get("Валюта баланса"), ("3", "4"))
        self.assertEqual(rows.get("Внеоборотные активы"), ("5", "6"))
        self.assertEqual(rows.get("Оборотные активы"), ("7", "8"))
        self.assertEqual(len(rows), 3, f"лишние показатели: {sorted(rows)}")

    def test_header_row_found(self):
        df = pd.read_excel(self.xlsx, header=None)
        self.assertEqual(_find_excel_header_row(df), 3)


if __name__ == "__main__":
    unittest.main()
