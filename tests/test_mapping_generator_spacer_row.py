"""Пустая строка-разделитель в шапке Excel не должна съедать последнюю колонку.

В t10Ved14 бюллетеня №2 между подзаголовками и строкой периодов («на
конец предыдущего / отчетного года») стоит пустая строка. Генератор мерил
ширину шапки по трём строкам под «Код / Наименование», подпись последней
колонки (2025 год) стояла только в строке периодов — и у последнего
показателя не оказалось колонки 2025 года: 88 ячеек таблицы 7 без данных.
"""
import contextlib
import io
import tempfile
import unittest
from pathlib import Path

import openpyxl

from config_v2 import _infer_mapping_from_excel


def _write(path: Path, spacer: bool) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Оборотные активы организаций"])
    ws.append([None, None, None, None, None, None, "тыс.руб."])
    ws.append(["Код", "Наименование", "Оборотные активы", None, "в том числе :", None, None])
    ws.append([None, None, None, None, "запасы", None, "финансовые и другие оборотные активы"])
    if spacer:
        ws.append([])
    ws.append([None, None, "на конец предыдущего года", "на конец отчетного года",
               "на конец предыдущего года", "на конец отчетного года",
               "на конец предыдущего года", "на конец отчетного года"])
    ws.append(["101.АГ", "Всего", 100, 110, 40, 45, 60, 65])
    ws.append(["A", "Сельское хозяйство", 10, 11, 4, 5, 6, 6])
    wb.save(path)


class SpacerRowTest(unittest.TestCase):
    def _rows(self, spacer):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "t10Ved14.xlsx"
            _write(path, spacer)
            with contextlib.redirect_stdout(io.StringIO()):
                rows = _infer_mapping_from_excel(path)
        return {name: (col_1, col_2) for _, name, _, col_1, col_2, _, _ in rows}

    def test_last_indicator_keeps_both_years_with_spacer_row(self):
        rows = self._rows(spacer=True)
        last = [cols for name, cols in rows.items() if "финансовые" in name]
        self.assertEqual(last, [("7", "8")], f"маппинг: {rows}")

    def test_without_spacer_row_unchanged(self):
        rows = self._rows(spacer=False)
        last = [cols for name, cols in rows.items() if "финансовые" in name]
        self.assertEqual(last, [("7", "8")], f"маппинг: {rows}")

    def test_data_numbers_do_not_widen_header(self):
        """Строки данных сразу под шапкой не расширяют её на свои числа."""
        rows = self._rows(spacer=False)
        used = {int(c) for cols in rows.values() for c in cols if c}
        self.assertLessEqual(max(used), 8)


if __name__ == "__main__":
    unittest.main()
