"""Строка данных не должна попадать в шапку «за компанию».

Шапку ищут и по склейке строки со следующей — на случай, когда название
показателя разбито на две строки. Раньше следующая строка объявлялась
частью шапки, как только склейка давала два совпадения, даже если их дала
одна первая строка, а следующая ничего не добавила. Строки шапки не
очищаются, и в бюллетене №2 в строке «азартные игры» таблицы 21, стоявшей
сразу под шапкой, остались цифры прошлого бюллетеня.
"""
import contextlib
import io
import unittest

from docx import Document

from logic import _find_header_rows


INDICATORS = {
    "Финансовый результат убыточных организаций в % к итогу актива": ["FinRez"],
    "Чистая прибыль (убыток) в % ко всем активам": ["ChiPri"],
    "Прибыль, убыток (-) до налогообложения": ["PriDoNal"],
}


def _table(rows):
    doc = Document()
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    for row, values in zip(table.rows, rows):
        for cell, value in zip(row.cells, values):
            cell.text = value
    return table


def _headers(table):
    with contextlib.redirect_stdout(io.StringIO()):
        return _find_header_rows(table, INDICATORS)


class HeaderRowMergeTest(unittest.TestCase):
    def test_data_row_under_complete_header_is_not_header(self):
        table = _table([
            ["", "Финансовый результат убыточных организаций в % к итогу актива",
             "Чистая прибыль (убыток) в % ко всем активам"],
            ["деятельность по организации азартных игр", "-", "3,8"],
            ["Образование", "1,2", "4,5"],
        ])
        self.assertEqual(_headers(table), [0])

    def test_header_split_over_two_rows_is_still_joined(self):
        """Настоящая разбитая шапка: вторая строка добавляет совпадения."""
        table = _table([
            ["", "Прибыль, убыток (-)", "Чистая прибыль (убыток)"],
            ["", "до налогообложения", "в % ко всем активам"],
            ["Образование", "1,2", "4,5"],
        ])
        self.assertEqual(_headers(table), [0, 1])


if __name__ == "__main__":
    unittest.main()
