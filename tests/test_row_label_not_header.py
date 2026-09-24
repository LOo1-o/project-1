"""Строка «в том числе:» под «Всего» — не шапка таблицы.

В бюллетене №2 показатели t03Ved14 называются «Внеоборотные активы в том
числе материальные…» и «Внеоборотные активы в том числе нематериальные…».
Строка «в том числе:» в таблице 6 набирала два совпадения по тексту и
считалась шапкой новой секции, в которой размечена одна колонка названий.
Вся первая страница таблицы после «Всего» осталась без тегов и без чисел.
"""
import contextlib
import io
import unittest

from docx import Document

from logic import _find_header_rows, _is_row_label_only


INDICATORS = {
    "Внеоборотные активы": ["VneAkt"],
    "Внеоборотные активы в том числе материальные внеоборотные активы": ["VneAktMat"],
    "Внеоборотные активы в том числе нематериальные, финансовые и другие внеоборотные активы": ["VneAktNem"],
}


def _table(rows):
    doc = Document()
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    for row, values in zip(table.rows, rows):
        for cell, value in zip(row.cells, values):
            cell.text = value
    return table


class RowLabelOnlyTest(unittest.TestCase):
    def test_label_row_is_not_header(self):
        table = _table([
            ["", "Внеоборотные активы", "", "в том числе", ""],
            ["", "", "", "материальные внеоборотные активы", "нематериальные, финансовые и другие внеоборотные активы"],
            ["", "2024", "2025", "2024", "2024"],
            ["Всего", "60 220 291", "71 126 628", "44 607 180", "15 613 111"],
            ["в том числе:", "", "", "", ""],
            ["Сельское, лесное хозяйство", "12 107 236", "14 764 568", "9 443 181", "2 664 055"],
        ])
        with contextlib.redirect_stdout(io.StringIO()):
            headers = _find_header_rows(table, INDICATORS)
        self.assertNotIn(4, headers, "«в том числе:» — строка данных, а не шапка")
        self.assertIn(2, headers, "настоящая шапка с годами по-прежнему находится")

    def test_detection(self):
        table = _table([
            ["в том числе:", "", ""],
            ["Всего", "1", "2"],
            ["", "Внеоборотные активы", ""],
            ["", "", ""],
        ])
        self.assertTrue(_is_row_label_only(table.rows[0]))
        self.assertFalse(_is_row_label_only(table.rows[1]), "строка с числами")
        self.assertFalse(_is_row_label_only(table.rows[2]), "подпись над колонкой с числами — это шапка")
        self.assertFalse(_is_row_label_only(table.rows[3]), "пустая строка")


if __name__ == "__main__":
    unittest.main()
