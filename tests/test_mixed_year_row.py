"""Смешанная строка годов в шапке таблицы Word.

Колонки таблицы размечаются по строке шапки, где стоят годы. Раньше она
считалась строкой ТОЛЬКО годов: для колонок, где в ней стоит не год, а
подпись, эта подпись выбрасывалась, и от шапки колонки оставалась лишь
групповая часть выше.

В бюллетене №1 строка годов всегда состояла из одних годов, и ошибка не
проявлялась. В бюллетене №2 у таблицы 4 появились колонки «Валюта баланса
убыточных организаций, на конец года» с подписями «итог актива убыточных
организаций» и «в % к валюте баланса» прямо в строке годов. Групповое
название по началу совпадало с «Валютой баланса», колонка убыточных
организаций размечалась как валюта баланса всех организаций, а колонка
процентов не размечалась вовсе.
"""
import contextlib
import io
import unittest

from docx import Document

from logic import _compute_section_mapping


INDICATORS = {
    "Капитал и резервы": ["KapIRez"],
    "Валюта баланса": ["ValBal"],
    "Валюта баланса убыточных организаций, на конец года итог актива убыточных организаций": ["ValBalUbyIt"],
    "Валюта баланса убыточных организаций, на конец года в % к валюте баланса": ["ValBalUbyPct"],
}


def _table(rows):
    doc = Document()
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    for row, values in zip(table.rows, rows):
        for cell, value in zip(row.cells, values):
            cell.text = value
    return table


def _mapping(table, header_idx):
    with contextlib.redirect_stdout(io.StringIO()):
        return _compute_section_mapping(table, header_idx, INDICATORS)


class MixedYearRowTest(unittest.TestCase):
    def test_label_in_year_row_completes_column_header(self):
        table = _table([
            ["", "Капитал и резервы", "", "Валюта баланса убыточных организаций, на конец года", ""],
            ["", "2024", "2025", "итог актива убыточных организаций", "в % к валюте баланса"],
            ["Всего", "", "", "", ""],
        ])
        mapping = _mapping(table, 1)
        self.assertEqual(mapping.get(1), ("KapIRez", "24"))
        self.assertEqual(mapping.get(2), ("KapIRez", "25"))
        self.assertEqual(mapping.get(3), ("ValBalUbyIt", None),
                         "колонка убыточных организаций не должна стать валютой баланса всех")
        self.assertEqual(mapping.get(4), ("ValBalUbyPct", None),
                         "колонка процентов должна получить свой показатель")

    def test_pure_year_row_unchanged(self):
        """Строка из одних годов — как в бюллетене №1 — размечается как раньше."""
        table = _table([
            ["", "Валюта баланса", "", "Капитал и резервы", ""],
            ["", "2024", "2025", "2024", "2025"],
            ["Всего", "", "", "", ""],
        ])
        self.assertEqual(_mapping(table, 1), {
            1: ("ValBal", "24"), 2: ("ValBal", "25"),
            3: ("KapIRez", "24"), 4: ("KapIRez", "25"),
        })


if __name__ == "__main__":
    unittest.main()
