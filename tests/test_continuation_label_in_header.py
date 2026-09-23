"""Метка «Продолжение таблицы N» не часть названия колонки.

На странице-продолжении метка стоит над шапкой справа, в ячейке,
объединённой над последними колонками, и раньше приклеивалась к их
названиям. Пока название показателя входило в текст шапки дословно, это
не мешало. В бюллетене №2 в Excel опечатка — «до налогооблажения», в Word
верно — «до налогообложения»: дословного вхождения нет, остаётся нечёткое
сравнение, а с приставкой «продолжение таблицы 13» сходство падало с 0,97
до 0,72 при пороге 0,80. Колонка «Прибыль (убыток) до налогообложения»
оставалась без тегов на всех страницах-продолжениях таблиц 13 и 14.
"""
import contextlib
import io
import unittest

from docx import Document

from logic import _compute_section_mapping, _strip_continuation_label


INDICATORS = {
    "Сальдо прочих доходов и расходов": ["SalProDohIRas"],
    "Прибыль (убыток) до налогооблажения": ["PriDoNal"],
    "Налог на прибыль": ["NalNaPri"],
    "Капитал и резервы": ["KapIRez"],
    "Оборотные активы": ["OboAkt"],
}


def _table(rows, merge_right_from=None):
    doc = Document()
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    for row, values in zip(table.rows, rows):
        for cell, value in zip(row.cells, values):
            cell.text = value
    if merge_right_from is not None:
        label_row = table.rows[0]
        merged = label_row.cells[merge_right_from].merge(label_row.cells[-1])
        merged.text = "Продолжение таблицы 13."
    return table


def _mapping(table, header_idx, run_rows=None):
    with contextlib.redirect_stdout(io.StringIO()):
        return _compute_section_mapping(table, header_idx, INDICATORS, run_rows=run_rows)


class StripContinuationLabelTest(unittest.TestCase):
    def test_label_removed(self):
        self.assertEqual(_strip_continuation_label("продолжение таблицы 13 налог на прибыль"),
                         "налог на прибыль")
        self.assertEqual(_strip_continuation_label("продолжение таблицы 4. оборотные активы"),
                         "оборотные активы")
        self.assertEqual(_strip_continuation_label("продолжение таблицы 13."), "")

    def test_other_text_unchanged(self):
        self.assertEqual(_strip_continuation_label("налог на прибыль"), "налог на прибыль")
        self.assertEqual(_strip_continuation_label(""), "")


class ContinuationLabelInHeaderTest(unittest.TestCase):
    def test_typo_indicator_found_under_label(self):
        table = _table([
            ["", "", "", ""],
            ["", "Сальдо прочих доходов и расходов",
             "Прибыль (убыток) до налого-обложения", "Налог на прибыль"],
            ["Добыча угля", "1", "2", "3"],
        ], merge_right_from=2)
        mapping = _mapping(table, 1, run_rows=[0, 1])
        self.assertEqual(mapping.get(1), ("SalProDohIRas", None))
        self.assertEqual(mapping.get(2), ("PriDoNal", None),
                         "метка продолжения не должна мешать узнать колонку с опечаткой в Excel")
        self.assertEqual(mapping.get(3), ("NalNaPri", None))

    def test_year_columns_under_label(self):
        """Шапка с годами: метка над правой группой колонок."""
        table = _table([
            ["", "", "", "", ""],
            ["", "Капитал и резервы", "", "Оборотные активы", ""],
            ["", "2024", "2025", "2024", "2025"],
            ["Всего", "", "", "", ""],
        ], merge_right_from=3)
        self.assertEqual(_mapping(table, 2), {
            1: ("KapIRez", "24"), 2: ("KapIRez", "25"),
            3: ("OboAkt", "24"), 4: ("OboAkt", "25"),
        })


if __name__ == "__main__":
    unittest.main()
