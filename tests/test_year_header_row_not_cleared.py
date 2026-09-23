"""Строка шапки с годами внутри таблицы не должна зачищаться.

Строка без узнанного названия считается строкой неизвестной отрасли, и её
числовые ячейки заменяются прочерком. Когда несколько страниц-продолжений
склеены в одну таблицу Word, их шапки с годами оказываются среди строк
данных, а год «2024» по признаку «нет букв» выглядит как число. В готовых
бюллетенях вместо «2023 | 2024» в шапке стояли прочерки, а в бюллетене №2
блок ниже ещё и размечался без годов.
"""
import unittest

from docx import Document

from logic import _is_year_header_row


def _row(values, merge_name_cells=False):
    doc = Document()
    table = doc.add_table(rows=1, cols=len(values))
    row = table.rows[0]
    for cell, value in zip(row.cells, values):
        cell.text = value
    if merge_name_cells:
        row.cells[0].merge(row.cells[1])
    return table.rows[0]


class YearHeaderRowTest(unittest.TestCase):
    def test_year_row_is_header(self):
        self.assertTrue(_is_year_header_row(_row(["", "2024", "2025", "2024", "2025"])))

    def test_year_row_with_merged_name_cell(self):
        """Как в таблице 29 бюллетеня №2: ячейка названия на две колонки."""
        self.assertTrue(_is_year_header_row(_row(["", "", "2024", "2025"], merge_name_cells=True)))

    def test_repeated_years_from_merged_cells(self):
        """«2023 | 2023 | 2024 | 2024» — как в таблице 33 бюллетеня №1."""
        self.assertTrue(_is_year_header_row(_row(["", "2023", "2023", "2024", "2024"])))

    def test_data_row_with_name_is_not_header(self):
        self.assertFalse(_is_year_header_row(_row(["Торговля оптовая", "2024", "2025"])))

    def test_data_values_are_not_header(self):
        self.assertFalse(_is_year_header_row(_row(["", "319 504", "322 324"])))
        self.assertFalse(_is_year_header_row(_row(["", "2024", "322 324"])))

    def test_single_year_is_not_enough(self):
        self.assertFalse(_is_year_header_row(_row(["", "2024", ""])))


if __name__ == "__main__":
    unittest.main()
