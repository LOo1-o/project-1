"""Что реально делают с файлами люди, которые готовят бюллетень.

Excel-выгрузки правят руками: вставляют пояснительную строку сверху,
переименовывают шапку, дописывают пробелы, сохраняют числа как текст,
удаляют ненужную строку. Опасны не падения — их видно сразу, — а случаи,
когда программа молча берёт ЧУЖОЕ значение и выдаёт его за правильное.
Эти тесты закрывают именно такие сценарии.
"""
import unittest

import pandas as pd

from data_filler_v2 import _derive_year_keywords, _find_column_smart, get_cell_value_safely
from data_normalizer import clean_excel_value_for_word


def _sheet(preamble_rows=0, group="Валюта баланса",
           prev="на конец предыдущего года", cur="на конец отчетного года"):
    rows = [[None, None, None, None]] * preamble_rows
    rows += [
        [None, "Баланс организаций за 2024 год", None, None],
        [None, None, None, None],
        ["Код", "Наименование", group, None],
        [None, None, prev, cur],
        ["А", "Б", 1, 2],
        ["101.АГ", "Всего", 731078261, 957828800],
        ["A", "Сельское хозяйство", 111, 222],
    ]
    return pd.DataFrame(rows)


class TestEditedExcelStructure(unittest.TestCase):
    def test_extra_explanatory_rows_above_the_header(self):
        """Кто-то дописал сверху пару строк с пояснением — номера столбцов
        не меняются, поиск по ключевому слову обязан остаться верным."""
        for preamble in (0, 1, 3):
            with self.subTest(preamble=preamble):
                df = _sheet(preamble_rows=preamble)
                derived = _derive_year_keywords(df, "3", "4", 8 + preamble)
                self.assertEqual(derived, ("на конец предыдущего года",
                                           "на конец отчетного года"))

    def test_column_inserted_before_the_indicator(self):
        """Вставили новый столбец — жёсткий номер устарел, спасать должно
        ключевое слово."""
        df = pd.DataFrame([
            [None, "Баланс организаций за 2024 год", None, None, None],
            [None, None, None, None, None],
            ["Код", "Наименование", "Новый столбец", "Валюта баланса", None],
            [None, None, None, "на конец предыдущего года", "на конец отчетного года"],
            ["А", "Б", 0, 1, 2],
            ["101.АГ", "Всего", 999, 731078261, 957828800],
        ])
        col = _find_column_smart(df, "3", "2022",
                                 ["на конец предыдущего года", "Валюта баланса"], "t", 6)
        self.assertEqual(df.iat[5, col], 731078261,
                         "после вставки столбца взято чужое значение")

    def test_renamed_group_header_still_separates_years(self):
        """Название группы переписали — годы всё равно должны разойтись,
        потому что различающий текст берётся из строки периода."""
        df = _sheet(group="ВАЛЮТА БАЛАНСА (всего по организациям)")
        derived = _derive_year_keywords(df, "3", "4", 8)
        first = _find_column_smart(df, "3", "2022", [derived[0]], "t", 8)
        second = _find_column_smart(df, "4", "2023", [derived[1]], "t", 8)
        self.assertEqual((df.iat[5, first], df.iat[5, second]),
                         (731078261, 957828800))

    def test_period_row_wording_replaced_entirely(self):
        df = _sheet(prev="2023", cur="2024")
        self.assertEqual(_derive_year_keywords(df, "3", "4", 8), ("2023", "2024"))


class TestValuesTypedByHand(unittest.TestCase):
    """Числа в выгрузке нередко оказываются текстом: с пробелами-разделителями
    тысяч, с неразрывными пробелами, с запятой вместо точки."""

    def test_thousand_separators_and_comma_decimal(self):
        cases = {
            "1 234": "1234",
            "1\xa0234": "1234",
            "42117373,7": "42117373,7",
            "-675 642": "-675642",
        }
        for raw, expected in cases.items():
            with self.subTest(raw=raw):
                self.assertEqual(clean_excel_value_for_word(raw), expected)

    def test_blank_like_values_do_not_become_zero(self):
        """Пустая ячейка должна остаться пустой, а не превратиться в 0 —
        иначе «нет данных» станет неотличимо от настоящего нуля."""
        for raw in (None, "", "   "):
            with self.subTest(raw=raw):
                self.assertIn(clean_excel_value_for_word(raw), ("", None))

    def test_dash_stays_a_dash(self):
        self.assertIn(clean_excel_value_for_word("-"), ("-", ""))


class TestMissingData(unittest.TestCase):
    def test_reading_past_the_last_column_returns_nothing(self):
        """Если номер столбца в маппинге больше, чем есть столбцов, нужно
        вернуть пусто, а не свалиться и не взять соседний столбец."""
        df = _sheet()
        row = df.iloc[5]
        self.assertIn(get_cell_value_safely(row, 99), ("", None))

    def test_missing_column_index_is_not_silently_replaced(self):
        df = _sheet()
        self.assertIsNone(_find_column_smart(df, "", "2022", [], "t", 8))


if __name__ == "__main__":
    unittest.main()
