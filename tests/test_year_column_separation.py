"""Годы двух соседних столбцов не должны схлопываться в один.

Реальный баг, найденный сверкой готового бюллетеня с заполненным вручную
оригиналом: в готовом документе колонки за 2023 и 2024 год местами
содержали одинаковые цифры.

Причина — объединённая шапка. Заголовок группы («Внеоборотные активы»)
в Excel растянут на оба года, и pandas кладёт его текст ТОЛЬКО в первый
из двух столбцов. Если ключевым словом для обоих годов оказывается это
название группы, у столбца второго года подтверждать нечем:
_find_column_smart считает структуру изменившейся, ищет слово заново и
возвращается в столбец первого года. Данные второго года не загружаются
вообще.

Лечится тем, что различающее слово берётся из самой шапки
(_derive_year_keywords), а не из заранее заданного словаря — иначе
разметка ломается каждый раз, когда Росстат меняет формулировку.
"""
import unittest

import pandas as pd

from data_filler_v2 import _derive_year_keywords, _find_column_smart


def _merged_header_df(period_prev="на конец предыдущего года",
                      period_cur="на конец отчетного года"):
    """Файл в форме t03Ved14: заголовок группы объединён на два года и
    поэтому виден только в первом столбце пары; строка с периодом стоит
    НЕ сразу под строкой «Наименование», а через уровень подгрупп."""
    return pd.DataFrame([
        [None, "Внеоборотные активы за 2024 год", None, None, None],
        [None, "Организации без субъектов малого предпринимательства", None, None, None],
        [None, "Камчатский край", None, None, None],
        [None, None, None, None, None],
        ["Наименование", "Внеоборотные активы", None, "в том числе:", None],
        [None, None, None, "нематериальные активы", None],
        [None, period_prev, period_cur, period_prev, period_cur],
        ["Б", 1, 2, 3, 4],
        [None, None, None, None, None],
        ["Всего", 42117373.7, 518590386, 41923539, 39129795],
    ])


class TestYearColumnSeparation(unittest.TestCase):
    def test_group_name_as_keyword_for_both_years_collapses_them(self):
        """Фиксируем саму поломку: одинаковое ключевое слово на два года
        уводит оба в один столбец. Если этот тест начнёт падать, значит
        поведение _find_column_smart изменилось и остальные тесты файла
        надо пересмотреть."""
        df = _merged_header_df()
        group_only = ["Внеоборотные активы"]

        first = _find_column_smart(df, "2", "2022", group_only, "test", 8)
        second = _find_column_smart(df, "3", "2023", group_only, "test", 8)

        self.assertEqual(first, second,
                         "ожидали воспроизвести схлопывание годов в один столбец")

    def test_derived_keywords_separate_the_two_years(self):
        df = _merged_header_df()
        derived = _derive_year_keywords(df, "2", "3", 8)
        self.assertIsNotNone(derived, "различающий текст в шапке не найден")

        first = _find_column_smart(df, "2", "2022", [derived[0]], "test", 8)
        second = _find_column_smart(df, "3", "2023", [derived[1]], "test", 8)

        self.assertEqual((first, second), (1, 2))
        self.assertEqual(df.iat[9, first], 42117373.7)
        self.assertEqual(df.iat[9, second], 518590386)

    def test_derivation_survives_changed_wording(self):
        """Ключевые слова не должны быть привязаны к словам
        «предыдущий»/«отчетный»: Росстат может написать годы или даты."""
        for prev, cur in [
            ("2023 год", "2024 год"),
            ("базисный период", "текущий период"),
            ("на 31.12.2023", "на 31.12.2024"),
        ]:
            with self.subTest(wording=f"{prev} / {cur}"):
                df = _merged_header_df(prev, cur)
                derived = _derive_year_keywords(df, "2", "3", 8)
                self.assertEqual(derived, (prev, cur))

                first = _find_column_smart(df, "2", "2022", [derived[0]], "t", 8)
                second = _find_column_smart(df, "3", "2023", [derived[1]], "t", 8)
                self.assertEqual((first, second), (1, 2))

    def test_derivation_ignores_column_number_row(self):
        """Строка с номерами столбцов («1», «2») годы тоже различает, но как
        ключевое слово бесполезна — короткое число совпадёт где угодно."""
        df = _merged_header_df()
        derived = _derive_year_keywords(df, "2", "3", 8)
        self.assertNotIn(derived[0], ("1", "2"))
        self.assertNotIn(derived[1], ("1", "2"))

    def test_no_distinguishing_text_returns_none_instead_of_guessing(self):
        """Если столбцы в шапке вообще ничем не различаются, лучше честно
        вернуть None (вызывающий код напечатает предупреждение), чем выдать
        случайное слово и молча увести год в чужой столбец."""
        df = pd.DataFrame([
            ["Наименование", "Показатель", None],
            [None, "одинаково", "одинаково"],
            ["Б", 1, 2],
            ["Всего", 10, 20],
        ])
        self.assertIsNone(_derive_year_keywords(df, "2", "3", 4))

    def test_out_of_range_columns_do_not_crash(self):
        df = _merged_header_df()
        self.assertIsNone(_derive_year_keywords(df, "99", "100", 8))
        self.assertIsNone(_derive_year_keywords(df, "", "", 8))
        self.assertIsNone(_derive_year_keywords(df, "не число", "3", 8))


if __name__ == "__main__":
    unittest.main()
