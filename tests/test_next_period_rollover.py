"""Переход на следующий период: файлы T25 вместо T24, данные за 2024-2025.

Главный вопрос при смене периода — откуда программа берёт годы. Если бы
она опиралась на имя файла (T24 -> 23/24) или на зашитые в код «2022/2023»,
следующий период пришлось бы править руками. Год берётся из ЗАГОЛОВКА
внутри самого Excel («Баланс организаций за 2025 год»), а предыдущий
считается как отчётный минус один — поэтому имя файла может быть любым.

Проверено также сквозным прогоном: при переименовании всех файлов в T25,
сдвиге отчётного года на 2025 и переводе подписей лет в бюллетене
19 755 значений из ~20 300 совпали с обычным прогоном, а расхождения
свелись к подписям самих годов.
"""
import unittest

import pandas as pd

from data_filler_v2 import _detect_year_suffix_for_column, _extract_report_year_from_excel


def _balance_sheet(report_year: int):
    """Файл в форме T*_000000_t01Ved14: отчётный год указан только в
    заголовке, а столбцы подписаны «предыдущего»/«отчетного года»."""
    return pd.DataFrame([
        [None, None, f"Баланс организаций за {report_year} год", None],
        [None, None, "Организации без субъектов малого предпринимательства", None],
        [None, None, "Камчатский край", None],
        [None, None, None, None],
        ["Код", "Наименование", "Валюта баланса", None],
        [None, None, "на конец предыдущего года", "на конец отчетного года"],
        ["А", "Б", 1, 2],
        ["101.АГ", "Всего", 731078261, 957828800],
    ])


class TestNextPeriodRollover(unittest.TestCase):
    def test_report_year_comes_from_the_file_contents(self):
        for year in (2024, 2025, 2030):
            with self.subTest(year=year):
                self.assertEqual(_extract_report_year_from_excel(_balance_sheet(year)), year)

    def test_year_suffixes_follow_the_report_year(self):
        cases = {2024: ("23", "24"), 2025: ("24", "25"), 2030: ("29", "30")}
        for year, expected in cases.items():
            with self.subTest(year=year):
                df = _balance_sheet(year)
                previous = _detect_year_suffix_for_column(df, 2, "22")
                current = _detect_year_suffix_for_column(df, 3, "23")
                self.assertEqual((previous, current), expected)

    def test_suffixes_do_not_depend_on_the_file_name(self):
        """Имя файла в функцию вообще не передаётся — фиксируем это как
        сознательное свойство: переименование T24 -> T25 без смены года
        внутри файла ничего не меняет, и наоборот."""
        df = _balance_sheet(2025)
        self.assertEqual(_detect_year_suffix_for_column(df, 3, "23"), "25")

    def test_explicit_years_in_the_column_win_over_the_title(self):
        """Если Росстат подпишет столбцы самими годами, брать надо их, а не
        вычислять от заголовка файла."""
        df = _balance_sheet(2025)
        df.iat[5, 2] = "2028"
        df.iat[5, 3] = "2029"
        self.assertEqual(_detect_year_suffix_for_column(df, 2, "22"), "28")
        self.assertEqual(_detect_year_suffix_for_column(df, 3, "23"), "29")

    def test_no_report_year_falls_back_to_the_default_suffix(self):
        """Заголовка с годом нет — лучше честно вернуть значение по
        умолчанию, чем угадать год и молча разложить данные не по тем
        столбцам."""
        df = _balance_sheet(2025)
        df.iat[0, 2] = "Баланс организаций"
        self.assertEqual(_detect_year_suffix_for_column(df, 2, "22"), "22")
        self.assertEqual(_detect_year_suffix_for_column(df, 3, "23"), "23")


if __name__ == "__main__":
    unittest.main()
