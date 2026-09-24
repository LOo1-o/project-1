"""Строка Word с неузнанным названием попадает в отчёт с подсказкой.

В бюллетене №3 на одной странице таблицы 4 написано «торговля оптовая и
розничная транспортными средствами…» вместо «…автотранспортными…».
Строка не узнавалась, очищалась прочерками, и об этом было только в
техническом журнале. Теперь — строка в отчёте «проверьте названия Word и
Excel» с ближайшим названием из справочника.
"""
import contextlib
import io
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from logic import _collect_unknown_row_entry, _write_name_check_report

DIRECTORY = {
    "торговля оптовая и розничная автотранспортными средствами и мотоциклами и их ремонт":
        ("45", "Торговля оптовая и розничная автотранспортными средствами и мотоциклами и их ремонт"),
    "торговля оптовая, кроме оптовой торговли автотранспортными средствами и мотоциклами":
        ("46", "Торговля оптовая, кроме оптовой торговли автотранспортными средствами и мотоциклами"),
}
TITLES = {"4. БАЛАНС ОРГАНИЗАЦИЙ ПО ВИДАМ ЭКОНОМИЧЕСКОЙ ДЕЯТЕЛЬНОСТИ": "t01.xlsx"}


class UnknownRowReportTest(unittest.TestCase):
    def _report(self, *rows):
        report = {}
        for row_name, table_title, continuation in rows:
            _collect_unknown_row_entry(report, table_title, continuation, TITLES, 3, row_name,
                                       "t01.xlsx", "ОКВЭД", DIRECTORY)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "report.xlsx"
            with contextlib.redirect_stdout(io.StringIO()):
                _write_name_check_report(report, path, Path(tmp) / "нет_маппинга.csv")
            return pd.read_excel(path)

    def test_typo_gets_suggestion(self):
        df = self._report(("торговля оптовая и розничная транспортными средствами и мотоциклами и их ремонт",
                           "4. баланс организаций по видам экономической деятельности", None))
        self.assertEqual(len(df), 1)
        row = df.iloc[0]
        self.assertEqual(row["Что случилось"], "Строка не узнана")
        self.assertEqual(row["Таблица в Word"], "Таблица 4")
        self.assertIn("автотранспортными", row["Что сделать"])
        self.assertIn("код 45", row["Что сделать"])

    def test_same_row_on_several_tables_is_one_line(self):
        name = "деятельность по организации и проведению азартных игр"
        df = self._report((name, None, 6), (name, None, 7), (name, None, 6))
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["Таблица в Word"], "Таблицы 6, 7")
        self.assertEqual(df.iloc[0]["Сколько страниц"], 3)
        self.assertIn("Похожей строки в справочнике нет", df.iloc[0]["Что сделать"])


    def test_not_row_names_are_skipped(self):
        report = {}
        for name, title, cont in (("", None, 4), ("6", None, 4), ("Продолжение таблицы 8.", None, 8),
                                  ("запасы", None, None)):
            _collect_unknown_row_entry(report, title, cont, TITLES, 1, name, "t01.xlsx", "ОКВЭД", DIRECTORY)
        self.assertEqual(report, {}, "пустое, номер, метка продолжения, таблица без номера — не в отчёт")


if __name__ == "__main__":
    unittest.main()
