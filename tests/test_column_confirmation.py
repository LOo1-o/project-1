"""Колонка из маппинга не должна подменяться чужой, если она правильная.

_find_column_smart доверяет номеру колонки из column_mapping_v2.csv, только
если шапка над ней подтверждается. Если не подтверждается, программа решает,
что структура Excel сдвинулась, и ищет колонку заново. Защита нужная (см.
test_column_shift_resilience), но подтверждение было слишком хрупким и
срабатывало там, где никакого сдвига не было. В готовом бюллетене №1 это
дало 60 ошибочных ячеек в двух колонках — каждый тест ниже воспроизводит
один из реальных случаев.
"""
import unittest

import pandas as pd

from data_filler_v2 import _find_column_smart


def _df(headers):
    """Шапка из одной строки над данными: колонки 0-1 — «Код»/«Наименование»."""
    return pd.DataFrame([["Код", "Наименование"] + headers,
                         ["101.АГ", "Всего"] + [100 + i for i in range(len(headers))]])


class ColumnConfirmationTest(unittest.TestCase):
    def test_hyphenated_header_confirms_mapped_column(self):
        """Таблица 17 бюллетеня №1: в шапке Excel перенос «управлен-ческие»,
        ключевое слово записано целым словом. Правее стоит колонка с тем же
        текстом, но про убыточные организации — её брать нельзя."""
        df = _df(["Себестоимость продаж",
                  "в том числе: коммерческие и управлен-ческие расходы",
                  "Организации, получившие убыток от продаж",
                  "в том числе: коммерческие и управленческие расходы"])
        col = _find_column_smart(df, "4", "2022",
                                 ["в том числе: коммерческие и управленческие расходы"])
        self.assertEqual(col, 3)

    def test_indicator_name_confirms_mapped_column_despite_wrong_keyword(self):
        """Таблица 20 бюллетеня №1: у «дивидендов» ключевое слово скопировано
        со строки «от продажи акций». Номер колонки при этом верный, и
        название показателя в шапке это подтверждает."""
        df = _df(["от продажи акций других организаций (долей участия)",
                  "дивидентов, процентов по долговым финансовым вложениям"])
        col = _find_column_smart(df, "4", "2022",
                                 ["от продажи акций других организаций (долей участия)"],
                                 indicator_name="дивидентов, процентов по долговым финансовым вложениям")
        self.assertEqual(col, 3)

    def test_keyword_does_not_match_inside_another_word(self):
        """Поиск без номера колонки: «оборотные активы» не должны находиться
        внутри «внеоборотные активы», которые в балансе стоят левее."""
        df = _df(["Внеоборотные активы", "Оборотные активы"])
        col = _find_column_smart(df, "", "2023", ["Оборотные активы"])
        self.assertEqual(col, 3)

    def test_real_shift_is_still_followed_and_reported(self):
        """Настоящий сдвиг: на месте колонки из маппинга теперь другой
        показатель. Колонку ищем заново — и пишем об этом в отчёт."""
        df = _df(["Новый показатель", "Валюта баланса"])
        events = []
        col = _find_column_smart(df, "3", "2022", ["Валюта баланса"],
                                 context_label="t01.xlsx: ValBal",
                                 indicator_name="Валюта баланса", events=events)
        self.assertEqual(col, 3)
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0]["Файл Excel"], "t01.xlsx")
        self.assertEqual(events[0]["Значение"], "столбец 3 → 4")

    def test_unconfirmed_column_kept_as_is_is_not_reported(self):
        """Колонку не подтвердить и заменить нечем — берём колонку из
        маппинга. Это не отмена маппинга, в отчёт не идёт: на реальных
        бюллетенях таких случаев под сотню, и почти все — ложные тревоги."""
        df = _df(["Совсем другое"])
        events = []
        col = _find_column_smart(df, "3", "2022", ["Валюта баланса"],
                                 context_label="t01.xlsx: ValBal", events=events)
        self.assertEqual(col, 2)
        self.assertEqual(events, [])

    def test_confirmed_column_is_not_reported(self):
        df = _df(["Валюта баланса"])
        events = []
        col = _find_column_smart(df, "3", "2022", ["Валюта баланса"], events=events)
        self.assertEqual(col, 2)
        self.assertEqual(events, [])


if __name__ == "__main__":
    unittest.main()
