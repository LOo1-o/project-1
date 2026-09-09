"""Название показателя из column_mapping_v2.csv нельзя резать по запятым.

Реальный баг: поле ключевых слов записано без кавычек, а разбор резал его
по запятым. Из названия «Организации, получившие убыток от продаж: в % к
общему количеству организаций» первым ключом получалось слово
«Организации» — настолько общее, что поиск столбца находил по нему ПЕРВЫЙ
попавшийся столбец с этим словом в шапке и подменял им правильный номер
из маппинга. В одной таблице бюллетеня семь разных показателей получили
из-за этого одно и то же значение из чужого столбца.
"""
import unittest

import pandas as pd

from config_v2 import _parse_keywords_from_field
from data_filler_v2 import _find_column_smart


class TestKeywordFieldParsing(unittest.TestCase):
    def test_unquoted_name_with_commas_stays_one_keyword(self):
        field = "Организации, получившие убыток от продаж: в % к общему количеству организаций"
        self.assertEqual(_parse_keywords_from_field(field), [field])

    def test_quoted_list_is_split_into_separate_keywords(self):
        field = '"2023","конец","Валюта баланса"'
        self.assertEqual(_parse_keywords_from_field(field),
                         ["2023", "конец", "Валюта баланса"])

    def test_quoted_keyword_may_itself_contain_a_comma(self):
        field = '"на конец отчетного года","Прибыль, убыток (-)"'
        self.assertEqual(_parse_keywords_from_field(field),
                         ["на конец отчетного года", "Прибыль, убыток (-)"])

    def test_empty_field_gives_no_keywords(self):
        for value in ("", "   ", None, float("nan")):
            with self.subTest(value=value):
                self.assertEqual(_parse_keywords_from_field(value), [])

    def test_generic_first_word_no_longer_steals_a_foreign_column(self):
        """Сквозная проверка: с прежним разбором столбец 12 подменялся на
        столбец 2, потому что «Организации» находилось в его шапке."""
        df = pd.DataFrame([
            [None, None, "Уровень рентабельности проданных товаров за 2024 год", None, None],
            [None, None, None, None, None],
            ["Код", "Наименование", "Организации, получившие прибыль", None,
             "Организации, получившие убыток от продаж"],
            [None, None, "количество организаций", "в % к количеству",
             "в % к общему количеству организаций"],
            ["А", "Б", 1, 2, 3],
            ["101.АГ", "Всего", 231, 55.5, 40.7],
        ])
        field = "Организации, получившие убыток от продаж: в % к общему количеству организаций"

        col = _find_column_smart(df, "5", "2022", _parse_keywords_from_field(field), "t", 5)

        self.assertEqual(col, 4, "номер столбца из маппинга подменён чужим")
        self.assertEqual(df.iat[5, col], 40.7)


if __name__ == "__main__":
    unittest.main()
