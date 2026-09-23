"""Проверяет, как «Продолжение таблицы N» находит свой Excel-источник.

Раньше источник брался по ПОЗИЦИИ строки в маппинге:

    list(table_source_mapping.values())[table_number - 1]

Это работает только пока номер таблицы в бюллетене совпадает с порядковым
номером строки. Совпадение держится не всегда: load_table_source_map
выбрасывает строки с пустым источником — так помечаются сводные таблицы,
которые ничем не заполняются. Каждая такая строка сдвигает все последующие
позиции на единицу.

В бюллетене с тремя сводными таблицами в начале сдвиг был равен 3:
«Продолжение таблицы 4» получало источник таблицы 7 и молча заполнялось
чужими данными, а продолжения таблиц 21-23 выходили за границу списка,
источника не получали вовсе — и потому не размечались тегами и не
очищались от данных прошлого периода.

Номер таблицы есть в её же названии («13. ФОРМИРОВАНИЕ ...»), поэтому
брать его надо оттуда, а не из порядка строк.
"""
import unittest

from logic import get_table_source_by_number


def _mapping_with_summary_tables():
    """Маппинг бюллетеня, где таблицы 1-3 сводные и в загруженный маппинг
    не попали. Ключи — нормализованные названия, как их отдаёт
    load_table_source_map."""
    return {
        "4. баланс организаций по видам экономической деятельности": "t01.xlsx",
        "5. баланс организаций по муниципальным образованиям": "t01mo8.xlsx",
        "6. внеоборотные активы организаций": "t03.xlsx",
        "7. оборотные активы организаций": "t10.xlsx",
        "21. уровень рентабельности активов организаций": "t27.xlsx",
        "22. уровень рентабельности активов по муниципальным образованиям": "t27mo8.xlsx",
        "23. группировка организаций по коэффициенту текущей ликвидности": "t28.xlsx",
    }


class ContinuationSourceByNumberTest(unittest.TestCase):
    def test_number_wins_over_position(self):
        """Номер из названия, а не позиция строки."""
        mapping = _mapping_with_summary_tables()

        # Позиционный поиск вернул бы здесь t10.xlsx — источник таблицы 7.
        self.assertEqual(get_table_source_by_number(mapping, 4), "t01.xlsx")
        self.assertEqual(get_table_source_by_number(mapping, 6), "t03.xlsx")
        self.assertEqual(get_table_source_by_number(mapping, 7), "t10.xlsx")

    def test_numbers_beyond_mapping_length_still_resolve(self):
        """Таблицы 21-23 при семи строках маппинга: позиционный поиск
        считал бы их выходящими за границу и возвращал None."""
        mapping = _mapping_with_summary_tables()

        self.assertGreater(21, len(mapping))
        self.assertEqual(get_table_source_by_number(mapping, 21), "t27.xlsx")
        self.assertEqual(get_table_source_by_number(mapping, 22), "t27mo8.xlsx")
        self.assertEqual(get_table_source_by_number(mapping, 23), "t28.xlsx")

    def test_unknown_number_returns_none(self):
        """Номера, которого нет в маппинге, быть не должно — и подставлять
        вместо него соседний источник нельзя."""
        mapping = _mapping_with_summary_tables()

        self.assertIsNone(get_table_source_by_number(mapping, 2))
        self.assertIsNone(get_table_source_by_number(mapping, 99))

    def test_non_positive_numbers_return_none(self):
        mapping = _mapping_with_summary_tables()

        self.assertIsNone(get_table_source_by_number(mapping, 0))
        self.assertIsNone(get_table_source_by_number(mapping, -1))

    def test_unnumbered_mapping_keeps_positional_behaviour(self):
        """Если названия не пронумерованы, опереться не на что — остаётся
        прежнее поведение по позиции."""
        mapping = {
            "баланс организаций": "t01.xlsx",
            "внеоборотные активы": "t03.xlsx",
            "оборотные активы": "t10.xlsx",
        }

        self.assertEqual(get_table_source_by_number(mapping, 1), "t01.xlsx")
        self.assertEqual(get_table_source_by_number(mapping, 3), "t10.xlsx")
        self.assertIsNone(get_table_source_by_number(mapping, 4))

    def test_continuous_numbering_unchanged(self):
        """Бюллетень без сводных таблиц: номера и позиции совпадают,
        результат обязан остаться прежним."""
        mapping = {
            "1. баланс организаций": "t01.xlsx",
            "2. внеоборотные активы": "t03.xlsx",
            "3. оборотные активы": "t10.xlsx",
        }

        for number, expected in ((1, "t01.xlsx"), (2, "t03.xlsx"), (3, "t10.xlsx")):
            with self.subTest(number=number):
                self.assertEqual(get_table_source_by_number(mapping, number), expected)


if __name__ == "__main__":
    unittest.main()
