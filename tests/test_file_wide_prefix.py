"""Приставка, общая для всех показателей файла, не требует подтверждения в Word.

В t23-пр бюллетеня №3 все показатели начинаются с «Предыдущий год
(сопоставимый круг)», а в Word этой надписи нет. Групповой префикс
подтверждался только по тексту таблицы, и три колонки таблицы 14 («Прибыль
(убыток) до налогообложения», «Налог на прибыль», «Чистая прибыль»)
оставались без тегов.
"""
import unittest

from logic import _file_wide_prefix, _normalize_match_text


def _norm(names):
    return [_normalize_match_text(n) for n in names]


class FileWidePrefixTest(unittest.TestCase):
    def test_prefix_common_to_all(self):
        names = _norm([
            "Предыдущий год (сопоставимый круг) Налог на прибыль",
            "Предыдущий год (сопоставимый круг) Чистая прибыль (убыток)",
            "Предыдущий год (сопоставимый круг) Доходы и расходы по обычным видам деятельности Выручка",
        ])
        self.assertEqual(_file_wide_prefix(names), ("предыдущий", "год", "сопоставимый", "круг"))

    def test_no_prefix_when_one_name_differs(self):
        names = _norm(["Предыдущий год Налог на прибыль", "Отчетный год Налог на прибыль"])
        self.assertEqual(_file_wide_prefix(names), ())

    def test_prefix_equal_to_a_whole_name_is_not_file_prefix(self):
        """«Оборотные активы» — сам показатель, а не заголовок файла."""
        names = _norm(["Оборотные активы", "Оборотные активы в том числе запасы"])
        self.assertEqual(_file_wide_prefix(names), ())

    def test_single_word_is_not_enough(self):
        names = _norm(["Прибыль от продаж", "Прибыль до налогообложения"])
        self.assertEqual(_file_wide_prefix(names), ())


if __name__ == "__main__":
    unittest.main()
