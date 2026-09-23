"""Колонка Word без короткого слова в скобках из названия Excel.

В Excel бюллетеня №2 — «Темп роста прибыли (убытка) в %», в Word —
«Темп роста прибыли, в %». Дословного совпадения нет, и нечёткое сравнение
выбирало более похожий, но чужой показатель «Темп роста суммы прибыли»
(0,87 против 0,85): колонка получила его числа, а настоящая колонка «Темп
роста суммы прибыли» осталась пустой как «дубликат».
"""
import contextlib
import io
import unittest

from docx import Document

from logic import _bracket_word_variants, _compute_section_mapping


INDICATORS = {
    "Прибыль, убыток (-) до налогообложения за предыдущий год": ["PriPreGo"],
    "Темп роста прибыли (убытка) в %": ["TeRosPri"],
    "Темп роста убытка в %": ["TeRosUby"],
    "Темп роста суммы прибыли в %": ["TeRosSumPri"],
}


def _table(rows):
    doc = Document()
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    for row, values in zip(table.rows, rows):
        for cell, value in zip(row.cells, values):
            cell.text = value
    return table


def _mapping(table, indicators, **sinks):
    with contextlib.redirect_stdout(io.StringIO()):
        return _compute_section_mapping(table, 0, indicators, run_rows=[0], **sinks)


class BracketWordVariantsTest(unittest.TestCase):
    def test_short_word_in_brackets_dropped(self):
        self.assertEqual(_bracket_word_variants("Темп роста прибыли (убытка) в %"),
                         {"темп роста прибыли в": "убытка"})

    def test_not_a_word_or_too_short_remainder(self):
        self.assertEqual(_bracket_word_variants("Прибыль, убыток (-) до налогообложения"), {})
        self.assertEqual(_bracket_word_variants("Прибыль (убыток)"), {},
                         "от «Прибыль (убыток)» осталось бы одно слово — это не правило, а угадывание")
        self.assertEqual(_bracket_word_variants("Выручка (в % к итогу)"), {})


class BracketWordRuleTest(unittest.TestCase):
    HEADER = ["", "Прибыль, убыток (-) до налогообложения за предыдущий год",
              "Темп роста прибыли, в %", "Темп роста убытка, в %", "Темп роста суммы прибыли, в %"]

    def test_each_growth_column_gets_its_own_indicator(self):
        table = _table([self.HEADER, ["Камчатский край", "1", "2", "3", "4"]])
        bracket_matches = []
        mapping = _mapping(table, INDICATORS, bracket_match_sink=bracket_matches)
        self.assertEqual(mapping.get(2), ("TeRosPri", None))
        self.assertEqual(mapping.get(3), ("TeRosUby", None))
        self.assertEqual(mapping.get(4), ("TeRosSumPri", None))
        self.assertEqual([(col, indicator) for col, _, _, indicator in bracket_matches], [(2, "TeRosPri")],
                         "совпадение по правилу должно попасть в отчёт")


class ExactMatchKeepsIndicatorTest(unittest.TestCase):
    def test_approximate_column_does_not_take_exact_columns_indicator(self):
        """Без правила (показателя «прибыли (убытка)» нет среди кандидатов)
        похожая колонка не должна забирать чужой показатель у колонки
        с дословным совпадением."""
        indicators = {k: v for k, v in INDICATORS.items() if v != ["TeRosPri"]}
        table = _table([self.HEADER, ["Камчатский край", "1", "2", "3", "4"]])
        taken = []
        mapping = _mapping(table, indicators, taken_column_sink=taken)
        self.assertEqual(mapping.get(4), ("TeRosSumPri", None), "дословное совпадение сохраняет показатель")
        self.assertNotIn(2, mapping, "похожая колонка остаётся пустой, а не с чужими числами")
        self.assertEqual([(col, indicator, winner) for col, _, indicator, winner in taken],
                         [(2, "TeRosSumPri", 4)])

    HEADER = BracketWordRuleTest.HEADER


if __name__ == "__main__":
    unittest.main()
