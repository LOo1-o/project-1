"""«от продолжающейся деятельности» в названии Excel, которого нет в Word.

В форме отчётности 2025 года Росстат пишет «Прибыль (убыток) от
продолжающейся деятельности до налогообложения», а в Word бюллетеня №3 —
по-прежнему «Прибыль, убыток (-) до налогообложения». Без правила эти
колонки оставались пустыми в таблицах 15–18 и 20.
"""
import contextlib
import io
import unittest

from docx import Document

from logic import _compute_section_mapping, _name_variants


INDICATORS = {
    "Прибыль (убыток) от продолжающейся деятельности до налогообложения в % к: всем активам": ["PriVsAkt"],
    "Прибыль (убыток) от продолжающейся деятельности до налогообложения в % к: оборотным активам": ["PriOboAkt"],
    "Финансовый результат убыточных организаций в % к итогу актива убыточных организаций": ["FinRez"],
}


class OptionalPhraseTest(unittest.TestCase):
    def test_variant_without_phrase(self):
        variants = _name_variants("Прибыль (убыток) от продолжающейся деятельности до налогообложения")
        self.assertEqual(variants.get("прибыль убыток до налогообложения"), "от продолжающейся деятельности")

    def test_too_short_remainder_is_not_a_variant(self):
        variants = _name_variants("Прибыль (убыток) от продолжающейся деятельности")
        self.assertNotIn("прибыль убыток", variants)

    def test_columns_matched_and_reported(self):
        doc = Document()
        rows = [
            ["", "Прибыль, убыток (-) до налогообложения в % к:", "Прибыль, убыток (-) до налогообложения в % к:",
             "Финансовый результат убыточных организаций в % к итогу актива убыточных организаций"],
            ["", "всем активам", "оборотным активам", ""],
            ["Всего", "4,8", "9,9", "8,5"],
        ]
        table = doc.add_table(rows=len(rows), cols=4)
        for row, values in zip(table.rows, rows):
            for cell, value in zip(row.cells, values):
                cell.text = value
        matches = []
        with contextlib.redirect_stdout(io.StringIO()):
            mapping = _compute_section_mapping(table, 1, INDICATORS, run_rows=[0, 1], bracket_match_sink=matches)
        self.assertEqual(mapping.get(1), ("PriVsAkt", None))
        self.assertEqual(mapping.get(2), ("PriOboAkt", None))
        self.assertEqual(mapping.get(3), ("FinRez", None))
        self.assertEqual(sorted(col for col, *_ in matches), [1, 2], "совпадения по правилу попадают в отчёт")


if __name__ == "__main__":
    unittest.main()
