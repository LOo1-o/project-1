"""Проверяет устойчивость к CSV-справочникам, пересохранённым из Excel на
Windows — обычно это даёт кодировку windows-1251, а не UTF-8. Раньше
несколько загрузчиков (table_source_data_mapping.csv, column_mapping_v2.csv,
unit_annotations.csv) читали файл только как utf-8-sig и падали
необработанным UnicodeDecodeError на первой же кириллице в такой кодировке.
"""
import tempfile
import unittest
from pathlib import Path

from config import load_table_source_map
from config_v2 import load_column_mapping_v2
from logic import load_unit_annotation_texts


def _cp1251_file(tmpdir: Path, name: str, text: str) -> Path:
    path = Path(tmpdir) / name
    path.write_bytes(text.encode("windows-1251"))
    return path


class TestCsvEncodingResilience(unittest.TestCase):
    def test_table_source_mapping_windows1251(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            content = (
                "Таблица;Файл\n"
                "1. Баланс организаций по видам экономической деятельности;file1.xlsx\n"
            )
            path = _cp1251_file(tmpdir, "table_source_data_mapping.csv", content)

            mapping = load_table_source_map(path)

            self.assertEqual(len(mapping), 1)
            self.assertIn("file1.xlsx", mapping.values())

    def test_column_mapping_v2_windows1251(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            content = (
                "Excel файл;Название показателя;Код показателя;"
                "Excel колонка 2022;Excel колонка 2023;"
                "Ключевое слово 2022;Ключевое слово 2023\n"
                "file1.xlsx;Валюта баланса;ValBal;3;4;;\n"
            )
            path = _cp1251_file(tmpdir, "column_mapping_v2.csv", content)

            word_to_indicator, _, indicator_to_file, _, _, _ = load_column_mapping_v2(path)

            self.assertIn("ValBal", word_to_indicator.values())
            self.assertEqual(indicator_to_file["ValBal"], "file1.xlsx")

    def test_unit_annotations_windows1251(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            content = "фраза;комментарий\nтыс.руб.;единица измерения\n"
            path = _cp1251_file(tmpdir, "unit_annotations.csv", content)

            texts = load_unit_annotation_texts(path)

            self.assertTrue(any("тыс руб" in t for t in texts))


if __name__ == "__main__":
    unittest.main()
