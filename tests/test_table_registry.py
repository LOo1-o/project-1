import csv
import tempfile
import unittest
from pathlib import Path

from table_registry import TableRegistry


class TestTableRegistry(unittest.TestCase):
    def test_get_by_position_and_fuzzy_name(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            mapping_path = Path(tmpdir) / "table_source_data_mapping.csv"
            with open(mapping_path, "w", encoding="utf-8-sig", newline="") as csvfile:
                writer = csv.writer(csvfile, delimiter=";")
                writer.writerow(["Таблица", "Источник"])
                writer.writerow(["1. Баланс организаций по видам экономической деятельности", "file1.xlsx"])
                writer.writerow(["2. Отчет о движении денежных средств", "file2.xlsx"])

            registry = TableRegistry(mapping_path)

            self.assertEqual(registry.get_source_by_position(1), "file1.xlsx")
            self.assertEqual(registry.get_source_by_position(2), "file2.xlsx")

            result_exact = registry.get_by_name(
                "1. Баланс организаций по видам экономической деятельности",
                threshold=1.0
            )
            self.assertIsNotNone(result_exact)
            self.assertEqual(result_exact["source"], "file1.xlsx")

            result_fuzzy = registry.get_by_name(
                "Баланс оргнизаций по видам деятельности",
                threshold=0.75
            )
            self.assertIsNotNone(result_fuzzy)
            self.assertEqual(result_fuzzy["source"], "file1.xlsx")


if __name__ == "__main__":
    unittest.main()
