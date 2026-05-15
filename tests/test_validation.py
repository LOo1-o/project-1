import csv
import tempfile
import unittest
from pathlib import Path

from validation import validate_table_registry


class TestValidation(unittest.TestCase):
    def test_validate_table_registry_missing_and_unused_files(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            excel_dir = tmpdir_path / "excel"
            excel_dir.mkdir()

            # Создаем только один нерелевантный excel-файл
            (excel_dir / "unused.xlsx").write_text("dummy")

            mapping_path = tmpdir_path / "table_source_data_mapping.csv"
            with open(mapping_path, "w", encoding="utf-8-sig", newline="") as csvfile:
                writer = csv.writer(csvfile, delimiter=";")
                writer.writerow(["Таблица", "Источник"])
                writer.writerow(["1. Баланс", "file1.xlsx"])

            validation_result = validate_table_registry(mapping_path, excel_dir)

            self.assertIn("file1.xlsx", validation_result["missing_files"])
            self.assertIn("unused.xlsx", validation_result["unused_files"])
            self.assertTrue(validation_result["errors"])
            self.assertTrue(validation_result["warnings"])


if __name__ == "__main__":
    unittest.main()
