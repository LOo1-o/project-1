"""Проверяет, что пайплайн не завязан на КОНКРЕТНЫЕ имена Excel-файлов
(например, "T24_..." или "T23_...").

Имена файлов от Росстата меняются год от года (T24 -> T25 и т.д.) и могут
отличаться между разными территориальными органами статистики — поэтому
поведение программы должно определяться СОДЕРЖИМЫМ файлов и названиями
таблиц из table_source_data_mapping.csv, а не жёстко заданными строками
имён файлов внутри Python-кода.
"""
import tempfile
import unittest
from pathlib import Path

import openpyxl

from category_mapping import detect_category_prefix


def _write_category_excel(path: Path, group_header: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Код", "Наименование", "Показатель"])
    ws.append(["", "", "тыс.руб."])
    ws.append(["А", "Б", "1"])
    ws.append(["", group_header])
    ws.append(["101.АГ", "Всего по обследуемым видам экономической деятельности", 123])
    wb.save(path)


class TestFilenameIndependence(unittest.TestCase):
    def test_opf_file_detected_regardless_of_name(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            # Совершенно произвольное, "не-T24" имя файла — как если бы его
            # так назвал другой территориальный орган статистики или как
            # если бы имя изменилось в следующем периоде.
            path = Path(tmpdir) / "some_other_region_export_2030.xlsx"
            _write_category_excel(
                path,
                "10000 - ОРГАНИЗАЦИОННО-ПРАВОВЫЕ ФОРМЫ ЮРИДИЧЕСКИХ ЛИЦ, ЯВЛЯЮЩИХСЯ КОММЕРЧЕСКИМИ",
            )
            self.assertEqual(detect_category_prefix(path), "OPF")

    def test_fs_file_detected_regardless_of_name(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "completely_renamed_next_year.xlsx"
            _write_category_excel(path, "10 - Российская собственность")
            self.assertEqual(detect_category_prefix(path), "FS")

    def test_ordinary_file_is_not_treated_as_category_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "regular_okved_source.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Код", "Наименование", "Показатель"])
            ws.append(["", "", "тыс.руб."])
            ws.append(["А", "Б", "1"])
            ws.append(["01", "Организация 1", 123])
            wb.save(path)
            self.assertIsNone(detect_category_prefix(path))


if __name__ == "__main__":
    unittest.main()
