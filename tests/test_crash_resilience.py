"""Краш-тесты пайплайна: проверяют, что "плохие" входные файлы (битый Excel,
повреждённый column_mapping_v2.csv и т.п.) не обрушивают весь процесс
необработанным исключением, а деградируют предсказуемо — с понятным
предупреждением в консоли и, где это возможно, пропуском только сломанного
источника.
"""
import tempfile
import unittest
from pathlib import Path

import openpyxl

from config_v2 import build_column_mapping_v2_from_excel, ensure_column_mapping_v2


def _write_good_excel(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Код", "Наименование", "Показатель А"])
    ws.append(["", "", "тыс.руб."])
    ws.append(["А", "Б", "1"])
    ws.append(["01", "Организация 1", "123"])
    wb.save(path)


def _write_corrupted_excel(path: Path) -> None:
    # Файл с расширением .xlsx, но не являющийся реальным Excel-файлом —
    # имитирует повреждённый/неверно сохранённый источник.
    path.write_text("это не excel файл, просто текст с расширением .xlsx")


class TestCrashResilience(unittest.TestCase):
    def test_one_corrupted_excel_file_does_not_crash_build(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            excel_dir = Path(tmpdir) / "excel"
            excel_dir.mkdir()
            _write_good_excel(excel_dir / "good.xlsx")
            _write_corrupted_excel(excel_dir / "corrupted.xlsx")

            table_mapping = {
                "1. Хорошая таблица": "good.xlsx",
                "2. Битая таблица": "corrupted.xlsx",
            }
            output_path = Path(tmpdir) / "column_mapping_v2.csv"

            # Не должно бросить исключение — битый файл пропускается,
            # хороший всё равно попадает в результат.
            build_column_mapping_v2_from_excel(excel_dir, table_mapping, output_path)

            self.assertTrue(output_path.exists())
            content = output_path.read_text(encoding="utf-8-sig")
            self.assertIn("good.xlsx", content)
            self.assertNotIn("corrupted.xlsx", content)

    def test_all_sources_corrupted_raises_clear_error(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            excel_dir = Path(tmpdir) / "excel"
            excel_dir.mkdir()
            _write_corrupted_excel(excel_dir / "corrupted.xlsx")

            table_mapping = {"1. Битая таблица": "corrupted.xlsx"}
            output_path = Path(tmpdir) / "column_mapping_v2.csv"

            with self.assertRaises(ValueError) as ctx:
                build_column_mapping_v2_from_excel(excel_dir, table_mapping, output_path)

            # Сообщение должно называть конкретный проблемный файл, а не
            # быть голым "не удалось сгенерировать".
            self.assertIn("corrupted.xlsx", str(ctx.exception))

    def test_missing_source_file_does_not_crash_build(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            excel_dir = Path(tmpdir) / "excel"
            excel_dir.mkdir()
            _write_good_excel(excel_dir / "good.xlsx")

            table_mapping = {
                "1. Хорошая таблица": "good.xlsx",
                "2. Отсутствующая таблица": "no_such_file.xlsx",
            }
            output_path = Path(tmpdir) / "column_mapping_v2.csv"

            build_column_mapping_v2_from_excel(excel_dir, table_mapping, output_path)

            self.assertTrue(output_path.exists())
            self.assertIn("good.xlsx", output_path.read_text(encoding="utf-8-sig"))

    def test_corrupted_existing_column_mapping_csv_is_regenerated_not_crashed(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            excel_dir = Path(tmpdir) / "excel"
            excel_dir.mkdir()
            _write_good_excel(excel_dir / "good.xlsx")

            table_mapping = {"1. Хорошая таблица": "good.xlsx"}
            output_path = Path(tmpdir) / "column_mapping_v2.csv"

            # Имитация повреждённого файла (например, сохранённого не в той
            # кодировке) — содержит невалидные для utf-8-sig байты.
            with open(output_path, "wb") as f:
                f.write(b"\xff\xfe\x00garbage bytes not valid utf-8-sig \xff\xff\xff")

            # ensure_column_mapping_v2 должна заметить, что файл не читается,
            # и молча пересобрать его, а не уронить весь пайплайн.
            ensure_column_mapping_v2(excel_dir, table_mapping, output_path)

            self.assertIn("good.xlsx", output_path.read_text(encoding="utf-8-sig"))


if __name__ == "__main__":
    unittest.main()
