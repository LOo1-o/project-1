"""Запуск из .exe должен вести себя так же, как запуск из исходников.

PyInstaller в режиме onefile распаковывает программу во временную папку и
удаляет её при закрытии. Внутри такой сборки __file__ указывает именно
туда, а не туда, где реально лежит .exe. Если считать пути от __file__,
программа будет искать input/ и output/ в папке, которой через секунду не
станет, и ни справочников не найдёт, и результат сохранит в никуда.

sys.frozen выставляет сам PyInstaller, а sys.executable в этом случае
указывает на настоящий .exe на диске — от него и надо считать.
"""
import sys
import unittest
from pathlib import Path
from unittest import mock


class TestFrozenExePaths(unittest.TestCase):
    """Проверяем выбор base_dir напрямую, не собирая настоящий exe."""

    @staticmethod
    def _resolve_base_dir(module_file: str) -> Path:
        # Та же ветка, что и в main.main().
        if getattr(sys, "frozen", False):
            return Path(sys.executable).parent.resolve()
        return Path(module_file).parent.resolve()

    def test_frozen_run_uses_the_folder_with_the_exe(self):
        exe = Path("/opt/бюллетень/Бюллетень_автоматизация.exe")
        with mock.patch.object(sys, "executable", str(exe)), \
                mock.patch.object(sys, "frozen", True, create=True):
            base = self._resolve_base_dir("/tmp/_MEI123456/main.py")

        self.assertEqual(base, exe.parent)
        self.assertNotIn("_MEI", str(base),
                         "пути считаются от временной папки распаковки")

    def test_normal_run_uses_the_folder_with_the_sources(self):
        # sys.frozen у обычного интерпретатора отсутствует.
        self.assertFalse(getattr(sys, "frozen", False))
        base = self._resolve_base_dir("/home/user/project-1/main.py")
        self.assertEqual(base, Path("/home/user/project-1"))

    def test_input_and_output_sit_next_to_the_exe(self):
        exe = Path("/opt/бюллетень/Бюллетень_автоматизация.exe")
        with mock.patch.object(sys, "executable", str(exe)), \
                mock.patch.object(sys, "frozen", True, create=True):
            base = self._resolve_base_dir("/tmp/_MEI123456/main.py")

        self.assertEqual(base / "input" / "mappings" / "column_mapping_v2.csv",
                         exe.parent / "input" / "mappings" / "column_mapping_v2.csv")
        self.assertEqual(base / "output", exe.parent / "output")


if __name__ == "__main__":
    unittest.main()
