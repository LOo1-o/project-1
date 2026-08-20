"""Проверяет, что поиск колонки не "слепнет", если преамбула перед реальной
шапкой Excel-файла длиннее нескольких строк.

Реальные файлы Росстата сейчас укладываются в 3-4 строки преамбулы перед
строкой "Код"/"Наименование", но это не гарантировано на будущее — в новом
периоде может добавиться ещё одна строка примечания/уточнения. Поиск
ключевых слов при подтверждении номера колонки просматривает только первые
несколько строк файла — если реальная шапка окажется глубже фиксированного
окна, подтверждение по ключевым словам перестаёт работать вообще. Если при
этом ЕЩЁ И структура сдвинулась (см. test_column_shift_resilience), номер
колонки остаётся неисправленным, и данные тихо берутся из чужого столбца.

_detect_header_scan_depth вычисляет реальную глубину шапки по буквенному
маркеру "А","Б",1,2... вместо фиксированного числа строк — эти тесты
проверяют, что это действительно спасает от такого сценария.
"""
import unittest
from pathlib import Path
import tempfile

import openpyxl
import pandas as pd

from data_filler_v2 import _detect_header_scan_depth, _find_column_smart


def _write_deep_and_shifted_excel(path: Path, preamble_lines: int) -> None:
    """Строит файл с длинной преамбулой (глубже фиксированного окна поиска)
    И со сдвинутым столбцом: перед "Валютой баланса" вставлен новый
    показатель, поэтому старый номер колонки (3) теперь указывает на НЕГО,
    а не на "Валюту баланса"."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for i in range(preamble_lines):
        ws.append([None, None, f"Строка примечания №{i + 1}"])
    ws.append(["Код", "Наименование", "Новый показатель", "Валюта баланса", None])
    ws.append([None, None, None, "на конец предыдущего года", "на конец отчетного года"])
    ws.append(["А", "Б", 0, 1, 2])
    ws.append([None] * 5)
    ws.append(["101.АГ", "Всего по обследуемым видам экономической деятельности", 999, 731078261, 957828800])
    wb.save(path)


class TestDeepHeaderResilience(unittest.TestCase):
    def test_dynamic_depth_recovers_shifted_column_behind_long_preamble(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "deep_and_shifted.xlsx"
            # 10 строк преамбулы — больше, чем старое фиксированное окно в 8 строк.
            _write_deep_and_shifted_excel(path, preamble_lines=10)

            df = pd.read_excel(path)
            data_row = len(df) - 1  # последняя строка — данные "101.АГ"

            depth = _detect_header_scan_depth(path)
            self.assertGreater(depth, 8, "глубина должна определяться по реальной структуре, а не быть фиксированной")

            # С фиксированным (старым) окном в 8 строк реальная шапка
            # оказывается вне видимости — подтвердить или пересчитать
            # столбец по ключевым словам НЕЧЕМ, и старый (уже неверный)
            # номер колонки используется как есть.
            shallow = _find_column_smart(df, "3", "2022", ["Валюта баланса"], "shallow", header_scan_rows=8)
            self.assertEqual(df.iat[data_row, shallow], 999, "слепое окно молча берёт данные чужого показателя")

            # С динамически определённой глубиной шапка видна целиком —
            # сдвиг обнаруживается и колонка пересчитывается верно.
            dynamic = _find_column_smart(df, "3", "2022", ["Валюта баланса"], "dynamic", header_scan_rows=depth)
            self.assertEqual(df.iat[data_row, dynamic], 731078261)


if __name__ == "__main__":
    unittest.main()
