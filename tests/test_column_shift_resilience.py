"""Проверяет ключевую идею мапинга: если структура исходного Excel-файла
сдвигается (Росстат вставил/убрал столбец в новом периоде, у другой
территории другой набор столбцов и т.п.), номера колонок из
column_mapping_v2.csv могут стать неверными, но при этом остаться в
допустимых границах таблицы — то есть программа не упадёт и не покажет
ошибку, а просто молча возьмёт данные из ЧУЖОГО столбца.

Ключевые слова в column_mapping_v2.csv существуют именно для восстановления
в такой ситуации. Раньше жёсткий индекс безусловно имел приоритет и
ключевые слова фактически никогда не проверялись, если старый номер просто
оставался в границах таблицы, — эти тесты фиксируют исправленное поведение.
"""
import unittest

import pandas as pd

from data_filler_v2 import _find_column_smart


def _real_shape_df(col_3_name="Валюта баланса", extra_col=None):
    """Строит DataFrame в форме реального росстатовского файла: строка
    заголовка на индексе 3, подзаголовок с годом на индексе 4, буквенный
    маркер А/Б/1/2 на индексе 5, данные с индекса 7."""
    header_row = ["Код ", "Наименование", col_3_name, None, "Внеоборотные активы", None]
    if extra_col is not None:
        header_row = ["Код ", "Наименование", extra_col, col_3_name, None, "Внеоборотные активы", None]
        subheader = [None, None, None, "на конец предыдущего года", "на конец отчетного года", "на конец предыдущего года", "на конец отчетного года"]
        letters = ["А", "Б", 0, 1, 2, 3, 4]
        data = ["101.АГ", "Всего", 111, 731078261, 957828800, 421173737, 518590386]
    else:
        subheader = [None, None, "на конец предыдущего года", "на конец отчетного года", "на конец предыдущего года", "на конец отчетного года"]
        letters = ["А", "Б", 1, 2, 3, 4]
        data = ["101.АГ", "Всего", 731078261, 957828800, 421173737, 518590386]

    rows = [
        [None, None, "Организации без субъектов малого предпринимательства (1-этап)"] + [None] * (len(header_row) - 3),
        [None, None, "Камчатский край"] + [None] * (len(header_row) - 3),
        [None] * len(header_row),
        header_row,
        subheader,
        letters,
        [None] * len(header_row),
        data,
    ]
    return pd.DataFrame(rows)


class TestColumnShiftResilience(unittest.TestCase):
    def test_stale_index_is_overridden_when_column_shifted(self):
        # Раньше "Валюта баланса" была в колонке 3 (1-based); в новом
        # периоде перед ней вставили новый столбец, и она сдвинулась на 4.
        df = _real_shape_df(extra_col="Новый показатель (вставлен в этом периоде)")
        stale_hardcode = "3"  # было верно ДО сдвига
        keywords = ["на конец предыдущего года", "Валюта баланса"]

        col_idx = _find_column_smart(df, stale_hardcode, "2022", keywords)

        # Правильный ответ теперь — колонка с реальным заголовком "Валюта
        # баланса" (индекс 3 в этой сдвинутой таблице), а не колонка 2,
        # где сейчас лежит "Новый показатель".
        self.assertEqual(col_idx, 3)
        self.assertEqual(df.iat[7, col_idx], 731078261)

    def test_unchanged_structure_still_uses_hardcoded_index_with_no_warning(self):
        df = _real_shape_df()
        col_idx = _find_column_smart(df, "3", "2022", ["на конец предыдущего года", "Валюта баланса"])
        self.assertEqual(col_idx, 2)
        self.assertEqual(df.iat[7, col_idx], 731078261)

    def test_no_keywords_falls_back_to_hardcoded_index_unconditionally(self):
        # Если для показателя ключевые слова вообще не заданы (старые
        # записи column_mapping_v2.csv могут быть без них), поведение как
        # раньше — доверяем номеру колонки без проверки.
        df = _real_shape_df()
        col_idx = _find_column_smart(df, "3", "2022", [])
        self.assertEqual(col_idx, 2)


if __name__ == "__main__":
    unittest.main()
