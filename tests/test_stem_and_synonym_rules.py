"""Колонки, названные в Word и в Excel разными словами.

Бюллетень №3: экономист удалила column_mapping_v2.csv, программа создала
его заново из Excel — и пропали ручные правки названий. Колонки «Уровень
рентабельности, убыточности (-), в % …» (Excel: «Рентабельность,
убыточность (-) в % …») и «Чистая прибыль (убыток) в % ко всем активам»
остались пустыми. Правило «совпали основы слов» узнаёт их само, а для
совсем разных слов («затраты на производство продаж» — «Себестоимость с
учетом коммерческих и управленческих расходов») есть синонимы.csv.
"""
import tempfile
import unittest
from pathlib import Path

from check_report import _group_problems
from logic import (_match_by_word_stems, _normalize_match_text, _synonym_suggestion, _synonym_variants,
                   load_name_synonyms)

T26 = {
    'Количество организаций, единиц': ['KolOrgEdi'],
    'Прибыль, убыток (-) от продаж': ['PriUbyOtPro'],
    'Себестоимость продаж с учетом коммерческих и управленческих расходов': ['SebPro'],
    'Рентабельность, убыточность (-) в % к себестоимости продаж с учетом коммерческих и управленческих '
    'расходов': ['RenSeb'],
    'Рентабельность, убыточность (-) в % к выручке': ['RenVyr'],
    'Рентабельность, убыточность (-) в % чистой прибыли (убытка) к выручке': ['RenChi'],
}
UNMATCHED = [
    (4, _normalize_match_text('Уровень рентабельности, убыточности (-), в % к себестоимости продаж с учетом '
                              'коммерческих и управленческих расходов'), None),
    (6, _normalize_match_text('Уровень рентабельности, убыточности (-), в % к выручке'), None),
    (7, _normalize_match_text('Уровень рентабельности, убыточности (-), в % чистой прибыли (убытка) к выручке'),
     None),
]


def _run(unmatched, taken=None):
    mapping = dict(taken or {1: ('KolOrgEdi', None), 2: ('PriUbyOtPro', None), 3: ('SebPro', None)})
    matched, left = [], []
    _match_by_word_stems(unmatched, T26, mapping, {}, {}, matched, left)
    return mapping, matched, left


class StemRuleTest(unittest.TestCase):
    def test_level_of_profitability_columns_found(self):
        mapping, matched, left = _run(UNMATCHED)
        self.assertEqual({col: mapping[col][0] for col in (4, 6, 7)},
                         {4: 'RenSeb', 6: 'RenVyr', 7: 'RenChi'})
        self.assertEqual(left, [])
        self.assertTrue(all(score >= 0.7 for *_rest, score in matched))

    def test_taken_indicator_is_not_given_twice(self):
        taken = {1: ('KolOrgEdi', None), 2: ('PriUbyOtPro', None), 3: ('SebPro', None), 5: ('RenVyr', None)}
        mapping, _matched, _left = _run(UNMATCHED[1:2], taken)
        self.assertNotIn(6, mapping, 'показатель уже в колонке 5 — вторую не даём')

    def test_unrelated_column_is_not_guessed(self):
        mapping, matched, left = _run([(8, _normalize_match_text('Темп роста выручки в %'), None)])
        self.assertNotIn(8, mapping)
        self.assertEqual(matched, [])

    def test_two_columns_for_one_indicator_are_not_guessed(self):
        same = [(4, UNMATCHED[1][1], None), (5, _normalize_match_text('Рентабельность, убыточность к выручке, процентов'), None)]
        mapping, matched, _left = _run(same)
        self.assertEqual(matched, [], 'две разные колонки Word похожи на один показатель')

    def test_years_keep_their_codes(self):
        cols = [(4, UNMATCHED[1][1], '24'), (5, UNMATCHED[1][1], '25')]
        mapping, _matched, _left = _run(cols)
        self.assertEqual((mapping[4], mapping[5]), (('RenVyr', '24'), ('RenVyr', '25')))


class SynonymsTest(unittest.TestCase):
    def test_file_and_variants(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'синонимы.csv'
            path.write_text('Как в Word;Как в Excel\n'
                            'затраты на производство продаж;Себестоимость с учетом коммерческих и '
                            'управленческих расходов\n', encoding='utf-8-sig')
            synonyms = load_name_synonyms(path)
        name = _normalize_match_text('Доходы и расходы по обычным видам деятельности Себестоимость с учетом '
                                     'коммерческих и управленческих расходов')
        self.assertEqual(list(_synonym_variants(name, synonyms)),
                         ['доходы и расходы по обычным видам деятельности затраты на производство продаж'])
        self.assertEqual(load_name_synonyms(Path('нет_такого_файла.csv')), [])

    def test_repo_file_is_readable(self):
        path = Path(__file__).resolve().parent.parent / 'input' / 'mappings' / 'синонимы.csv'
        self.assertGreaterEqual(len(load_name_synonyms(path)), 2)


class SynonymSuggestionTest(unittest.TestCase):
    def test_only_differing_words(self):
        self.assertEqual(_synonym_suggestion(UNMATCHED[1][1], _normalize_match_text('Рентабельность, убыточность (-) '
                                                                                   'в % к выручке')),
                         'уровень рентабельности убыточности;рентабельность убыточность')

    def test_whole_names_when_difference_is_only_prepositions(self):
        line = _synonym_suggestion('в процентах чистая прибыль убыток в ко всем активам',
                                   'чистая прибыль убыток до налогообложения в к всем активам')
        self.assertEqual(line, 'чистая прибыль убыток в ко всем активам;'
                               'чистая прибыль убыток до налогообложения в к всем активам')


class GroupProblemsTest(unittest.TestCase):
    def _problem(self, row, col, why, color='оранжевый'):
        return {'Таблица': 19, 'Строка в Word': row, 'Колонка': col, 'Показатель': '', 'Почему пусто': why,
                'Цвет в Бюллетень_ПРОВЕРКА.docx': color}

    def test_whole_column_is_one_line(self):
        problems = [self._problem(f'строка {n}', col, 'Колонка не узнана') for n in range(84) for col in (5, 6, 7)]
        rows = _group_problems(problems)
        self.assertEqual(len(rows), 1)
        self.assertEqual((rows[0]['Колонки'], rows[0]['Сколько ячеек']), ('5, 6, 7', 252))
        self.assertTrue(rows[0]['Строка в Word'].startswith('84 строк'))

    def test_row_level_reasons_stay_per_row(self):
        problems = [self._problem(f'строка {n}', 2, 'Этой строки нет в Excel-файле', 'жёлтый') for n in range(5)]
        self.assertEqual(len(_group_problems(problems)), 5)


if __name__ == '__main__':
    unittest.main()
