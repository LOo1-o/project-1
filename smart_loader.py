# smart_loader.py - Модуль семантического поиска и поиска по году
"""
Умный поиск данных в Excel таблицах:
1. Поиск колонок по году (динамический)
2. Поиск строк по нечеткому совпадению (fuzzy match)
3. Нормализация текста для поиска
"""

from difflib import SequenceMatcher
import pandas as pd
import numpy as np
from typing import Tuple, Optional, List


def normalize_text(text: str) -> str:
    """
    Нормализует текст для сравнения:
    - Приводит к нижнему регистру
    - Удаляет лишние пробелы
    - Удаляет скобки и содержимое в скобках
    """
    if not isinstance(text, str):
        return ""
    
    # Удаляем скобки и содержимое внутри них: "активы (тыс.руб)" → "активы"
    import re
    text = re.sub(r'\([^)]*\)', '', text)
    
    # Нижний регистр и лишние пробелы
    text = text.strip().lower()
    text = ' '.join(text.split())
    
    return text


'''def fuzzy_match(target: str, candidates: List[str], threshold: float = 0.80) -> Tuple[Optional[str], float]:
    """
    Нечеткое совпадение: ищет лучшее совпадение в списке кандидатов.
    
    Args:
        target: Искомая строка (например, "Внеоборотные активы")
        candidates: Список кандидатов из Excel
        threshold: Минимальный порог сходства (0-1)
    
    Returns:
        (Лучший кандидат, Коэффициент сходства) или (None, 0.0)
    
    Пример:
        >>> fuzzy_match("Внеоборотные активы", ["Внеоборотные активы (тыс.руб)", "Оборотные активы"])
        ("Внеоборотные активы (тыс.руб)", 0.95)
    """
    best_match = None
    best_score = 0.0
    
    target_norm = normalize_text(target)
    
    if not target_norm:
        return None, 0.0
    
    for candidate in candidates:
        candidate_norm = normalize_text(candidate)
        
        if not candidate_norm:
            continue
        
        # Используем SequenceMatcher для расчета сходства
        score = SequenceMatcher(None, target_norm, candidate_norm).ratio()
        
        if score > best_score:
            best_score = score
            best_match = candidate
    
    # Возвращаем результат только если выше порога
    if best_score >= threshold:
        return best_match, best_score
    
    return None, 0.0
'''


def fuzzy_match(target: str, candidates: List[str], threshold: float = 0.80) -> Tuple[Optional[str], float]:
    if not target or not candidates:
        return None, 0.0

    target_norm = normalize_text(target)
    target_tokens = set(target_norm.split())
    if not target_norm:
        return None, 0.0

    best_match = None
    best_score = 0.0

    for cand in candidates:
        cand_norm = normalize_text(cand)
        if not cand_norm:
            continue

        # Standard structural matcher
        match_score = SequenceMatcher(None, target_norm, cand_norm).ratio()

        # Profile token-based fallback match
        cand_tokens = set(cand_norm.split())
        if target_tokens and cand_tokens:
            intersection = target_tokens.intersection(cand_tokens)
            token_score = len(intersection) / max(len(target_tokens), len(cand_tokens))
            # Blend metrics or use max safely
            combined_score = max(match_score, token_score)
        else:
            combined_score = match_score

        if combined_score > best_score:
            best_score = combined_score
            best_match = cand

    if best_score >= threshold:
        return best_match, best_score
    return None, 0.0

def find_column_by_year(df: pd.DataFrame, year: str, year_markers: List[str] = None) -> Tuple[Optional[int], Optional[str]]:
    """
    Находит колонку в Excel по году (динамический поиск вместо жесткого индекса).
    
    Args:
        df: DataFrame из Excel (первая строка - заголовки)
        year: Год для поиска (например, "2023")
        year_markers: Дополнительные маркеры периода (например, ["На конец", "На начало"])
    
    Returns:
        (Индекс колонки, Заголовок колонки) или (None, None)
    
    Логика:
        1. Ищет ячейку с точным совпадением года ("2023")
        2. Если не найдена, ищет маркеры периода ("На конец", "На начало")
        3. Возвращает индекс первой совпавшей колонки
    
    Пример:
        >>> df = pd.read_excel("file.xlsx")
        >>> col_idx, col_name = find_column_by_year(df, "2023")
        >>> # Можно теперь использовать df.iloc[:, col_idx]
    """
    if df.empty:
        return None, None
    
    # Получаем первую строку как заголовок
    headers = df.iloc[0].astype(str) if len(df) > 0 else pd.Series()
    
    # Ищем точное совпадение года
    for col_idx, header in enumerate(headers):
        header_str = str(header).strip()
        if year in header_str:
            return col_idx, header_str
    
    # Если маркеры периода переданы, ищем их
    if year_markers:
        for col_idx, header in enumerate(headers):
            header_str = str(header).strip()
            for marker in year_markers:
                if marker.lower() in header_str.lower():
                    return col_idx, header_str
    
    return None, None


def find_row_by_fuzzy_match(df: pd.DataFrame, target_name: str, threshold: float = 0.80) -> Tuple[Optional[int], Optional[str], float]:
    """
    Находит строку в Excel по нечеткому совпадению названия показателя.
    
    Args:
        df: DataFrame из Excel
        target_name: Название показателя для поиска (например, "Внеоборотные активы")
        threshold: Минимальный порог сходства (0-1)
    
    Returns:
        (Индекс строки, Название из Excel, Коэффициент сходства) или (None, None, 0.0)
    
    Логика:
        1. Нормализует все ячейки первого столбца
        2. Применяет fuzzy_match для каждой ячейки
        3. Возвращает лучшее совпадение выше порога
    
    Пример:
        >>> row_idx, row_name, score = find_row_by_fuzzy_match(df, "Внеоборотные активы")
        >>> print(f"Найдена строка {row_idx}: '{row_name}' (сходство: {score:.2%})")
    """
    if df.empty:
        return None, None, 0.0
    
    # Получаем первый столбец как список названий показателей
    first_column = df.iloc[:, 0].astype(str).tolist()
    
    best_match, best_score = fuzzy_match(target_name, first_column, threshold=threshold)
    
    if best_match is None:
        return None, None, 0.0
    
    # Находим индекс этого совпадения
    for idx, row_val in enumerate(first_column):
        if normalize_text(row_val) == normalize_text(best_match):
            return idx, best_match, best_score
    
    return None, None, 0.0


def get_cell_value_safely(df: pd.DataFrame, row_idx: Optional[int], col_idx: Optional[int]) -> str:
    """
    Безопасно получает значение ячейки, обработав возможные ошибки индексов.
    
    Returns:
        Строковое значение ячейки или пустая строка
    """
    if row_idx is None or col_idx is None:
        return ""
    
    try:
        value = df.iloc[row_idx, col_idx]
        return str(value).strip() if pd.notna(value) else ""
    except (IndexError, KeyError):
        return ""
