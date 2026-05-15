"""
table_registry.py - Реестр таблиц с сигнатурами

Все таблицы в маппинге получают уникальный ID, независимый от порядка.
Это позволяет безопасно добавлять/удалять/переупорядочивать таблицы
без влияния на определение источников и метки "Продолжение таблицы".

ПРОБЛЕМА ЭТО РЕШАЕТ:
- При добавлении новой таблицы в маппинг все индексы сдвигаются
- Метка "Продолжение таблицы 2" может указывать на неправильный источник
- Результат: неправильный выбор источника для целых таблиц

РЕШЕНИЕ:
- Каждой таблице присваивается уникальный ID (T001, T002, ...) по позиции в маппинге
- Поиск по названию, нормализованному названию или позиции
- Независимо от переупорядочивания в маппинге система работает корректно
"""

from typing import Dict, Optional, Tuple
from pathlib import Path
import csv
from difflib import SequenceMatcher


def _normalize_text(text: str) -> str:
    """Базовая нормализация для сравнения."""
    if not text:
        return ""
    text = text.lower().strip()
    text = text.replace('ё', 'е')
    import re
    text = re.sub(r'\s+', ' ', text).strip()
    return text


class TableRegistry:
    """
    Реестр таблиц с безопасным доступом по ID, названию или позиции.
    
    Формат данных:
    - table_id: "T001", "T002", ... (уникальный ID по позиции при создании)
    - name: "1. БАЛАНС ОРГАНИЗАЦИЙ..."
    - source: "T24_000000_t01Ved14.xlsx"
    - position: 1, 2, 3... (порядковый номер в маппинге)
    """
    
    def __init__(self, table_source_mapping_path: Path):
        """
        Инициализирует реестр таблиц из CSV файла.
        
        Args:
            table_source_mapping_path: Путь к файлу table_source_data_mapping.csv
        """
        self.registry = {}  # table_id → {name, source, position, name_normalized}
        self.name_to_id = {}  # нормализованное имя → table_id
        self.position_to_id = {}  # порядковый номер → table_id
        self.source_to_ids = {}  # source_file → [table_ids]
        
        # Загружаем маппинг из CSV
        with open(table_source_mapping_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f, delimiter=';')
            next(reader)  # skip header "Таблица;Источник"
            
            for position, row in enumerate(reader, 1):
                if not row or len(row) < 2:
                    continue
                
                table_name = str(row[0]).strip().strip('"')
                source_file = str(row[1]).strip().strip('"')
                
                if not table_name or not source_file:
                    continue
                
                # Создаем уникальный ID на основе позиции
                table_id = f"T{position:03d}"
                name_normalized = _normalize_text(table_name)
                
                # Регистрируем таблицу
                self.registry[table_id] = {
                    'name': table_name,
                    'source': source_file,
                    'position': position,
                    'name_normalized': name_normalized
                }
                
                # Индексы для быстрого поиска
                self.name_to_id[name_normalized] = table_id
                self.position_to_id[position] = table_id
                
                if source_file not in self.source_to_ids:
                    self.source_to_ids[source_file] = []
                self.source_to_ids[source_file].append(table_id)
    
    def get_by_id(self, table_id: str) -> Optional[Dict]:
        """
        Получить таблицу по ID (T001, T002, ...).
        
        Args:
            table_id: Идентификатор таблицы в формате "T001"
            
        Returns:
            Словарь с данными таблицы или None
        """
        return self.registry.get(table_id)
    
    def get_by_position(self, position: int) -> Optional[Dict]:
        """
        Получить таблицу по порядковому номеру (для меток продолжения).
        
        Args:
            position: Порядковый номер таблицы (1-based)
            
        Returns:
            Словарь с данными таблицы или None
        """
        table_id = self.position_to_id.get(position)
        return self.registry.get(table_id) if table_id else None
    
    def get_source_by_position(self, position: int) -> Optional[str]:
        """
        Получить файл-источник по номеру таблицы (для меток продолжения).
        
        НОВЫЙ СПОСОБ: безопасен при изменении маппинга!
        Вместо list(dict)[index-1], используем реестр.
        
        Args:
            position: Номер таблицы из метки "Продолжение таблицы N"
            
        Returns:
            Имя файла источника (T24_000000_...) или None
        """
        table = self.get_by_position(position)
        return table['source'] if table else None
    
    def get_by_name(self, table_name: str, threshold: float = 0.95) -> Optional[Dict]:
        """
        Получить таблицу по названию с поддержкой нечеткого совпадения.
        
        Args:
            table_name: Название таблицы из документа Word
            threshold: Минимальный коэффициент сходства (0.0 - 1.0)
                      1.0 = точное совпадение после нормализации
                      0.95 = почти точное
                      0.85 = нечеткое совпадение
            
        Returns:
            Словарь с данными таблицы или None
        """
        name_norm = _normalize_text(table_name)
        
        if not name_norm:
            return None
        
        # Точное совпадение после нормализации
        if name_norm in self.name_to_id:
            table_id = self.name_to_id[name_norm]
            return self.registry[table_id]
        
        # Нечеткое совпадение (если требуется)
        if threshold < 1.0:
            best_match = None
            best_score = 0.0
            
            for existing_name_norm, table_id in self.name_to_id.items():
                score = SequenceMatcher(None, name_norm, existing_name_norm).ratio()
                if score > best_score:
                    best_score = score
                    best_match = self.registry[table_id]
            
            if best_score >= threshold:
                return best_match
        
        return None
    
    def get_all(self) -> Dict[str, Dict]:
        """Получить все таблицы реестра (для диагностики)."""
        return dict(self.registry)
    
    def print_summary(self):
        """Печать сводки реестра для диагностики."""
        print("\n📋 === РЕЕСТР ТАБЛИЦ ===")
        print(f"Всего таблиц: {len(self.registry)}\n")
        
        for table_id in sorted(self.registry.keys()):
            data = self.registry[table_id]
            pos = data['position']
            name = data['name']
            source = data['source']
            
            # Форматируем вывод
            print(f"{table_id} (позиция {pos}): {name}")
            print(f"     └─ Источник: {source}")
        
        print("\n✓ Реестр готов к использованию")
    
    def validate(self) -> Tuple[list, list]:
        """
        Проверить консистентность реестра.
        
        Returns:
            (errors, warnings) - списки ошибок и предупреждений
        """
        errors = []
        warnings = []
        
        # Проверка на дублирование позиций
        positions = [data['position'] for data in self.registry.values()]
        if len(positions) != len(set(positions)):
            errors.append("❌ Найдены дублирующиеся позиции в реестре")
        
        # Проверка на дублирование ID
        ids = list(self.registry.keys())
        if len(ids) != len(set(ids)):
            errors.append("❌ Найдены дублирующиеся ID в реестре")
        
        # Проверка на дублирование нормализованных имен
        names = [data['name_normalized'] for data in self.registry.values()]
        dup_names = [name for name in names if names.count(name) > 1]
        if dup_names:
            warnings.append(f"⚠️ Найдены дублирующиеся нормализованные имена: {set(dup_names)}")
        
        return errors, warnings


# ============================================================
# ПРИМЕР ИСПОЛЬЗОВАНИЯ
# ============================================================
if __name__ == "__main__":
    from pathlib import Path
    
    # Инициализируем реестр
    registry_path = Path(__file__).parent / "input" / "mappings" / "table_source_data_mapping.csv"
    
    if registry_path.exists():
        registry = TableRegistry(registry_path)
        
        # Печатаем сводку
        registry.print_summary()
        
        # Проверяем консистентность
        errors, warnings = registry.validate()
        if errors:
            print("\n❌ ОШИБКИ:")
            for error in errors:
                print(f"  {error}")
        if warnings:
            print("\n⚠️ ПРЕДУПРЕЖДЕНИЯ:")
            for warning in warnings:
                print(f"  {warning}")
        
        # Примеры использования
        print("\n🔍 ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ:")
        
        # По позиции (для меток продолжения)
        source = registry.get_source_by_position(2)
        print(f"Источник таблицы #2: {source}")
        
        # По названию
        name = "БАЛАНС ОРГАНИЗАЦИЙ ПО ВИДАМ ЭКОНОМИЧЕСКОЙ ДЕЯТЕЛЬНОСТИ"
        table = registry.get_by_name(name, threshold=0.9)
        if table:
            print(f"Таблица '{name}': {table['source']}")
        
        # По ID
        table = registry.get_by_id("T001")
        if table:
            print(f"Таблица T001: {table['name']}")
    else:
        print(f"❌ Файл не найден: {registry_path}")
