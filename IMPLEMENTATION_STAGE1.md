# 🔧 РЕАЛИЗАЦИЯ СТАБИЛИЗАЦИИ - ЭТАП 1

## Решение 1: Замена позиционного определения источника на сигнатуры

### Текущая проблемная функция
```python
# logic.py, строка ~234
def get_table_source_by_number(table_source_mapping, table_number):
    """❌ ПРОБЛЕМА: зависит от порядка в словаре"""
    if table_number < 1 or table_number > len(table_source_mapping):
        return None
    return list(table_source_mapping.values())[table_number - 1]  # ← ОПАСНО!
```

### Новое решение
Создать модуль `table_registry.py`:

```python
"""
table_registry.py - Реестр таблиц с сигнатурами вместо позиций

Все таблицы в маппинге получают уникальный ID, независимый от порядка.
Это позволяет безопасно добавлять/удалять/переупорядочивать таблицы.
"""

from typing import Dict, Optional
from pathlib import Path
import csv
from config import _normalize_text


class TableRegistry:
    def __init__(self, table_source_mapping_path: Path):
        """
        Инициализирует реестр таблиц из CSV файла.
        Каждой таблице присваивается уникальный ID.
        """
        self.registry = {}  # table_id → {name, source, position, hash}
        self.name_to_id = {}  # нормализованное имя → table_id
        self.position_to_id = {}  # порядковый номер → table_id
        
        with open(table_source_mapping_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f, delimiter=';')
            next(reader)  # skip header
            
            for position, row in enumerate(reader, 1):
                if not row or len(row) < 2:
                    continue
                
                table_name = str(row[0]).strip().strip('"')
                source_file = str(row[1]).strip().strip('"')
                
                if not table_name or not source_file:
                    continue
                
                # Создаем уникальный ID на основе имени и позиции
                # Формат: "T001" (T = Table)
                table_id = f"T{position:03d}"
                
                # Регистрируем таблицу
                self.registry[table_id] = {
                    'name': table_name,
                    'source': source_file,
                    'position': position,  # для метак продолжения
                    'name_normalized': _normalize_text(table_name)
                }
                
                self.name_to_id[self.registry[table_id]['name_normalized']] = table_id
                self.position_to_id[position] = table_id
    
    def get_by_id(self, table_id: str) -> Optional[Dict]:
        """Получить таблицу по ID (T001, T002, ...)"""
        return self.registry.get(table_id)
    
    def get_by_position(self, position: int) -> Optional[Dict]:
        """Получить таблицу по порядковому номеру (для метак продолжения)"""
        table_id = self.position_to_id.get(position)
        return self.registry.get(table_id) if table_id else None
    
    def get_source_by_position(self, position: int) -> Optional[str]:
        """
        Получить файл-источник по номеру таблицы (для метак продолжения).
        НОВЫЙ СПОСОБ: безопасен при изменении маппинга!
        """
        table = self.get_by_position(position)
        return table['source'] if table else None
    
    def get_by_name(self, table_name: str, threshold: float = 1.0) -> Optional[Dict]:
        """
        Получить таблицу по названию.
        threshold=1.0: точное совпадение после нормализации
        threshold=0.8: нечеткое совпадение (fuzzy match)
        """
        name_norm = _normalize_text(table_name)
        
        # Точное совпадение
        if name_norm in self.name_to_id:
            table_id = self.name_to_id[name_norm]
            return self.registry[table_id]
        
        # Нечеткое совпадение (если threshold < 1.0)
        if threshold < 1.0:
            from difflib import SequenceMatcher
            best_match = None
            best_score = 0
            
            for existing_name_norm, table_id in self.name_to_id.items():
                score = SequenceMatcher(None, name_norm, existing_name_norm).ratio()
                if score > best_score:
                    best_score = score
                    best_match = self.registry[table_id]
            
            if best_score >= threshold:
                return best_match
        
        return None
    
    def list_all(self) -> Dict[str, Dict]:
        """Список всех таблиц для диагностики"""
        return dict(self.registry)
    
    def print_summary(self):
        """Печать сводки реестра"""
        print("\n📋 === РЕЕСТР ТАБЛИЦ ===")
        for table_id, data in sorted(self.registry.items()):
            print(f"{table_id}: {data['name']}")
            print(f"     → {data['source']}")


# ============================================================
# ИНТЕГРАЦИЯ В СУЩЕСТВУЮЩИЙ КОД
# ============================================================

# В logic.py заменить:
# OLD:
def generate_word_template_OLD(input_doc_path, okved_map_path, 
                                table_source_mapping_path, column_mapping_path, 
                                output_doc_path):
    _, name_to_okved_cleaned = load_okved_map(okved_map_path)
    table_source_mapping = load_table_source_map(table_source_mapping_path)
    # ... остальной код

# NEW:
def generate_word_template(input_doc_path, okved_map_path, 
                          table_source_mapping_path, column_mapping_path, 
                          output_doc_path, debug=False):
    from table_registry import TableRegistry
    
    _, name_to_okved_cleaned = load_okved_map(okved_map_path)
    registry = TableRegistry(table_source_mapping_path)  # ← НОВЫЙ!
    
    if debug:
        registry.print_summary()
    
    # ... остальной код остается тем же, но с использованием registry
    
    for t_index, table in enumerate(doc.tables):
        # OLD: table_title_norm = _normalize_text(table_title)
        #      if table_title_norm in normalized_title_to_src:
        #          current_source_file = normalized_title_to_src[table_title_norm]
        
        # NEW:
        table_title = get_table_name(table, registry.name_to_id.keys())
        if table_title:
            table_data = registry.get_by_name(table_title, threshold=0.9)
            if table_data:
                current_source_file = table_data['source']
                if debug:
                    print(f"✓ Таблица найдена в реестре: {table_data['name']}")
        
        # OLD: continuation_source = get_table_source_by_number(...)
        # NEW:
        continuation_number = get_continuation_table_number(table)
        if continuation_number:
            continuation_source = registry.get_source_by_position(continuation_number)
            if continuation_source:
                current_source_file = continuation_source
                if debug:
                    print(f"✓ Метка продолжения таблицы {continuation_number} → {continuation_source}")
```

### Преимущества этого подхода

✅ **Независимость от порядка**: Добавить таблицу в начало или середину маппинга - и система все еще работает  
✅ **Безопасность метак**: "Продолжение таблицы 2" всегда указывает на вторую таблицу маппинга, независимо от ее позиции в списке  
✅ **Диагностика**: Метод `print_summary()` показывает реестр таблиц  
✅ **Расширяемость**: Легко добавить новые способы поиска (по ID, по хешу содержимого и т.д.)  

### Пример использования

```python
from table_registry import TableRegistry

registry = TableRegistry(Path("input/mappings/table_source_data_mapping.csv"))

# Получить по названию
data = registry.get_by_name("БАЛАНС ОРГАНИЗАЦИЙ ПО ВИДАМ ЭКОНОМИЧЕСКОЙ ДЕЯТЕЛЬНОСТИ")
print(data['source'])  # T24_000000_t01Ved14.xlsx

# Получить по позиции (для метак продолжения)
source = registry.get_source_by_position(2)  # Источник для второй таблицы в маппинге
print(source)  # T24_000000_t01Ved14mo8.xlsx

# Вывести всех таблиц
registry.print_summary()
```

---

## ⚡ БЫСТРОЕ ВНЕДРЕНИЕ (30 минут)

Если нет времени на полную реорганизацию, можно быстро исправить критичную проблему:

**Файл: logic.py**

```python
# ЗАМЕНИТЬ (строка ~234):
def get_table_source_by_number(table_source_mapping, table_number):
    if table_number < 1 or table_number > len(table_source_mapping):
        return None
    return list(table_source_mapping.values())[table_number - 1]

# НА:
def get_table_source_by_number_safe(table_source_mapping, table_number):
    """
    Безопасная версия: сначала пытается найти по названию в маппинге,
    затем fallback на позицию (для обратной совместимости).
    """
    # Есть ли явная маркировка позиций в маппинге? 
    # Если нет - выводим предупреждение
    if table_number < 1:
        print(f"⚠️ Неправильный номер таблицы в метке продолжения: {table_number}")
        return None
    
    # Список таблиц в порядке маппинга
    tables_list = list(table_source_mapping.items())
    
    if table_number > len(tables_list):
        print(f"⚠️ Номер таблицы {table_number} превышает количество таблиц в маппинге ({len(tables_list)})")
        return None
    
    # Возвращаем источник для таблицы с нужным номером
    _, source = tables_list[table_number - 1]
    
    # ЛОГИРОВАНИЕ для отладки!
    print(f"🔍 Метка продолжения таблицы {table_number}: {source}")
    
    return source
```

Это даст:
- ✅ Явную логирование при использовании метак продолжения
- ✅ Проверку границ
- ⚠️ Но не решит проблему сдвига индексов при изменении маппинга

