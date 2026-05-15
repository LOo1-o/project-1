# 🔍 АНАЛИЗ СТАБИЛЬНОСТИ АЛГОРИТМА ПРОСТАНОВКИ ТЕГОВ

## Дата анализа: 14 мая 2026 г.
**Текущее состояние:** Алгоритм работает нестабильно при изменении источников и расширении маппингов.

---

## 🚨 ВЫЯВЛЕННЫЕ ПРОБЛЕМЫ

### **Проблема 1: Жесткая зависимость от позиций в словаре маппинга**

**Локация:** `logic.py`, функции `get_table_source_by_number()` и логика с "меткой продолжения таблицы"

**Суть проблемы:**
```python
def get_table_source_by_number(table_source_mapping, table_number):
    if table_number < 1 or table_number > len(table_source_mapping):
        return None
    return list(table_source_mapping.values())[table_number - 1]  # ← ПРОБЛЕМА!
```

- Используется позиция в списке значений словаря
- Если добавить новую таблицу в `table_source_data_mapping.csv`, все индексы сдвинутся
- При новых источниках и маппингах индексы могут перепутаться

**Пример:**
```
Было:
1. Таблица A → источник 1
2. Таблица B → источник 2
3. Таблица C → источник 3

Станет (добавили таблицу в начало):
1. Таблица NEW → источник NEW
2. Таблица A → источник 1 (индекс сменился с 1 на 2)
3. Таблица B → источник 2 (индекс сменился с 2 на 3)
4. Таблица C → источник 3 (индекс сменился с 3 на 4)

Результат: метка "Продолжение таблицы 1" будет указывать на Таблицу NEW вместо Таблицы A
```

**Критичность:** 🔴 ВЫСОКАЯ (приводит к неправильному выбору источника данных)

---

### **Проблема 2: Неустойчивый поиск названий таблиц**

**Локация:** `logic.py`, функция `get_table_name()` (config.py)

**Суть проблемы:**
- Нормализация текста может работать непредсказуемо с новыми форматами
- Не учитывает вариативность в оформлении названий
- Полагается на точные совпадения после нормализации

**Примеры конфликтов:**
```
"1. БАЛАНС ..." vs "1.БАЛАНС ..." (пробел после цифры)
"БАЛАНС  ОРГАНИЗАЦИЙ" vs "БАЛАНС ОРГАНИЗАЦИЙ" (двойной пробел)
"БАЛАНС ОРГАНИЗАЦИЙ ПО ВИДАМ..." vs "БАЛАНС ОРГАНИЗАЦИЙ по видам..." (разный регистр)
```

**Текущая логика fallback:**
```python
if not current_source_file:
    print("ℹ️ Заголовок таблицы не определён, используем предыдущий источник")
```
Использование **предыдущего** источника опасно: если предыдущая таблица тоже не определилась, может возникнуть каскадный сбой.

**Критичность:** 🟠 СРЕДНЯЯ (неправильное определение источника для отдельной таблицы)

---

### **Проблема 3: Сложная логика определения источника с нечеткими приоритетами**

**Локация:** `logic.py`, функция `generate_word_template()`, строки 250-300

**Суть проблемы:**

Три способа определения источника:
1. По названию таблицы (`get_table_name()`)
2. По метке продолжения (`get_continuation_table_number()`)
3. Авто-определение по содержимому (`auto_detect_table_source()`)

**Приоритет логики:**
```python
# 1. Пытаемся определить по названию
table_title = get_table_name(table, table_source_mapping.keys())
if table_title:
    current_source_file = normalized_title_to_src[table_title_norm]

# 2. Если есть метка продолжения, переопределяем даже если уже определили
if continuation_number:
    current_source_file = get_table_source_by_number(...)  # ПЕРЕЗАТИРАЕТ!

# 3. Если все еще не определили, пытаемся авто-определить
if not current_source_file:
    detected = auto_detect_table_source(...)
```

**Проблемы:**
- Метка продолжения **перезатирает** правильное определение по названию
- Auto-detect зависит от наличия показателей в `file_word_to_indicator`
- Fallback на "предыдущий источник" может быть неправильным

**Критичность:** 🟠 СРЕДНЯЯ (может привести к использованию неправильного источника)

---

### **Проблема 4: Непредсказуемый поиск заголовков в новых структурах таблиц**

**Локация:** `logic.py`, функции:
- `_find_header_rows()` (ищет строки с годами или показателями)
- `_compute_section_mapping()` (вычисляет маппинг столбцов)
- `_find_year_header_row()` (ищет строку с годами)

**Суть проблемы:**

```python
def _find_header_rows(table, source_word_to_indicator, max_search_rows=80):
    for row_idx, row in enumerate(table.rows[:max_search_rows]):
        row_years = [get_cleaned_cell_text(cell).strip() for cell in row.cells]
        year_count = sum(1 for year in row_years if _extract_year(year))
        row_texts = [_normalize_match_text(get_cleaned_cell_text(cell)) for cell in row.cells]
        indicator_score = sum(1 for cell_text in row_texts for name in normalized_names if name and name in cell_text)
        if year_count >= 2 or indicator_score >= 2:  # ← Жесткий критерий!
            headers.append(row_idx)
```

**Проблемы:**
- Критерий `year_count >= 2 or indicator_score >= 2` жестко закодирован
- Если новая структура таблицы содержит только 1 год или 1 показатель, это не будет найдено
- Поиск ограничен `max_search_rows=80` - может не найти заголовок в больших таблицах

**Примеры сбоев:**
```
- Таблица с одним годом (2023) не будет распознана как заголовок
- Таблица с новыми названиями показателей не совпадёт при нормализации
- Если структура таблицы отличается, поиск может найти ложный заголовок
```

**Критичность:** 🔴 ВЫСОКАЯ (приводит к неправильной простановке тегов или отсутствию тегов)

---

### **Проблема 5: Чувствительность к изменениям индексов колонок в Excel**

**Локация:** `data_filler_v2.py`, функция `_find_column_smart()`

**Суть проблемы:**

```python
def _find_column_smart(df, hardcode_idx, year, keywords):
    # 1. Сначала используем жесткий индекс
    if hardcode_idx and hardcode_idx.strip():
        try:
            idx = int(hardcode_idx) - 1
            if 0 <= idx < len(df.columns):
                return idx  # ← Если данные сдвинулись, будет неправильная колонка!
        except ValueError:
            pass
    
    # 2. Пытаемся найти по ключевым словам
    if keywords:
        for keyword in keywords:
            for col_idx, header in enumerate(df.iloc[0]):
                if keyword.lower() in header.lower():
                    return col_idx
```

**Проблемы:**
- Приоритет жесткого индекса выше, чем поиск по ключевым словам
- Если в Excel добавили/удалили колонку, жесткий индекс будет неправильным
- Нет валидации: код не проверяет, что в найденной колонке действительно нужные данные

**Пример:**
```
Было:
Колонка 3 (индекс 2): Валюта баланса (2022)
Колонка 4 (индекс 3): Валюта баланса (2023)

Стало (добавили колонку "Комментарии"):
Колонка 2 (индекс 1): Комментарии
Колонка 3 (индекс 2): Валюта баланса (2022) ← по-прежнему индекс 2!
Колонка 4 (индекс 3): Валюта баланса (2023)

Результат: все остается правильным (совпадение случайное!)
```

**Но если добавить колонку в начало:**
```
Колонка 1 (индекс 0): Номер
Колонка 2 (индекс 1): Код ОКВЭД
Колонка 3 (индекс 2): Комментарии ← НЕПРАВИЛЬНО!
Колонка 4 (индекс 3): Валюта баланса (2022)
Колонка 5 (индекс 4): Валюта баланса (2023)

Результат: индекс 2 теперь указывает на "Комментарии", а не на данные!
```

**Критичность:** 🟡 ВЫСОКАЯ (скрытые ошибки, сложно отследить при изменениях)

---

### **Проблема 6: Отсутствие механизма валидации и диагностики**

**Локация:** Весь проект

**Суть проблемы:**

Нет инструментов для:
- Проверки консистентности маппингов перед запуском
- Логирования на уровне DEBUG с информацией о выборе источников
- Валидации найденных колонок в Excel перед заполнением
- Сравнения ожидаемого количества тегов с реальным

**Критичность:** 🔴 ВЫСОКАЯ (сложно диагностировать проблемы при сбое)

---

## ✅ РЕКОМЕНДАЦИИ ПО СТАБИЛИЗАЦИИ

### **Рекомендация 1: Заменить позиционное определение источника на хеширование (КРИТИЧНО)**

**Текущий подход:**
```python
# Опасно: зависит от порядка в словаре
return list(table_source_mapping.values())[table_number - 1]
```

**Новый подход:**
```python
# Создавать хеш каждой таблицы в маппинге
table_signatures = {}
for i, (table_name, source_file) in enumerate(table_source_mapping.items(), 1):
    table_signatures[f"TABLE_{i:02d}"] = {
        'name': table_name,
        'source': source_file,
        'position': i
    }

def get_table_source_by_signature(table_signatures, table_name):
    # Сначала ищем по точному названию
    for sig_id, data in table_signatures.items():
        if normalize(data['name']) == normalize(table_name):
            return data['source']
    
    # Fallback на частичное совпадение
    best_match = None
    best_score = 0
    for sig_id, data in table_signatures.items():
        score = similarity(normalize(table_name), normalize(data['name']))
        if score > best_score:
            best_score = score
            best_match = data['source']
    
    return best_match if best_score > 0.8 else None
```

**Преимущества:**
- ✅ Независимо от порядка в маппинге
- ✅ Работает с частичными совпадениями
- ✅ Можно добавить индекс для метки "Продолжение таблицы"

---

### **Рекомендация 2: Реорганизовать логику определения источника с ясными приоритетами**

**Текущий код:**
```python
# Три способа, без ясной иерархии
table_title = get_table_name(...)
if table_title:
    current_source_file = ...  # Способ 1

continuation_number = get_continuation_table_number(table)
if continuation_number:
    current_source_file = ...  # Способ 2 (ПЕРЕЗАТИРАЕТ!)

if not current_source_file:
    detected = auto_detect_table_source(...)  # Способ 3
```

**Новый подход:**
```python
def determine_table_source(table, table_source_mapping, file_word_to_indicator, 
                          debug=False):
    """
    Определяет источник таблицы с логированием каждого шага.
    
    Приоритеты:
    1. Метка "Продолжение таблицы N" (самый надежный способ)
    2. Точное совпадение названия таблицы
    3. Частичное совпадение названия (>80% сходства)
    4. Авто-определение по содержимому
    5. Возврат None (таблица должна быть пропущена явно)
    """
    source = None
    method = None
    
    # 1. Метка продолжения (самая надежная)
    continuation_number = get_continuation_table_number(table)
    if continuation_number:
        source = table_signatures.get(f"TABLE_{continuation_number:02d}", {}).get('source')
        if source:
            method = f"continuation_marker(table_{continuation_number})"
            if debug:
                print(f"🔍 [1] Метка продолжения → {source}")
            return source, method
    
    # 2. Точное совпадение названия
    table_title = get_table_name(table, table_source_mapping.keys())
    if table_title:
        source = get_table_source_by_signature(table_signatures, table_title)
        if source:
            method = f"exact_title_match({table_title})"
            if debug:
                print(f"🔍 [2] Точное совпадение названия → {source}")
            return source, method
    
    # 3. Частичное совпадение названия
    if table_title:
        source = get_table_source_by_fuzzy_match(table_signatures, table_title, threshold=0.8)
        if source:
            method = f"fuzzy_title_match({table_title})"
            if debug:
                print(f"🔍 [3] Частичное совпадение названия → {source}")
            return source, method
    
    # 4. Авто-определение по содержимому
    source = auto_detect_table_source(table, table_source_mapping, file_word_to_indicator)
    if source:
        method = f"content_auto_detection"
        if debug:
            print(f"🔍 [4] Авто-определение по содержимому → {source}")
        return source, method
    
    # 5. Не определяем
    if debug:
        print(f"❌ [5] Не удалось определить источник")
    return None, "failed"
```

**Преимущества:**
- ✅ Четкие приоритеты
- ✅ Логирование каждого шага
- ✅ Не зависит от порядка обработки

---

### **Рекомендация 3: Добавить интеллектуальный поиск заголовков таблиц**

**Текущий код:**
```python
if year_count >= 2 or indicator_score >= 2:
    headers.append(row_idx)
```

**Новый подход:**
```python
def find_header_rows_robust(table, source_word_to_indicator, max_search_rows=150):
    """
    Интеллектуальный поиск заголовков:
    - Любое количество лет (включая 1)
    - Адаптивный поиск показателей
    - Логирование найденных заголовков
    """
    headers = []
    normalized_names = [_normalize_match_text(name) for name in source_word_to_indicator.keys()]
    
    for row_idx in range(min(max_search_rows, len(table.rows))):
        row = table.rows[row_idx]
        row_years = [get_cleaned_cell_text(cell).strip() for cell in row.cells]
        year_count = sum(1 for year in row_years if _extract_year(year))
        
        row_texts = [_normalize_match_text(get_cleaned_cell_text(cell)) for cell in row.cells]
        indicator_score = sum(1 for cell_text in row_texts for name in normalized_names 
                             if name and name in cell_text)
        
        # Адаптивные критерии:
        # - Хотя бы 1 год И хотя бы 1 показатель
        # - ИЛИ много показателей (даже без года)
        # - ИЛИ много лет (даже без показателей)
        is_header = (
            (year_count >= 1 and indicator_score >= 1) or  # Новый критерий!
            indicator_score >= 3 or  # Много показателей
            year_count >= 3  # Много лет
        )
        
        if is_header:
            headers.append(row_idx)
            print(f"   🔍 Заголовок найден в строке {row_idx + 1}: "
                  f"{year_count} год(ов), {indicator_score} показателей")
    
    return sorted(set(headers))
```

**Преимущества:**
- ✅ Находит заголовки с 1 годом
- ✅ Адаптируется к новым структурам
- ✅ Логирует найденные заголовки

---

### **Рекомендация 4: Валидировать индексы Excel и предпочитать поиск по ключевым словам**

**Текущий код:**
```python
if hardcode_idx and hardcode_idx.strip():
    idx = int(hardcode_idx) - 1
    if 0 <= idx < len(df.columns):
        return idx  # Ненадежно!
```

**Новый подход:**
```python
def find_column_smart_v2(df, hardcode_idx, year, keywords, column_name="", force_keywords=False):
    """
    Находит колонку с валидацией и приоритизацией:
    
    1. Если force_keywords=True: ищем только по ключевым словам
    2. Если keywords предоставлены: ищем по ключевым словам (надежнее!)
    3. Fallback на жесткий индекс (с валидацией содержимого)
    """
    
    # Способ 1: Поиск по ключевым словам (ПРИОРИТЕТ!)
    if keywords:
        for keyword in keywords:
            for col_idx, header in enumerate(df.iloc[0]):
                header_str = str(header).strip().lower()
                if keyword.lower() in header_str:
                    # ВАЛИДАЦИЯ: проверяем, что колонка содержит числа
                    col_data = df.iloc[1:, col_idx]
                    numeric_count = sum(1 for x in col_data if isinstance(x, (int, float)) or 
                                       (isinstance(x, str) and x.replace('.', '', 1).replace('-', '', 1).isdigit()))
                    if numeric_count > 0:  # Есть числовые данные
                        print(f"   ✓ Найдена колонка '{header}' по ключевому слову '{keyword}'")
                        return col_idx
    
    # Способ 2: Жесткий индекс (только если keywords не помогли)
    if hardcode_idx and hardcode_idx.strip() and not keywords:
        try:
            idx = int(hardcode_idx) - 1
            if 0 <= idx < len(df.columns):
                # ВАЛИДАЦИЯ: проверяем, что это действительно нужная колонка
                header = df.iloc[0, idx]
                if column_name.lower() in str(header).lower():
                    print(f"   ✓ Используем жесткий индекс {hardcode_idx} (валидировано)")
                    return idx
                else:
                    print(f"   ⚠️ Жесткий индекс {hardcode_idx} не совпадает с ожидаемым '{column_name}', "
                          f"найдено: '{header}'")
        except ValueError:
            pass
    
    # Способ 3: Поиск по году (финальный fallback)
    if year:
        for col_idx, header in enumerate(df.iloc[0]):
            if year in str(header):
                print(f"   ⚠️ Колонка найдена по году '{year}' как последний вариант")
                return col_idx
    
    print(f"   ❌ Колонка не найдена для '{column_name}'")
    return None
```

**Преимущества:**
- ✅ Поиск по ключевым словам предпочтительнее
- ✅ Валидация найденных данных
- ✅ Явное логирование проблем

---

### **Рекомендация 5: Добавить механизм валидации и диагностики**

**Создать новый модуль: `validation.py`**

```python
def validate_mappings(table_source_mapping, column_mapping, excel_dir):
    """
    Полная валидация маппингов перед запуском.
    """
    issues = []
    warnings = []
    
    # 1. Проверка файлов Excel
    for table_name, excel_file in table_source_mapping.items():
        filepath = excel_dir / excel_file
        if not filepath.exists():
            issues.append(f"❌ Файл не найден: {excel_file} (для таблицы '{table_name}')")
        else:
            try:
                df = pd.read_excel(filepath)
                print(f"   ✓ {excel_file}: {df.shape[0]} строк, {df.shape[1]} колонок")
            except Exception as e:
                issues.append(f"❌ Ошибка чтения {excel_file}: {e}")
    
    # 2. Проверка маппинга показателей
    for excel_file, word_name, indicator, col_22, col_23 in column_mapping:
        if not word_name:
            warnings.append(f"⚠️ Пустое название показателя для {indicator}")
        if not col_22 and not col_23:
            issues.append(f"❌ Нет колонок для {indicator} в {excel_file}")
    
    # 3. Проверка на дублирование индикаторов
    indicators = [ind for _, _, ind, _, _ in column_mapping]
    duplicates = [ind for ind in indicators if indicators.count(ind) > 1]
    if duplicates:
        warnings.append(f"⚠️ Дублирующиеся коды индикаторов: {set(duplicates)}")
    
    # 4. Проверка консистентности файлов
    excel_files_in_mapping = set(table_source_mapping.values())
    excel_files_actual = {f.name for f in excel_dir.glob('*.xlsx')}
    unused = excel_files_actual - excel_files_in_mapping
    if unused:
        warnings.append(f"⚠️ Файлы не в маппинге: {unused}")
    
    return issues, warnings


def validate_template_against_mappings(doc_path, table_source_mapping, column_mapping):
    """
    Проверка шаблона Word:
    - Есть ли таблицы для каждого маппинга?
    - Совпадают ли названия таблиц?
    """
    doc = Document(doc_path)
    found_tables = []
    
    for idx, table in enumerate(doc.tables, 1):
        title = get_table_name(table, table_source_mapping.keys())
        found_tables.append((idx, title))
        
        if not title:
            print(f"⚠️ Таблица {idx}: название не распознано")
    
    for expected_title, _ in table_source_mapping.items():
        if not any(expected_title.lower() in (title or "").lower() for _, title in found_tables):
            print(f"⚠️ Ожидаемая таблица не найдена: {expected_title}")
    
    print(f"\n📊 Найдено {len(found_tables)} таблиц в шаблоне")
    for idx, title in found_tables:
        print(f"   {idx}. {title or '[не определено]'}")
```

---

### **Рекомендация 6: Добавить DEBUG режим с подробным логированием**

**В `main.py` добавить параметр:**

```python
def main(input_template_override=None, output_file_override=None, debug_mode=False, debug_table=None):
    """
    debug_mode: выводить подробные логи каждого шага
    debug_table: проанализировать только конкретную таблицу (для отладки)
    """
    
    # Валидация перед запуском
    if debug_mode:
        print("\n🔍 === РЕЖИМ ОТЛАДКИ: Валидация маппингов ===")
        issues, warnings = validate_mappings(...)
        for issue in issues:
            print(issue)
        for warning in warnings:
            print(warning)
```

---

## 📋 ПЛАН ВНЕДРЕНИЯ (приоритеты)

| Приоритет | Рекомендация | Сложность | Время | Критичность |
|-----------|-------------|----------|-------|-------------|
| **1** | Заменить позиционное определение источника на сигнатуры | ⭐⭐ | 2ч | 🔴 ВЫСОКАЯ |
| **2** | Реорганизовать логику определения источника | ⭐⭐⭐ | 3ч | 🔴 ВЫСОКАЯ |
| **3** | Добавить валидацию маппингов | ⭐⭐ | 2ч | 🔴 ВЫСОКАЯ |
| **4** | Улучшить поиск заголовков таблиц | ⭐⭐ | 2ч | 🔴 ВЫСОКАЯ |
| **5** | Переделать поиск колонок в Excel | ⭐⭐⭐ | 3ч | 🟡 СРЕДНЯЯ |
| **6** | Добавить DEBUG режим | ⭐ | 1ч | 🟢 НИЗКАЯ |

**Общее время для полной стабилизации:** ~13 часов разработки

---

## 🎯 РЕЗУЛЬТАТЫ ПОСЛЕ ВНЕДРЕНИЯ

✅ Алгоритм будет **независим от порядка таблиц** в маппинге  
✅ Добавление новых таблиц **не сломает существующие**  
✅ Изменение колонок в Excel **будет обнаружено и сообщено**  
✅ Новые структуры таблиц **будут обработаны адаптивно**  
✅ Проблемы будут **явно логированы и диагностированы**  
✅ Код будет **более поддерживаемым** и **понятным**

