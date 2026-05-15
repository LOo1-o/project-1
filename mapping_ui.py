import streamlit as st
import pandas as pd
from docx import Document
from pathlib import Path
import json
from thefuzz import fuzz, process
import streamlit.components.v1 as components

# ─── Пути ───────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"
EXCEL_DIR = INPUT_DIR / "excel"
MAPPINGS_DIR = INPUT_DIR / "mappings"
DOCX_PATH = INPUT_DIR / "Бюллетень_17.2.8 раздел 1-8.docx"

TABLE_CSV = MAPPINGS_DIR / "table_source_data_mapping.csv"
COLUMN_CSV = MAPPINGS_DIR / "column_mapping_v2.csv"

# ─── Загрузка данных ────────────────────────────────────
@st.cache_data
def load_tables_from_docx(docx_path):
    """Парсит DOCX: возвращает {заголовок_таблицы: [список_показателей]}"""
    doc = Document(docx_path)
    tables_dict = {}
    
    for table in doc.tables:
        if not table.rows:
            continue
        
        header_cells = [cell.text.strip() for cell in table.rows[0].cells]
        table_name = " ".join(header_cells) or f"Таблица {len(tables_dict)+1}"
        
        indicators = []
        for row in table.rows[1:]:
            first_cell = row.cells[0].text.strip() if row.cells else ""
            if first_cell:
                indicators.append(first_cell)
        
        if indicators:
            tables_dict[table_name] = indicators
    
    return tables_dict

@st.cache_data
def load_table_source_mapping():
    """Загружает привязку таблица → excel-файл"""
    if TABLE_CSV.exists():
        df = pd.read_csv(TABLE_CSV, sep=';')
        mapping = {}
        for _, row in df.iterrows():
            mapping[row['Таблица']] = row['Источник']
        return mapping
    return {}

@st.cache_data
def load_column_mapping():
    """Загружает привязку показатель → колонка_excel"""
    if COLUMN_CSV.exists():
        return pd.read_csv(COLUMN_CSV, sep=';')
    return pd.DataFrame(columns=['table', 'indicator', 'column_from_excel'])

@st.cache_data
def load_excel_full(excel_file):
    """Загружает Excel полностью (для автоопределения)"""
    path = EXCEL_DIR / excel_file
    if path.exists():
        return pd.read_excel(path)
    return pd.DataFrame()

@st.cache_data
def load_excel_preview(excel_file, rows=10):
    """Загружает первые N строк Excel"""
    return load_excel_full(excel_file).head(rows)

def save_column_mapping(updated_df):
    """Сохраняет обновлённый column_mapping_v2.csv"""
    updated_df.to_csv(COLUMN_CSV, sep=';', index=False)

# ─── Функция автоопределения ────────────────────────────
def auto_determine_mapping(indicators, excel_columns):
    """
    Автоматически определяет привязку показателей к колонкам Excel.
    Использует fuzzy matching (thefuzz).
    Возвращает {показатель: колонка}
    """
    mapping = {}
    
    for indicator in indicators:
        # Ищем лучшее совпадение среди колонок Excel
        best_match, score = process.extractOne(
            indicator, 
            excel_columns, 
            scorer=fuzz.token_sort_ratio
        )
        
        # Порог схожести: если меньше 60% — не привязываем
        if score >= 60:
            mapping[indicator] = best_match
        else:
            mapping[indicator] = "❌ Не привязано"
    
    return mapping

# ─── Drag-and-drop HTML компонент ───────────────────────
def drag_drop_mapping_ui(indicators, excel_columns, current_mapping):
    """
    Создаёт drag-and-drop интерфейс для привязки колонок.
    Возвращает обновлённый словарь {показатель: колонка}.
    """
    # Подготавливаем данные для JavaScript
    indicators_json = json.dumps(indicators)
    excel_cols_json = json.dumps(excel_columns)
    current_json = json.dumps(current_mapping)
    
    html_code = f"""
    <style>
        .dd-container {{
            display: flex;
            gap: 20px;
            font-family: -apple-system, BlinkMacSystemFont, sans-serif;
        }}
        .dd-column {{
            flex: 1;
            background: #1e1e1e;
            border-radius: 8px;
            padding: 15px;
            min-height: 400px;
        }}
        .dd-column h3 {{
            color: #4fc3f7;
            margin-top: 0;
            font-size: 16px;
            border-bottom: 1px solid #333;
            padding-bottom: 10px;
        }}
        .dd-item {{
            background: #2d2d2d;
            color: #e0e0e0;
            padding: 10px 15px;
            margin: 8px 0;
            border-radius: 6px;
            cursor: grab;
            border-left: 3px solid #4fc3f7;
            font-size: 14px;
            transition: all 0.2s;
        }}
        .dd-item:hover {{
            background: #3d3d3d;
            border-left-color: #81d4fa;
        }}
        .dd-item.dragging {{
            opacity: 0.5;
            cursor: grabbing;
        }}
        .dd-item.mapped {{
            border-left-color: #66bb6a;
            background: #1b3a1b;
        }}
        .dd-item.unmapped {{
            border-left-color: #ef5350;
            background: #3a1b1b;
        }}
        .dd-dropzone {{
            min-height: 50px;
            border: 2px dashed #555;
            border-radius: 8px;
            padding: 10px;
            margin: 5px 0;
            transition: all 0.2s;
        }}
        .dd-dropzone.hover {{
            border-color: #4fc3f7;
            background: rgba(79, 195, 247, 0.1);
        }}
        .dd-pair {{
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 5px 0;
        }}
        .dd-arrow {{
            color: #4fc3f7;
            font-weight: bold;
        }}
        .dd-remove {{
            color: #ef5350;
            cursor: pointer;
            font-weight: bold;
            padding: 2px 8px;
            border-radius: 4px;
        }}
        .dd-remove:hover {{
            background: #ef5350;
            color: white;
        }}
        .dd-button {{
            background: #4fc3f7;
            color: #1e1e1e;
            border: none;
            padding: 8px 16px;
            border-radius: 6px;
            cursor: pointer;
            font-weight: bold;
            margin-top: 10px;
        }}
        .dd-button:hover {{
            background: #81d4fa;
        }}
    </style>
    
    <div class="dd-container">
        <div class="dd-column">
            <h3>📋 Показатели (перетащите к колонке)</h3>
            <div id="indicators-list"></div>
        </div>
        <div class="dd-column">
            <h3>📊 Колонки Excel</h3>
            <div id="excel-columns-list"></div>
        </div>
        <div class="dd-column">
            <h3>🔗 Текущие привязки</h3>
            <div id="mappings-list"></div>
            <button class="dd-button" onclick="saveMappings()">💾 Сохранить</button>
        </div>
    </div>
    
    <script>
        const indicators = {indicators_json};
        const excelColumns = {excel_cols_json};
        let currentMapping = {current_json};
        
        // Отрисовка показателей
        function renderIndicators() {{
            const container = document.getElementById('indicators-list');
            container.innerHTML = '';
            indicators.forEach(ind => {{
                const isMapped = currentMapping[ind] && currentMapping[ind] !== '❌ Не привязано';
                const div = document.createElement('div');
                div.className = 'dd-item ' + (isMapped ? 'mapped' : 'unmapped');
                div.textContent = ind;
                div.draggable = true;
                div.dataset.indicator = ind;
                
                div.addEventListener('dragstart', (e) => {{
                    e.dataTransfer.setData('text/plain', ind);
                    div.classList.add('dragging');
                }});
                div.addEventListener('dragend', () => {{
                    div.classList.remove('dragging');
                }});
                
                container.appendChild(div);
            }});
        }}
        
        // Отрисовка колонок Excel как drop-зон
        function renderExcelColumns() {{
            const container = document.getElementById('excel-columns-list');
            container.innerHTML = '';
            excelColumns.forEach(col => {{
                const dropzone = document.createElement('div');
                dropzone.className = 'dd-dropzone';
                dropzone.dataset.column = col;
                
                const label = document.createElement('div');
                label.className = 'dd-item';
                label.textContent = col;
                label.style.cursor = 'default';
                dropzone.appendChild(label);
                
                // Показываем привязанные показатели
                const mappedIndicators = Object.entries(currentMapping)
                    .filter(([ind, c]) => c === col)
                    .map(([ind]) => ind);
                
                mappedIndicators.forEach(ind => {{
                    const pair = document.createElement('div');
                    pair.className = 'dd-pair';
                    pair.innerHTML = `<span class="dd-arrow">↳</span> ${{ind}} <span class="dd-remove" onclick="removeMapping('${{ind}}')">×</span>`;
                    dropzone.appendChild(pair);
                }});
                
                dropzone.addEventListener('dragover', (e) => {{
                    e.preventDefault();
                    dropzone.classList.add('hover');
                }});
                dropzone.addEventListener('dragleave', () => {{
                    dropzone.classList.remove('hover');
                }});
                dropzone.addEventListener('drop', (e) => {{
                    e.preventDefault();
                    dropzone.classList.remove('hover');
                    const indicator = e.dataTransfer.getData('text/plain');
                    if (indicator) {{
                        currentMapping[indicator] = col;
                        renderAll();
                    }}
                }});
                
                container.appendChild(dropzone);
            }});
            
            // Добавляем зону "Не привязано"
            const unmappedZone = document.createElement('div');
            unmappedZone.className = 'dd-dropzone';
            unmappedZone.style.borderColor = '#ef5350';
            unmappedZone.dataset.column = '❌ Не привязано';
            
            const unmappedLabel = document.createElement('div');
            unmappedLabel.className = 'dd-item unmapped';
            unmappedLabel.textContent = '❌ Не привязано';
            unmappedLabel.style.cursor = 'default';
            unmappedZone.appendChild(unmappedLabel);
            
            const unmappedIndicators = Object.entries(currentMapping)
                .filter(([ind, c]) => c === '❌ Не привязано')
                .map(([ind]) => ind);
            
            unmappedIndicators.forEach(ind => {{
                const pair = document.createElement('div');
                pair.className = 'dd-pair';
                pair.innerHTML = `<span class="dd-arrow">↳</span> ${{ind}} <span class="dd-remove" onclick="removeMapping('${{ind}}')">×</span>`;
                unmappedZone.appendChild(pair);
            }});
            
            unmappedZone.addEventListener('dragover', (e) => {{
                e.preventDefault();
                unmappedZone.classList.add('hover');
            }});
            unmappedZone.addEventListener('dragleave', () => {{
                unmappedZone.classList.remove('hover');
            }});
            unmappedZone.addEventListener('drop', (e) => {{
                e.preventDefault();
                unmappedZone.classList.remove('hover');
                const indicator = e.dataTransfer.getData('text/plain');
                if (indicator) {{
                    currentMapping[indicator] = '❌ Не привязано';
                    renderAll();
                }}
            }});
            
            container.appendChild(unmappedZone);
        }}
        
        // Отрисовка сводки привязок
        function renderMappings() {{
            const container = document.getElementById('mappings-list');
            container.innerHTML = '';
            
            const mapped = Object.entries(currentMapping).filter(([ind, col]) => col !== '❌ Не привязано');
            const unmapped = Object.entries(currentMapping).filter(([ind, col]) => col === '❌ Не привязано');
            
            container.innerHTML += `<p style="color: #66bb6a;">✅ Привязано: ${{mapped.length}}</p>`;
            container.innerHTML += `<p style="color: #ef5350;">❌ Не привязано: ${{unmapped.length}}</p>`;
            
            if (mapped.length > 0) {{
                container.innerHTML += '<hr style="border-color: #333;">';
                mapped.forEach(([ind, col]) => {{
                    const div = document.createElement('div');
                    div.className = 'dd-item mapped';
                    div.innerHTML = `${{ind}} <span class="dd-arrow">→</span> ${{col}} <span class="dd-remove" onclick="removeMapping('${{ind}}')">×</span>`;
                    container.appendChild(div);
                }});
            }}
        }}
        
        function removeMapping(indicator) {{
            currentMapping[indicator] = '❌ Не привязано';
            renderAll();
        }}
        
        function saveMappings() {{
            // Отправляем данные в Streamlit
            window.parent.postMessage({{
                type: 'streamlit:setComponentValue',
                value: currentMapping
            }}, '*');
        }}
        
        function renderAll() {{
            renderIndicators();
            renderExcelColumns();
            renderMappings();
        }}
        
        renderAll();
    </script>
    """
    
    return components.html(html_code, height=600)

# ─── UI ─────────────────────────────────────────────────
st.set_page_config(layout="wide")
st.title("✏️ Редактор привязки колонок к показателям")

# Загружаем все данные
tables_dict = load_tables_from_docx(DOCX_PATH)
table_source = load_table_source_mapping()
column_mapping_df = load_column_mapping()

if not tables_dict:
    st.error("❌ Не найдены таблицы в DOCX. Проверьте путь к файлу.")
    st.stop()

# ─── Сайдбар: выбор таблицы ────────────────────────────
st.sidebar.header("📋 Таблицы в документе")
table_names = list(tables_dict.keys())
selected_table = st.sidebar.selectbox("Выберите таблицу для редактирования", table_names)

current_excel = table_source.get(selected_table, None)

if not current_excel:
    st.error(f"❌ Таблица «{selected_table}» не привязана к Excel-файлу!")
    st.stop()

st.success(f"📁 Источник данных: **{current_excel}**")

# Загружаем данные Excel
full_df = load_excel_full(current_excel)
preview_df = load_excel_preview(current_excel)
excel_columns = list(full_df.columns) if not full_df.empty else []

indicators = tables_dict[selected_table]

# Фильтруем существующие привязки
existing_for_table = column_mapping_df[column_mapping_df['table'] == selected_table]
current_mapping = {}
for _, row in existing_for_table.iterrows():
    current_mapping[row['indicator']] = row['column_from_excel']

# Заполняем отсутствующие показатели
for ind in indicators:
    if ind not in current_mapping:
        current_mapping[ind] = "❌ Не привязано"

# ─── Табы: выбор режима редактирования ─────────────────
tab1, tab2, tab3 = st.tabs(["📋 Табличный режим", "🖱️ Drag-and-Drop", "⚡ Массовые операции"])

# ─── ТАБ 1: Табличный режим ────────────────────────────
with tab1:
    with st.expander("📊 Предпросмотр Excel (первые 10 строк)", expanded=True):
        if not preview_df.empty:
            st.dataframe(preview_df, use_container_width=True)
    
    st.subheader("Привязка показателей")
    
    excel_columns_with_none = ["❌ Не привязано"] + excel_columns
    edited_mapping = {}
    
    for indicator in indicators:
        current_col = current_mapping.get(indicator, "❌ Не привязано")
        
        col1, col2, col3 = st.columns([4, 3, 1])
        
        with col1:
            st.markdown(f"**{indicator}**")
        
        with col2:
            try:
                default_idx = excel_columns_with_none.index(current_col)
            except ValueError:
                default_idx = 0
            
            new_col = st.selectbox(
                f"Колонка",
                excel_columns_with_none,
                index=default_idx,
                key=f"table_{selected_table}_{indicator}",
                label_visibility="collapsed"
            )
            edited_mapping[indicator] = new_col
        
        with col3:
            if new_col == "❌ Не привязано":
                st.error("⚠️")
            elif new_col != current_col:
                st.warning("✏️")
            else:
                st.success("✅")
    
    # Статистика
    total = len(indicators)
    mapped = sum(1 for v in edited_mapping.values() if v != "❌ Не привязано")
    unmapped = total - mapped
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Всего", total)
    col2.metric("Привязано", mapped)
    col3.metric("Не привязано", unmapped, delta=f"-{unmapped}" if unmapped > 0 else "0")
    
    if st.button("💾 Сохранить изменения (табличный режим)", type="primary"):
        column_mapping_df = column_mapping_df[column_mapping_df['table'] != selected_table]
        new_rows = []
        for ind, col in edited_mapping.items():
            if col != "❌ Не привязано":
                new_rows.append({'table': selected_table, 'indicator': ind, 'column_from_excel': col})
        if new_rows:
            column_mapping_df = pd.concat([column_mapping_df, pd.DataFrame(new_rows)], ignore_index=True)
        save_column_mapping(column_mapping_df)
        st.cache_data.clear()
        st.rerun()

# ─── ТАБ 2: Drag-and-Drop ──────────────────────────────
with tab2:
    st.subheader("🖱️ Перетащите показатели к нужным колонкам")
    st.caption("Перетаскивайте показатели из левой колонки на колонки Excel в центре. Для отвязки — перетащите в зону «Не привязано».")
    
    # Drag-and-drop компонент
    dd_result = drag_drop_mapping_ui(indicators, excel_columns, current_mapping)
    
    if st.button("💾 Сохранить изменения (drag-and-drop)", type="primary", key="save_dd"):
        if dd_result:
            column_mapping_df = column_mapping_df[column_mapping_df['table'] != selected_table]
            new_rows = []
            for ind, col in dd_result.items():
                if col != "❌ Не привязано":
                    new_rows.append({'table': selected_table, 'indicator': ind, 'column_from_excel': col})
            if new_rows:
                column_mapping_df = pd.concat([column_mapping_df, pd.DataFrame(new_rows)], ignore_index=True)
            save_column_mapping(column_mapping_df)
            st.cache_data.clear()
            st.rerun()

# ─── ТАБ 3: Массовые операции ──────────────────────────
with tab3:
    st.subheader("⚡ Массовые операции")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🔍 Автоопределение")
        st.caption("Автоматически находит соответствия по названиям показателей и колонок")
        
        threshold = st.slider("Порог схожести (%)", 30, 100, 60, 
                            help="Минимальный процент совпадения для автоматической привязки")
        
        if st.button("🤖 Запустить автоопределение", use_container_width=True):
            auto_mapping = auto_determine_mapping(indicators, excel_columns)
            
            # Фильтруем по порогу
            for ind in indicators:
                if ind in auto_mapping:
                    best_match, score = process.extractOne(ind, excel_columns, scorer=fuzz.token_sort_ratio)
                    if score >= threshold:
                        current_mapping[ind] = best_match
                    else:
                        current_mapping[ind] = "❌ Не привязано"
            
            st.success("✅ Автоопределение выполнено!")
            st.info("Перейдите на вкладку «Табличный режим» чтобы проверить и сохранить результат")
    
    with col2:
        st.markdown("### 🔄 Массовая привязка")
        st.caption("Привязать все похожие показатели одним кликом")
        
        # Находим непривязанные показатели
        unmapped_indicators = [ind for ind, col in current_mapping.items() if col == "❌ Не привязано"]
        
        if unmapped_indicators:
            st.write(f"Не привязано: **{len(unmapped_indicators)}** показателей")
            
            # Показываем лучшие предположения
            st.markdown("#### Лучшие предположения:")
            suggestions = []
            for ind in unmapped_indicators:
                best_match, score = process.extractOne(ind, excel_columns, scorer=fuzz.token_sort_ratio)
                suggestions.append({
                    'Показатель': ind,
                    'Предлагаемая колонка': best_match,
                    'Схожесть': f"{score}%"
                })
            
            suggestions_df = pd.DataFrame(suggestions)
            st.dataframe(suggestions_df, use_container_width=True)
            
            if st.button("⚡ Привязать все предложенные (схожесть > 60%)", use_container_width=True):
                for ind in unmapped_indicators:
                    best_match, score = process.extractOne(ind, excel_columns, scorer=fuzz.token_sort_ratio)
                    if score >= 60:
                        current_mapping[ind] = best_match
                
                # Сохраняем сразу
                column_mapping_df = column_mapping_df[column_mapping_df['table'] != selected_table]
                new_rows = []
                for ind, col in current_mapping.items():
                    if col != "❌ Не привязано":
                        new_rows.append({'table': selected_table, 'indicator': ind, 'column_from_excel': col})
                if new_rows:
                    column_mapping_df = pd.concat([column_mapping_df, pd.DataFrame(new_rows)], ignore_index=True)
                save_column_mapping(column_mapping_df)
                st.cache_data.clear()
                st.success("✅ Массовая привязка сохранена!")
                st.rerun()
        else:
            st.success("✅ Все показатели привязаны!")
    
    # Разделитель
    st.divider()
    
    # Очистка
    st.markdown("### 🧹 Очистка привязок")
    if st.button("🗑️ Отвязать все показатели этой таблицы", type="secondary", use_container_width=True):
        column_mapping_df = column_mapping_df[column_mapping_df['table'] != selected_table]
        save_column_mapping(column_mapping_df)
        st.cache_data.clear()
        st.warning("Все привязки для этой таблицы удалены")
        st.rerun()

# ─── Футер ──────────────────────────────────────────────
st.divider()
st.caption("""
**Режимы:**  
📋 **Табличный** — классический выбор из выпадающих списков  
🖱️ **Drag-and-Drop** — перетаскивание показателей мышкой  
⚡ **Массовые операции** — автоопределение и групповая привязка
""")