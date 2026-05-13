# main.py
from pathlib import Path
from docx import Document

from config import (
    load_okved_map,
    load_table_source_map,
)
from config_v2 import (
    load_column_mapping_v2,
    build_column_mapping_v2_from_excel,
    validate_table_source_mapping,
)
from logic import (
    generate_word_template,
    collect_okved_codes_from_template,
    compare_okved_sets
)
from data_filler_v2 import (
    pre_load_all_excel_data_v2,
    fill_word_template_by_tags_v2
)


def main(input_template_override=None, output_file_override=None):
    print("\n🚀 === НАЧАЛО РАБОТЫ (v2.0 с умными тегами) ===")

    # 📁 Пути
    # base_dir = Path(r"C:\Users\41.Bogatyrevaee\PycharmProjects\pythonProject18")
    base_dir = Path(__file__).parent.resolve()
    input_dir = base_dir / "input"
    mappings_dir = input_dir / "mappings"
    excel_dir = input_dir / "excel"
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)

    # 📄 Файлы
    okved_file = mappings_dir / "okved_mapping.csv"
    table_mapping_file = mappings_dir / "table_source_data_mapping.csv"
    column_mapping_file = mappings_dir / "column_mapping_v2.csv"

    input_word_file = input_template_override or input_dir / "Бюллетень_17.2.8 раздел 1-8.docx"
    template_word_file = output_dir / "Бюллетень_ШАБЛОН_С_ТЕГАМИ_v2.docx"
    final_word_file = Path(output_file_override) if output_file_override else output_dir / "Бюллетень_17.2.8_ГОТОВЫЙ.docx"

    # === ШАГ 1: Загрузка справочников ===
    print("\n=== ШАГ 1: Загрузка справочников ===")
    okved_to_name, name_to_okved_cleaned = load_okved_map(okved_file)
    table_mapping = load_table_source_map(table_mapping_file)
    validate_table_source_mapping(table_mapping, excel_dir)

    if not column_mapping_file.exists():
        print(f"⚠️ Файл {column_mapping_file} не найден. Генерируем column_mapping_v2.csv из Excel...")
        build_column_mapping_v2_from_excel(excel_dir, table_mapping, column_mapping_file)

    word_to_indicator, indicator_to_excel, indicator_to_file, file_word_to_indicator, indicator_keywords = load_column_mapping_v2(column_mapping_file)
    okved_codes_set = set(okved_to_name.keys())
    print("✅ Справочники успешно загружены.")

    # === ШАГ 2: Генерация шаблона Word с тегами ===
    print("\n=== ШАГ 2: Генерация шаблона Word с умными тегами ===")
    generate_word_template(input_word_file, okved_file, table_mapping_file, column_mapping_file, template_word_file)
    print(f"📄 Шаблон с тегами сохранен: {template_word_file}")

    # === ШАГ 3: Предварительная загрузка данных из Excel ===
    print("\n=== ШАГ 3: Предварительная загрузка данных из Excel (УМНЫЙ поиск) ===")
    master_data, stats = pre_load_all_excel_data_v2(
        excel_dir=excel_dir,
        table_source_mapping=table_mapping,
        okved_codes_set=okved_codes_set,
        column_mapping_path=column_mapping_file,
        use_fuzzy_match=True,
        fuzzy_threshold=0.80
    )
    print(f"📊 Статистика загрузки: {stats}")
    print(f"✅ Загружено данных для {len(master_data)} кодов ОКВЭД")

    # === ШАГ 4: Заполнение шаблона по тегам ===
    print("\n=== ШАГ 4: Заполнение шаблона по умным тегам (с нормализацией) ===")
    doc = Document(template_word_file)
    unfilled_tags = fill_word_template_by_tags_v2(
        doc,
        master_data,
        log_path=output_dir / "fill_log.txt",
        report_path=output_dir / "unfilled_tags.xlsx"
    )
    doc.save(final_word_file)
    print(f"📘 Заполненный документ сохранён: {final_word_file}")

    # === ШАГ 5: Диагностика и сверка ===
    print("\n=== ШАГ 5: Диагностика и сверка ===")
    template_codes = collect_okved_codes_from_template(template_word_file)
    excel_codes = set(master_data.keys())
    compare_okved_sets(template_codes, excel_codes)

    # ✅ Финальный отчёт
    print("\n✅ === ПРОЦЕСС ЗАВЕРШЁН ===")
    print(f"📄 Шаблон с тегами: {template_word_file}")
    print(f"📘 Заполненный документ: {final_word_file}")
    if unfilled_tags:
        print(f"🔍 Найдено {len(unfilled_tags)} тегов, для которых не нашлось данных (заменены на '—').")
        print(f"📄 Отчёт: {output_dir / 'unfilled_tags.xlsx'}")
        print(f"📝 Лог: {output_dir / 'fill_log.txt'}")


if __name__ == "__main__":
    main()
