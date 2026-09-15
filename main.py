# main.py
import sys
from pathlib import Path
from docx import Document

from config import (
    load_okved_map,
    load_table_source_map,
    normalize_dash_bold_in_document,
    autofit_tables_to_window,
)
from config_v2 import (
    load_column_mapping_v2,
    ensure_column_mapping_v2,
    validate_table_source_mapping,
    load_indicator_display_names,
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
    # Внутри exe, собранного PyInstaller-ом, __file__ указывает не туда, где
    # реально лежит .exe, а во временную папку распаковки (sys._MEIPASS) —
    # она удаляется сразу после закрытия программы. Если считать base_dir от
    # неё, папки input/output будут искаться там, где их никогда не положат.
    # sys.frozen выставляется самим PyInstaller именно для этого случая;
    # sys.executable тогда указывает на настоящий .exe-файл на диске.
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).parent.resolve()
    else:
        base_dir = Path(__file__).parent.resolve()
    input_dir = base_dir / "input"
    mappings_dir = input_dir / "mappings"
    excel_dir = input_dir / "excel"
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)

    # 📄 Файлы
    okved_file = mappings_dir / "okved_mapping.csv"
    mo_file = mappings_dir / "mo.csv"
    table_mapping_file = mappings_dir / "table_source_data_mapping.csv"
    column_mapping_file = mappings_dir / "column_mapping_v2.csv"
    unit_annotations_file = mappings_dir / "unit_annotations.csv"

    input_word_file = input_template_override or input_dir / "Бюллетень.docx"
    cleared_word_file = output_dir / "Бюллетень_ОЧИЩЕННЫЙ.docx"
    template_word_file = output_dir / "Бюллетень_ШАБЛОН_С_ТЕГАМИ_v2.docx"
    final_word_file = Path(output_file_override) if output_file_override else output_dir / "Бюллетень_ГОТОВЫЙ.docx"

    # === ШАГ 1: Загрузка справочников ===
    print("\n=== ШАГ 1: Загрузка справочников ===")
    okved_to_name, name_to_okved_cleaned = load_okved_map(okved_file)
    table_mapping = load_table_source_map(table_mapping_file)
    validate_table_source_mapping(table_mapping, excel_dir)

    ensure_column_mapping_v2(excel_dir, table_mapping, column_mapping_file)

    word_to_indicator, indicator_to_excel, indicator_to_file, file_word_to_indicator, indicator_keywords, _ = load_column_mapping_v2(column_mapping_file)
    okved_codes_set = set(okved_to_name.keys())
    print("✅ Справочники успешно загружены.")

    # === ШАГ 1.5: Очистка шаблона от данных прошлого периода ===
    # Сначала убираем старые числа из исходного бюллетеня (кроме шапки с
    # годами) — до того, как на их место встанут теги. Так в репозитории
    # остаётся явный промежуточный "пустой бланк", и специалистам, которые
    # открывают файл между этим шагом и следующим, не приходится гадать,
    # откуда цифры прошлого периода или что означают {{...}}.
    print("\n=== ШАГ 1.5: Очистка шаблона от данных ===")
    generate_word_template(
        input_word_file,
        okved_file,
        table_mapping_file,
        column_mapping_file,
        cleared_word_file,
        mo_map_path=mo_file,
        excel_dir=excel_dir,
        unit_annotations_path=unit_annotations_file,
        clear_only=True,
    )
    print(f"🧹 Очищенный бланк сохранён: {cleared_word_file}")

    # === ШАГ 2: Генерация шаблона Word с тегами ===
    print("\n=== ШАГ 2: Генерация шаблона Word с умными тегами ===")
    generate_word_template(
        cleared_word_file,
        okved_file,
        table_mapping_file,
        column_mapping_file,
        template_word_file,
        mo_map_path=mo_file,
        validation_log_path=output_dir / "mapping_validation_log.txt",
        near_miss_report_path=output_dir / "нужна_проверка.xlsx",
        squish_match_report_path=output_dir / "склеенные_совпадения.xlsx",
        typo_match_report_path=output_dir / "опечатки_в_одну_букву.xlsx",
        excel_dir=excel_dir,
        unit_annotation_report_path=output_dir / "отфильтрованные_разметки_единиц.xlsx",
        unit_annotations_path=unit_annotations_file,
        group_prefix_match_report_path=output_dir / "совпадения_после_отбрасывания_префикса.xlsx",
        unused_indicator_report_path=output_dir / "неиспользованные_показатели.xlsx",
        duplicate_year_report_path=output_dir / "повторяющиеся_года_в_шапке.xlsx",
    )
    print(f"📄 Шаблон с тегами сохранен: {template_word_file}")

    # === ШАГ 3: Предварительная загрузка данных из Excel ===
    print("\n=== ШАГ 3: Предварительная загрузка данных из Excel (УМНЫЙ поиск) ===")
    master_data, stats = pre_load_all_excel_data_v2(
        excel_dir=excel_dir,
        table_source_mapping=table_mapping,
        okved_codes_set=okved_codes_set,
        okved_name_to_code=name_to_okved_cleaned,
        column_mapping_path=column_mapping_file,
        mo_map_path=mo_file,
        use_fuzzy_match=True,
        fuzzy_threshold=0.80,
        errors_report_path=output_dir / "ошибки_загрузки_excel.txt",
        implausible_report_path=output_dir / "подозрительные_значения.xlsx"
    )
    print(f"📊 Статистика загрузки: {stats}")
    print(f"✅ Загружено данных для {len(master_data)} кодов ОКВЭД")

    # === ШАГ 4: Заполнение шаблона по тегам ===
    print("\n=== ШАГ 4: Заполнение шаблона по умным тегам (с нормализацией) ===")
    doc = Document(template_word_file)
    # code показателя -> человекочитаемое название (для отчёта о незаполненных
    # тегах — чтобы его понимал человек, который готовит бюллетень, а не
    # только разработчик).
    indicator_display_names = load_indicator_display_names(column_mapping_file)
    unfilled_tags = fill_word_template_by_tags_v2(
        doc,
        master_data,
        log_path=output_dir / "fill_log.txt",
        report_path=output_dir / "unfilled_tags.xlsx",
        indicator_display_names=indicator_display_names
    )
    fixed_dashes = normalize_dash_bold_in_document(doc)
    if fixed_dashes:
        print(f"🔧 Выровнена жирность {fixed_dashes} прочерков, не совпадавшая с соседями по строке")
    autofit_tables_to_window(doc)
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