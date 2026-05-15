from pathlib import Path
from typing import Dict

from table_registry import TableRegistry


def validate_table_registry(table_source_mapping_path: Path, excel_dir: Path) -> Dict:
    """Выполняет проверку маппинга таблиц и наличия Excel-файлов."""
    registry = TableRegistry(table_source_mapping_path)
    errors, warnings = registry.validate()

    excel_dir = Path(excel_dir)
    available_files = {p.name for p in excel_dir.glob('*.xlsx')}
    mapped_files = {table_data['source'] for table_data in registry.get_all().values()}

    missing_files = sorted(mapped_files - available_files)
    unused_files = sorted(available_files - mapped_files)

    if missing_files:
        errors.append(
            f"Отсутствуют Excel-файлы из маппинга: {', '.join(missing_files)}"
        )

    if unused_files:
        warnings.append(
            f"В папке input/excel есть файлы, не указанные в маппинге: {', '.join(unused_files[:20])}"
        )
        if len(unused_files) > 20:
            warnings.append(f"...и еще {len(unused_files) - 20} файлов.")

    return {
        'errors': errors,
        'warnings': warnings,
        'missing_files': missing_files,
        'unused_files': unused_files,
        'registry': registry,
    }


def print_validation_report(validation_result: Dict) -> None:
    """Выводит результат валидации в консоль."""
    errors = validation_result.get('errors', [])
    warnings = validation_result.get('warnings', [])
    missing_files = validation_result.get('missing_files', [])
    unused_files = validation_result.get('unused_files', [])

    if errors:
        print("\n❌ ОШИБКИ ВАЛИДАЦИИ:")
        for error in errors:
            print(f"   - {error}")

    if warnings:
        print("\n⚠️ ПРЕДУПРЕЖДЕНИЯ ВАЛИДАЦИИ:")
        for warning in warnings:
            print(f"   - {warning}")

    if not errors and not warnings:
        print("\n✅ Валидация маппинга и реестра прошла успешно.")
