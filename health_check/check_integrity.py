#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
health_check/check_integrity.py

Проверяет целостность системы перед запуском main.py.
Убедитесь, что все необходимые файлы и папки на месте.
"""

import sys
from pathlib import Path


def check_directories():
    """Проверяет наличие необходимых директорий."""
    required_dirs = [
        "input",
        "input/mappings",
        "input/excel",
        "output",
    ]
    
    print("📁 Проверка директорий...")
    missing = []
    
    for dir_name in required_dirs:
        dir_path = Path(dir_name)
        if dir_path.exists():
            print(f"   ✅ {dir_name}")
        else:
            print(f"   ❌ {dir_name} - НЕ НАЙДЕНА")
            missing.append(dir_name)
    
    return len(missing) == 0, missing


def check_files():
    """Проверяет наличие необходимых конфигурационных файлов."""
    required_files = [
        "input/mappings/okved_mapping.csv",
        "input/mappings/table_source_data_mapping.csv",
        "input/mappings/column_mapping_v2.csv",
        "input/Бюллетень.docx",
    ]
    
    print("\n📄 Проверка файлов...")
    missing = []
    
    for file_name in required_files:
        file_path = Path(file_name)
        if file_path.exists():
            size = file_path.stat().st_size
            print(f"   ✅ {file_name} ({size} bytes)")
        else:
            print(f"   ❌ {file_name} - НЕ НАЙДЕН")
            missing.append(file_name)
    
    return len(missing) == 0, missing


def check_python_imports():
    """Проверяет наличие необходимых Python пакетов."""
    required_modules = [
        "docx",
        "pandas",
        "openpyxl",
    ]
    
    print("\n📦 Проверка Python пакетов...")
    missing = []
    
    for module_name in required_modules:
        try:
            __import__(module_name)
            print(f"   ✅ {module_name}")
        except ImportError:
            print(f"   ❌ {module_name} - НЕ УСТАНОВЛЕН")
            missing.append(module_name)
    
    return len(missing) == 0, missing


def check_input_files():
    """Проверяет наличие Excel файлов в input/excel/"""
    print("\n📊 Проверка Excel файлов в input/excel/...")
    excel_dir = Path("input/excel")
    
    if not excel_dir.exists():
        print("   ❌ Папка input/excel не существует")
        return False
    
    excel_files = list(excel_dir.glob("*.xlsx"))
    
    if not excel_files:
        print("   ⚠️  В папке input/excel нет Excel файлов!")
        return False
    
    for file in excel_files[:10]:
        size = file.stat().st_size / (1024 * 1024)  # MB
        print(f"   ✅ {file.name} ({size:.1f} MB)")
    
    if len(excel_files) > 10:
        print(f"   ... и еще {len(excel_files) - 10} файлов")
    
    return True


def main():
    """Главная функция проверки."""
    print("\n" + "="*80)
    print("🏥 HEALTH CHECK - ПРОВЕРКА ЦЕЛОСТНОСТИ СИСТЕМЫ")
    print("="*80)
    
    all_ok = True
    
    # Проверяем директории
    dirs_ok, missing_dirs = check_directories()
    if not dirs_ok:
        all_ok = False
        print(f"   ⚠️  Отсутствуют директории: {missing_dirs}")
    
    # Проверяем файлы
    files_ok, missing_files = check_files()
    if not files_ok:
        all_ok = False
        print(f"   ⚠️  Отсутствуют файлы: {missing_files}")
    
    # Проверяем импорты
    imports_ok, missing_modules = check_python_imports()
    if not imports_ok:
        all_ok = False
        print(f"\n   ⚠️  Не установлены модули: {missing_modules}")
        print("   Установите: pip install -r requirements.txt")
    
    # Проверяем Excel файлы
    excel_ok = check_input_files()
    if not excel_ok:
        all_ok = False
    
    print("\n" + "="*80)
    if all_ok:
        print("✅ ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ - СИСТЕМА ГОТОВА К ЗАПУСКУ!")
        return 0
    else:
        print("❌ ОБНАРУЖЕНЫ ПРОБЛЕМЫ - ИСПРавьте ИХ ПЕРЕД ЗАПУСКОМ main.py")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
