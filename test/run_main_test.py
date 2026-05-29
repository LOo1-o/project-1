#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test/run_main_test.py

Полный тест основного конвейера с отчетом о результатах.
"""

import sys
from pathlib import Path

# Добавляем корневую директорию в путь
sys.path.insert(0, str(Path(__file__).parent.parent))

import main as main_module


def run_test():
    """Запускает основной конвейер в режиме тестирования."""
    
    print("\n" + "="*80)
    print("🧪 ТЕСТ: Полный конвейер обработки (main.py)")
    print("="*80)
    
    # Проверяем входные файлы
    input_file = Path("input/Бюллетень_17.2.8 раздел 1-8.docx")
    if not input_file.exists():
        print(f"❌ Входной файл не найден: {input_file}")
        return 1
    
    print(f"\n📄 Входной файл: {input_file}")
    print(f"   Размер: {input_file.stat().st_size / (1024*1024):.1f} MB")
    
    try:
        # Запускаем основной конвейер
        print("\n▶️  Запускаем конвейер...")
        main_module.main()
        
        print("\n" + "="*80)
        print("✅ ТЕСТ УСПЕШНО ЗАВЕРШЁН")
        print("="*80)
        
        # Проверяем выходные файлы
        output_dir = Path("output")
        output_files = [
            "Бюллетень_ШАБЛОН_С_ТЕГАМИ_v2.docx",
            "Бюллетень_17.2.8_ГОТОВЫЙ.docx",
            "fill_log.txt",
        ]
        
        print("\n📁 Выходные файлы:")
        for file_name in output_files:
            file_path = output_dir / file_name
            if file_path.exists():
                size = file_path.stat().st_size / (1024*1024)
                print(f"   ✅ {file_name} ({size:.1f} MB)")
            else:
                print(f"   ⚠️  {file_name} - не создан")
        
        return 0
    
    except Exception as e:
        print("\n" + "="*80)
        print(f"❌ ТЕСТ ПРОВАЛЕН")
        print("="*80)
        print(f"\nОшибка: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit_code = run_test()
    sys.exit(exit_code)
