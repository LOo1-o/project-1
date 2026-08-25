#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test/run_analysis.py

Тестирует анализ таблиц перед запуском основного процесса.
Используется для проверки структуры документа и выявления проблем.
"""

import sys
from pathlib import Path

# Добавляем корневую директорию в путь
sys.path.insert(0, str(Path(__file__).parent.parent))

from analyze_table_tags import analyze_document_structure, print_analysis, generate_detailed_report


def main():
    """Запускает анализ документа."""
    
    doc_path = Path("input/Бюллетень.docx")
    
    print("\n" + "="*80)
    print("🧪 ТЕСТ: Анализ заполнения тегов в документе")
    print("="*80)
    
    if not doc_path.exists():
        print(f"❌ Файл не найден: {doc_path}")
        print("   Убедитесь, что файл находится в папке input/")
        return 1
    
    print(f"\n📄 Анализируем: {doc_path}")
    print(f"   Размер: {doc_path.stat().st_size / (1024*1024):.1f} MB")
    
    try:
        # Анализируем
        analysis = analyze_document_structure(str(doc_path))
        
        # Выводим результаты
        print_analysis(analysis)
        
        # Сохраняем отчет
        report_path = "output/analysis_tag_coverage.txt"
        generate_detailed_report(analysis, report_path)
        print(f"\n✅ Отчет сохранен: {report_path}")
        
        # Возвращаем результат
        summary = analysis['summary']
        if summary['tables_without_tags'] > 0:
            print(f"\n⚠️  ВНИМАНИЕ: Обнаружено {summary['tables_without_tags']} таблиц без тегов!")
            print("   Эти таблицы не будут заполняться автоматически.")
        
        if summary['coverage_percent'] < 50:
            print(f"\n🔴 КРИТИЧЕСКОЕ: Общее покрытие только {summary['coverage_percent']}%!")
            print("   Большинство таблиц не подготовлены к заполнению.")
            return 1
        
        return 0
    
    except Exception as e:
        print(f"\n❌ Ошибка при анализе: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
