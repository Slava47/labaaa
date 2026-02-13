#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Мастер-скрипт для выполнения всех лабораторных работ 5-10
"""

import os
import sys
import subprocess
from datetime import datetime

def print_header(text):
    """Печать красивого заголовка"""
    print("\n" + "="*80)
    print(f"  {text}")
    print("="*80)

def run_lab(lab_number, script_name, description):
    """Запуск отдельной лабораторной работы"""
    print_header(f"ЛАБОРАТОРНАЯ РАБОТА №{lab_number}: {description}")
    
    try:
        result = subprocess.run(
            ['python3', f'labs/{script_name}'],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode == 0:
            print(result.stdout)
            print(f"\n✅ Лабораторная работа №{lab_number} выполнена успешно!")
            return True
        else:
            print(f"❌ Ошибка при выполнении лабораторной работы №{lab_number}")
            print(result.stderr)
            return False
    except subprocess.TimeoutExpired:
        print(f"❌ Timeout при выполнении лабораторной работы №{lab_number}")
        return False
    except Exception as e:
        print(f"❌ Исключение при выполнении лабораторной работы №{lab_number}: {e}")
        return False

def main():
    """Главная функция"""
    print_header("ВЫПОЛНЕНИЕ ЛАБОРАТОРНЫХ РАБОТ №5-10")
    print(f"Дата и время запуска: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    labs = [
        (5, 'lab5_word_formatting.py', 'Microsoft Word - Форматирование текста'),
        (6, 'lab6_word_tables.py', 'Microsoft Word - Таблицы'),
        (7, 'lab7_part1_salary.py', 'Excel - Расчет зарплаты'),
        (7, 'lab7_part2_graphs.py', 'Excel - Графики функций'),
        (7, 'lab7_part3_sorting.py', 'Excel - Сортировка и фильтры'),
        (8, 'lab8_powerpoint.py', 'PowerPoint - Презентация'),
        (9, 'lab9_graph_problems.py', 'Задачи с графами'),
        (10, 'lab10_game_theory.py', 'Игровые задачи'),
    ]
    
    results = []
    
    for lab_num, script, desc in labs:
        success = run_lab(lab_num, script, desc)
        results.append((lab_num, desc, success))
    
    # Итоговый отчет
    print_header("ИТОГОВЫЙ ОТЧЕТ")
    
    successful = sum(1 for _, _, success in results if success)
    total = len(results)
    
    print(f"\nВыполнено работ: {successful} из {total}")
    print("\nДетальный отчет:")
    
    for lab_num, desc, success in results:
        status = "✅ Успешно" if success else "❌ Ошибка"
        print(f"  Работа №{lab_num} ({desc}): {status}")
    
    # Список созданных файлов
    print("\n" + "="*80)
    print("СОЗДАННЫЕ ФАЙЛЫ:")
    print("="*80)
    
    if os.path.exists('labs'):
        files = sorted(os.listdir('labs'))
        for f in files:
            if f.endswith(('.docx', '.xlsx', '.pptx', '.png', '.txt', '.py')):
                file_path = os.path.join('labs', f)
                size = os.path.getsize(file_path)
                print(f"  📄 {f} ({size:,} bytes)")
    
    print("\n" + "="*80)
    print("🎉 ВСЕ ЛАБОРАТОРНЫЕ РАБОТЫ ВЫПОЛНЕНЫ!")
    print("="*80)
    
    return successful == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
