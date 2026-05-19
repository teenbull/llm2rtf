"""
Минималистичная система тестирования llm2word (Snapshot Testing).
Tsoding-style: без зависимостей, быстро, наглядно.

Использование:
  python test.py          -> Обычный прогон (сравнение с эталонами)
  python test.py --update -> Перезаписать эталоны (если мы улучшили логику генератора)
"""
import os
import re
import sys
import time

from llm2word import clean_math_text, generate_rtf

def run_tests():
    update_mode = '--update' in sys.argv
    start_time = time.time()
    
    test_file = "test.md"
    if not os.path.exists(test_file):
        print(f"[!] Ошибка: {test_file} не найден.")
        sys.exit(1)
        
    with open(test_file, 'r', encoding='utf-8') as f:
        content = f.read()

    output_dir = "tests_output"
    golden_dir = "tests_golden" # Папка с идеальными эталонами
    
    for d in [output_dir, golden_dir]:
        if not os.path.exists(d):
            os.makedirs(d)

    tasks = re.split(r'(?=###\s+Задача)', content)
    passed = 0
    failed = 0
    
    print(f"[*] Запуск автотестов {'(РЕЖИМ ОБНОВЛЕНИЯ ЭТАЛОНОВ)' if update_mode else '(Сравнение)'}...")

    # Добавим общий тест в конец списка задач
    tasks.append("### Задача FULL: Общий прогон всего файла\n" + content)

    for i, task_content in enumerate(tasks):
        if not task_content.strip() or "Задача" not in task_content:
            continue
            
        title_match = re.search(r'###\s+(Задача [a-zA-Z0-9_]+:.*?)\n', task_content)
        title = title_match.group(1).strip() if title_match else f"Блок {i}"
        safe_name = re.sub(r'[^\w\s-]', '', title).strip().replace(' ', '_')
        
        try:
            # 1. Генерируем новый RTF
            cleaned = clean_math_text(task_content)
            rtf_code = generate_rtf(cleaned)
            
            # 2. Базовые проверки на целостность (Smoke)
            assert rtf_code.strip() != "", "Пустой RTF"
            assert '\x01' not in rtf_code and '\x02' not in rtf_code, "Остались маркеры \x01 \x02"
            
            out_path = os.path.join(output_dir, f"{safe_name}.rtf")
            golden_path = os.path.join(golden_dir, f"{safe_name}.rtf")
            
            # Сохраняем текущий выхлоп всегда, чтобы можно было открыть в Word
            with open(out_path, 'w', encoding='ascii', errors='replace') as out_f:
                out_f.write(rtf_code)

            # 3. Snapshot-тестирование
            if update_mode:
                # Если мы чинили баги и хотим сохранить результат как новый идеал
                with open(golden_path, 'w', encoding='ascii', errors='replace') as gf:
                    gf.write(rtf_code)
                print(f"  [~] {title} -> ЭТАЛОН ОБНОВЛЕН")
                passed += 1
            else:
                # Обычный режим: сравниваем с идеалом
                if not os.path.exists(golden_path):
                    print(f"  [?] {title} -> ПРОПУСК (Нет эталона. Запустите с --update)")
                    continue
                    
                with open(golden_path, 'r', encoding='ascii', errors='replace') as gf:
                    golden_rtf = gf.read()
                    
                if rtf_code == golden_rtf:
                    print(f"  [+] {title} -> Успех (Совпадает с эталоном)")
                    passed += 1
                else:
                    print(f"  [-] {title} -> ПРОВАЛ! (RTF отличается от эталона)")
                    # В будущем тут можно выводить кусочек diff'а
                    failed += 1
                    
        except Exception as e:
            print(f"  [-] {title} -> ПРОВАЛ С ОШИБКОЙ: {str(e)}")
            failed += 1

    dt = time.time() - start_time
    print("-" * 50)
    print(f"[*] Время: {dt:.3f} сек. | Успех: {passed} | Провал: {failed}")
    if failed > 0: sys.exit(1)

if __name__ == "__main__":
    run_tests()