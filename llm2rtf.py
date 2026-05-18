import re
import sys
import os
import subprocess

def get_clipboard_text():
    if sys.platform == "win32":
        import ctypes
        user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
        if not user32.OpenClipboard(0): return ""
        try:
            if not user32.IsClipboardFormatAvailable(13): return "" # 13 = CF_UNICODETEXT
            handle = user32.GetClipboardData(13)
            ptr = kernel32.GlobalLock(handle)
            text = ctypes.c_wchar_p(ptr).value
            kernel32.GlobalUnlock(handle)
            return text or ""
        finally:
            user32.CloseClipboard()
    elif sys.platform == "darwin":
        return subprocess.run(['pbpaste'], capture_output=True, text=True).stdout
    else:
        # Пытаемся использовать Wayland или X11
        for cmd in [['wl-paste'], ['xclip', '-selection', 'clipboard', '-o'], ['xsel', '--clipboard', '--output']]:
            try: return subprocess.run(cmd, capture_output=True, text=True).stdout
            except FileNotFoundError: continue
        return ""

def clean_math_text(text):
    if not text: return ""
    
    # 0. Убиваем скрытые Windows-переносы каретки (\r)
    text = text.replace('\r', '')

    # 1. Очистка от визуального мусора LaTeX
    text = re.sub(r'\\(?:left|right)\b\s*', '', text)
    text = re.sub(r'\\(?:begin|end)\{[^}]+\}', '', text)
    text = re.sub(r'\\q?quad\b', ' ', text)

    # Снимаем экранирование с %, $, & и _, которые LLM ставит для Markdown
    text = re.sub(r'\\([%$&_])', r'\1', text)

    # 2. Словарь констант и символов
    replacements = {
        # Греческие буквы
        r'\\?alpha\b': 'α', r'\\?beta\b': 'β', r'\\?gamma\b': 'γ', r'\\?delta\b': 'δ', r'\\?epsilon\b': 'ε',
        r'\\?Delta\b': 'Δ', r'\\?Gamma\b': 'Γ', r'\\?Theta\b': 'Θ', r'\\?Lambda\b': 'Λ', r'\\?Pi\b': 'Π', r'\\?pi\b': 'π',
        # Геометрия и стрелки
        r'\\?triangle\b': '△', r'\\?angle\b': '∠', r'\\?perp\b': '⊥',
        r'\\?Rightarrow\b': '⇒', r'\\?rightarrow\b': '→', r'\\?Leftarrow\b': '⇐', r'\\?Leftrightarrow\b': '⇔',
        # Операции
        r'\\?cdot\b': '⋅', r'\\?times\b': '×', r'\\?div\b': '÷', r'\\?pm\b': '±', r'\\?mp\b': '∓',
        # Отношения и множества
        r'\\?ge\b': '≥', r'\\?geq\b': '≥', r'\\?le\b': '≤', r'\\?leq\b': '≤', r'\\?neq\b': '≠', r'\\?equiv\b': '≡',
        r'\\?approx\b': '≈', r'\\?in\b': '∈', r'\\?notin\b': '∉', r'\\?subset\b': '⊂', r'\\?cup\b': '∪', r'\\?cap\b': '∩',
        r'\\?emptyset\b': '∅', r'\\?forall\b': '∀', r'\\?exists\b': '∃',
        # Матанализ и прочее
        r'\\?infty\b': '∞', r'\\?circ\b': '°', r'\\?int\b': '∫', r'\\?sum\b': '∑', r'\\?prod\b': '∏',
        # Текстовые функции
        r'\\?sin\b': 'sin', r'\\?cos\b': 'cos', r'\\?tan\b': 'tan', r'\\?cot\b': 'cot',
        r'\\?arcsin\b': 'arcsin', r'\\?arccos\b': 'arccos', r'\\?arctan\b': 'arctan',
        r'\\?ln\b': 'ln', r'\\?log\b': 'log', r'\\?lim\b': 'lim', r'\\?max\b': 'max', r'\\?min\b': 'min'
    }
    for pattern, repl in replacements.items():
        text = re.sub(pattern, repl, text)

    # 3. Степени и градусы
    # ^\circ или ^{°} это обычный градус
    text = re.sub(r'\^\{?(?:\\?circ|°)\}?', '°', text)
    # Исправление частой опечатки учеников/ИИ: "120^ " -> "120° "
    text = re.sub(r'(\d+)\^([^\w\d]|$)', r'\1°\2', text)
    
    # Транслятор любых цифровых степеней в Unicode (x^2 -> x², 10^{12} -> 10¹²)
    superscript_map = str.maketrans("0123456789+-=()n", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ")
    def super_replacer(match):
        return match.group(1).translate(superscript_map)
    text = re.sub(r'\^\{([0-9+\-=()n]+)\}', super_replacer, text) # Для ^{12}
    text = re.sub(r'\^([0-9n])', super_replacer, text)            # Для ^2

    # --- Начало блока "изнутри наружу" ---
    # Общий цикл для парсинга вложенных структур.
    # Выедаем самые глубокие команды без фигурных скобок на каждом проходе.
    # Гарантирует правильный парсинг любой вложенности (\mathbf{\frac{1}{2}}).
    def sqrt_replacer(match):
        inner = match.group(1)
        # Если внутри корня есть сложение/вычитание или пробел, берем в скобки для надежности
        if any(c in inner for c in ['+', '-', '=', ' ']): return f"√({inner})"
        return f"√{inner}"

    prev_text = None
    while text != prev_text:
        prev_text = text
        # 4. Корни (вложенные)
        text = re.sub(r'\\?sqrt\{([^{}]*)\}', sqrt_replacer, text)

        # 5. Жирный шрифт и обычный текст (внутри математики)
        # Используем обычные строки с '\\1' вместо raw-строк, чтобы обмануть тупой линтер IDE
        text = re.sub(r'\\?(?:mathbf|textbf)\{([^{}]*)\}', '**\\1**', text)
        text = re.sub(r'\\?(?:mathrm|text)\{([^{}]*)\}', '\\1', text)

        # 6. Дроби
        text = re.sub(r'\\?frac\{([^{}]*)\}\{([^{}]*)\}', '(\\1)/(\\2)', text)

    # Корни без скобок (например \sqrt3)
    text = re.sub(r'\\?sqrt\s*(\d+)', r'√\1', text)
    text = re.sub(r'\\?sqrt\b', '√', text)
    # --- конец блока ---

    # 7. Убираем лишние скобки вокруг "атомарных" выражений (числа, переменные, корни).
    # Не удаляем скобки, если внутри есть умножение или минус, 
    # так как это может изменить приоритет операций при делении.
    # Безопасно удаляем для: корней, степеней, десятичных дробей и градусов.
    text = re.sub(r'\(([A-Za-z0-9А-Яа-яα-ωΑ-Ω√⁰¹²³⁴⁵⁶⁷⁸⁹°.,]+)\)', r'\1', text)

    # 8. Формулы $...$ превращаем в Markdown-курсив *...*
    text = re.sub(r'\$(.*?)\$', r'*\1*', text)

    # 9. Финальная зачистка
    text = text.replace('\\', '').replace('{}', '')

    return text.strip()

def generate_rtf(text):
    escaped = ""
    for char in text:
        if char in ('\\', '{', '}'): 
            escaped += f"\\{char}"
        elif ord(char) > 127:
            code = ord(char)
            # RTF требует 16-битные signed int. Обходим краш Word'а на символах > 32767
            if 32767 < code <= 65535:
                code -= 65536
            elif code > 65535:
                code = 63 # Заглушка '?' для эмодзи
            escaped += f"\\u{code}?"
        else: 
            escaped += char

    # 1. Заголовки Markdown
    escaped = re.sub(r'^###\s+(.*?)$', r'\\b\\fs28 \1\\b0\\fs24 ', escaped, flags=re.MULTILINE)
    escaped = re.sub(r'^##\s+(.*?)$', r'\\b\\fs32 \1\\b0\\fs24 ', escaped, flags=re.MULTILINE)
    escaped = re.sub(r'^#\s+(.*?)$', r'\\b\\fs36 \1\\b0\\fs24 ', escaped, flags=re.MULTILINE)

    # 2. Жирный шрифт и курсив
    escaped = re.sub(r'\*\*(.*?)\*\*', r'\\b \1\\b0 ', escaped)
    escaped = re.sub(r'\*(.*?)\*', r'\\i \1\\i0 ', escaped)

    # 3. Таблицы Markdown
    def rtf_table_replacer(match):
        lines = match.group(0).strip().split('\n')
        rtf_out = []
        for line in lines:
            # Пропускаем Markdown-разделитель (например, |---|---|)
            if re.match(r'^[ \t]*\|(?:[-: ]+\|)+[ \t]*$', line):
                continue
            
            # Парсим ячейки, удаляя крайние пайпы
            cells = [c.strip() for c in line.strip().strip('|').split('|')]
            if not cells: continue
            
            num_cols = len(cells)
            # 9000 твипов — это примерно 16 см (на всю ширину листа А4)
            cell_width = 9000 // num_cols 
            
            # \trowd = начало строки таблицы
            row_def = "\\trowd\\trgaph108\\trleft-108 "
            for i in range(num_cols):
                # Рисуем рамки: верх (t), лево (l), низ (b), право (r) и задаем ширину ячейки (\cellx)
                row_def += f"\\clbrdrt\\brdrs\\brdrw10 \\clbrdrl\\brdrs\\brdrw10 \\clbrdrb\\brdrs\\brdrw10 \\clbrdrr\\brdrs\\brdrw10 \\cellx{(i+1)*cell_width} "
            
            # Собираем контент ячеек и закрываем строку (\row)
            row_content = " ".join(f"{c} \\cell" for c in cells)
            rtf_out.append(f"{row_def} {row_content} \\row")
        
        # Оборачиваем таблицу в \pard (сброс абзаца). 
        # Внутри нет символов \n, чтобы следующий шаг не сломал RTF-код.
        # Обязательно добавляем \n в конце, так как регулярка "съела" его при поиске!
        return r" \pard " + " ".join(rtf_out) + r" \pard " + "\n"

    # Ищем блоки, где от 2 строк и больше начинаются и заканчиваются на пайп "|"
    escaped = re.sub(r'(?:^[ \t]*\|.*\|[ \t]*(?:\n|$)){2,}', rtf_table_replacer, escaped, flags=re.MULTILINE)

    # 4. Переносы строк
    escaped = escaped.replace('\n', r' \par ' + '\n')

    # Заголовок RTF (Times / Arial 12pt)
    rtf_header = r"{\rtf1\ansi\ansicpg1251\uc1\f0\fs24 "
    return rtf_header + "\n" + escaped + "\n}"

def set_clipboard(plain_text, rtf_text):
    if sys.platform == "win32":
        import ctypes
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        CF_RTF = user32.RegisterClipboardFormatW("Rich Text Format")
        CF_UNICODETEXT = 13

        user32.OpenClipboard(0)
        user32.EmptyClipboard()

        rtf_bytes = rtf_text.encode('ascii')
        hGlobalRtf = kernel32.GlobalAlloc(0x0002, len(rtf_bytes) + 1)
        pRtf = kernel32.GlobalLock(hGlobalRtf)
        ctypes.memmove(pRtf, rtf_bytes, len(rtf_bytes))
        kernel32.GlobalUnlock(hGlobalRtf)
        user32.SetClipboardData(CF_RTF, hGlobalRtf)

        plain_bytes = (plain_text + '\0').encode('utf-16-le')
        hGlobalPlain = kernel32.GlobalAlloc(0x0002, len(plain_bytes))
        pPlain = kernel32.GlobalLock(hGlobalPlain)
        ctypes.memmove(pPlain, plain_bytes, len(plain_bytes))
        kernel32.GlobalUnlock(hGlobalPlain)
        user32.SetClipboardData(CF_UNICODETEXT, hGlobalPlain)

        user32.CloseClipboard()
    else:
        # Fallback для Mac / Linux: кладем обычный текст в буфер через терминал
        if sys.platform == "darwin":
            subprocess.run(['pbcopy'], input=plain_text, text=True)
        else:
            for cmd in [['wl-copy'], ['xclip', '-selection', 'clipboard'], ['xsel', '--clipboard', '--input']]:
                try:
                    subprocess.run(cmd, input=plain_text, text=True)
                    break
                except FileNotFoundError: continue
        
        # RTF открываем как файл, так как не все X11 менеджеры умеют держать RTF в буфере
        filename = "converted_math.rtf"
        with open(filename, 'w', encoding='ascii') as f:
            f.write(rtf_text)
        os.system(f'open "{filename}"' if sys.platform == "darwin" else f'xdg-open "{filename}"')

def main():
    raw_text = get_clipboard_text()
    if not raw_text.strip():
        print("Ошибка: Буфер обмена пуст или не содержит текста.")
        return

    raw_kb = len(raw_text.encode('utf-8')) / 1024

    clean_text = clean_math_text(raw_text)
    rtf_code = generate_rtf(clean_text)
    set_clipboard(clean_text, rtf_code)

    # RTF состоит только из ASCII, энкодинг в utf-8 здесь избыточен
    rtf_kb = len(rtf_code) / 1024

    print("Текст сконвертирован и находится в буфере обмена!")
    print(f"   Было:  {raw_kb:.2f} КБ")
    print(f"   Стало: {rtf_kb:.2f} КБ (RTF)")
    print("   -> Зайдите в Word/Wordpad и нажмите [Ctrl+V]")

if __name__ == "__main__":
    main()