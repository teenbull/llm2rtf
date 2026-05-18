"""
llm2rtf - Конвертер LaTeX/Markdown в RTF.
Философия: Код должен быть простым, быстрым и без зависимостей (Tsoding-style).
Главный трюк (v2.0): Мы используем нативные поля уравнений Microsoft Word (EQ fields). 
Вместо построения сложных деревьев (AST), мы парсим математику "изнутри-наружу": 
регулярные выражения сворачивают самые глубокие формулы в промежуточные маркеры \x01...\x02, 
сбрасывая внутренние маркеры "на лету". В конце все это превращается в один общий 
EQ field для всего блока. Это защищает от крашей Word'а и делает код компактным.
"""
import re
import sys
import os
import subprocess

def get_list_separator():
    if sys.platform == "win32":
        try:
            import ctypes
            buffer = ctypes.create_unicode_buffer(4)
            ctypes.windll.kernel32.GetLocaleInfoW(0x0400, 0x000C, buffer, 4)
            return buffer.value or ";"
        except Exception: pass
    return ";"

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

    # 1. Очистка от визуального мусора LaTeX и форматирование блоков
    text = re.sub(r'\\(?:left|right)\b\s*', '', text)
    text = re.sub(r'\\q?quad\b', ' ', text)

    # Обработка матриц и систем уравнений "изнутри наружу"
    sep = get_list_separator()
    prev_text = None
    while text != prev_text:
        prev_text = text
        def env_replacer(m):
            env, content = m.group(1), m.group(2)
            content = content.strip()
            if 'matrix' in env:
                # Создаем матрицу EQ \a. Считаем колонки по первой строке
                cols = content.split(r'\\')[0].count('&') + 1
                items = [c.strip() for c in re.split(r'\\\\|&', content)]
                return f"\x01\\b\\bc\\[(\\a\\ac\\co{cols}(" + sep.join(items) + "))\x02"
            elif 'cases' in env:
                # Системы уравнений объединяем одной левой скобкой (\lc\{)
                items = [c.strip() for c in re.split(r'\\\\', content)]
                return f"\x01\\b\\lc\\{{(\\a\\al\\co1(" + sep.join(items) + "))\x02"
            else:
                # align, equation - просто убираем окружение
                return re.sub(r'\\\\(?:\s*\n)?', '\n', content).replace('&', ' ')
        # Ищем самые глубокие блоки (без \begin внутри), чтобы матрицы внутри align обрабатывались первыми
        text = re.sub(r'\\begin\{([a-zA-Z*]+)\}((?:(?!\\begin\{).)*?)\\end\{\1\}', env_replacer, text, flags=re.DOTALL)

    # Добиваем оставшийся мусор
    text = re.sub(r'\\(?:begin|end)\{[^}]+\}', '', text)
    text = re.sub(r'\\\\(?:\s*\n)?', '\n', text)
    text = text.replace('&', ' ')

    # Снимаем экранирование с %, $, _ которые LLM ставит для Markdown
    text = re.sub(r'\\([%$_])', r'\1', text)

    # 2. Словарь констант и символов
    replacements = {
        # Греческие буквы (добавлены недостающие)
        r'alpha': 'α', r'beta': 'β', r'gamma': 'γ', r'delta': 'δ', r'epsilon': 'ε',
        r'zeta': 'ζ', r'eta': 'η', r'theta': 'θ', r'iota': 'ι', r'kappa': 'κ', r'lambda': 'λ',
        r'mu': 'μ', r'nu': 'ν', r'xi': 'ξ', r'rho': 'ρ', r'sigma': 'σ', r'tau': 'τ',
        r'upsilon': 'υ', r'phi': 'φ', r'chi': 'χ', r'psi': 'ψ', r'omega': 'ω',
        r'Delta': 'Δ', r'Gamma': 'Γ', r'Theta': 'Θ', r'Lambda': 'Λ', r'Xi': 'Ξ',
        r'Pi': 'Π', r'pi': 'π', r'Sigma': 'Σ', r'Phi': 'Φ', r'Psi': 'Ψ', r'Omega': 'Ω',
        # Базовые дроби
        r'frac\{1\}\{2\}': '½', r'frac\{1\}\{3\}': '⅓', r'frac\{2\}\{3\}': '⅔',
        r'frac\{1\}\{4\}': '¼', r'frac\{3\}\{4\}': '¾', r'frac\{1\}\{5\}': '⅕',
        # Геометрия и стрелки
        r'triangle': '△', r'angle': '∠', r'perp': '⊥',
        r'Rightarrow': '⇒', r'rightarrow': '→', r'to': '→', r'Leftarrow': '⇐', r'Leftrightarrow': '⇔',
        # Операции
        r'cdot': '⋅', r'times': '×', r'div': '÷', r'pm': '±', r'mp': '∓',
        # Отношения и множества
        r'ge': '≥', r'geq': '≥', r'le': '≤', r'leq': '≤', r'neq': '≠', r'equiv': '≡',
        r'approx': '≈', r'in': '∈', r'notin': '∉', r'subset': '⊂', r'cup': '∪', r'cap': '∩',
        r'emptyset': '∅', r'forall': '∀', r'exists': '∃',
        # Матанализ и прочее
        r'infty': '∞', r'circ': '°', r'int': '∫', r'sum': '∑', r'prod': '∏',
        # Текстовые функции
        r'sin': 'sin', r'cos': 'cos', r'tan': 'tan', r'cot': 'cot',
        r'arcsin': 'arcsin', r'arccos': 'arccos', r'arctan': 'arctan',
        r'ln': 'ln', r'log': 'log', r'lim': 'lim', r'max': 'max', r'min': 'min'
    }
    for pattern, repl in replacements.items():
        # (?<![a-zA-Z]) защищает от замены внутри других слов (s\in -> s∈)
        # (?![a-zA-Z]) разрешает замену, если дальше идет _, ^, ( - то есть не буква. Это чинит \int_a^b!
        text = re.sub(rf'(?<![a-zA-Z])\\?{pattern}(?![a-zA-Z])', repl, text)

    # 3. Степени и градусы
    text = re.sub(r'\^\{?(?:\\?circ|°)\}?', '°', text)
    text = re.sub(r'(\d+)\^([^\w\d]|$)', r'\1°\2', text)
    
    # Умный трюк: переводим простые числа в степенях в Unicode.
    # Это решает баг вложенности (e^{x^2}), так как RTF не умеет делать двойной \super.
    # Букву 'n' мы исключили из конвертации, чтобы не получить 'квадратики' в Word (шрифты плохо поддерживают ₙ)
    sup_map = str.maketrans("0123456789+-=()", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾")
    sub_map = str.maketrans("0123456789+-=()", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎")
    text = re.sub(r'\^\{([0-9+\-=()]+)\}', lambda m: m.group(1).translate(sup_map), text)
    text = re.sub(r'_\{([0-9+\-=()]+)\}', lambda m: m.group(1).translate(sub_map), text)
    text = re.sub(r'\^([0-9])', lambda m: m.group(1).translate(sup_map), text)
    text = re.sub(r'_([0-9])', lambda m: m.group(1).translate(sub_map), text)

    # Приводим оставшиеся индексы (переменные) к формату ^{x} для парсинга в RTF
    text = re.sub(r'\^([a-zA-Z0-9А-Яа-яα-ωΑ-Ω∞°])', r'^{\1}', text)
    text = re.sub(r'_([a-zA-Z0-9А-Яа-яα-ωΑ-Ω∞°])', r'_{\1}', text)

    # Корни без скобок (например \sqrt3 -> \sqrt{3})
    text = re.sub(r'\\?sqrt\s*(\d+)', r'\\sqrt{\1}', text)
    # Одиночные корни без скобок превращаем в символ (только если после них нет { или [)
    text = re.sub(r'\\?sqrt(?!\s*(?:\{|\[))', '√', text)
    
    # Удаление служебных директив, мешающих парсингу
    text = re.sub(r'\\(?:limits|displaystyle)\b\s*', '', text)

    # Убираем лишние пробелы вокруг знаков умножения (по просьбе UI)
    text = re.sub(r'\s*⋅\s*', '⋅', text)
    text = re.sub(r'\s*×\s*', '×', text)

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

        # 3. Парсинг математики (дроби, корни) изнутри-наружу через маркеры \x01 \x02
    # Идея (трюк): вместо создания вложенных RTF полей {\field...}, мы сворачиваем
    # внутренние формулы в \f(a;b) и удаляем у них маркеры, собирая всё в один общий EQ field.
    sep = get_list_separator()
    nb = r'(?:(?!\\[{}]).)*?'  # Жестко запрещаем экранированные скобки \{ и \} внутри аргументов
    
    def strip_m(s): return s.replace('\x01', '').replace('\x02', '')

    prev = None
    while escaped != prev:
        prev = escaped
        # Дроби \frac{A}{B} -> \x01\f(A;B)\x02
        escaped = re.sub(r'\\\\frac\s*\\\{(' + nb + r')\\\}\s*\\\{(' + nb + r')\\\}', 
                         lambda m: f"\x01\\\\f({strip_m(m.group(1))}{sep}{strip_m(m.group(2))})\x02", 
                         escaped)
        # Корни с индексом \sqrt[A]{B}
        escaped = re.sub(r'\\\\sqrt\s*\[(' + nb + r')\]\s*\\\{(' + nb + r')\\\}', 
                         lambda m: f"\x01\\\\r({strip_m(m.group(1))}{sep}{strip_m(m.group(2))})\x02", 
                         escaped)
        # Обычные корни \sqrt{A}
        escaped = re.sub(r'\\\\sqrt\s*\\\{(' + nb + r')\\\}', 
                         lambda m: f"\x01\\\\r({sep}{strip_m(m.group(1))})\x02", 
                         escaped)
        # Степени и индексы (RTF форматирование работает прямо внутри EQ полей!)
        escaped = re.sub(r'\^\s*\\\{(' + nb + r')\\\}', r'{\\super \1}', escaped)
        escaped = re.sub(r'_\s*\\\{(' + nb + r')\\\}', r'{\\sub \1}', escaped)
        # Очистка шрифтовых тегов
        escaped = re.sub(r'\\\\(?:mathrm|text)\s*\\\{(' + nb + r')\\\}', r'\1', escaped)
        escaped = re.sub(r'\\\\(?:mathbf|textbf)\s*\\\{(' + nb + r')\\\}', r'\\b \1\\b0 ', escaped)
        
    # Если маркеры вложены (например, дробь внутри матрицы), снимаем внутренние,
    # чтобы не спровоцировать ошибку вложенных EQ полей в Word
    while re.search(r'\x01[^\x02]*\x01', escaped):
        escaped = re.sub(r'(\x01[^\x01\x02]*)\x01([^\x02]*)\x02', r'\1\2', escaped)
        
    # Превращаем маркеры в полноценные поля EQ
    escaped = escaped.replace('\x01', r'{\field{\*\fldinst EQ ').replace('\x02', r'}{\fldrslt}}')

    # 4. Формулы $...$ превращаем в курсив
    escaped = re.sub(r'\$+(.*?)\$+', r'\\i \1\\i0 ', escaped)

    # 5. Убираем оставшиеся "бесхозные" фигурные скобки (например, от \sin{\beta})
    # Делаем это только в самом конце, чтобы не сломать парсинг вложенных структур!
    prev = None
    while escaped != prev:
        prev = escaped
        escaped = re.sub(r'\\\{(' + nb + r')\\\}', r'\1', escaped)

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

    # Заголовок RTF с таблицей шрифтов (защита от кракозябр в старых Word)
    rtf_header = r"{\rtf1\ansi\ansicpg1251\deff0{\fonttbl{\f0 Times New Roman;}{\f1 Arial;}}\uc1\f0\fs24 "
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