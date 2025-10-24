import re
import random
import string
from collections import Counter
from itertools import combinations
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

EXTRA_SERVICE_WORDS = {
    'Раунд', 'Тема', 'Полуоткрытый', 'Открытый', 'Закрытый', 'Блок',
    'Четвертьфинал', 'Полуфинал', 'Финал', 'Автор', 'Авторы', 'Редакторы'
}


def read_input_file(path):
    with open(path, encoding='utf-8') as f:
        return f.read().splitlines()


def clean_theme_name(line):
    s = line.strip()
    if re.match(r'^(?:Ответ|Комментарий|Источник|Форма|Зачет|Зачёт|http|www)\b', s, re.IGNORECASE):
        return None
    m = re.match(r'^(?:Тема(?:[-\s]*\d+)?|\d+)[\.:–\-\s]*(.+)$', s, re.IGNORECASE)
    title = m.group(1).strip() if m else s
    if not title or title.lower() in (w.lower() for w in EXTRA_SERVICE_WORDS):
        return None
    return title


def is_question_line(line):
    """
    Распознаёт строку, начинающую вопрос:
    - "10. ..."     (обычный)
    - "10 ..."      (иногда)
    - "Вопрос 10:"  / "Вопрос: 10" / "Вопрос за 10" / "Вопрос №10"
    Возвращает True только если найдено число-стоимость, кратное 10 в разумном диапазоне.
    """
    if not line:
        return False
    s = line.strip()
    # 1) стандартный случай: число в начале: "10." или "10 "
    m = re.match(r'^\s*(\d+)[\.\)\s]', s)
    if m:
        try:
            price = int(m.group(1))
        except:
            return False
        return price % 10 == 0 and 10 <= price <= 10000

    # 2) варианты с словом "Вопрос" (включая "Вопрос за", "Вопрос №", "Вопрос:")
    m2 = re.match(r'^\s*вопрос\b(?:\s+за)?[\s\:\-\—\–]*#?\s*(\d+)', s, re.IGNORECASE)
    if m2:
        try:
            price = int(m2.group(1))
        except:
            return False
        return price % 10 == 0 and 10 <= price <= 10000

    # 3) общая запасная эвристика: "Вопрос ... N" (если есть)
    m3 = re.match(r'^\s*вопрос[^\d]*(\d+)', s, re.IGNORECASE)
    if m3:
        try:
            price = int(m3.group(1))
        except:
            return False
        return price % 10 == 0 and 10 <= price <= 10000

    return False


def is_service_line(line):
    L = line.lower()
    return any(k in L for k in ['источник', 'комментарий', 'зачет', 'незачет', 'http'])


def is_probably_question(line):
    if not is_question_line(line):
        return False
    content = re.sub(r'^\s*\d+[\.\s]', '', line).strip()
    words = content.split()
    if not words:
        return False
    first, second = words[0], words[1] if len(words) > 1 else ""
    if is_service_line(content):
        return False
    if first.istitle() and not content.endswith('?') and len(words) < 4:
        return False
    return True


def has_answer_nearby(lines, idx, search_limit=10):
    pattern = re.compile(r'(?:^|\s)Ответ\s*[:\-–]', re.IGNORECASE)
    for i in range(idx, min(len(lines), idx + 1 + search_limit)):
        line = lines[i].strip()
        if i > idx and is_question_line(lines[i]):
            return False
        if pattern.search(line):
            return True
    return False


def normalize_prices(prices):
    sorted_prices = sorted(set(prices))[:5]
    mapping = {old: new for old, new in zip(sorted_prices, [10, 20, 30, 40, 50])}
    return mapping


def extract_answer_block(lines, start_idx, end_idx):
    answer_idx = None
    pattern = re.compile(r'(?:^|\s)Ответ\s*[:\-–]', re.IGNORECASE)
    for i in range(start_idx, end_idx):
        if pattern.search(lines[i]):
            answer_idx = i
            break
    if answer_idx is None:
        return ("", "")
    q_lines = [lines[j] for j in range(start_idx, answer_idx)]
    a_lines = [lines[answer_idx]]
    i = answer_idx + 1
    while i < end_idx and not is_question_line(lines[i]):
        if lines[i].strip() == '':
            break
        a_lines.append(lines[i])
        i += 1
    q_text = ' '.join(l.strip() for l in q_lines if l.strip())
    a_text = '\n'.join(l for l in a_lines if l.strip())
    return (q_text, a_text)


def score_question_block(questions):
    prices = [q['price'] for q in questions]
    if len(prices) < 2:
        return 0
    diffs = [prices[i] - prices[i - 1] for i in range(1, len(prices))]
    step = diffs[0]
    penalty = sum(abs(d - step) for d in diffs)
    return 100 - penalty - abs(len(prices) - 5) * 10


def extract_questions(theme_lines):
    questions = []
    i = 0
    while i < len(theme_lines):
        if is_question_line(theme_lines[i]):
            # Попробуем извлечь число несколькими способами
            m = re.match(r'^\s*(\d+)[\.\)\s]', theme_lines[i])
            if not m:
                m = re.match(r'^\s*вопрос\b(?:\s+за)?[\s\:\-\—\–]*#?\s*(\d+)', theme_lines[i], re.IGNORECASE)
            if not m:
                m = re.match(r'^\s*вопрос[^\d]*(\d+)', theme_lines[i], re.IGNORECASE)
            if not m:
                i += 1
                continue
            price = int(m.group(1))
            start = i
            i += 1
            while i < len(theme_lines) and not is_question_line(theme_lines[i]):
                i += 1
            questions.append({'price': price, 'start': start, 'end': i})
        else:
            i += 1
    return questions


def extract_theme_candidates(lines):
    """
    Поиск кандидатов на темы в тексте.
    Логика:
      - сканируем вниз от позиции i пока не соберём достаточно вопросов (>=4) или не выйдем за разумный предел;
      - выше первого найденного вопроса собираем несколько предыдущих не-пустых строк (candidate lines);
      - выбираем заголовок по приоритету:
          1) ближайшая строка, содержащая явный маркер 'Тема' или нумерованный заголовок (N. Название с N не кратным 10)
          2) ближайшая непарентетическая, не-служебная строка (т.е. обычный текст)
          3) fallback: "Тема N"
      - всё, что между выбранным заголовком и первым вопросом — считается комментарием темы (comment) и сохраняется.
    """
    themes = []
    i = 0
    num_theme = 1
    # сколько строк максимум просканируем вперёд при попытке собрать один блок вопросов
    FORWARD_SCAN_LIMIT = 200
    # сколько непустых строк вверх от первого вопроса возьмём в кандидаты заголовка
    BACK_CANDIDATES = 40

    end_idx = -1
    while i < len(lines):
        block = []
        j = i
        scanned = 0
        # сканируем вперед, ищем вопросы, без жёсткой маленькой "высоты"
        while j < len(lines) and scanned < FORWARD_SCAN_LIMIT:
            if is_question_line(lines[j]) and has_answer_nearby(lines, j):
                m = re.match(r'^\s*(\d+)', lines[j])
                if not m:
                    m = re.match(r'^\s*вопрос\b(?:\s+за)?[\s\:\-\—\–]*#?\s*(\d+)', lines[j], re.IGNORECASE)
                if not m:
                    m = re.match(r'^\s*вопрос[^\d]*(\d+)', lines[j], re.IGNORECASE)
                if m:
                    block.append({'line': j, 'price': int(m.group(1))})
            # если достаточно вопросов — можно остановиться ранее
            if len(block) >= 5:
                break
            j += 1
            scanned += 1

        # если найдено достаточное количество вопросов — считаем это темой-кандидатом
        if len(block) >= 4:
            first_q_line = block[0]['line']

            # Новый, более надёжный выбор названия темы и комментария:
            # Сканируем вверх от первого вопроса, пропуская:
            #  - пустые строки,
            #  - строки, начинающиеся с "Форма:" (они относятся к вопросам),
            #  - собираем блочные комментарии (скобочные или начинающиеся с "Комментарий:")
            #  и затем берём первую подходящую строку над ними как заголовок.
            title_idx = None
            title_text = None
            theme_comment = ""

            # --- НОВАЯ, надёжная логика выбора названия темы и комментария ---
            # 1) двигаемся вверх от первого вопроса, пропуская подряд пустые строки и строки 'Форма:' прямо перед вопросом
            pos = first_q_line - 1
            while pos >= end_idx and (lines[pos].strip() == "" or lines[pos].strip().lower().startswith('форма:')):
                pos -= 1

            # 2) соберём подряд идущие (соседние) непустые строки вверх — это блок (сверху->вниз после reverse)
            block_start = pos
            block_lines = []
            while block_start >= max(0, end_idx) and lines[block_start].strip() != "":
                block_lines.append(lines[block_start].rstrip('\n'))
                block_start -= 1
            block_lines.reverse()
            # теперь block_start указывает на строку выше блока (пустая или -1)

            title_idx = None
            title_text = None
            theme_comment = ""

            def _block_text(bl):
                return "\n".join([ln.strip() for ln in bl]).strip()

            if block_lines:
                block_text = _block_text(block_lines)
                first_line = block_lines[0].strip()

                # CASE B/C: если блок — комментарий в скобках или начинается со слова "Комментарий:"
                if (block_text.startswith("(") and block_text.endswith(")")) or first_line.lower().startswith(
                        'комментарий:'):
                    theme_comment = block_text

                    # если над скобками может быть пустая строка — поднимемся через неё, чтобы найти название темы
                    prev = block_start
                    while prev >= max(0, end_idx) and lines[prev].strip() == "":
                        prev -= 1

                    title_idx = None
                    title_text = None

                    # если над блоком есть непустая строка, попробуем её интерпретировать
                    if prev >= max(0, end_idx):
                        cand_line = lines[prev].strip()
                        cand_name = clean_theme_name(cand_line)
                        if cand_name:
                            title_idx = prev
                            title_text = cand_name
                        else:
                            # немного поднимемся выше (до BACK_CANDIDATES)
                            scan_up = prev - 1
                            scanned = 0
                            while scan_up >= max(0, end_idx) and scanned < BACK_CANDIDATES:
                                if lines[scan_up].strip():
                                    cand2 = lines[scan_up].strip()
                                    cand2_name = clean_theme_name(cand2)
                                    if cand2_name:
                                        title_idx = scan_up
                                        title_text = cand2_name
                                        break
                                scan_up -= 1
                                scanned += 1
                            # fallback — хотя бы возьмём строку prev как заголовок
                            if title_idx is None:
                                title_idx = prev
                                title_text = clean_theme_name(lines[prev].strip()) or lines[prev].strip()
                    else:
                        # нет непустых строк выше — fallback
                        title_text = f"Тема {num_theme}"
                        title_idx = max(0, first_q_line - 1)
                else:
                    # CASE A: обычный случай — первая строка блока = заголовок, остальные — комментарий
                    title_text = clean_theme_name(first_line) or first_line
                    # индекс первой строки блока — block_start + 1 (потому что loop отнял 1 лишний шаг)
                    title_idx = block_start + 1
                    # собираем коммент (остальные строки блока), исключая строки 'Форма:' и служебные метки
                    if len(block_lines) > 1:
                        comment_list = []
                        for ln in block_lines[1:]:
                            st = ln.strip()
                            if st.lower().startswith('форма:'):
                                continue
                            if is_service_line(st):
                                # сохраняем авторов и пр. (они не входят в is_service_line), но исключим явные 'Комментарий/Источник/Зачет'
                                continue
                            comment_list.append(ln.rstrip('\n'))
                        if comment_list:
                            theme_comment = "\n".join(comment_list)
            else:
                # Ничего подрядого не нашлось — небольшой lookup вверх (без длинного lookback)
                SMALL_LOOKUP = 10
                up = first_q_line - 1
                steps = 0
                found = False
                while up >= max(0, end_idx) and steps < SMALL_LOOKUP:
                    if lines[up].strip() == "" or lines[up].strip().lower().startswith('форма:'):
                        up -= 1
                        steps += 1
                        continue
                    cand = clean_theme_name(lines[up].strip())
                    title_text = cand if cand else lines[up].strip()
                    title_idx = up
                    found = True
                    break
                if not found:
                    title_text = f"Тема {num_theme}"
                    title_idx = max(0, first_q_line - 1)
            # --- конец новой логики ---

            # определяем корректный конец темы: ищем следующий явный заголовок (Тема ...) или
            # числовой заголовок (например "1. Австралия") — тогда останавливаемся прямо перед ним.
            end_idx = len(lines)
            # безопасно определим номер последнего вопроса в block (словаря ожидаются как {'line':..., 'price':...})
            last_q_line = None
            for itm in reversed(block):
                if isinstance(itm, dict) and 'line' in itm:
                    last_q_line = itm['line']
                    break
            # защитный fallback
            if last_q_line is None:
                last_q_line = first_q_line

            lookahead_limit = min(len(lines), last_q_line + FORWARD_SCAN_LIMIT)
            for t2 in range(last_q_line + 1, lookahead_limit):
                ln_t = lines[t2].strip()
                if re.match(r'^\s*Тема\b', ln_t, re.IGNORECASE):
                    end_idx = t2
                    break
                mnum = re.match(r'^\s*(\d+)\.(?:\s+)(.+)$', ln_t)
                if mnum and (int(mnum.group(1)) % 10 != 0):
                    end_idx = t2
                    break
                cleaned_candidate = clean_theme_name(ln_t)
                if cleaned_candidate and len(ln_t.split()) <= 12:
                    end_idx = t2
                    break

            warnings = []

            # Проверка названия
            if not title_text.strip():
                warnings.append(f"⚠️ Тема {num_theme}: пустое название")
            if ":" in title_text:
                warnings.append(f"⚠️ Тема {num_theme}: содержит ':' в названии")
            if title_text.lower().startswith("комментарий"):
                warnings.append(f"⚠️ Тема {num_theme}: начинается со слова 'Комментарий'")
            if title_text.strip().startswith("(") and title_text.strip().endswith(")"):
                warnings.append(f"⚠️ Тема {num_theme}: название заключено в скобки")
            if len(title_text.split()) > 10:
                warnings.append(f"⚠️ Тема {num_theme}: слишком длинное название ({len(title_text.split())} слов)")

            # Проверка вопросов
            prices = [b["price"] for b in block if "price" in b]
            if len(prices) != 5:
                warnings.append(f"⚠️ Тема {num_theme} '{title_text}': найдено {len(prices)} вопросов (ожидалось 5)")

            expected = [10, 20, 30, 40, 50]
            if sorted(prices) != expected:
                warnings.append(f"⚠️ Тема {num_theme} '{title_text}': некорректные стоимости вопросов ({prices})")

            # Сохраняем предупреждения в саму тему
            theme_entry = {
                "num": num_theme,
                "name": title_text.strip(),
                "comment": theme_comment,
                "start": title_idx,
                "end": end_idx,
                "warnings": warnings
            }
            themes.append(theme_entry)
            num_theme += 1
            i = end_idx
        else:
            i += 1

    return themes


def process_theme(lines, theme, theme_num):
    theme_lines = lines[theme['start']+1:theme['end']]
    questions = extract_questions(theme_lines)
    if not questions:
        return ""
    prices = [q['price'] for q in questions]
    mapping = normalize_prices(prices)
    result = [f"{theme_num}. {theme['name']}"]

    # Пересоберём комментарий из оригинальных строк, чтобы иметь доступ к индексам
    # и удалить строки 'Форма:' которые мы перенесём внутрь ответов.
    first_q_abs = theme['start'] + 1 + (questions[0]['start'] if questions else 0)
    comment_abs = []
    for ci in range(theme['start'] + 1, first_q_abs):
        ln = lines[ci].strip()
        if ln and not is_service_line(ln):
            comment_abs.append((ci, ln))

    # Найдём все индексы строк "Форма:" сразу перед вопросами и запомним их
    form_indices = set()
    for q in questions:
        abs_form_idx = theme['start'] + 1 + q['start'] - 1
        if abs_form_idx >= theme['start'] + 1 and abs_form_idx < theme['end']:
            if lines[abs_form_idx].strip().lower().startswith('форма:'):
                form_indices.add(abs_form_idx)

    # построим окончательный текст комментария без тех "Форма:" строк
    comment_lines = [ln for (ci, ln) in comment_abs if ci not in form_indices]
    if comment_lines:
        result.append("\n".join(comment_lines))
    result.append("")

    # Обрабатываем вопросы, переносим при необходимости 'Форма:' в ответы
    for q in questions:
        q_text, a_text = extract_answer_block(theme_lines, q['start'], q['end'])
        q_text = re.sub(r'^\s*\d+[\.\s]', '', q_text).strip()
        if not q_text and not a_text:
            continue

        # проверим, есть ли перед вопросом строка 'Форма:'
        abs_form_idx = theme['start'] + 1 + q['start'] - 1
        form_line = None
        if abs_form_idx >= theme['start'] + 1 and abs_form_idx < theme['end']:
            if lines[abs_form_idx].strip().lower().startswith('форма:'):
                form_line = lines[abs_form_idx].strip()

        norm_price = mapping.get(q['price'], q['price'])
        result.append(f"{norm_price}. {q_text}")
        result.append("")

        if a_text:
            if form_line:
                # Вставляем form_line сразу после строки "Ответ:" в блоке a_text (если там есть)
                a_lines = a_text.split('\n')
                inserted = False
                for idx_ln, ln in enumerate(a_lines):
                    if re.match(r'^\s*Ответ\s*[:\-–]', ln, re.IGNORECASE):
                        # если уже есть форма прямо после — не дублируем
                        if idx_ln + 1 < len(a_lines) and a_lines[idx_ln + 1].strip().lower().startswith('форма:'):
                            inserted = True
                            break
                        a_lines.insert(idx_ln + 1, form_line)
                        inserted = True
                        break
                if not inserted:
                    a_lines.append(form_line)
                result.append('\n'.join(a_lines))
            else:
                result.append(a_text)
        else:
            if form_line:
                result.append(form_line)

        result.append("")

    text = '\n'.join(result)
    # нормализуем подряд идущие пустые строки: не более одной пустой строки между блоками
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


# === Разбиение ===
def even_split(total, min_size=9, max_size=12):
    for blocks in range(total // max_size, total // min_size + 2):
        if blocks == 0:
            continue
        base = total // blocks
        plus = total % blocks
        sizes = [base+1]*plus + [base]*(blocks-plus)
        if all(min_size <= s <= max_size for s in sizes):
            return sorted(sizes)
    sizes = [max_size] * (total // max_size)
    if total % max_size:
        sizes.append(total % max_size)
    return sorted(sizes)

def split_themes(themes, user_split=None, min_size=9, max_size=12):
    total = len(themes)
    sizes = user_split if user_split else even_split(total, min_size, max_size)
    blocks, idx = [], 0
    for sz in sizes:
        blocks.append(themes[idx:idx + sz])
        idx += sz
    return blocks


# === Вывод ===
def random_package_name(length=4):
    return ''.join(random.choices(string.ascii_lowercase, k=length))

def save_to_docx(block_texts, package_name=None, theme_counts=None):
    package_name = package_name or random_package_name()

    # Подсчёт номеров тем внутри блоков
    current_theme = 1
    for i, text in enumerate(block_texts, 1):
        count = theme_counts[i - 1] if theme_counts and i - 1 < len(theme_counts) else text.count('\n\n10.')
        first_theme = current_theme
        last_theme = current_theme + count - 1
        current_theme += count

        # Формируем имя файла
        filename = f"{package_name}-{i}-{first_theme}..{last_theme}.docx"

        # Сохраняем файл
        doc = Document()
        for line in text.split('\n'):
            p = doc.add_paragraph(line)
            run = p.runs[0] if p.runs else p.add_run('')
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.0
        doc.save(filename)
        print(f"✅ Saved {filename}")


def main():
    lines = read_input_file('vuelta-autumn-2024.txt')
    themes = extract_theme_candidates(lines)

    blocks = split_themes(themes, user_split=[9]*2+[10]*5)
    theme_counter = 1
    output_blocks = []

    for i, block in enumerate(blocks, 1):
        block_texts = ['\n'.join([theme['name'] for theme in block])]
        block_warnings = []

        # Собираем текст и предупреждения по темам
        for theme in block:
            block_texts.append(process_theme(lines, theme, theme_counter))
            theme_counter += 1
            if "warnings" in theme and theme["warnings"]:
                block_warnings.extend(theme["warnings"])

        # Сохраняем этот блок в отдельный файл
        first_theme = block[0]["num"]
        last_theme = block[-1]["num"]
        filename = f"Вуэльта.2024.Осень-{i}-{first_theme}..{last_theme}.docx"
        doc = Document()
        for line in '\n\n'.join(block_texts).split('\n'):
            p = doc.add_paragraph(line)
            run = p.runs[0] if p.runs else p.add_run('')
            run.font.name = 'Arial'
            run.font.size = Pt(11)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.0
        doc.save(filename)

        # Выводим отчёт
        if block_warnings:
            print(f"⚠️ Saved {filename} (предупреждений: {len(block_warnings)})")
            for w in block_warnings:
                print("   ", w)
            return
        else:
            print(f"✅ Saved {filename}")


if __name__ == '__main__':
    main()
