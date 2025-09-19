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
    m = re.match(r'^(\d+)[\.\s]', line.strip())
    if not m:
        return False
    price = int(m.group(1))
    return price % 10 == 0 and 10 <= price <= 10000


def is_service_line(line):
    return any(word in line for word in ['Источник', 'Комментарий', 'Зачет', 'Незачет', 'http'])


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
        if i > idx and re.match(r'^\s*\d+[\.\s]', line):
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
    while i < end_idx and not re.match(r'^\s*\d+[\.\s]', lines[i]):
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
            m = re.match(r'^\s*(\d+)[\.\s]', theme_lines[i])
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

    while i < len(lines):
        block = []
        j = i
        scanned = 0
        # сканируем вперед, ищем вопросы, без жёсткой маленькой "высоты"
        while j < len(lines) and scanned < FORWARD_SCAN_LIMIT:
            if is_question_line(lines[j]) and has_answer_nearby(lines, j):
                m = re.match(r'^\s*(\d+)', lines[j])
                if m:
                    block.append({'line': j, 'price': int(m.group(1))})
            # если достаточно вопросов — можно остановиться раннее
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

            pos = first_q_line - 1
            # накопитель для фрагментов комментария, которые ближе к вопросам
            tail_comments = []

            while pos >= 0:
                ln = lines[pos].rstrip('\n')
                stripped = ln.strip()

                # пропускаем пустые строки
                if stripped == "":
                    pos -= 1
                    continue

                # пропускаем строки "Форма:" — они относятся к вопросу, а не к теме
                if stripped.lower().startswith('форма:'):
                    pos -= 1
                    continue

                # если это parenthetical блок (скобки) или явный "Комментарий:", соберём его как часть comment
                # и НИКОГДА не включаем в комментарий строку с названием темы выше.
                if (stripped.startswith('(') and stripped.endswith(')')) or stripped.lower().startswith('комментарий:'):
                    # 1) Случай: однострочные скобки "(...)" — возьмём только эту строку как комментарий
                    if stripped.startswith('(') and stripped.endswith(')'):
                        paren_block = [lines[pos].rstrip('\n')]
                        pos -= 1
                        tail_comments = paren_block + tail_comments
                        continue

                    # 2) Случай: конец многострочного скобочного блока (строка может заканчиваться ')')
                    if stripped.endswith(')'):
                        paren_block = [lines[pos].rstrip('\n')]
                        pos -= 1
                        # поднимаемся вверх, собирая только те строки, которые являются частью скобочного блока
                        # (останавливаемся, если встречаем строку, начинающуюся с "(", или видим пустую строку,
                        # или наткнулись на "Форма:" / заголовок темы / номер вопроса)
                        while pos >= 0:
                            up = lines[pos].strip()
                            if up == "" or up.lower().startswith('форма:') or re.match(r'^\s*Тема\b', up, re.IGNORECASE) or re.match(r'^\s*\d+[\.\)]', up):
                                break
                            paren_block.append(lines[pos].rstrip('\n'))
                            if up.startswith('('):
                                pos -= 1
                                break
                            pos -= 1
                        paren_block.reverse()
                        tail_comments = paren_block + tail_comments
                        continue

                    # 3) Случай: блок, начинающийся со слова "Комментарий:" — собираем только строки этого блока,
                    #    но прекратим, если выше идёт явный заголовок/пустая строка/Форма/номер вопроса
                    if stripped.lower().startswith('комментарий:'):
                        com_block = [lines[pos].rstrip('\n')]
                        pos -= 1
                        while pos >= 0:
                            up = lines[pos].strip()
                            if up == "" or up.lower().startswith('форма:') or re.match(r'^\s*Тема\b', up, re.IGNORECASE) or re.match(r'^\s*\d+[\.\)]', up):
                                break
                            com_block.append(lines[pos].rstrip('\n'))
                            pos -= 1
                        com_block.reverse()
                        tail_comments = com_block + tail_comments
                        continue

                # иначе — найден потенциальный заголовок
                title_idx = pos
                title_text = clean_theme_name(stripped) or stripped

                # соберём строки между этим заголовком и первым вопросом, исключая 'Форма:' и служебные строки
                add_comments = []
                cp = pos + 1
                while cp < first_q_line:
                    s = lines[cp].strip()
                    if s == "" or s.lower().startswith('форма:') or is_service_line(s):
                        cp += 1
                        continue
                    add_comments.append(lines[cp].rstrip('\n'))
                    cp += 1

                # итоговый комментарий — сначала строки ниже заголовка (если они есть), затем tail_comments
                if add_comments or tail_comments:
                    theme_comment = "\n".join(add_comments + tail_comments)
                else:
                    theme_comment = ""

                break  # мы нашли заголовок — выходим

            # если не нашли ничего — fallback
            if title_text is None:
                title_text = f"Тема {num_theme}"
                title_idx = max(0, first_q_line - 1)

            # комментарий к теме — все непустые строки между title_idx и первым вопросом
            # (исключаем служебные строки и строки "Форма:")
            comment_lines = []
            for ci in range(title_idx + 1, first_q_line):
                ln = lines[ci].strip()
                if not ln:
                    continue
                # не включаем 'Форма:' в комментарий темы — это форма вопроса, она относится к вопросу
                if ln.lower().startswith('форма:'):
                    continue
                if is_service_line(ln):
                    continue
                comment_lines.append(ln)
            theme_comment = "\n".join(comment_lines) if comment_lines else ""

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
                if cleaned_candidate and len(ln_t.split()) <= 8:
                    end_idx = t2
                    break

            themes.append({
                'num': num_theme,
                'name': title_text,
                'start': title_idx,
                'end': end_idx,
                'comment': theme_comment
            })
            num_theme += 1
            i = last_q_line + 1
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
def random_prefix(length=4):
    return ''.join(random.choices(string.ascii_lowercase, k=length))

def save_to_docx(block_texts, prefix=None):
    prefix = prefix or random_prefix()
    for i, text in enumerate(block_texts, 1):
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
        doc.save(f"{prefix}_{i}.docx")


def main():
    lines = read_input_file('umlaut-2025.txt')
    themes = extract_theme_candidates(lines)
    blocks = split_themes(themes, user_split=[10]*8)
    theme_counter = 1
    output_blocks = []
    for block in blocks:
        block_texts = []
        for theme in block:
            block_texts.append(process_theme(lines, theme, theme_counter))
            theme_counter += 1
        output_blocks.append('\n\n'.join(block_texts))
    save_to_docx(output_blocks)


if __name__ == '__main__':
    main()
