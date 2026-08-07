"""Парсер структуры тем и вопросов из txt файлов."""
import os
import re
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

INPUT_DIR = "input/Проба пера_txt"

FILE_ORDER = [
    "Круг 1.txt",
    "Круг 2.txt",
    "Круг 3.txt",
    "Круг 4.txt",
    "Круг 5.txt",
    "Круг 6.txt",
    "Круг 7.txt",
    "Бонус 1. Совковый.txt",
    "Бонус 2. Наркоманский.txt",
]


def read_file(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def parse_file(path):
    """Парсит один файл и возвращает список тем с вопросами."""
    text = read_file(path)
    lines = text.split("\n")

    file_title = lines[0].strip() if lines else ""

    # Стратегия: ищем все строки вида "N. ..." где N от 1 до 8
    # Затем определяем, является ли это заголовком темы или началом вопроса
    # по наличию "Форма:" или "Ответ:" в следующих 10 строках

    candidates = []  # (line_idx, num, text)
    for i in range(1, len(lines)):
        line = lines[i].strip()
        m = re.match(r"^(\d+)\.\s+(.+)$", line)
        if m:
            n = int(m.group(1))
            if 1 <= n <= 8:
                candidates.append((i, n, m.group(2).strip()))

    # Для каждого кандидата определяем, тема это или вопрос
    # Вопрос: после него в течение 15 строк идёт "Форма:" или "Ответ:"
    # Тема: после него идёт описание темы или сразу следующий кандидат

    items = []  # (line_idx, type, num, text) где type = 'theme' или 'question'
    for idx, (line_idx, n, text) in enumerate(candidates):
        # Проверяем, есть ли в следующих 15 строках "Форма:" или "Ответ:"
        is_question = False
        for j in range(line_idx + 1, min(line_idx + 20, len(lines))):
            stripped = lines[j].strip()
            if stripped.startswith("Форма:") or stripped.startswith("Ответ:"):
                is_question = True
                break
            # Если встретили следующий кандидат - это точно не вопрос
            if j > line_idx + 1:
                next_m = re.match(r"^(\d+)\.\s+(.+)$", stripped)
                if next_m:
                    next_n = int(next_m.group(1))
                    if 1 <= next_n <= 8:
                        break

        if is_question:
            items.append((line_idx, 'question', n, text))
        else:
            items.append((line_idx, 'theme', n, text))

    # Теперь группируем: темы содержат вопросы до следующей темы
    themes = []
    current_theme = None
    current_question = None
    current_field = None
    in_theme_comment = False

    for item_idx, (line_idx, item_type, n, text) in enumerate(items):
        if item_type == 'theme':
            # Сохраняем предыдущую тему
            if current_theme:
                if current_question:
                    current_theme["questions"].append(current_question)
                themes.append(current_theme)
            current_theme = {
                "number": n,
                "name": text,
                "comment": None,
                "questions": []
            }
            current_question = None
            current_field = None
            in_theme_comment = True
            continue

        if item_type == 'question':
            in_theme_comment = False
            # Сохраняем предыдущий вопрос
            if current_question:
                current_theme["questions"].append(current_question)
            current_question = {
                "number": n,
                "cost": n * 10,
                "text": text,
                "form": None,
                "answer": None,
                "zachet": None,
                "nezachet": None,
                "comment": None,
                "source": None
            }
            current_field = "question"
            continue

    # Определяем границы для последней темы
    if items:
        last_line_idx = items[-1][0]
    else:
        last_line_idx = 0

    # Теперь обрабатываем содержимое между элементами
    # Для каждого элемента определяем диапазон строк
    for item_idx, (line_idx, item_type, n, text) in enumerate(items):
        # Определяем конец диапазона
        if item_idx + 1 < len(items):
            end_idx = items[item_idx + 1][0]
        else:
            end_idx = len(lines)

        if item_type == 'theme':
            # Комментарий к теме - это строки после заголовка до первого вопроса
            # Найдём первый вопрос этой темы
            first_question_idx = None
            for j in range(item_idx + 1, len(items)):
                if items[j][1] == 'question':
                    first_question_idx = items[j][0]
                    break
                if items[j][1] == 'theme':
                    break

            if first_question_idx is not None:
                # Комментарий - строки между line_idx+1 и first_question_idx
                comment_lines = []
                for j in range(line_idx + 1, first_question_idx):
                    stripped = lines[j].strip()
                    if stripped:
                        comment_lines.append(stripped)
                if comment_lines:
                    current_theme["comment"] = " ".join(comment_lines)

        elif item_type == 'question':
            # Обрабатываем поля вопроса
            for j in range(line_idx + 1, end_idx):
                stripped = lines[j].strip()
                if not stripped:
                    continue

                if stripped.startswith("Форма:"):
                    current_field = "form"
                    current_question["form"] = stripped[6:].strip()
                elif stripped.startswith("Ответ:"):
                    current_field = "answer"
                    current_question["answer"] = stripped[6:].strip()
                elif stripped.startswith("Зачёт:"):
                    current_field = "zachet"
                    current_question["zachet"] = stripped[5:].strip()
                elif stripped.startswith("Незачёт:"):
                    current_field = "nezachet"
                    current_question["nezachet"] = stripped[8:].strip()
                elif stripped.startswith("Комментарий"):
                    current_field = "comment"
                    m = re.match(r"^Комментарий(?:\s+\d+)?:\s*(.*)$", stripped)
                    if m:
                        current_question["comment"] = m.group(1).strip()
                    else:
                        current_question["comment"] = ""
                elif stripped.startswith("Источник"):
                    current_field = "source"
                    m = re.match(r"^Источник(?:и)?:\s*(.*)$", stripped)
                    if m:
                        current_question["source"] = m.group(1).strip()
                    else:
                        current_question["source"] = ""
                else:
                    if current_field == "question":
                        current_question["text"] += " " + stripped
                    elif current_field == "form":
                        current_question["form"] += " " + stripped
                    elif current_field == "answer":
                        current_question["answer"] += " " + stripped
                    elif current_field == "zachet":
                        current_question["zachet"] += " " + stripped
                    elif current_field == "nezachet":
                        current_question["nezachet"] += " " + stripped
                    elif current_field == "comment":
                        if current_question["comment"]:
                            current_question["comment"] += " " + stripped
                        else:
                            current_question["comment"] = stripped
                    elif current_field == "source":
                        if current_question["source"]:
                            current_question["source"] += " " + stripped
                        else:
                            current_question["source"] = stripped

    # Сохраняем последнюю тему
    if current_theme:
        if current_question:
            current_theme["questions"].append(current_question)
        themes.append(current_theme)

    return file_title, themes


def main():
    all_data = []
    for fname in FILE_ORDER:
        path = os.path.join(INPUT_DIR, fname)
        if not os.path.exists(path):
            print(f"Файл не найден: {path}")
            continue
        file_title, themes = parse_file(path)
        print(f"\n=== {fname} ({file_title}) ===")
        print(f"Тем: {len(themes)}")
        for t in themes:
            print(f"  Тема {t['number']}: {t['name']} ({len(t['questions'])} вопросов)")
            for q in t['questions']:
                ans = q['answer'][:50] if q['answer'] else "НЕТ ОТВЕТА"
                print(f"    Вопрос {q['number']} ({q['cost']}): {ans}...")
        all_data.append({
            "file": fname,
            "title": file_title,
            "themes": themes
        })

    with open("parsed_data.json", "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    print(f"\n\nСохранено в parsed_data.json")

    total_themes = sum(len(d["themes"]) for d in all_data)
    total_questions = sum(len(t["questions"]) for d in all_data for t in d["themes"])
    print(f"Всего тем: {total_themes}")
    print(f"Всего вопросов: {total_questions}")


if __name__ == "__main__":
    main()
