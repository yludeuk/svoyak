"""Патч для parse_docx.py: добавляет пост-обработку для нумерации тем.

Логика: внутри одного source_file границы тем определяются по стоимости
вопросов. Если текущий вопрос имеет price=10, а предыдущий имел price=50
(или это первый вопрос) — это начало новой темы.

К названию темы добавляется префикс с порядковым номером: "1. Яблоки",
"2. Пока секрет", "3. Пока секрет" и т.д.
"""
with open("parse_docx.py", "r", encoding="utf-8") as f:
    content = f.read()

# Добавляем функцию пост-обработки перед process_package
old = '''def process_package(pkg: dict, model: str) -> list:'''

new = '''def renumber_themes_by_price(questions: list) -> list:
    """Пост-обработка: добавляет порядковый номер к названию темы.

    Границы тем определяются по стоимости вопросов внутри одного source_file:
    если текущий вопрос имеет price=10, а предыдущий имел price=50 (или это
    первый вопрос) — это начало новой темы.

    К названию темы добавляется префикс: "1. Яблоки", "2. Пока секрет" и т.д.
    Это решает проблему неуникальных названий (несколько "Пока секрет" в пакете).
    """
    if not questions:
        return questions

    # Группируем по source_file, сохраняя порядок
    by_source = OrderedDict()
    for q in questions:
        sf = q.get("_source_file", "?")
        by_source.setdefault(sf, []).append(q)

    result = []
    for sf, qs in by_source.items():
        theme_counter = 0
        prev_price = None
        for q in qs:
            price = q.get("price")
            # Новая тема начинается, если:
            # - это первый вопрос
            # - текущий price=10, а предыдущий был 50
            # - текущий price=None (нет цены)
            is_new_theme = (
                prev_price is None
                or price == 10
                or price is None
                or (prev_price is not None and prev_price >= 50 and price < prev_price)
            )
            if is_new_theme:
                theme_counter += 1

            theme = q.get("theme")
            if theme is not None and theme != "":
                # Добавляем порядковый номер к названию темы
                q["theme"] = f"{theme_counter}. {theme}"

            prev_price = price
            result.append(q)

    return result


def process_package(pkg: dict, model: str) -> list:'''

if old not in content:
    print("ОШИБКА: process_package не найден")
    raise SystemExit(1)

content = content.replace(old, new)

# Добавляем импорт OrderedDict
old2 = '''import json
import sys
import os
import time

import requests'''

new2 = '''import json
import sys
import os
import time
from collections import OrderedDict

import requests'''

if old2 not in content:
    print("ОШИБКА: импорты не найдены")
    raise SystemExit(1)

content = content.replace(old2, new2)

# Вставляем вызов пост-обработки после получения questions
old3 = '''    questions = process_docx_files(input_files, model, name, prompt_hash)

    # Сохраняем промежуточный JSON'''

new3 = '''    questions = process_docx_files(input_files, model, name, prompt_hash)

    # Пост-обработка: нумерация тем по границам (50→10)
    print(f"\\n[*] Пост-обработка: нумерация тем...")
    questions = renumber_themes_by_price(questions)

    # Сохраняем промежуточный JSON'''

if old3 not in content:
    print("ОШИБКА: блок после process_docx_files не найден")
    raise SystemExit(1)

content = content.replace(old3, new3)

with open("parse_docx.py", "w", encoding="utf-8") as f:
    f.write(content)

print("OK: пост-обработка добавлена")
