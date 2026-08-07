"""Патч: исправляет двойную нумерацию тем.

Проблема: LLM иногда возвращает тему уже с номером ("1. Смех."),
а пост-обработка добавляет ещё один номер, получается "1. 1. Смех.".

Решение: проверяем, начинается ли тема с числа и точки. Если да — не добавляем номер.
"""
import re

with open("parse_docx.py", "r", encoding="utf-8") as f:
    content = f.read()

old = '''            theme = q.get("theme")
            if theme is not None and theme != "":
                # Добавляем порядковый номер к названию темы
                q["theme"] = f"{theme_counter}. {theme}"

            prev_price = price
            result.append(q)'''

new = '''            theme = q.get("theme")
            if theme is not None and theme != "":
                # Проверяем, не начинается ли тема уже с номера ("1. Смех.")
                if re.match(r"^\\d+\\.\\s+", theme):
                    # Тема уже пронумерована — не добавляем номер
                    pass
                else:
                    # Добавляем порядковый номер к названию темы
                    q["theme"] = f"{theme_counter}. {theme}"

            prev_price = price
            result.append(q)'''

if old not in content:
    print("ОШИБКА: блок не найден")
    raise SystemExit(1)

content = content.replace(old, new)

with open("parse_docx.py", "w", encoding="utf-8") as f:
    f.write(content)

print("OK: двойная нумерация исправлена")
