"""Исправляет двойную нумерацию тем в JSON файле.

Проблема: LLM иногда возвращает тему уже с номером ("1. Смех."),
а пост-обработка добавляет ещё один номер, получается "1. 1. Смех.".

Этот скрипт убирает дублирующийся номер.
"""
import json
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")


def fix_theme(theme: str) -> str:
    """Убирает дублирующийся номер в начале темы.

    "1. 1. Смех." -> "1. Смех."
    "2. 23. Яблоки" -> "23. Яблоки"
    """
    if not theme:
        return theme
    # Паттерн: "N. M. Название" -> "M. Название"
    m = re.match(r"^(\d+)\.\s+(\d+\.\s+.+)$", theme)
    if m:
        return m.group(2)
    return theme


def fix_json(path: str):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    fixed_count = 0
    for q in data:
        theme = q.get("theme")
        if theme:
            new_theme = fix_theme(theme)
            if new_theme != theme:
                q["theme"] = new_theme
                fixed_count += 1

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Исправлено {fixed_count} тем в {path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Использование: python fix_double_numbering.py <path_to_json>")
        sys.exit(1)

    for path in sys.argv[1:]:
        fix_json(path)
