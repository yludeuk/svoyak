"""Патч для docx_generator.py: исправляет split_themes для работы с OrderedDict."""
with open("docx_generator.py", "r", encoding="utf-8") as f:
    content = f.read()

old = '''def split_themes(themes, user_split):
    """Разбивает OrderedDict тем на блоки согласно user_split.

    user_split — список целых чисел, например [10, 10, 10, 10, 10, 10, 12].
    Сумма элементов должна быть равна len(themes).
    """
    total = len(themes)
    blocks, idx = [], 0
    for sz in user_split:
        blocks.append(themes[idx : idx + sz])
        idx += sz
    if idx != total:
        raise ValueError(
            f"Сумма user_split ({idx}) не равна количеству тем ({total})"
        )
    return blocks'''

new = '''def split_themes(themes, user_split):
    """Разбивает OrderedDict тем на блоки согласно user_split.

    user_split — список целых чисел, например [10, 10, 10, 10, 10, 10, 12].
    Сумма элементов должна быть равна len(themes).
    Возвращает список OrderedDict-ов.
    """
    total = len(themes)
    items = list(themes.items())
    blocks, idx = [], 0
    for sz in user_split:
        blocks.append(OrderedDict(items[idx : idx + sz]))
        idx += sz
    if idx != total:
        raise ValueError(
            f"Сумма user_split ({idx}) не равна количеству тем ({total})"
        )
    return blocks'''

if old not in content:
    print("ОШИБКА: старая функция не найдена")
    raise SystemExit(1)

content = content.replace(old, new)

with open("docx_generator.py", "w", encoding="utf-8") as f:
    f.write(content)

print("OK: патч применён")
