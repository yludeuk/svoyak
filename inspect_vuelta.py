import sys, re
sys.stdout.reconfigure(encoding='utf-8')
with open('input/vuelta-summer-2026.txt', encoding='utf-8') as f:
    text = f.read()
lines = text.split('\n')
print(f"Всего строк: {len(lines)}")
print(f"Всего символов: {len(text)}")

# Найдём строки, начинающиеся с "N. " где N - число
theme_candidates = []
for i, line in enumerate(lines):
    m = re.match(r'^\s*(\d+)\.\s+(.+)$', line.strip())
    if m:
        n = int(m.group(1))
        name = m.group(2).strip()
        # Тема - это строка, после которой идёт "Автор:" или просто пустая строка
        # Вопрос - это строка, после которой идёт "Форма:" или "Ответ:"
        # Проверим следующие 5 строк
        is_theme = False
        for j in range(i + 1, min(i + 6, len(lines))):
            nxt = lines[j].strip()
            if nxt.startswith('Автор:') or nxt.startswith('Автор '):
                is_theme = True
                break
            if nxt.startswith('Форма:') or nxt.startswith('Ответ:'):
                is_theme = False
                break
        if is_theme:
            theme_candidates.append((i, n, name))

print(f"\nНайдено тем: {len(theme_candidates)}")
print("\nПервые 5 тем:")
for i, n, name in theme_candidates[:5]:
    print(f"  [{i}] {n}. {name}")
print("\nПоследние 5 тем:")
for i, n, name in theme_candidates[-5:]:
    print(f"  [{i}] {n}. {name}")
