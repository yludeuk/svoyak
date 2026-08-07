"""Анализ тем в all_questions.json."""
import json
import sys
from collections import OrderedDict

sys.stdout.reconfigure(encoding="utf-8")

with open("all_questions.json", "r", encoding="utf-8") as f:
    data = json.load(f)

print(f"Всего вопросов: {len(data)}")

# Группируем по source_file и theme
by_file = OrderedDict()
for q in data:
    sf = q.get("_source_file", "?")
    by_file.setdefault(sf, []).append(q)

for sf, qs in by_file.items():
    themes = OrderedDict()
    for q in qs:
        t = q.get("theme") or "<NULL>"
        themes.setdefault(t, 0)
        themes[t] += 1
    print(f"\n=== {sf} ({len(qs)} вопросов) ===")
    for t, c in themes.items():
        print(f"  {t}: {c}")
