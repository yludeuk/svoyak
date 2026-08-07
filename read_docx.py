"""Чтение docx файлов из папки Проба пера и сохранение в txt для анализа."""
import os
from docx import Document

INPUT_DIR = "input/Проба пера"
OUTPUT_DIR = "input/Проба пера_txt"
os.makedirs(OUTPUT_DIR, exist_ok=True)

for fname in sorted(os.listdir(INPUT_DIR)):
    if not fname.lower().endswith(".docx"):
        continue
    src = os.path.join(INPUT_DIR, fname)
    doc = Document(src)
    lines = []
    for p in doc.paragraphs:
        text = p.text
        if text.strip() == "":
            lines.append("")
        else:
            lines.append(text)
    out_name = os.path.splitext(fname)[0] + ".txt"
    with open(os.path.join(OUTPUT_DIR, out_name), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"OK: {fname} -> {out_name} ({len(lines)} paragraphs)")
