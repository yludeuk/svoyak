"""Генерация docx из уже сохранённого JSON."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
from docx_generator import generate_docx_files
import json

with open("Vuelta-summer-2026_questions.json", "r", encoding="utf-8") as f:
    questions = json.load(f)

print(f"Загружено {len(questions)} вопросов")

# 103 темы: 10 блоков по 10 + 1 блок на 3 = 11 файлов
generate_docx_files(
    questions=questions,
    package_name="Vuelta-summer-2026",
    user_split=[10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 3],
    output_dir="output",
)
