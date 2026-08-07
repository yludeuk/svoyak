"""Пакетная обработка файлов через LLM API с последующей генерацией docx.

Поддерживает:
- Разбиение больших файлов на блоки по строкам с перекрытием
- Дедупликацию вопросов на стыках блоков
- Генерацию docx файлов с заданным разбиением по темам
"""
import json
import sys
import os
import time
import re

import requests

from docx_generator import generate_docx_files

sys.stdout.reconfigure(encoding="utf-8")

# Максимальное количество строк в одном запросе к LLM
MAX_LINES_PER_CHUNK = 1000
# Перекрытие между блоками (строк)
CHUNK_OVERLAP = 50

# Конфигурация пакетов для обработки
# Каждый пакет: (имя_пакета, путь_к_txt, разбиение_на_блоки_для_docx)
PACKAGES = [
    {
        "name": "Vuelta-summer-2026",
        "input_file": "input/vuelta-summer-2026.txt",
        "split": [10] * 10,
    },
]

SYSTEM_PROMPT = """Ты парсер пакетов "Своя игра".

Твоя задача — разобрать ВСЕ вопросы из переданного текста и вернуть строго JSON.

---

## Формат входа

Тебе даётся сырой текст пакета. Он может:
- содержать лишние пробелы, переносы, мусор
- иметь или не иметь поля (Комментарий, Источник, Зачёт)
- содержать несколько источников
- содержать квадратные скобки в ответе

Темы могут идти сначала списком:
    1. Название темы (Автор)
    2. Другая тема (Автор)
    ...
либо просто по одной.

После названия темы может идти комментарий к теме или указание автора.
    1. Название темы
Комментарий: ...
(без автора в скобках)

Тема может быть секретной, тогда она может называться "Пока секрет", а может и по другому.
После какого-либо из вопросов тема может менять название

Затем идут ВОПРОСЫ, относящиеся к теме. В каждой теме обычно 5 вопросов
со стоимостью 10, 20, 30, 40, 50 (по порядку). В разных пакетах стоимости могут отличаться, но как правило всегда кратны этой последовательности.

Вопрос может иметь вид:
    10. Текст вопроса
    Форма: ... (опционально)
    Ответ: ...
    Комментарий: ... (опционально)
    Источник: ... (опционально)

---

## Что нужно вернуть

Верни JSON-МАССИВ объектов (по одному на каждый вопрос) со следующими полями:

{
  "theme": string | null,
  "theme_comment": string | null,
  "author": string | null,
  "price": number | null,
  "question": string,
  "answer": string,
  "accepted_answers": string[] | null,
  "comment": string | null,
  "sources": string[] | null
}

---

## Правила

1. theme:
- название темы (без номера и автора в скобках)
- определяется по позиции вопроса в раунде
- если определить невозможно — null

2. price:
- число (10, 20, 30, 40, 50...) если есть
- иначе null

3. question:
- полный текст вопроса
- БЕЗ строки с ценой
- БЕЗ "Ответ:", "Комментарий:", "Источник:"
- сохраняй формулировку как есть

4. answer:
- основной ответ

5. accepted_answers:
- из блока "Зачёт"
- разбей на список
- если нет — null

6. comment:
- текст после "Комментарий:" (имеется в виду комментарий К ВОПРОСУ, не к теме)
- если нет — null

7. sources:
- все источники
- если несколько — массив
- если один — массив из одного элемента
- если нет — null

---

## ВАЖНО

- НЕ выдумывай данные
- НЕ теряй данные
- если поле отсутствует — ставь null
- верни ТОЛЬКО JSON-массив без пояснений
- JSON должен быть валидным
- разбери ВСЕ вопросы из входа, не пропускай

---

Теперь разбери ВСЕ вопросы из следующего текста:

{input_text}
"""


def read_file_lines(path: str) -> list:
    """Читает файл и возвращает список строк."""
    with open(path, "r", encoding="utf-8") as f:
        return f.read().split("\n")


def list_models() -> list:
    url = f"{API_BASE}/v1/models"
    headers = {"Authorization": f"Bearer {API_KEY}"}
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        return [m.get("id", "") for m in data.get("data", [])]
    except Exception as e:
        print(f"[!] Не удалось получить список моделей: {e}")
        return []


def call_llm(system_prompt: str, user_text: str, model: str,
             max_retries: int = 3, timeout: int = 900) -> str:
    url = f"{API_BASE}/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_text},
        ],
        "temperature": 0.0,
    }
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
            return data["choices"][0]["message"]["content"]
        except (requests.exceptions.Timeout, requests.exceptions.HTTPError) as e:
            last_error = e
            print(f"      [!] Попытка {attempt}/{max_retries} не удалась: {e}")
            if attempt < max_retries:
                wait = 15 * attempt
                print(f"      Жду {wait} сек перед повтором...")
                time.sleep(wait)
    raise last_error


def extract_json(raw: str):
    cleaned = raw.strip()
    if cleaned.startswith("```"):
        lines = cleaned.split("\n")
        lines = [ln for ln in lines if not ln.strip().startswith("```")]
        cleaned = "\n".join(lines).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass
    start = None
    end = None
    for i, ch in enumerate(cleaned):
        if ch in "[{" and start is None:
            start = i
        if ch in "]}":
            end = i
    if start is not None and end is not None and end > start:
        candidate = cleaned[start : end + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            pass
    raise json.JSONDecodeError("Не удалось извлечь JSON", cleaned, 0)


def make_chunks(lines: list, max_lines: int, overlap: int) -> list:
    """Разбивает список строк на чанки с перекрытием.

    Возвращает список кортежей (start_idx, end_idx, text).
    """
    if len(lines) <= max_lines:
        return [(0, len(lines), "\n".join(lines))]

    chunks = []
    start = 0
    while start < len(lines):
        end = min(start + max_lines, len(lines))
        chunk_text = "\n".join(lines[start:end])
        chunks.append((start, end, chunk_text))
        if end >= len(lines):
            break
        start = end - overlap
    return chunks


def question_signature(q: dict) -> str:
    """Создаёт сигнатуру вопроса для дедупликации."""
    # Используем комбинацию темы, цены и первых 100 символов вопроса
    theme = q.get("theme") or ""
    price = q.get("price")
    question = (q.get("question") or "").strip()[:100]
    return f"{theme}|{price}|{question}"


def deduplicate_questions(questions: list) -> list:
    """Удаляет дубликаты вопросов, сохраняя порядок."""
    seen = set()
    result = []
    for q in questions:
        # sig = question_signature(q)
        sig = (q['question'], q['answer'], q['comment'])
        if sig not in seen:
            seen.add(sig)
            result.append(q)
    return result


def process_file(path: str, model: str, source_label: str) -> list:
    """Обрабатывает файл, при необходимости разбивая на чанки."""
    print(f"\n[*] Обрабатываю {path}...")
    lines = read_file_lines(path)
    print(f"    Всего строк: {len(lines)}")

    chunks = make_chunks(lines, MAX_LINES_PER_CHUNK, CHUNK_OVERLAP)
    print(f"    Разбито на чанков: {len(chunks)}")

    all_questions = []
    for chunk_idx, (start, end, chunk_text) in enumerate(chunks, 1):
        print(f"    Чанк {chunk_idx}/{len(chunks)} (строки {start + 1}–{end})...")
        try:
            raw = call_llm(SYSTEM_PROMPT, chunk_text, model)
            parsed = extract_json(raw)
            if isinstance(parsed, dict):
                parsed = [parsed]
            print(f"      Получено вопросов: {len(parsed)}")
            for q in parsed:
                q["_source_file"] = source_label
                q["_chunk"] = chunk_idx
            all_questions.extend(parsed)
        except Exception as e:
            print(f"      [!] Ошибка: {e}")
            raise

        # Пауза между чанками
        if chunk_idx < len(chunks):
            time.sleep(2)

    # Дедупликация
    before = len(all_questions)
    all_questions = deduplicate_questions(all_questions)
    after = len(all_questions)
    if before != after:
        print(f"    Дедупликация: {before} → {after} (удалено {before - after})")

    print(f"    Итого уникальных вопросов: {len(all_questions)}")
    return all_questions


def process_package(pkg: dict, model: str) -> list:
    """Обрабатывает один пакет: парсит через LLM и генерирует docx."""
    name = pkg["name"]
    input_file = pkg["input_file"]
    user_split = pkg["split"]

    print(f"\n{'=' * 60}")
    print(f"[*] Пакет: {name}")
    print(f"    Файл: {input_file}")
    print(f"    Разбиение: {user_split}")
    print(f"{'=' * 60}")

    if not os.path.exists(input_file):
        print(f"[!] Файл {input_file} не найден, пропускаю")
        return []

    questions = process_file(input_file, model, name)

    # Сохраняем промежуточный JSON
    json_path = f"{name}_questions.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(questions, f, ensure_ascii=False, indent=2)
    print(f"[*] JSON сохранён в {json_path}")

    # Генерируем docx файлы
    print(f"\n[*] Генерирую docx файлы для пакета «{name}»...")
    created = generate_docx_files(
        questions=questions,
        package_name=name,
        user_split=user_split,
        output_dir="output",
    )
    print(f"[*] Создано файлов: {len(created)}")
    return questions


def main():
    print("[*] Запрашиваю список моделей...")
    models = list_models()
    if models:
        print(f"[*] Доступные модели: {models}")
        model_to_use = MODEL if MODEL in models else models[0]
        if model_to_use != MODEL:
            print(f"[!] Модель {MODEL} не найдена, использую {model_to_use}")
    else:
        model_to_use = MODEL

    for pkg in PACKAGES:
        try:
            process_package(pkg, model_to_use)
        except Exception as e:
            print(f"[!] Ошибка при обработке пакета {pkg['name']}: {e}")
            import traceback
            traceback.print_exc()
            continue

    print("\n[*] Всё готово!")


if __name__ == "__main__":
    main()
