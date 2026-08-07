"""Пакетная обработка всех файлов из input/Проба пера_txt через LLM API."""
import json
import sys
import os
import time

import requests

sys.stdout.reconfigure(encoding="utf-8")

INPUT_DIR = "input/Проба пера_txt"
OUTPUT_FILE = "all_questions.json"

# Порядок обработки: сначала 7 кругов, потом Совковый, потом Наркоманский
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

SYSTEM_PROMPT = """Ты парсер пакетов "Своя игра".

Твоя задача — разобрать ВСЕ вопросы из переданного текста и вернуть строго JSON.

---

## Формат входа

Тебе даётся сырой текст одного раунда пакета. Он может:
- содержать лишние пробелы, переносы, мусор
- иметь или не иметь поля (Комментарий, Источник, Зачёт)
- содержать несколько источников
- содержать квадратные скобки в ответе
- содержать указания типа "В ответе два слова" — их НЕ удалять из вопроса

В начале текста идёт СПИСОК ТЕМ — строки вида:
    1. Название темы (Автор)
    2. Другая тема (Автор)
    ...

В некоторых раундах (бонусных) формат может быть:
    1. Название темы
(без автора в скобках)

Тема может быть секретной — тогда в её описании указано "Пока секрет" или
"Темой станет ответ на пятый вопрос". В таком случае в поле theme поставь null.

Затем идут ВОПРОСЫ, сгруппированные по темам. В каждой теме обычно 5 вопросов
со стоимостью 10, 20, 30, 40, 50 (по порядку). Тема вопроса определяется
порядком: первые 5 вопросов относятся к теме 1, следующие 5 — к теме 2 и т.д.

Каждый вопрос имеет вид:
    10. Текст вопроса
    Форма: ...
    Ответ: ...
    Комментарий: ...
    Источник: ...

Между темами может быть "Комментарий к теме:" — это описание темы, не вопроса.

---

## Что нужно вернуть

Верни JSON-МАССИВ объектов (по одному на каждый вопрос) со следующими полями:

{
  "theme": string | null,
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
- если тема секретная — null
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
- основной ответ (убери квадратные скобки, но текст внутри оставь)
Пример:
"[Альберт] Эйнштейн" → "Альберт Эйнштейн"

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
- если поле отсутствует — ставь null
- верни ТОЛЬКО JSON-массив без пояснений
- JSON должен быть валидным
- разбери ВСЕ вопросы из входа, не пропускай

---

Теперь разбери ВСЕ вопросы из следующего текста:

{input_text}
"""


def read_file(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


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


def call_llm(system_prompt: str, user_text: str, model: str) -> str:
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
    resp = requests.post(url, headers=headers, json=payload, timeout=600)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]


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


def process_file(path: str, model: str) -> list:
    print(f"\n[*] Обрабатываю {path}...")
    text = read_file(path)
    print(f"    Размер: {len(text)} символов")

    raw = call_llm(SYSTEM_PROMPT, text, model)
    parsed = extract_json(raw)
    if isinstance(parsed, dict):
        parsed = [parsed]
    print(f"    Получено вопросов: {len(parsed)}")
    return parsed


def main():
    if not os.path.isdir(INPUT_DIR):
        print(f"[!] Папка {INPUT_DIR} не найдена")
        sys.exit(1)

    print("[*] Запрашиваю список моделей...")
    models = list_models()
    if models:
        print(f"[*] Доступные модели: {models}")
        model_to_use = MODEL if MODEL in models else models[0]
        if model_to_use != MODEL:
            print(f"[!] Модель {MODEL} не найдена, использую {model_to_use}")
    else:
        model_to_use = MODEL

    all_questions = []
    for fname in FILE_ORDER:
        path = os.path.join(INPUT_DIR, fname)
        if not os.path.exists(path):
            print(f"[!] Файл {path} не найден, пропускаю")
            continue
        try:
            questions = process_file(path, model_to_use)
            # Добавим метаинформацию
            for q in questions:
                q["_source_file"] = fname
            all_questions.extend(questions)
        except Exception as e:
            print(f"[!] Ошибка при обработке {fname}: {e}")
            # Сохраним то, что уже есть
            with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
                json.dump(all_questions, f, ensure_ascii=False, indent=2)
            raise

        # Промежуточное сохранение
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(all_questions, f, ensure_ascii=False, indent=2)
        print(f"    Промежуточно сохранено в {OUTPUT_FILE} (всего {len(all_questions)})")

        # Пауза между запросами
        time.sleep(2)

    print(f"\n[*] Готово! Всего вопросов: {len(all_questions)}")
    print(f"[*] Результат сохранён в {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
