"""Скрипт для отправки запроса к LLM API для парсинга пакетов 'Своя игра'."""
import json
import sys
import os

import requests

sys.stdout.reconfigure(encoding="utf-8")

INPUT_FILE = "input/as-14.txt"
MAX_LINES = 100
OUTPUT_FILE = "llm_result.json"

SYSTEM_PROMPT = """Ты парсер пакетов "Своя игра".

Твоя задача — разобрать ВСЕ вопросы из переданного текста и вернуть строго JSON.

---

## Формат входа

Тебе даётся сырой текст пакета. Он может:
- содержать лишние пробелы, переносы, мусор
- иметь или не иметь поля (Комментарий, Источник, Зачёт)
- содержать несколько источников
- содержать квадратные скобки в ответе
- содержать указания типа "В ответе два слова" — их НЕ удалять из вопроса

В начале текста идёт СПИСОК ТЕМ — строки вида:
    1. Название темы (Автор)
    2. Другая тема (Автор)
    ...

Затем идут ВОПРОСЫ, сгруппированные по темам. В каждой теме обычно 5 вопросов
со стоимостью 10, 20, 30, 40, 50 (по порядку). Тема вопроса определяется
порядком: первые 5 вопросов относятся к теме 1, следующие 5 — к теме 2 и т.д.

Каждый вопрос имеет вид:
    10. Текст вопроса
    Форма: ...
    Ответ: ...
    Комментарий: ...
    Источник: ...

---

## Что нужно вернуть

Верни JSON-МАССИВ объектов (по одному на каждый вопрос) со следующими полями:

{
  "theme": string,
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
- определяется по позиции вопроса в пакете
- если тему определить невозможно — null

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
- текст после "Комментарий:"
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

## Пример

Вход:
1. Животные (Иванов)

10. Название ЭТОЙ ЧАСТИ ЗЕМНОГО ШАРА...
Ответ: Арктика
Комментарий: От греческого «арктос».
Источник: http://example.com

20. По сюжету ЭТОГО ФИЛЬМА...
Ответ: «Анаконда».

Выход:
[
  {
    "theme": "Животные",
    "price": 10,
    "question": "Название ЭТОЙ ЧАСТИ ЗЕМНОГО ШАРА...",
    "answer": "Арктика",
    "accepted_answers": null,
    "comment": "От греческого «арктос».",
    "sources": ["http://example.com"]
  },
  {
    "theme": "Животные",
    "price": 20,
    "question": "По сюжету ЭТОГО ФИЛЬМА...",
    "answer": "«Анаконда»",
    "accepted_answers": null,
    "comment": null,
    "sources": null
  }
]

---

Теперь разбери ВСЕ вопросы из следующего текста:

{input_text}
"""


def read_input(path: str, max_lines: int) -> str:
    """Читает файл и возвращает первые max_lines строк."""
    with open(path, "r", encoding="utf-8") as f:
        lines = []
        for i, line in enumerate(f):
            if i >= max_lines:
                break
            lines.append(line.rstrip("\n"))
    return "\n".join(lines)


def list_models() -> list:
    """Получает список доступных моделей."""
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
    """Отправляет запрос к LLM и возвращает текст ответа."""
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
    resp = requests.post(url, headers=headers, json=payload, timeout=300)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]


def extract_json(raw: str):
    """Извлекает JSON-массив или объект из ответа модели."""
    cleaned = raw.strip()

    # Удалим markdown-обёртки ```json ... ```
    if cleaned.startswith("```"):
        lines = cleaned.split("\n")
        lines = [ln for ln in lines if not ln.strip().startswith("```")]
        cleaned = "\n".join(lines).strip()

    # Попробуем распарсить как есть
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # Попробуем найти первый [ или { и последний ] или }
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


def main():
    if not os.path.exists(INPUT_FILE):
        print(f"[!] Файл {INPUT_FILE} не найден")
        sys.exit(1)

    print(f"[*] Читаю {INPUT_FILE} (первые {MAX_LINES} строк)...")
    input_text = read_input(INPUT_FILE, MAX_LINES)
    print(f"[*] Прочитано {len(input_text)} символов")

    print("[*] Запрашиваю список моделей...")
    models = list_models()
    if models:
        print(f"[*] Доступные модели: {models[:10]}{'...' if len(models) > 10 else ''}")
        if MODEL not in models and models:
            new_model = models[0]
            print(f"[!] Модель {MODEL} не найдена, использую {new_model}")
            model_to_use = new_model
        else:
            model_to_use = MODEL
    else:
        print(f"[!] Не удалось получить список моделей, использую {MODEL}")
        model_to_use = MODEL

    print(f"[*] Отправляю запрос к модели {model_to_use}...")
    try:
        raw_response = call_llm(SYSTEM_PROMPT, input_text, model_to_use)
    except requests.exceptions.HTTPError as e:
        print(f"[!] HTTP ошибка: {e}")
        print(f"[!] Тело ответа: {e.response.text if e.response else 'нет'}")
        sys.exit(1)
    except Exception as e:
        print(f"[!] Ошибка запроса: {e}")
        sys.exit(1)

    print("[*] Сырой ответ модели:")
    print("-" * 60)
    print(raw_response)
    print("-" * 60)

    try:
        parsed = extract_json(raw_response)
        if isinstance(parsed, dict):
            parsed = [parsed]
        print(f"[*] JSON успешно распарсен, получено объектов: {len(parsed)}")
        print(json.dumps(parsed[:3], ensure_ascii=False, indent=2))
        if len(parsed) > 3:
            print(f"    ... и ещё {len(parsed) - 3} объектов")
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(parsed, f, ensure_ascii=False, indent=2)
        print(f"[*] Результат сохранён в {OUTPUT_FILE}")
    except json.JSONDecodeError as e:
        print(f"[!] Не удалось распарсить JSON: {e}")
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(raw_response)
        print(f"[*] Сырой ответ сохранён в {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
