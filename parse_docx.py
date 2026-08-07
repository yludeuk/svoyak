"""Пакетная обработка docx/pdf файлов через LLM API с последующей генерацией docx.

Копия parse_all_llm.py, но на вход подаются docx или pdf файлы, которые
скрипт сам конвертирует в txt перед отправкой в LLM.

Каждый файл обрабатывается ОТДЕЛЬНО. Если файл большой (> MAX_LINES_PER_CHUNK),
он разбивается на чанки с перекрытием. Файлы между собой не склеиваются.

Результаты обработки кешируются в папке cache/ по SHA256 хэшу
(содержимое + системный промпт + модель). При повторном запуске
уже обработанные файлы/чанки не отправляются в LLM повторно.

Поддерживает:
- Автоматическую конвертацию docx → txt и pdf → txt
- Разбиение больших файлов на блоки по строкам с перекрытием
- Кеширование результатов по хэшу
- Дедупликацию вопросов на стыках чанков одного файла
- Генерацию docx файлов с заданным разбиением по темам
"""
import re
import json
import sys
import os
import time
from collections import OrderedDict

import requests

from docx import Document

from docx_generator import generate_docx_files
from cache import (
    compute_text_hash,
    compute_prompt_hash,
    make_cache_key,
    load_from_cache,
    save_to_cache,
    CACHE_DIR,
)
from llm_connection import API_KEY, API_BASE, MODEL

sys.stdout.reconfigure(encoding="utf-8")

# Максимальное количество строк в одном запросе к LLM
MAX_LINES_PER_CHUNK = 1000
# Перекрытие между блоками (строк)
CHUNK_OVERLAP = 50

# Конфигурация пакетов для обработки
# Каждый пакет: (имя_пакета, список_путей_к_docx/pdf, разбиение_на_блоки_для_docx)
PACKAGES = [
    # {
    #     "name": "Дровушки-2024",
    #     "input_files": [
    #         "input/Дровушки-2024/Дровушки 24. СИ. Отбор.docx",
    #         "input/Дровушки-2024/Дровушки 24. СИ. 1_16_1.docx",
    #         "input/Дровушки-2024/Дровушки 24. СИ. 1_16_2.docx",
    #         "input/Дровушки-2024/Дровушки 24. СИ. 1_8.docx",
    #         "input/Дровушки-2024/Дровушки 24. СИ. 1_4.docx",
    #         "input/Дровушки-2024/Дровушки 24. СИ. 1_2.docx",
    #         "input/Дровушки-2024/Дровушки 24. СИ. Финал.docx",
    #     ],
    #     # Разбиение на блоки для финальных docx (заполнить после оценки количества тем)
    #     "split": [10] + [9]* 4 + [10]*2,
    # }
    # {
    #     "name": "БФИИ-2025",
    #     "input_files": [
    #         "input/Байкалфест-2025/СИ. 1_16 финала №1. XXI БФИИ .pdf",
    #         "input/Байкалфест-2025/СИ. 1_16 финала №2. XXI БФИИ .pdf",
    #         "input/Байкалфест-2025/СИ. 1_8 Финала. XXI БФИИ.pdf",
    #         "input/Байкалфест-2025/СИ. 1_4 финала. XXI БФИИ .pdf",
    #         "input/Байкалфест-2025/СИ. 1_2 финала. XXI БФИИ .pdf",
    #         "input/Байкалфест-2025/СИ. Финал. XXI БФИИ .pdf",
    #     ],
    #     # Разбиение на блоки для финальных docx (заполнить после оценки количества тем)
    #     "split": [9] * 4 + [10] * 2,
    # {
    #     "name": "ЧР-2026",
    #     "input_files": [
    #         "input/ЧР-2026/chr26_eq_1_16.docx",
    #         "input/ЧР-2026/chr26_eq_1_8.docx",
    #         "input/ЧР-2026/chr26_eq_reserve.docx",
    #         "input/ЧР-2026/chr26_eq_1_4.docx",
    #         "input/ЧР-2026/chr26_eq_1_2.docx",
    #         "input/ЧР-2026/chr26_eq_final.docx",
    #     ],
    #     # Разбиение на блоки для финальных docx (заполнить после оценки количества тем)
    #     "split": [12] * 2 + [10] * 2 + [11, 12],
    # }

    {
        "name": "Love in December-2020",
        "input_files": [
            "input/LiD - 2020/Групповой этап. Круг 1 (9 тем).docx",
            "input/LiD - 2020/Групповой этап. Круг 2 (9 тем).docx",
            "input/LiD - 2020/Групповой этап. Круг 3 (9 тем).docx",
            "input/LiD - 2020/Групповой этап. Круг 4 (9 тем).docx",
            "input/LiD - 2020/Групповой этап. Круг 5 (9 тем).docx",
            "input/LiD - 2020/Одна восьмая финала (12 тем).docx",
            "input/LiD - 2020/Одна четвертая финала (12 тем).docx",
            "input/LiD - 2020/Полуфинал(12 тем).docx",
            "input/LiD - 2020/Финал (15 тем).docx",
            "input/LiD - 2020/Запас (3 темы).docx"
        ],
        # Разбиение на блоки для финальных docx (заполнить после оценки количества тем)
        "split": [9] + [10] * 9,
    }
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
  "form": string | null,
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

2. theme_comment:
- текст комментария к теме (идёт после названия темы, обычно начинается с "Комментарий к теме:" или просто описательный текст перед первым вопросом)
- если нет — null

3. author:
- автор темы (из строки "Автор: ..." которая идёт сразу после названия темы)
- если нет — null

4. price:
- число (10, 20, 30, 40, 50...) если есть
- иначе null

5. question:
- полный текст вопроса
- БЕЗ строки с ценой
- БЕЗ "Ответ:", "Комментарий:", "Источник:"
- сохраняй формулировку как есть

6. form:
- форма вопроса, указывает на то, что нужно сдать в ответе, чаще всего это местоимение, которое относится к ответу.
- может быть напрямую указана до или после вопроса, в таком случае её нужно убрать из вопроса и перенести в отдельное поле
- может быть не указана напрямую, в таком случае нужно попробовать её понять самому:
    часто форма полностью или частично указана в тексте вопроса капсом -
      "ОН", "ОНА", "ОНО", "ЕЁ", "ЕГО", "ЭТО", "КТО", "ЧТО", "ЭТОТ ЧЕЛОВЕК", "ЭТОТ ФРАНЦУЗСКИЙ ГОРОД", "ЧТО МЫ ПРОПУСТИЛИ" и другие подобные конструкции;
    иногда капсом указана только часть формы - "ЭТОТ", "ЭТУ" и другие похожие местоимения. В таком случае нужно попробовать самому дополнить её существительным из текста вопроса;
    если форма не указана капсом, нужно по смыслу попробовать её достать из текста самому, если задача сложная - оставить null и форму не указывать.

7. answer:
- основной ответ

8. accepted_answers:
- из блока "Зачёт"
- разбей на список
- если нет — null

9. comment:
- текст после "Комментарий:" (имеется в виду комментарий К ВОПРОСУ, не к теме)
- если нет — null

10. sources:
- все источники
- если несколько — массив
- если один — массив из одного элемента
- если нет — null

---

## ВАЖНО

- НЕ выдумывай данные
- НЕ теряй данные, если видишь ошибку - исправь либо обойди ограничение, но ничего не удаляй
- если поле отсутствует — ставь null
- верни ТОЛЬКО JSON-массив без пояснений
- JSON должен быть валидным
- разбери ВСЕ вопросы из входа, не пропускай

---

Теперь разбери ВСЕ вопросы из следующего текста:

{input_text}
"""


def docx_to_text(path: str) -> str:
    """Конвертирует docx файл в текст (по параграфам)."""
    doc = Document(path)
    lines = []
    for p in doc.paragraphs:
        text = p.text
        if text.strip() == "":
            lines.append("")
        else:
            lines.append(text)
    return "\n".join(lines)


def pdf_to_text(path: str) -> str:
    """Конвертирует pdf файл в текст (постранично, с сохранением переносов строк)."""
    import pdfplumber

    pages_text = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            # extract_text возвращает текст с переносами строк внутри страницы
            text = page.extract_text() or ""
            pages_text.append(text)
    # Между страницами вставляем пустую строку для лучшего разделения
    return "\n\n".join(pages_text)


def extract_text(path: str) -> str:
    """Извлекает текст из файла в зависимости от расширения.

    Поддерживает .docx и .pdf.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return docx_to_text(path)
    elif ext == ".pdf":
        return pdf_to_text(path)
    else:
        raise ValueError(
            f"Неподдерживаемый формат файла: {ext} "
            f"(ожидается .docx или .pdf): {path}"
        )


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
    theme = q.get("theme") or ""
    price = q.get("price")
    question = (q.get("question") or "").strip()[:100]
    return f"{theme}|{price}|{question}"


def deduplicate_questions(questions: list) -> list:
    """Удаляет дубликаты вопросов, сохраняя порядок."""
    seen = set()
    result = []
    for q in questions:
        sig = question_signature(q)
        if sig not in seen:
            seen.add(sig)
            result.append(q)
    return result


def process_chunk_with_cache(chunk_text: str, chunk_idx: int, model: str,
                             source_label: str, source_file: str,
                             prompt_hash: str, chunk_start: int = 0,
                             chunk_end: int = 0) -> list:
    """Обрабатывает один чанк с использованием кеша.

    Возвращает список вопросов.
    """
    text_hash = compute_text_hash(chunk_text)
    cache_key = make_cache_key(text_hash, prompt_hash, model)

    # Проверяем кеш
    cached = load_from_cache(CACHE_DIR, cache_key)
    if cached is not None:
        print(f"          [cache] Чанк {chunk_idx}: найден в кеше ({len(cached)} вопросов)")
        for q in cached:
            q["_source_file"] = source_label
            q["_source_docx"] = source_file
            q["_chunk"] = chunk_idx
            q["_from_cache"] = True
        return cached

    # Кеша нет — отправляем в LLM
    print(f"          [api]   Чанк {chunk_idx}: отправляю в LLM...")
    raw = call_llm(SYSTEM_PROMPT, chunk_text, model)
    parsed = extract_json(raw)
    if isinstance(parsed, dict):
        parsed = [parsed]
    print(f"          Получено вопросов: {len(parsed)}")

    # Сохраняем в кеш (без служебных полей)
    questions_for_cache = []
    for q in parsed:
        q_clean = {k: v for k, v in q.items() if not k.startswith("_")}
        questions_for_cache.append(q_clean)
    save_to_cache(
        CACHE_DIR, cache_key, text_hash, prompt_hash, model,
        questions_for_cache,
        source_file=source_file,
        chunk_index=chunk_idx,
        chunk_start=chunk_start,
        chunk_end=chunk_end,
    )

    # Возвращаем с метаданными
    for q in parsed:
        q["_source_file"] = source_label
        q["_source_docx"] = source_file
        q["_chunk"] = chunk_idx
        q["_from_cache"] = False
    return parsed


def process_single_file(path: str, model: str, source_label: str,
                        prompt_hash: str) -> list:
    """Обрабатывает один входной файл (docx или pdf).

    Если файл большой — разбивает на чанки с перекрытием.
    Каждый чанк кешируется отдельно.
    Возвращает список вопросов.
    """
    print(f"\n    [*] Файл: {path}")
    if not os.path.exists(path):
        print(f"        [!] Не найден, пропускаю")
        return []

    try:
        text = extract_text(path)
    except ValueError as e:
        print(f"        [!] {e}")
        return []
    except Exception as e:
        print(f"        [!] Ошибка при извлечении текста: {e}")
        return []

    lines = text.split("\n")
    print(f"        Конвертировано: {len(lines)} строк, {len(text)} символов")

    chunks = make_chunks(lines, MAX_LINES_PER_CHUNK, CHUNK_OVERLAP)
    if len(chunks) > 1:
        print(f"        Разбит на {len(chunks)} чанков (перекрытие {CHUNK_OVERLAP} строк)")

    file_questions = []
    for chunk_idx, (start, end, chunk_text) in enumerate(chunks, 1):
        if len(chunks) > 1:
            print(f"        Чанк {chunk_idx}/{len(chunks)} (строки {start + 1}–{end})...")
        try:
            chunk_questions = process_chunk_with_cache(
                chunk_text=chunk_text,
                chunk_idx=chunk_idx,
                model=model,
                source_label=source_label,
                source_file=os.path.basename(path),
                prompt_hash=prompt_hash,
                chunk_start=start,
                chunk_end=end,
            )
            file_questions.extend(chunk_questions)
        except Exception as e:
            print(f"          [!] Ошибка: {e}")
            raise

        # Пауза между чанками одного файла
        if chunk_idx < len(chunks):
            time.sleep(2)

    # Дедупликация внутри файла (на стыках чанков)
    before = len(file_questions)
    file_questions = deduplicate_questions(file_questions)
    after = len(file_questions)
    if before != after:
        print(f"        Дедупликация: {before} → {after} (удалено {before - after})")

    print(f"        Итого из файла: {len(file_questions)} вопросов")
    return file_questions


def process_input_files(file_paths: list, model: str, source_label: str,
                        prompt_hash: str) -> list:
    """Обрабатывает список входных файлов (docx/pdf): каждый файл отдельно.

    Файлы между собой не склеиваются. Каждый файл обрабатывается
    независимо (с возможным разбиением на чанки внутри).
    """
    print(f"\n[*] Обрабатываю {len(file_paths)} входных файлов (каждый отдельно)...")

    all_questions = []
    for path in file_paths:
        file_questions = process_single_file(
            path, model, source_label, prompt_hash
        )
        all_questions.extend(file_questions)
        # Пауза между файлами
        if path != file_paths[-1]:
            time.sleep(2)

    print(f"\n    Всего вопросов из всех файлов: {len(all_questions)}")
    return all_questions


def renumber_themes_by_price(questions: list) -> list:
    """Пост-обработка: добавляет порядковый номер к названию темы.

    Границы тем определяются по стоимости вопросов внутри одного source_file:
    если текущий вопрос имеет price=10, а предыдущий имел price=50 (или это
    первый вопрос) — это начало новой темы.

    К названию темы добавляется префикс: "1. Яблоки", "2. Пока секрет" и т.д.
    Это решает проблему неуникальных названий (несколько "Пока секрет" в пакете).
    """
    if not questions:
        return questions

    # Группируем по source_file, сохраняя порядок
    by_source = OrderedDict()
    for q in questions:
        sf = q.get("_source_file", "?")
        by_source.setdefault(sf, []).append(q)

    result = []
    for sf, qs in by_source.items():
        theme_counter = 0
        prev_price = None
        for q in qs:
            price = q.get("price")
            # Новая тема начинается, если:
            # - это первый вопрос
            # - текущий price=10, а предыдущий был 50
            # - текущий price=None (нет цены)
            is_new_theme = (
                prev_price is None
                or price == 10
                or price is None
                or (prev_price is not None and prev_price >= 50 and price < prev_price)
            )
            if is_new_theme:
                theme_counter += 1

            theme = q.get("theme")
            if theme is not None and theme != "":
                # Проверяем, не начинается ли тема уже с номера ("1. Смех.")
                if re.match(r"^\d+\.\s+", theme):
                    # Тема уже пронумерована — не добавляем номер
                    pass
                else:
                    # Добавляем порядковый номер к названию темы
                    q["theme"] = f"{theme_counter}. {theme}"

            prev_price = price
            result.append(q)

    return result


def process_package(pkg: dict, model: str) -> list:
    """Обрабатывает один пакет: парсит через LLM и генерирует docx."""
    name = pkg["name"]
    input_files = pkg["input_files"]
    user_split = pkg["split"]

    print(f"\n{'=' * 60}")
    print(f"[*] Пакет: {name}")
    print(f"    Файлов: {len(input_files)}")
    print(f"    Разбиение: {user_split}")
    print(f"{'=' * 60}")

    if not input_files:
        print("[!] Не указаны входные файлы, пропускаю")
        return []

    # Хэш системного промпта — часть ключа кеша
    prompt_hash = compute_prompt_hash(SYSTEM_PROMPT)

    questions = process_input_files(input_files, model, name, prompt_hash)

    # Пост-обработка: нумерация тем по границам (50→10)
    print(f"\n[*] Пост-обработка: нумерация тем...")
    questions = renumber_themes_by_price(questions)

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
