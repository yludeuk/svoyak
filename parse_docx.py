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
    load_cache_entry,
    load_raw_from_cache,
    save_to_cache,
    CACHE_DIR,
)
from llm_connection import API_KEY, API_BASE, MODEL

sys.stdout.reconfigure(encoding="utf-8")

# Максимальное количество строк в одном запросе к LLM
MAX_LINES_PER_CHUNK = 1000
# Перекрытие между блоками (строк)
# Должно быть достаточно большим, чтобы заголовок темы вместе со всеми её
# вопросами гарантированно попадал в следующий чанк. Иначе LLM может не
# распознать тему вопросов из зоны перекрытия (theme=null), и такие вопросы
# не будут удалены дедупликацией (см. deduplicate_questions).
CHUNK_OVERLAP = 150

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

    # {
    #     "name": "Love in December-2020",
    #     "input_files": [
    #         "input/LiD - 2020/Групповой этап. Круг 1 (9 тем).docx",
    #         "input/LiD - 2020/Групповой этап. Круг 2 (9 тем).docx",
    #         "input/LiD - 2020/Групповой этап. Круг 3 (9 тем).docx",
    #         "input/LiD - 2020/Групповой этап. Круг 4 (9 тем).docx",
    #         "input/LiD - 2020/Групповой этап. Круг 5 (9 тем).docx",
    #         "input/LiD - 2020/Одна восьмая финала (12 тем).docx",
    #         "input/LiD - 2020/Одна четвертая финала (12 тем).docx",
    #         "input/LiD - 2020/Полуфинал(12 тем).docx",
    #         "input/LiD - 2020/Финал (15 тем).docx",
    #         "input/LiD - 2020/Запас (3 темы).docx"
    #     ],
    #     # Разбиение на блоки для финальных docx (заполнить после оценки количества тем)
    #     "split": [9] + [10] * 9,
    # }

    # {
    #     "name": "ЛВТ-2024",
    #     "input_files": [
    #         "input/ЛВТ-2024/СВОЯ ИГРА ЛВТ 2024.docx"
    #     ],
    #     # Разбиение на блоки для финальных docx (заполнить после оценки количества тем)
    #     "split": [9] * 5 + [10] * 4,
    # }

    # {
    #     "name": "Чарка-2026",
    #     "input_files": [
    #         "input/Чарка/Чарка 1, 2026, тур 1.docx",
    #         "input/Чарка/Чарка 1, 2026, тур 2.docx",
    #         "input/Чарка/Чарка 1, 2026, тур 3.docx",
    #         "input/Чарка/Чарка 1, 2026, тур 4.docx",
    #         "input/Чарка/Чарка 1, 2026, тур 5.docx",
    #         "input/Чарка/Чарка 1, 2026, тур 6.docx",
    #         "input/Чарка/Чарка 1, 2026, тур 7.docx",
    #     ],
    #     # Разбиение на блоки для финальных docx (заполнить после оценки количества тем)
    #     "split": [9] + [10] * 4,
    # }

    {
        "name": "Шкврка-2026",
        "input_files": [
            "input/Чарка/Шкварка 1, 2026, этап 1.docx",
            "input/Чарка/Шкварка 1, 2026, этап 2.docx",
            "input/Чарка/Шкварка 1, 2026, этап 3.docx",
            "input/Чарка/Шкварка 1, 2026, этап 4.docx",
            "input/Чарка/Шкварка 1, 2026, этап 5.docx",
            "input/Чарка/Шкварка_1,_2026,_этап_6_четвертьфинал.docx",
            "input/Чарка/Шкварка 1, 2026, этап 7 (полуфинал).docx",
            "input/Чарка/Шкварка 1, 2026, этап 8 (финал).docx",
            "input/Чарка/Запас. Чарка и Шкварка 1, 2026.docx",
        ],
        # Разбиение на блоки для финальных docx (заполнить после оценки количества тем)
        "split": [9]*3 + [10] * 3,
    }Чарка. Зеркало СЧРБ-2025 (СИ/КСИ/ЭК). Подборка посложнее. Авторы - Карпиевич, Копочель, Кулешов, Макаревич, Руденко,  Шевела. 49 тем.
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
- если вопрос явно продолжает предыдущую тему (например, в тексте видна
  последовательность цен 10, 20, 30, 40, 50), но название темы в этом
  фрагменте не видно — НАСЛЕДУЙ название темы из предыдущего вопроса,
  НЕ ставь null
- только если тему действительно невозможно определить — null

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
- форма вопроса — это краткая подсказка, ЧТО ИМЕННО нужно сдать в ответе (чаще всего местоимение или короткая конструкция, относящаяся к ответу).
- ВАЖНО: НЕ удаляй форму из текста вопроса! Форма (обычно выделена КАПСОМ) остаётся в вопросе на своём месте — это часть формулировки вопроса.
- в поле form укажи ОДНУ краткую форму, которая помогает понять, что нужно ответить.
- НЕ перечисляй все вхождения формы через запятую: если в вопросе несколько одинаковых местоимений, относящихся к одному и тому же (например, "ОНИ ... ИХ" — оба про один объект), укажи форму ОДИН раз ("ОНИ").
- если в вопросе несколько РАЗНЫХ форм, которые вместе образуют вопрос (например, "ОНИ ДЕЛАЮТ ЭТО" — нужно ответить, что именно ОНИ делают), объедини их в одну осмысленную конструкцию ("ОНИ ДЕЛАЮТ ЭТО").
- если капсом указана только часть формы ("ЭТОТ", "ЭТУ"), дополни её существительным из текста вопроса, чтобы форма стала понятной ("ЭТОТ ГОРОД", "ЭТУ ДОЛЖНОСТЬ").
- если форма не указана капсом, попробуй по смыслу достать её из текста самому.
- если форма сложная, неоднозначная или её невозможно кратко сформулировать — поставь null (лучше null, чем запутывающий список).

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
            content = data["choices"][0]["message"]["content"]
            print(f"          Размер полученного текста: {len(content)}")
            return content
        except (requests.exceptions.Timeout, requests.exceptions.HTTPError) as e:
            last_error = e
            print(f"      [!] Попытка {attempt}/{max_retries} не удалась: {e}")
            if attempt < max_retries:
                wait = 15 * attempt
                print(f"      Жду {wait} сек перед повтором...")
                time.sleep(wait)
    raise last_error


def _save_raw_for_debug(raw: str, prefix: str = "llm_raw") -> str:
    """Сохраняет сырой ответ LLM в файл для отладки. Возвращает путь к файлу."""
    import datetime
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = f"{prefix}_{ts}.txt"
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(raw)
        return path
    except Exception as e:
        print(f"          [!] Не удалось сохранить сырой ответ: {e}")
        return ""


def _find_matching_bracket(text: str, start: int, open_ch: str, close_ch: str) -> int:
    """Ищет соответствующую закрывающую скобку с учётом вложенности и строк.

    Возвращает индекс закрывающей скобки или -1, если не найдена.
    Корректно обрабатывает строки в кавычках и экранированные символы.
    """
    depth = 0
    in_string = False
    escape = False
    for i in range(start, len(text)):
        ch = text[i]
        if escape:
            escape = False
            continue
        if ch == "\\":
            escape = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == open_ch:
            depth += 1
        elif ch == close_ch:
            depth -= 1
            if depth == 0:
                return i
    return -1


def _repair_unescaped_quotes(text: str) -> str:
    """Чинит неэкранированные кавычки внутри JSON-строк.

    LLM иногда возвращает JSON, где внутри строкового значения встречается
    обычная кавычка " (например, «...нашел!"."»), что делает JSON невалидным.
    Эта функция проходит по тексту, отслеживая состояние "внутри строки",
    и экранирует кавычки, которые стоят НЕ на границе строки (т.е. после них
    не идёт , ] } : или пробел/конец — признак конца значения).

    Возвращает исправленный текст. Если ничего не изменилось — исходный текст.
    """
    out = []
    in_string = False
    escape = False
    changed = False
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if escape:
            out.append(ch)
            escape = False
            i += 1
            continue
        if ch == "\\":
            out.append(ch)
            escape = True
            i += 1
            continue
        if ch == '"':
            if in_string:
                # Смотрим, что идёт после кавычки
                nxt = text[i + 1] if i + 1 < n else ""
                # Если после кавычки идёт структурный символ — это конец строки
                if nxt in (",", "]", "}", ":", "\n", "\r", " ", "\t") or nxt == "":
                    in_string = False
                    out.append(ch)
                else:
                    # Кавычка внутри строки — экранируем
                    out.append("\\\"")
                    changed = True
            else:
                in_string = True
                out.append(ch)
            i += 1
            continue
        out.append(ch)
        i += 1
    if changed:
        return "".join(out)
    return text


def extract_json(raw: str):
    """Извлекает JSON-массив или объект из ответа модели.

    Устойчив к:
    - пустому ответу
    - markdown-обёрткам ```json ... ```
    - пояснительному тексту до/после JSON
    - обрезанному JSON (если найдено начало, но нет конца)
    - неэкранированным кавычкам внутри строк (чинит их)

    ВАЖНО: при поиске границ JSON учитывает вложенность скобок и строки,
    чтобы не спутать [ и ] внутри текста вопросов с границами массива.
    """
    if raw is None:
        raise json.JSONDecodeError("Пустой ответ от LLM (None)", "", 0)

    cleaned = raw.strip()
    if not cleaned:
        raise json.JSONDecodeError("Пустой ответ от LLM", "", 0)

    # Убираем markdown-обёртку ```json ... ``` или ``` ... ```
    if cleaned.startswith("```"):
        lines = cleaned.split("\n")
        if lines and lines[0].strip().startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip().startswith("```"):
            lines = lines[:-1]
        cleaned = "\n".join(lines).strip()

    # Попытка 1: парсим как есть
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # Попытка 2: чиним неэкранированные кавычки и пробуем снова
    repaired = _repair_unescaped_quotes(cleaned)
    if repaired != cleaned:
        try:
            return json.loads(repaired)
        except json.JSONDecodeError:
            pass

    # Попытка 3: ищем JSON-массив с правильным учётом скобок и строк
    # Сначала ищем все позиции [ (приоритет массиву)
    array_starts = [i for i, ch in enumerate(cleaned) if ch == "["]
    # Затем позиции { (для одиночного объекта)
    object_starts = [i for i, ch in enumerate(cleaned) if ch == "{"]

    # Пробуем массивы от самых ВНЕШНИХ к внутренним.
    # ВАЖНО: если ответ начинается с "[" (это и есть основной массив вопросов),
    # пробуем его ПЕРВЫМ. Иначе при повреждённом JSON (например, неэкранированная
    # кавычка внутри строки) мы можем случайно распарсить внутренний массив
    # (например, sources: ["..."]), вернув список строк вместо вопросов.
    # Сортируем стартовые позиции: сначала те, что ближе к началу текста,
    # но при этом отдаём приоритет самому первому "[" (внешнему массиву).
    array_starts_sorted = sorted(
        array_starts,
        key=lambda i: (i != array_starts[0] if array_starts else False, i),
    )
    for start in array_starts_sorted:
        end = _find_matching_bracket(cleaned, start, "[", "]")
        if end == -1:
            continue
        candidate = cleaned[start : end + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            continue

    # Пробуем объекты от самых ранних к поздним
    for start in object_starts:
        end = _find_matching_bracket(cleaned, start, "{", "}")
        if end == -1:
            continue
        candidate = cleaned[start : end + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            continue

    # Сохраняем сырой ответ для отладки
    debug_path = _save_raw_for_debug(raw)
    preview = cleaned[:300].replace("\n", " ")
    msg = (
        f"Не удалось извлечь JSON. Длина ответа: {len(cleaned)}. "
        f"Начало: {preview!r}"
    )
    if debug_path:
        msg += f". Сырой ответ сохранён в {debug_path}"
    raise json.JSONDecodeError(msg, cleaned, 0)


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


def _normalize_text(s) -> str:
    """Нормализует текст для нечёткого сравнения.

    - приводит к строке
    - lowercase
    - заменяет ё на е
    - убирает знаки препинания
    - схлопывает пробелы
    - убирает слова-заглушки (ЭТОТ/ЭТОГО/ОН/ОНА/ОНИ/ЕГО/ЕЁ и т.п.),
      которые LLM по-разному подставляет в один и тот же вопрос при
      обработке разных чанков (перекрытие). Это позволяет дедупликации
      находить перефразированные дубликаты на стыках чанков.
    """
    if s is None:
        return ""
    s = str(s).lower().replace("ё", "е").replace("Ё", "Е")
    # Убираем все символы, кроме букв, цифр и пробелов
    s = re.sub(r"[^\w\s]", " ", s)
    # Убираем слова-заглушки (местоимения-указатели), которые LLM
    # подставляет по-разному в разных чанках
    placeholder_words = {
        "этот", "этого", "этому", "этим", "этом", "эти", "этих", "этим",
        "эта", "этой", "эту", "это", "этими",
        "он", "она", "они", "оно", "его", "её", "ее", "их", "ему", "ей",
        "им", "ними", "него", "нее", "неё", "них",
        "та", "тот", "того", "тому", "тем", "том", "те", "тех", "теми",
        "такой", "такого", "такому", "таким", "таком", "такая", "такую",
        "такие", "таких", "такими",
    }
    words = s.split()
    s = " ".join(w for w in words if w not in placeholder_words)
    return s


def _normalize_theme(theme) -> str:
    """Нормализует название темы для группировки."""
    return _normalize_text(theme)


def _fuzzy_ratio(a: str, b: str) -> float:
    """Степень схожести двух строк (0..100).

    Использует rapidfuzz.fuzz.token_sort_ratio, если доступен.
    Fallback на difflib.SequenceMatcher.
    """
    if not a and not b:
        return 100.0
    if not a or not b:
        return 0.0
    try:
        from rapidfuzz import fuzz
        return float(fuzz.token_sort_ratio(a, b))
    except ImportError:
        from difflib import SequenceMatcher
        return SequenceMatcher(None, a, b).ratio() * 100.0


def deduplicate_questions(questions: list, threshold: float = 90.0,
                          verbose: bool = False) -> list:
    """Удаляет дубликаты вопросов, сохраняя порядок.

    Алгоритм:
    1. Группируем вопросы по (нормализованная_тема, price).
    2. Внутри каждой группы сравниваем нормализованный текст вопроса
       нечётким сравнением (token_sort_ratio).
    3. Если схожесть >= threshold — считаем дубликатом.
    4. ДОПОЛНИТЕЛЬНО: вопросы с одинаковой ценой сравниваются глобально
       (без учёта темы). Это ловит дубликаты на стыках чанков, когда LLM
       не распознала тему (theme=null) и вопрос не попал в свою группу.

    Это устойчиво к:
    - разному регистру ("Смех" vs "смех")
    - замене ё/е
    - лишним пробелам и знакам препинания
    - перестановке слов в вопросе
    - мелким отличиям в формулировке

    threshold=95 — почти точное совпадение (рекомендуется 90-98).
    """
    if not questions:
        return questions

    # groups[key] = [(normalized_question_text, original_question), ...]
    groups: dict = {}
    # price_groups[price] = [(normalized_question_text, original_question), ...]
    # Глобальное сравнение по цене — для вопросов с нераспознанной темой.
    price_groups: dict = {}
    result = []
    duplicates_count = 0

    for q in questions:
        theme_key = _normalize_theme(q.get("theme"))
        price = q.get("price")
        key = (theme_key, price)
        text = _normalize_text(q.get("question"))

        bucket = groups.setdefault(key, [])

        is_duplicate = False
        for existing_text, existing_q in bucket:
            score = _fuzzy_ratio(text, existing_text)
            if score >= threshold:
                is_duplicate = True
                duplicates_count += 1
                if verbose:
                    print(
                        f"          [dedup] Дубликат (score={score:.1f}): "
                        f"theme={q.get('theme')!r}, price={price}\n"
                        f"            существующий: {existing_q.get('question', '')[:80]!r}\n"
                        f"            новый:        {q.get('question', '')[:80]!r}"
                    )
                break

        # Если в своей группе не нашли — ищем глобально по цене.
        # Это покрывает случай, когда у дубликата theme=null (LLM не
        # распознала тему на стыке чанков), а у оригинала тема есть.
        if not is_duplicate and price is not None:
            price_bucket = price_groups.setdefault(price, [])
            for existing_text, existing_q in price_bucket:
                score = _fuzzy_ratio(text, existing_text)
                if score >= threshold:
                    is_duplicate = True
                    duplicates_count += 1
                    if verbose:
                        print(
                            f"          [dedup] Дубликат по цене (score={score:.1f}): "
                            f"theme={q.get('theme')!r}, price={price}\n"
                            f"            существующий: {existing_q.get('theme')!r} "
                            f"{existing_q.get('question', '')[:80]!r}\n"
                            f"            новый:        {q.get('question', '')[:80]!r}"
                        )
                    break

        if not is_duplicate:
            bucket.append((text, q))
            if price is not None:
                price_groups.setdefault(price, []).append((text, q))
            result.append(q)

    if verbose and duplicates_count > 0:
        print(f"          [dedup] Удалено дубликатов: {duplicates_count}")

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

    # В кеше нет вопросов, но может быть raw_response от предыдущей неудачной попытки
    cache_entry = load_cache_entry(CACHE_DIR, cache_key)
    if cache_entry and cache_entry.get("raw_response"):
        raw = cache_entry["raw_response"]
        prev_error = cache_entry.get("parse_error", "")
        print(f"          [cache] Чанк {chunk_idx}: найден raw_response в кеше "
              f"(предыдущая ошибка: {prev_error[:80] if prev_error else 'нет'!r})")
        print(f"          [reparse] Пробую перепарсить без запроса к LLM...")
        # Парсим (логика ниже)
        return _parse_and_cache(
            raw=raw,
            cache_key=cache_key,
            text_hash=text_hash,
            prompt_hash=prompt_hash,
            model=model,
            source_label=source_label,
            source_file=source_file,
            chunk_idx=chunk_idx,
            chunk_start=chunk_start,
            chunk_end=chunk_end,
        )

    # Кеша нет — отправляем в LLM
    print(f"          [api]   Чанк {chunk_idx}: отправляю в LLM...")
    raw = call_llm(SYSTEM_PROMPT, chunk_text, model)
    return _parse_and_cache(
        raw=raw,
        cache_key=cache_key,
        text_hash=text_hash,
        prompt_hash=prompt_hash,
        model=model,
        source_label=source_label,
        source_file=source_file,
        chunk_idx=chunk_idx,
        chunk_start=chunk_start,
        chunk_end=chunk_end,
    )


def _parse_and_cache(raw: str, cache_key: str, text_hash: str,
                     prompt_hash: str, model: str, source_label: str,
                     source_file: str, chunk_idx: int,
                     chunk_start: int, chunk_end: int) -> list:
    """Парсит сырой ответ LLM, нормализует и сохраняет в кэш.

    Вынесено в отдельную функцию для переиспользования:
    - при первом получении ответа от LLM
    - при перепарсинге из кэша (raw_response без questions)
    """
    # СРАЗУ сохраняем сырой ответ в кэш, чтобы не потерять его при ошибке парсинга
    save_to_cache(
        CACHE_DIR, cache_key, text_hash, prompt_hash, model,
        questions=[],
        source_file=source_file,
        chunk_index=chunk_idx,
        chunk_start=chunk_start,
        chunk_end=chunk_end,
        raw_response=raw,
        parse_error="in_progress",
    )

    # Парсим ответ
    try:
        parsed = extract_json(raw)
    except Exception as e:
        save_to_cache(
            CACHE_DIR, cache_key, text_hash, prompt_hash, model,
            questions=[],
            source_file=source_file,
            chunk_index=chunk_idx,
            chunk_start=chunk_start,
            chunk_end=chunk_end,
            raw_response=raw,
            parse_error=str(e),
        )
        raise

    # Нормализуем результат к списку словарей
    if isinstance(parsed, list):
        pass
    elif isinstance(parsed, dict):
        found_list = None
        for key in ("questions", "data", "items", "results"):
            if key in parsed and isinstance(parsed[key], list):
                found_list = parsed[key]
                break
        if found_list is None:
            for v in parsed.values():
                if isinstance(v, list):
                    found_list = v
                    break
        if found_list is not None:
            parsed = found_list
        else:
            parsed = [parsed]
    else:
        err_msg = (
            f"Неожиданный тип результата от LLM: {type(parsed).__name__}. "
            f"Значение: {str(parsed)[:200]!r}"
        )
        save_to_cache(
            CACHE_DIR, cache_key, text_hash, prompt_hash, model,
            questions=[],
            source_file=source_file,
            chunk_index=chunk_idx,
            chunk_start=chunk_start,
            chunk_end=chunk_end,
            raw_response=raw,
            parse_error=err_msg,
        )
        raise ValueError(err_msg)

    # Дополнительная проверка: все элементы должны быть словарями
    if not all(isinstance(q, dict) for q in parsed):
        bad = [i for i, q in enumerate(parsed) if not isinstance(q, dict)]
        err_msg = (
            f"Некоторые элементы результата не являются словарями. "
            f"Индексы: {bad[:10]}. Типы: {[type(parsed[i]).__name__ for i in bad[:5]]}"
        )
        save_to_cache(
            CACHE_DIR, cache_key, text_hash, prompt_hash, model,
            questions=[],
            source_file=source_file,
            chunk_index=chunk_idx,
            chunk_start=chunk_start,
            chunk_end=chunk_end,
            raw_response=raw,
            parse_error=err_msg,
        )
        raise ValueError(err_msg)

    print(f"          Получено вопросов: {len(parsed)}")

    # Сохраняем в кэш финальный результат
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
        raw_response=raw,
        parse_error=None,
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

    Вопросы с нераспознанной темой (theme=null/пусто) НЕ создают новую тему:
    им наследуется название текущей темы. Это чинит ситуацию, когда LLM не
    распознала тему на стыке чанков и вернула theme=null для части вопросов.
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
        current_theme = None  # название текущей темы (без номера)
        for q in qs:
            price = q.get("price")
            theme = q.get("theme")

            # Новая тема начинается, если:
            # - это первый вопрос
            # - текущий price=10, а предыдущий был 50
            # - текущий price=None (нет цены) И у вопроса есть своя тема
            # - цена упала (50→20 и т.п.) И у вопроса есть своя тема
            # Вопрос без темы НИКОГДА не начинает новую тему — он наследует
            # текущую (иначе на стыке чанков создаётся фантомная тема).
            has_own_theme = theme is not None and str(theme).strip() != ""
            is_new_theme = (
                prev_price is None
                or price == 10
                or (price is None and has_own_theme)
                or (
                    has_own_theme
                    and prev_price is not None
                    and prev_price >= 50
                    and price < prev_price
                )
            )
            if is_new_theme:
                theme_counter += 1

            if has_own_theme:
                # Проверяем, не начинается ли тема уже с номера ("1. Смех.")
                if re.match(r"^\d+\.\s+", theme):
                    # Тема уже пронумерована — не добавляем номер
                    pass
                else:
                    # Добавляем порядковый номер к названию темы
                    q["theme"] = f"{theme_counter}. {theme}"
                # В current_theme храним ЧИСТОЕ название (без номера),
                # чтобы наследование не давало "85. 86. Дома".
                current_theme = re.sub(r"^\d+\.\s+", "", theme)
            else:
                # Вопрос без темы — наследуем название текущей темы,
                # чтобы не создавать фантомную тему и не ломать нумерацию.
                if current_theme is not None:
                    q["theme"] = f"{theme_counter}. {current_theme}"

            prev_price = price
            result.append(q)

    return result


def validate_questions(questions: list, package_name: str) -> None:
    """Пост-проверка результата парсинга.

    Проверяет:
    - нет ли вопросов с пустой/неопределённой темой (theme=null)
    - нет ли тем с количеством вопросов, не кратным 5
    - нет ли тем с нестандартным набором цен

    Проблемы выводятся в консоль, но не прерывают обработку.
    """
    if not questions:
        print(f"[!] Валидация {package_name}: пустой список вопросов")
        return

    problems = []

    # 1. Вопросы без темы
    no_theme = [q for q in questions if not (q.get("theme") or "").strip()]
    if no_theme:
        problems.append(
            f"вопросов без темы: {len(no_theme)} "
            f"(пример: {no_theme[0].get('question', '')[:60]!r})"
        )

    # 2. Темы с количеством вопросов, не кратным 5
    themes = OrderedDict()
    for q in questions:
        themes.setdefault(q.get("theme") or "(NULL)", []).append(q)
    bad_counts = {t: len(qs) for t, qs in themes.items() if len(qs) % 5 != 0}
    if bad_counts:
        problems.append(
            "тем с числом вопросов, не кратным 5: "
            + ", ".join(f"{t!r} ({n})" for t, n in list(bad_counts.items())[:5])
        )

    # 3. Темы с нестандартным набором цен
    bad_prices = []
    for t, qs in themes.items():
        prices = sorted(q.get("price") for q in qs)
        if prices != [10, 20, 30, 40, 50]:
            bad_prices.append(f"{t!r} ({prices})")
    if bad_prices:
        problems.append(
            "тем с нестандартными ценами: " + ", ".join(bad_prices[:5])
        )

    # 4. Вопросы, где форма удалена из текста (осталась "дыра")
    #    Например: "мороженое в ." вместо "мороженое в ЭТОТ ДЕНЬ НЕДЕЛИ"
    import re as _re
    hole_qs = []
    for q in questions:
        question = q.get("question") or ""
        form = q.get("form")
        if form and str(form).strip():
            form_lower = str(form).strip().lower()
            if form_lower not in question.lower():
                # Форма не найдена в вопросе — проверяем, есть ли "дыра"
                if _re.search(r"\s+\.\s", question) or _re.search(r"\s{2,}\.", question):
                    hole_qs.append(q.get("question", "")[:60])
    if hole_qs:
        problems.append(
            f"вопросов с удалённой формой (дыра в тексте): {len(hole_qs)} "
            f"(пример: {hole_qs[0]!r})"
        )

    if problems:
        print(f"[!] Валидация {package_name}: найдены проблемы:")
        for p in problems:
            print(f"    - {p}")
    else:
        print(f"[*] Валидация {package_name}: OK "
              f"({len(questions)} вопросов, {len(themes)} тем)")


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

    # Пост-проверка: валидация результата
    validate_questions(questions, name)

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
