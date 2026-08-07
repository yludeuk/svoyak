"""Патч для parse_docx.py: добавляет метаинформацию в кеш."""
with open("parse_docx.py", "r", encoding="utf-8") as f:
    content = f.read()

# Обновляем save_to_cache — добавляем метаинформацию
old = '''    # Сохраняем в кеш (без служебных полей)
    questions_for_cache = []
    for q in parsed:
        q_clean = {k: v for k, v in q.items() if not k.startswith("_")}
        questions_for_cache.append(q_clean)
    save_to_cache(CACHE_DIR, cache_key, text_hash, prompt_hash, model,
                  questions_for_cache)'''

new = '''    # Сохраняем в кеш (без служебных полей)
    questions_for_cache = []
    for q in parsed:
        q_clean = {k: v for k, v in q.items() if not k.startswith("_")}
        questions_for_cache.append(q_clean)
    save_to_cache(
        CACHE_DIR, cache_key, text_hash, prompt_hash, model,
        questions_for_cache,
        source_file=source_docx,
        chunk_index=chunk_idx,
    )'''

if old not in content:
    print("ОШИБКА: блок save_to_cache не найден")
    raise SystemExit(1)

content = content.replace(old, new)

# Также нужно передать chunk_start и chunk_end в process_chunk_with_cache
# Для этого нужно изменить сигнатуру и вызов
old2 = '''def process_chunk_with_cache(chunk_text: str, chunk_idx: int, model: str,
                             source_label: str, source_docx: str,
                             prompt_hash: str) -> list:
    """Обрабатывает один чанк с использованием кеша.

    Возвращает список вопросов.
    """
    text_hash = compute_text_hash(chunk_text)
    cache_key = make_cache_key(text_hash, prompt_hash, model)'''

new2 = '''def process_chunk_with_cache(chunk_text: str, chunk_idx: int, model: str,
                             source_label: str, source_docx: str,
                             prompt_hash: str, chunk_start: int = 0,
                             chunk_end: int = 0) -> list:
    """Обрабатывает один чанк с использованием кеша.

    Возвращает список вопросов.
    """
    text_hash = compute_text_hash(chunk_text)
    cache_key = make_cache_key(text_hash, prompt_hash, model)'''

if old2 not in content:
    print("ОШИБКА: сигнатура process_chunk_with_cache не найдена")
    raise SystemExit(1)

content = content.replace(old2, new2)

# Обновляем вызов process_chunk_with_cache — передаём start/end
old3 = '''            chunk_questions = process_chunk_with_cache(
                chunk_text=chunk_text,
                chunk_idx=chunk_idx,
                model=model,
                source_label=source_label,
                source_docx=os.path.basename(path),
                prompt_hash=prompt_hash,
            )'''

new3 = '''            chunk_questions = process_chunk_with_cache(
                chunk_text=chunk_text,
                chunk_idx=chunk_idx,
                model=model,
                source_label=source_label,
                source_docx=os.path.basename(path),
                prompt_hash=prompt_hash,
                chunk_start=start,
                chunk_end=end,
            )'''

if old3 not in content:
    print("ОШИБКА: вызов process_chunk_with_cache не найден")
    raise SystemExit(1)

content = content.replace(old3, new3)

# Обновляем save_to_cache — добавляем chunk_start и chunk_end
old4 = '''    save_to_cache(
        CACHE_DIR, cache_key, text_hash, prompt_hash, model,
        questions_for_cache,
        source_file=source_docx,
        chunk_index=chunk_idx,
    )'''

new4 = '''    save_to_cache(
        CACHE_DIR, cache_key, text_hash, prompt_hash, model,
        questions_for_cache,
        source_file=source_docx,
        chunk_index=chunk_idx,
        chunk_start=chunk_start,
        chunk_end=chunk_end,
    )'''

if old4 not in content:
    print("ОШИБКА: блок save_to_cache (с метаданными) не найден")
    raise SystemExit(1)

content = content.replace(old4, new4)

with open("parse_docx.py", "w", encoding="utf-8") as f:
    f.write(content)

print("OK: патч применён")
