"""Модуль кеширования результатов обработки файлов через LLM.

Кеш хранится в отдельной папке (по умолчанию `cache/`).
Ключ кеша — SHA256 хэш содержимого чанка + хэш системного промпта + имя модели.
Это гарантирует, что при изменении файла, промпта или модели кеш будет
пересоздан.

Формат кеш-файла (JSON):
{
    "source_hash": "sha256...",
    "prompt_hash": "sha256...",
    "model": "model-name",
    "source_file": "имя_исходного_файла",
    "chunk_index": 1,
    "chunk_start": 0,
    "chunk_end": 1000,
    "created_at": "ISO 8601 timestamp",
    "questions": [ ... ]
}
"""
import hashlib
import json
import os
import sys
from datetime import datetime, timezone

sys.stdout.reconfigure(encoding="utf-8")


CACHE_DIR = "cache"


def compute_file_hash(path: str) -> str:
    """Вычисляет SHA256 хэш содержимого файла."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def compute_text_hash(text: str) -> str:
    """Вычисляет SHA256 хэш строки."""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def compute_prompt_hash(prompt: str) -> str:
    """Вычисляет SHA256 хэш системного промпта."""
    return hashlib.sha256(prompt.encode("utf-8")).hexdigest()


def make_cache_key(source_hash: str, prompt_hash: str, model: str) -> str:
    """Формирует составной ключ кеша."""
    combined = f"{source_hash}|{prompt_hash}|{model}"
    return hashlib.sha256(combined.encode("utf-8")).hexdigest()


def get_cache_path(cache_dir: str, cache_key: str) -> str:
    """Возвращает путь к кеш-файлу."""
    return os.path.join(cache_dir, f"{cache_key}.json")


def load_from_cache(cache_dir: str, cache_key: str) -> list | None:
    """Загружает вопросы из кеша. Возвращает None, если кеш отсутствует."""
    path = get_cache_path(cache_dir, cache_key)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("questions", [])
    except (json.JSONDecodeError, OSError) as e:
        print(f"      [!] Ошибка чтения кеша {path}: {e}")
        return None


def save_to_cache(cache_dir: str, cache_key: str, source_hash: str,
                  prompt_hash: str, model: str, questions: list,
                  source_file: str = "", chunk_index: int = 0,
                  chunk_start: int = 0, chunk_end: int = 0) -> str:
    """Сохраняет вопросы в кеш. Возвращает путь к кеш-файлу."""
    os.makedirs(cache_dir, exist_ok=True)
    path = get_cache_path(cache_dir, cache_key)
    payload = {
        "source_hash": source_hash,
        "prompt_hash": prompt_hash,
        "model": model,
        "source_file": source_file,
        "chunk_index": chunk_index,
        "chunk_start": chunk_start,
        "chunk_end": chunk_end,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "questions": questions,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return path


def get_cache_metadata(cache_dir: str, cache_key: str) -> dict | None:
    """Возвращает метаинформацию о кеш-файле без загрузки вопросов."""
    path = get_cache_path(cache_dir, cache_key)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {
            "source_file": data.get("source_file", ""),
            "chunk_index": data.get("chunk_index", 0),
            "chunk_start": data.get("chunk_start", 0),
            "chunk_end": data.get("chunk_end", 0),
            "created_at": data.get("created_at", ""),
            "questions_count": len(data.get("questions", [])),
        }
    except (json.JSONDecodeError, OSError):
        return None


def list_cache_files(cache_dir: str = CACHE_DIR) -> list:
    """Возвращает список всех кеш-файлов с метаинформацией."""
    if not os.path.isdir(cache_dir):
        return []
    result = []
    for fname in sorted(os.listdir(cache_dir)):
        if not fname.endswith(".json"):
            continue
        cache_key = fname[:-5]  # убираем .json
        meta = get_cache_metadata(cache_dir, cache_key)
        if meta:
            meta["cache_key"] = cache_key
            meta["path"] = os.path.join(cache_dir, fname)
            result.append(meta)
    return result


def clear_cache(cache_dir: str = CACHE_DIR) -> int:
    """Удаляет все кеш-файлы. Возвращает количество удалённых."""
    if not os.path.isdir(cache_dir):
        return 0
    count = 0
    for fname in os.listdir(cache_dir):
        if fname.endswith(".json"):
            os.remove(os.path.join(cache_dir, fname))
            count += 1
    return count


def delete_cache_for_source(cache_dir: str, source_file: str) -> int:
    """Удаляет кеш-файлы для конкретного исходного файла.

    Возвращает количество удалённых файлов.
    """
    if not os.path.isdir(cache_dir):
        return 0
    count = 0
    for fname in os.listdir(cache_dir):
        if not fname.endswith(".json"):
            continue
        cache_key = fname[:-5]
        meta = get_cache_metadata(cache_dir, cache_key)
        if meta and meta.get("source_file") == source_file:
            os.remove(os.path.join(cache_dir, fname))
            count += 1
    return count
