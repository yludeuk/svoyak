"""Скрипт для управления кешем обработки файлов.

Использование:
    python cache_manager.py list                    # показать все кеш-файлы
    python cache_manager.py clear                   # очистить весь кеш
    python cache_manager.py delete <source_file>    # удалить кеш для конкретного файла
    python cache_manager.py info <source_file>      # показать инфо о кеше файла
"""
import sys
import os

from cache import (
    list_cache_files,
    clear_cache,
    delete_cache_for_source,
    CACHE_DIR,
)

sys.stdout.reconfigure(encoding="utf-8")


def cmd_list():
    """Показывает список всех кеш-файлов с метаинформацией."""
    files = list_cache_files(CACHE_DIR)
    if not files:
        print(f"Кеш пуст ({CACHE_DIR}/)")
        return

    print(f"Найдено {len(files)} кеш-файлов в {CACHE_DIR}/:\n")
    print(f"{'Source file':<50} {'Chunk':<8} {'Lines':<12} {'Questions':<10} {'Created'}")
    print("-" * 110)

    # Группируем по source_file
    by_source = {}
    for f in files:
        sf = f.get("source_file", "?")
        by_source.setdefault(sf, []).append(f)

    for sf, items in sorted(by_source.items()):
        for item in sorted(items, key=lambda x: x.get("chunk_index", 0)):
            sf_display = sf[:48] + ".." if len(sf) > 50 else sf
            chunk = item.get("chunk_index", 0)
            start = item.get("chunk_start", 0)
            end = item.get("chunk_end", 0)
            lines = f"{start + 1}-{end}" if start or end else "-"
            qcount = item.get("questions_count", 0)
            created = item.get("created_at", "")[:19]
            print(f"{sf_display:<50} {chunk:<8} {lines:<12} {qcount:<10} {created}")


def cmd_clear():
    """Очищает весь кеш."""
    count = clear_cache(CACHE_DIR)
    print(f"Удалено {count} кеш-файлов из {CACHE_DIR}/")


def cmd_delete(source_file: str):
    """Удаляет кеш для конкретного исходного файла."""
    count = delete_cache_for_source(CACHE_DIR, source_file)
    if count == 0:
        print(f"Кеш для файла '{source_file}' не найден")
    else:
        print(f"Удалено {count} кеш-файлов для '{source_file}'")


def cmd_info(source_file: str):
    """Показывает информацию о кеше для конкретного файла."""
    files = list_cache_files(CACHE_DIR)
    matching = [f for f in files if f.get("source_file") == source_file]
    if not matching:
        print(f"Кеш для файла '{source_file}' не найден")
        return

    print(f"Кеш для '{source_file}': {len(matching)} чанков\n")
    for item in sorted(matching, key=lambda x: x.get("chunk_index", 0)):
        chunk = item.get("chunk_index", 0)
        start = item.get("chunk_start", 0)
        end = item.get("chunk_end", 0)
        qcount = item.get("questions_count", 0)
        created = item.get("created_at", "")[:19]
        cache_key = item.get("cache_key", "")
        print(f"  Чанк {chunk}: строки {start + 1}-{end}, {qcount} вопросов, {created}")
        print(f"    cache_key: {cache_key}")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    cmd = sys.argv[1].lower()

    if cmd == "list":
        cmd_list()
    elif cmd == "clear":
        confirm = input("Точно очистить весь кеш? (yes/no): ")
        if confirm.lower() == "yes":
            cmd_clear()
        else:
            print("Отменено")
    elif cmd == "delete":
        if len(sys.argv) < 3:
            print("Использование: python cache_manager.py delete <source_file>")
            return
        source_file = sys.argv[2]
        cmd_delete(source_file)
    elif cmd == "info":
        if len(sys.argv) < 3:
            print("Использование: python cache_manager.py info <source_file>")
            return
        source_file = sys.argv[2]
        cmd_info(source_file)
    else:
        print(f"Неизвестная команда: {cmd}")
        print(__doc__)


if __name__ == "__main__":
    main()
