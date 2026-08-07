"""Модуль генерации docx файлов из списка вопросов.

Содержит функции для группировки вопросов по темам, разбиения на блоки
и рендеринга в docx с правильной кодировкой и языком.
"""
import re
import sys
from collections import OrderedDict

from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn

from doc_stuff import (
    fix_settings_theme_lang,
    set_docdefaults_language_ru,
    set_normal_style_language_ru,
)

sys.stdout.reconfigure(encoding="utf-8")


def group_by_theme(questions: list) -> "OrderedDict[str, list]":
    """Группирует вопросы по темам, сохраняя порядок.

    Использует (source_file, theme) как ключ, чтобы темы с одинаковым
    названием из разных файлов не объединялись.

    NULL-темы разбиваются на подтемы по 5 вопросов (так как в исходных
    файлах NULL — это секретные темы, каждая по 5 вопросов).
    """
    themes = OrderedDict()
    null_counter = 0
    null_buffer = []

    def flush_null(source_file: str):
        nonlocal null_counter, null_buffer
        if null_buffer:
            for i in range(0, len(null_buffer), 5):
                chunk = null_buffer[i : i + 5]
                null_counter += 1
                key = f"{source_file}::__secret_{null_counter}"
                themes[key] = chunk
            null_buffer = []

    current_source = None
    for q in questions:
        sf = q.get("_source_file", "?")
        if sf != current_source:
            if null_buffer:
                flush_null(current_source or "?")
            current_source = sf

        theme = q.get("theme")
        if theme is None or theme == "":
            null_buffer.append(q)
        else:
            if null_buffer:
                flush_null(current_source)
                null_buffer = []
            key = f"{sf}::{theme}"
            themes.setdefault(key, []).append(q)

    flush_null(current_source or "?")
    return themes


def split_themes(themes, user_split):
    """Разбивает OrderedDict тем на блоки согласно user_split.

    user_split — список целых чисел, например [10, 10, 10, 10, 10, 10, 12].
    Сумма элементов должна быть равна len(themes).
    Возвращает список OrderedDict-ов.
    """
    total = len(themes)
    items = list(themes.items())
    blocks, idx = [], 0
    for sz in user_split:
        blocks.append(OrderedDict(items[idx : idx + sz]))
        idx += sz
    if idx != total:
        raise ValueError(
            f"Сумма user_split ({idx}) не равна количеству тем ({total})"
        )
    return blocks


def add_paragraph(doc: Document, text: str, bold: bool = False, size: int = 11):
    """Добавляет параграф с заданным текстом и форматированием."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(11)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    run.bold = bold


def get_display_name(theme_key: str, theme_num: int) -> str:
    """Возвращает отображаемое имя темы."""
    if "::" in theme_key:
        _, theme_name = theme_key.split("::", 1)
    else:
        theme_name = theme_key
    if theme_name.startswith("__secret_"):
        return f"Тема-{theme_num} (секрет)"
    return theme_name


def render_block(doc: Document, block, block_num: int,
                 total_blocks: int, theme_start_num: int, package_name: str):
    """Рендерит один блок тем в документ."""
    # Заголовок документа
    add_paragraph(
        doc,
        f"Пакет «{package_name}». Темы {theme_start_num}–"
        f"{theme_start_num + len(block) - 1}.",
        bold=True,
        size=14,
    )
    add_paragraph(doc, "")

    # Список тем в начале
    add_paragraph(doc, "Список тем:", bold=True, size=12)
    for i, theme_key in enumerate(block.keys()):
        theme_num = theme_start_num + i
        display_name = get_display_name(theme_key, theme_num)
        # Если display_name уже начинается с номера, не добавляем ещё один
        if re.match(r"^\d+\.\s+", display_name):
            add_paragraph(doc, display_name)
        else:
            add_paragraph(doc, f"{theme_num}. {display_name}")
    add_paragraph(doc, "")

    # Темы с вопросами
    for i, (theme_key, questions) in enumerate(block.items()):
        theme_num = theme_start_num + i
        display_name = get_display_name(theme_key, theme_num)

        # Берём автора и комментарий к теме из первого вопроса темы
        author = None
        theme_comment = None
        if questions:
            author = questions[0].get("author")
            theme_comment = questions[0].get("theme_comment")

        # Заголовок темы
        add_paragraph(doc, f"{display_name}", bold=True, size=12)

        # Автор (опционально)
        if author:
            add_paragraph(doc, f"Автор: {author}")

        # Комментарий к теме (опционально)
        if theme_comment:
            add_paragraph(doc, f"Комментарий к теме: {theme_comment}")

        add_paragraph(doc, "")

        for q in questions:
            price = q.get("price")
            question_text = q.get("question", "").strip()
            answer = q.get("answer", "").strip()
            accepted = q.get("accepted_answers")
            comment = q.get("comment")
            sources = q.get("sources")
            form = (q.get("form") or "").strip()

            if price is not None:
                add_paragraph(doc, f"{price}. {question_text}")
            else:
                add_paragraph(doc, question_text)

            # Форма (опционально) — пустая строка перед формой,
            # между формой и ответом пустой строки нет
            if form:
                add_paragraph(doc, "")
                add_paragraph(doc, f"Форма: {form}")
                if answer:
                    add_paragraph(doc, f"Ответ: {answer}")
            elif answer:
                # Пустая строка перед ответом (если формы нет)
                add_paragraph(doc, "")
                add_paragraph(doc, f"Ответ: {answer}")

            if accepted and len(accepted) > 0:
                add_paragraph(doc, f"Зачёт: {', '.join(accepted)}")

            if comment:
                add_paragraph(doc, f"Комментарий: {comment}")

            if sources and len(sources) > 0:
                if len(sources) == 1:
                    add_paragraph(doc, f"Источник: {sources[0]}")
                else:
                    add_paragraph(doc, f"Источники: {', '.join(sources)}")

            add_paragraph(doc, "")

        add_paragraph(doc, "")


def generate_docx_files(questions: list, package_name: str, user_split,
                        output_dir: str = "output") -> list:
    """Генерирует docx файлы из списка вопросов.

    Возвращает список путей к созданным файлам.
    """
    import os

    os.makedirs(output_dir, exist_ok=True)

    print(f"[*] Группирую {len(questions)} вопросов по темам...")
    themes = group_by_theme(questions)
    print(f"[*] Всего тем: {len(themes)}")

    print(f"[*] Разбиваю на блоки: {user_split}")
    blocks = split_themes(themes, user_split)
    print(f"[*] Получено блоков: {len(blocks)}")
    for i, b in enumerate(blocks, 1):
        print(f"    Блок {i}: {len(b)} тем")

    created_files = []
    theme_counter = 1
    for i, block in enumerate(blocks, 1):
        first = theme_counter
        last = theme_counter + len(block) - 1
        fname = f"{package_name}-{i}-{first}..{last}.docx"
        fpath = os.path.join(output_dir, fname)

        doc = Document()
        fix_settings_theme_lang(doc)
        set_docdefaults_language_ru(doc)
        set_normal_style_language_ru(doc)
        section = doc.sections[0]
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
        render_block(doc, block, i, len(blocks), first, package_name)
        doc.save(fpath)
        print(f"[*] Сохранён {fpath}")
        created_files.append(fpath)
        theme_counter = last + 1

    print(f"\n[*] Готово! Создано {len(blocks)} файлов в {output_dir}/")
    return created_files
