"""Патч для docx_generator.py: обновляет render_block и group_by_theme."""
with open("docx_generator.py", "r", encoding="utf-8") as f:
    content = f.read()

# 1. Обновляем render_block — добавляем автора и комментарий к теме,
#    и пустую строку перед ответом
old = '''    # Темы с вопросами
    for i, (theme_key, questions) in enumerate(block.items()):
        theme_num = theme_start_num + i
        display_name = get_display_name(theme_key, theme_num)

        add_paragraph(doc, f"{theme_num}. {display_name}", bold=True, size=12)
        add_paragraph(doc, "")

        for q in questions:
            price = q.get("price")
            question_text = q.get("question", "").strip()
            answer = q.get("answer", "").strip()
            accepted = q.get("accepted_answers")
            comment = q.get("comment")
            sources = q.get("sources")

            if price is not None:
                add_paragraph(doc, f"{price}. {question_text}")
            else:
                add_paragraph(doc, question_text)

            if answer:
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

        add_paragraph(doc, "")'''

new = '''    # Темы с вопросами
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
        add_paragraph(doc, f"{theme_num}. {display_name}", bold=True, size=12)

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

            if price is not None:
                add_paragraph(doc, f"{price}. {question_text}")
            else:
                add_paragraph(doc, question_text)

            # Пустая строка перед ответом
            if answer:
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

        add_paragraph(doc, "")'''

if old not in content:
    print("ОШИБКА: блок render_block не найден")
    raise SystemExit(1)

content = content.replace(old, new)

with open("docx_generator.py", "w", encoding="utf-8") as f:
    f.write(content)

print("OK: render_block обновлён")
