from .utils.text_style import apply_text_style


def add_text_slide(self, data):
    layout = self.prs.slide_layouts[1]  # Title and Content
    slide = self.prs.slides.add_slide(layout)

    # === Заголовок ===
    title_shape = slide.shapes.title
    title_shape.text = data.get("title", "")
    for paragraph in title_shape.text_frame.paragraphs:
        apply_text_style(paragraph, self.text_config.title)

    # === Подзаголовок/основной текст ===
    subtitle_shape = slide.placeholders[1]
    subtitle_shape.text = data.get("subtitle", "")
    for paragraph in subtitle_shape.text_frame.paragraphs:
        apply_text_style(paragraph, self.text_config.body)

    # === Заметки ===
    if "notes" in data:
        notes_shape = slide.notes_slide.notes_text_frame
        notes_shape.text = data["notes"]
        for paragraph in notes_shape.paragraphs:
            apply_text_style(paragraph, self.text_config.body)
