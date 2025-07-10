from .utils.text_style import apply_text_style


def add_text_slide(self, data):
    layout = self.prs.slide_layouts[1]  # Title and Content
    slide = self.prs.slides.add_slide(layout)

    # === Заголовок ===
    title_shape = slide.shapes.title
    title_shape.text = data.get("title", "")
    for paragraph in title_shape.text_frame.paragraphs:
        apply_text_style(paragraph, self.text_config.title)

    # === Блоки в подзаголовке ===
    subtitle_shape = slide.placeholders[1]
    text_frame = subtitle_shape.text_frame
    text_frame.clear()  # Очищаем placeholder

    blocks = data.get("blocks", [])

    for block in blocks:
        paragraph = text_frame.add_paragraph()
        paragraph.text = block.get("text", "")
        style_name = block.get("style", "body")
        style = getattr(self.text_config, style_name, self.text_config.body)
        apply_text_style(paragraph, style)
