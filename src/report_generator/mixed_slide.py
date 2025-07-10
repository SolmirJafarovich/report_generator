import os

from .utils.table_loader import load_table
from .utils.text_style import apply_text_style
from .utils.units import mm


def add_mixed_slide(self, data):
    layout = self.prs.slide_layouts[6]  # Blank slide
    slide = self.prs.slides.add_slide(layout)

    # === Заголовок ===
    title_shape = slide.shapes.add_textbox(mm(10), mm(5), mm(277), mm(12))
    title_tf = title_shape.text_frame
    title_tf.text = data.get("title", "")
    for paragraph in title_tf.paragraphs:
        apply_text_style(paragraph, self.text_config.title)

    # === Картинка слева ===
    images = data.get("images", [])
    if images:
        image = images[0]
        if os.path.exists(image["path"]):
            slide.shapes.add_picture(image["path"], mm(10), mm(20), width=mm(100))

            caption = image.get("caption")
            if caption:
                tx = slide.shapes.add_textbox(mm(10), mm(85), mm(100), mm(6))
                caption_tf = tx.text_frame
                caption_tf.text = caption
                for paragraph in caption_tf.paragraphs:
                    apply_text_style(paragraph, self.text_config.body)

    # === Таблица справа ===
    tables = data.get("tables", [])
    if tables:
        table_data = tables[0]
        path = table_data.get("path")
        sheet = table_data.get("sheet")
        if os.path.exists(path):
            df = load_table(path, sheet)
            rows, cols = df.shape
            height = mm(10 + rows * 7)

            table_shape = slide.shapes.add_table(
                rows + 1, cols, mm(120), mm(20), mm(150), height
            )
            table = table_shape.table

            # Заголовок таблицы
            for col_idx, col_name in enumerate(df.columns):
                cell = table.cell(0, col_idx)
                cell.text = str(col_name)
                for paragraph in cell.text_frame.paragraphs:
                    apply_text_style(paragraph, self.text_config.table_header)

            # Данные таблицы
            for row_idx, row in enumerate(df.values):
                for col_idx, val in enumerate(row):
                    cell = table.cell(row_idx + 1, col_idx)
                    cell.text = str(val)
                    for paragraph in cell.text_frame.paragraphs:
                        apply_text_style(paragraph, self.text_config.table_cell)
