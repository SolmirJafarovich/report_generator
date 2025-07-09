import math
import os

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

from .utils.table_config import (
    HEADER_ROWS,
    get_max_rows_per_slide,
)
from .utils.table_loader import load_table


def add_table_slide(self, data):
    layout = self.prs.slide_layouts[1]  # Title and Content

    for table_data in data.get("tables", []):
        path = table_data.get("path")
        sheet = table_data.get("sheet")
        if not os.path.exists(path):
            print(f"[!] Таблица не найдена: {path}")
            continue

        df = load_table(path, sheet)
        max_rows_per_slide = get_max_rows_per_slide(df)

        rows, cols = df.shape
        total_parts = math.ceil(rows / max_rows_per_slide)

        for part_idx in range(total_parts):
            slide = self.prs.slides.add_slide(layout)
            title_shape = slide.shapes.title
            title_text = data.get("title", "Таблица")
            if total_parts > 1:
                title_text += f" (Часть {part_idx + 1}/{total_parts})"

            # Заголовок слайда
            title_shape.text = title_text
            for paragraph in title_shape.text_frame.paragraphs:
                paragraph.alignment = {
                    "center": PP_ALIGN.CENTER,
                    "left": PP_ALIGN.LEFT,
                    "right": PP_ALIGN.RIGHT,
                }.get(self.text_config.title.text_align, PP_ALIGN.LEFT)
                for run in paragraph.runs:
                    run.font.name = self.text_config.title.font_name
                    run.font.size = Pt(self.text_config.title.font_size)
                    run.font.bold = self.text_config.title.bold
                    run.font.italic = self.text_config.title.italic
                    run.font.color.rgb = RGBColor.from_string(
                        self.text_config.title.color
                    )

            # Подтаблица
            start = part_idx * max_rows_per_slide
            end = min(start + max_rows_per_slide, rows)
            df_chunk = df.iloc[start:end]

            placeholder = slide.placeholders[1]
            table_shape = slide.shapes.add_table(
                len(df_chunk) + HEADER_ROWS,
                cols,
                placeholder.left,
                placeholder.top,
                placeholder.width,
                placeholder.height,
            ).table

            # Заголовки
            for col_idx, col_name in enumerate(df.columns):
                cell = table_shape.cell(0, col_idx)
                cell.text = str(col_name)
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.alignment = {
                        "center": PP_ALIGN.CENTER,
                        "left": PP_ALIGN.LEFT,
                        "right": PP_ALIGN.RIGHT,
                    }.get(self.text_config.table_header.text_align, PP_ALIGN.LEFT)
                    for run in paragraph.runs:
                        style = getattr(
                            self.text_config, "table_header", self.text_config.title
                        )
                        run.font.name = style.font_name
                        run.font.size = Pt(style.font_size)
                        run.font.bold = style.bold
                        run.font.italic = style.italic
                        run.font.color.rgb = RGBColor.from_string(style.color)

            # Ячейки
            for row_idx, row in enumerate(df_chunk.values):
                for col_idx, val in enumerate(row):
                    cell = table_shape.cell(row_idx + HEADER_ROWS, col_idx)
                    cell.text = str(val)
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.alignment = {
                            "center": PP_ALIGN.CENTER,
                            "left": PP_ALIGN.LEFT,
                            "right": PP_ALIGN.RIGHT,
                        }.get(self.text_config.table_cell.text_align, PP_ALIGN.LEFT)
                        for run in paragraph.runs:
                            style = getattr(
                                self.text_config, "table_cell", self.text_config.body
                            )
                            run.font.name = style.font_name
                            run.font.size = Pt(style.font_size)
                            run.font.bold = style.bold
                            run.font.italic = style.italic
                            run.font.color.rgb = RGBColor.from_string(style.color)

            col_max_char = [0] * cols
            for row in df_chunk.values:
                for col_idx, val in enumerate(row):
                    col_max_char[col_idx] = max(col_max_char[col_idx], len(str(val)))
            for col_idx, col_name in enumerate(df.columns):
                col_max_char[col_idx] = max(col_max_char[col_idx], len(str(col_name)))

            col_weights = [math.log(1 + c) for c in col_max_char]
            total_weight = sum(col_weights)
            MIN_COL_WIDTH_PCT = 0.08

            for col_idx in range(cols):
                ratio = col_weights[col_idx] / total_weight
                ratio = max(ratio, MIN_COL_WIDTH_PCT)
                table_shape.columns[col_idx].width = int(placeholder.width * ratio)
