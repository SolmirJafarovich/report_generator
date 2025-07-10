from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt


def apply_text_style(paragraph, style):
    """
    Применяет стиль к абзацу: шрифт, размер, жирность, курсив, цвет, выравнивание.
    """
    paragraph.alignment = {
        "center": PP_ALIGN.CENTER,
        "left": PP_ALIGN.LEFT,
        "right": PP_ALIGN.RIGHT,
    }.get(getattr(style, "text_align", "left"), PP_ALIGN.LEFT)

    for run in paragraph.runs:
        run.font.name = style.font_name
        run.font.size = Pt(style.font_size)
        run.font.bold = style.bold
        run.font.italic = style.italic
        run.font.color.rgb = RGBColor.from_string(style.color)
