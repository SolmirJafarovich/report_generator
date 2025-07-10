# src/report_generator/utils/table_config.py

# Высота заголовков таблицы (заранее 1 строка)
HEADER_ROWS = 1
CHARS_PER_LINE = 30  # Средняя длина строки

EMU_PER_INCH = 914400
PT_PER_INCH = 72


def emu_to_pt(emu):
    return emu / EMU_PER_INCH * PT_PER_INCH


def estimate_cell_lines(text: str, chars_per_line=CHARS_PER_LINE):
    text = str(text)
    return max(1, text.count("\n") + len(text) // chars_per_line + 1)


def estimate_row_height(row, line_height_pt: float):
    """
    Оценивает высоту строки на основе максимального количества строк текста в ячейке.
    """
    max_lines = max(estimate_cell_lines(cell) for cell in row)
    return max_lines * line_height_pt


def get_max_rows_per_slide(df, placeholder_height, text_config):
    """
    Определяет, сколько строк таблицы влезает в указанный placeholder,
    учитывая фактический размер шрифта в текстовой конфигурации.
    """
    available_height_pt = emu_to_pt(placeholder_height)

    # Получаем высоту строки из text_config
    style = getattr(text_config, "table_cell", text_config.body)
    line_height_pt = style.font_size

    total = line_height_pt  # заголовок
    count = 0

    for row in df.values:
        row_height = estimate_row_height(row, line_height_pt)
        if total + row_height > available_height_pt:
            break
        total += row_height
        count += 1

    return max(count, 1)
