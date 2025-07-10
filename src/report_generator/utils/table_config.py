HEADER_ROWS = 1

EMU_PER_INCH = 914400
PT_PER_INCH = 72


def emu_to_pt(emu):
    return emu / EMU_PER_INCH * PT_PER_INCH


def estimate_cell_lines(text: str, chars_per_line: int):
    text = str(text)
    return max(1, text.count("\n") + len(text) // chars_per_line + 1)


def estimate_row_height(row, line_height_pt: float, chars_per_line: int):
    """
    Оценивает высоту строки на основе максимального количества строк текста в ячейке.
    """
    max_lines = max(estimate_cell_lines(cell, chars_per_line) for cell in row)
    return max_lines * line_height_pt


def get_max_rows_per_slide(df, placeholder, text_config):
    """
    Определяет, сколько строк таблицы влезает в указанный placeholder,
    учитывая фактический размер шрифта и ширину ячеек.
    """
    available_width_pt = emu_to_pt(placeholder["width"])
    num_columns = df.shape[1]

    col_width_pt = available_width_pt / max(num_columns, 1)

    style = getattr(text_config, "table_cell", text_config.body)
    font_size = style.font_size
    approx_char_width_pt = font_size * 0.53

    chars_per_line = max(int(col_width_pt / approx_char_width_pt), 10)

    return 1, chars_per_line
