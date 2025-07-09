from .units import mm

# Размеры placeholder'а "Content" в layout "Title and Content"
PLACEHOLDER_LEFT = mm(25)
PLACEHOLDER_TOP = mm(55)
PLACEHOLDER_WIDTH = mm(260)
PLACEHOLDER_HEIGHT = mm(135)

# Настройки строки таблицы
ROW_HEIGHT_PT = 18
HEADER_ROWS = 1
CHARS_PER_LINE = 30  # Средняя ширина строки
LINE_HEIGHT_PT = 18  # Высота одной строки текста

# Преобразование EMU в pt (используется в расчётах)
EMU_PER_INCH = 914400
PT_PER_INCH = 72


def emu_to_pt(emu):
    return emu / EMU_PER_INCH * PT_PER_INCH


def estimate_cell_lines(text: str, chars_per_line=CHARS_PER_LINE):
    text = str(text)
    return max(1, text.count("\n") + len(text) // chars_per_line + 1)


def estimate_row_height(row, line_height_pt=LINE_HEIGHT_PT):
    max_lines = max(estimate_cell_lines(cell) for cell in row)
    return max_lines * line_height_pt


def get_max_rows_per_slide(df):
    """
    Оценка максимального количества строк таблицы, которые могут поместиться на слайде,
    с учётом многострочного текста.
    """
    available_height_pt = emu_to_pt(PLACEHOLDER_HEIGHT)
    total = ROW_HEIGHT_PT  # заголовок
    count = 0
    for row in df.values:
        row_height = estimate_row_height(row)
        if total + row_height > available_height_pt:
            break
        total += row_height
        count += 1

    return max(count, 1)
