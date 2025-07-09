class TextStyle:
    """
    Один стиль текста: шрифт, размер, жирность, курсив, цвет.
    """

    def __init__(
        self,
        font_name="Arial",
        font_size=18,
        bold=False,
        italic=False,
        color="000000",
        text_align="left",
    ):
        self.font_name = font_name
        self.font_size = font_size  # в поинтах
        self.bold = bold
        self.italic = italic
        self.color = color  # HEX без #
        self.text_align = text_align


class TextConfig:
    """
    Общий конфиг для всех текстовых элементов с разумными дефолтами.
    """

    def __init__(
        self,
        title: TextStyle = None,
        body: TextStyle = None,
        table_header: TextStyle = None,
        table_cell: TextStyle = None,
        notes: TextStyle = None,
    ):
        # Заголовки слайдов и таблиц — крупнее и жирнее, цвет тёмно-синий
        self.title = (
            title
            if title
            else TextStyle(
                font_name="Calibri",
                font_size=28,
                bold=True,
                italic=False,
                color="002060",
                text_align="center",
            )
        )

        # Основной текст, подписи — читаемый средний размер, чёрный или тёмно-серый
        self.body = (
            body
            if body
            else TextStyle(
                font_name="Calibri",
                font_size=16,
                bold=False,
                italic=False,
                color="333333",
            )
        )

        # Заголовки таблиц — чуть меньше и жирнее, черный цвет
        self.table_header = (
            table_header
            if table_header
            else TextStyle(
                font_name="Calibri",
                font_size=14,
                bold=True,
                italic=False,
                color="000000",
            )
        )

        # Ячейки таблиц — стандартный читаемый размер, без жирности, черный цвет
        self.table_cell = (
            table_cell
            if table_cell
            else TextStyle(
                font_name="Calibri",
                font_size=12,
                bold=False,
                italic=False,
                color="000000",
            )
        )

        # Заметки слайда — обычно курсив, мелкий шрифт, серый цвет
        self.notes = (
            notes
            if notes
            else TextStyle(
                font_name="Calibri",
                font_size=12,
                bold=False,
                italic=True,
                color="555555",
            )
        )
