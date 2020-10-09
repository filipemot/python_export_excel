class CellSheetExcel:
    value = None
    type_column = None
    format_value = None
    font_bold = False
    font_size = 10

    def __init__(self, value, type_column, format_value, font_bold=False, font_size=10):
        self.value = value
        self.type_column = type_column
        self.format_value = format_value
        self.bold = font_bold
        self.font_size = font_size