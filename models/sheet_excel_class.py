class SheetExcel:
    row_sheets = []
    title = None

    def __init__(self, row_sheets, title):
        if(row_sheets is None):
            row_sheets = []
        self.row_sheets = row_sheets
        self.title = title