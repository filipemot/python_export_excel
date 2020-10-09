class RowSheetExcel:
    cell_sheets = []

    def __init__(self, cell_sheets):
        if(cell_sheets is None):
            cell_sheets = []

        self.cell_sheets = cell_sheets