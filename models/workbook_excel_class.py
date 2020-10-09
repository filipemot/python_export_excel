class WorkbookExcel:
    filename = None
    sheets = []

    def __init__(self, filename, sheets):
        self.filename = filename

        if(sheets is None):
            sheets = []

        self.sheets = sheets