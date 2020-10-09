from models.row_sheet_excel_class import RowSheetExcel
from models.cell_sheet_excel_class import CellSheetExcel
from models.sheet_excel_class import SheetExcel
from models.workbook_excel_class import WorkbookExcel

from models.type_column_enum import TypeColumn

from services.workbook_excel_service import create_workbook

from datetime import datetime

rowSheet = RowSheetExcel([
    CellSheetExcel('teste', TypeColumn.TEXT, None, False, 20),
    CellSheetExcel(50000, TypeColumn.DECIMAL, '#,##0.00', True, 10),
    CellSheetExcel(datetime.now(), TypeColumn.DATE, 'DD/mm/YYYY', True, 10)
])
sheet = SheetExcel([rowSheet], "teste")
workbook = WorkbookExcel("relatorio.xlsx", [sheet, sheet])


create_workbook(workbook)
