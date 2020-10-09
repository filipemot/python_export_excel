### Introdução 

Biblioteca para a criação de planilhas excel utilizando python. Demonstrarei como fazer a codificação e disponibilizarei uma pequena biblioteca para ajudar o desenvolvimento.

### Introdução 
`pip install openpyxl`

**Se estiver utilizando Python 2**
`pip install enum34`

### Codificação

#### Criei um enum para os tipos de campos que permitiremos criar

```python
class TypeColumn(Enum):
    TEXT = 1
    DECIMAL = 2
    DATE = 3
```

#### Crie uma classe para representar um objeto de workbook

**Atributos**
- Caminho onde será gravado a planilha
- Array de objeto do tipo SheetExcel com os dados das sheets que serão criadas


```python
class WorkbookExcel:
    filename = None
    sheets = []
    def __init__(self, filename, sheets):
        self.filename = filename
        if(sheets is None):
            sheets = []
        self.sheets = sheets
```
Crie uma classe para representar um objeto de sheet

**Atributos**
- Array de objeto do tipo RowSheetExcel, com a representação das linhas da planilha
- Campo para o título da sheet

```python
class SheetExcel:
    row_sheets = []
    title = None
    def __init__(self, row_sheets, title):
        if(row_sheets is None):
            row_sheets = []
        self.row_sheets = row_sheets
        self.title = title
```

#### Crie uma classe para representar um objeto as linhas da planilha

**Atributos**
- Array de objeto do tipoCellSheetExcel, com a representação das células de cada linha

```python
class RowSheetExcel:
    cell_sheets = []
    def __init__(self, cell_sheets):
        if(cell_sheets is None):
            cell_sheets = []
        self.cell_sheets = cell_sheets
```

#### Crie uma classe para representar um objeto das células de cada linha

**Atributos**
- Valor,
- Tipo da coluna representado pelo enum TypeColumn,
- Formato da célula esse parâmetro é utilizado apenas se o tipo for Decimal ou Data
- True/False para campo Negrito
- Tamanho da Fonte

```python
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
```
#### Função para buscar a "aba" ativa
```python
def get_active_sheet(active_sheet, sheet, file_workbook):
   """
    Get active sheet
    :param Object | active_sheet:
        The object with the instance of Active Sheet
    :param Object{SheetExcel} | sheet:
        The object of type SheetExcel with information of sheet
    :param Object{Workbook} | file_workbook:
        The object of type Workbook with information of woorkbook
    """
    try:
        if(active_sheet is None):
            active_sheet = file_workbook.active
            active_sheet.title = sheet.title
        else:
            active_sheet = file_workbook.create_sheet(sheet.title)


        return active_sheet
    except:
        print('get_active_sheet() - workbook_excel_service', 'Error')
```

#### Função para criar uma célula

```python
def create_cell(active_sheet, cell, id_row, id_column):
    """
    Create a cell of sheet
    :param Object | active_sheet:
        The object with the instance of Active Sheet
    :param Object{CellSheetExcel} | cell:
        The object of type CellSheetExcel with information of cell
    :param int | id_row:
        The index of row
    :param int | id_column:
        The index of column
    """
    try:
        active_cell = active_sheet.cell(row=id_row, column=id_column)
        create_value(active_cell, cell)
    except:
        print('create_cell() - workbook_excel_service', 'Error')
```
#### Função para criar um valor em uma célula
```python
def create_value(active_cell, cell):
    """
    Create value of cell
    :param Object | active_sheet:
        The object with the instance of Active Sheet
    :param Object{CellSheetExcel} | cell:
        The object of type CellSheetExcel with information of cell
    """
    try:
        active_cell.font = (Font(bold=cell.font_bold, size=cell.font_size))
        if(cell.type_column == TypeColumn.DATE or cell.type_column == TypeColumn.DECIMAL):
            active_cell.number_format = cell.format_value
        active_cell.value = cell.value
    except:
        print('create_value() - workbook_excel_service', 'Error')
```
#### Função que cria uma linha
```python
def create_row(row, active_sheet, id_row):
    """
    Create a row of sheet
    :param Object{RowSheetExcel} | row:
        The object of type RowSheetExcel with information of row  
    :param Object | active_sheet:
        The object with the instance of Active Sheet
    :param int | id_row:
        The index of row
    """
    try:
        id_column = 1
        for cell in row.cell_sheets:
            create_cell(active_sheet, cell, id_row, id_column)
            id_column = id_column + 1
    except:
        print('create_row() - workbook_excel_service', 'Error')
```
#### Função que cria uma "aba"

```python
def create_sheet(file_workbook, active_sheet, sheet, id_row):
    """
    Create a sheet
    :param Object{Workbook} | file_workbook:
        The object of type Workbook with information of woorkbook    
    :param Object | active_sheet:
        The object with the instance of Active Sheet
    :param Object{SheetExcel} | sheet:
        The object of type SheetExcel with information of sheet        
    :param int | id_row:
        The index of row
    """
    try:
        active_sheet = get_active_sheet(active_sheet, sheet, file_workbook)


        for row in sheet.row_sheets:
            create_row(row, active_sheet, id_row)


        return active_sheet
    except:
        print('create_sheet() - workbook_excel_service', 'Error')
```

#### Função que cria um workbook
```python
def create_workbook(workbook):
    """
    Create a workbook
    :param Object{WorkbookExcel} | workbook:
        The object of type WorkbookExcel with information of woorkbook    
    """
    try:
        file_workbook = Workbook()
        id_row = 0
        active_sheet = None
        for sheet in workbook.sheets:
            id_row = 1
            active_sheet = create_sheet(file_workbook, active_sheet, sheet, id_row)
        file_workbook.save(workbook.filename)
    except:
        print('create_workbook() - workbook_excel_service', 'Error')

```

# Exemplo
```python
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

```
