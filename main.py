import win32com.client
import pythoncom
import time
import openpyxl as opxl
from PyQt5 import QtWidgets
from math import modf
import pyacadcom


# * functions
def POINT(x, y, z):
    return pyacadcom.AcadPoint(x, y, z).coordinates


def check_and_round(cell):
    if sheet[cell].value is None:
        return ''
    elif type(sheet[cell].value) is float:
        return round(sheet[cell].value, 2)
    else:
        return sheet[cell].value


# * read and store total data of panel from excel sheet named 'AutoCAD'
def get_total_data(excel_file):
    #wb = xw.Book(excel_file)
    sheet = wb['AutoCAD']
    total_data = [check_and_round('A1'),            # get Panel Name
                  check_and_round('B2'),            # get Py
                  check_and_round('B3'),            # get Kc
                  check_and_round('B4'),            # get Pp
                  check_and_round('B5'),            # get cos f
                  check_and_round('B6'),            # get Ip
                  check_and_round('B7')             # get Isc(3)
                  ]
    return total_data


# * read and store balance data of panel from excel sheet named 'AutoCAD'
def get_balance_data(excel_file):
    #wb = xw.Book(excel_file)
    sheet = wb['AutoCAD']
    balance_data = [check_and_round('A9'),          # is balance nedd to be inserted (Yes/No)
                    check_and_round('B10'),         # get L1_Pp
                    check_and_round('B11'),         # get L2_Pp
                    check_and_round('B12'),         # get L3_Pp
                    check_and_round('C10'),         # get L1_Ip
                    check_and_round('C11'),         # get L2_Ip
                    check_and_round('C12'),         # get L3_Ip
                    check_and_round('D10'),         # get L1_KNF
                    check_and_round('D11'),         # get L2_KNF
                    check_and_round('D12')          # get L3_KNF
                    ]
    return balance_data


def get_incomer_data(excel_file):
    #wb = xw.Book(excel_file)
    sheet = wb['AutoCAD']
    incomer_data = [check_and_round('A21'),
                    check_and_round('A22'),
                    check_and_round('A23'),
                    check_and_round('A24'),
                    check_and_round('C21'),
                    check_and_round('C22'),
                    check_and_round('C23'),
                    check_and_round('C24'),
                    check_and_round('F21'),
                    check_and_round('F22'),
                    check_and_round('B21'),
                    check_and_round('B22'),
                    check_and_round('B23'),
                    check_and_round('B24'),
                    check_and_round('D1')
                    ]
    return incomer_data


def insert_incomer():
    pt1 = POINT(0.0, 0.0, 0.0)
    incomer_blk = ms.InsertBlock(pt1, "INCOMER", 1.0, 1.0, 1.0, 0)
    incomer_blk_atts = incomer_blk.GetAttributes()
    incomer_data = get_incomer_data(excel_file)
    visibility_props = incomer_blk.GetDynamicBlockProperties()
    for prop in visibility_props:
        if prop.PropertyName == "IncomerType":
            prop.Value = "SWITCH_3pole"
    i = 0
    for attr in incomer_blk_atts:
        attr.TextString = incomer_data[i]
        attr.Update()
        i += 1


def insert_total():  # * inserting block 'TOTAL' in Model Space of dwg file and inserting block 'KNF' if user selected it
    pt1 = POINT(0.0, 0.0, 0.0)
    total_blk = ms.InsertBlock(pt1, "TOTAL", 1.0, 1.0, 1.0, 0)
    total_blk_atts = total_blk.GetAttributes()
    total_data = get_total_data(excel_file)
    i = 0
    for attr in total_blk_atts:
        attr.TextString = total_data[i]
        attr.Update()
        i += 1
    balance_data = get_balance_data(excel_file)
    if balance_data[0].lower() == 'да':
        balance_blk = ms.InsertBlock(pt1, "KNF", 1.0, 1.0, 1.0, 0)
        balance_blk_atts = balance_blk.GetAttributes()
        j = 1
        for attr in balance_blk_atts:
            attr.TextString = balance_data[j]
            attr.Update()
            j += 1


def round_data(source_range):  # * rounding floats in data from 'AutoCAD' sheet from excel
    c = 0
    columns = len(source_range[0]) # количество ячеек в строке 0
    rounded_data = []
    while c < columns:
        
        column_c = []
        for i in range(len(source_range)):
            column_c.append(source_range[i][c])

        rounded_data.append(column_round(column_c))
        c += 1
    return rounded_data


def column_round(column):  # * part of previus function
    rounded_column = []
    for _item in column:
        item = _item.value
        if item is None:
            rounded_column.append('')
        elif type(item) is float:
            if modf(item)[0] == 0:
                rounded_column.append(round(item))
            else:
                rounded_column.append(round(item, 2))
        else:
            rounded_column.append(item)
    return rounded_column


# ! selecting source excel file and store data for groups
app = QtWidgets.QApplication([])
excel_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите исходный файл Excel... ", filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем исхоный файл

# ! selecting source dwg file and deleting existing target blocks (if exist)
acad = pyacadcom.AutoCAD()
dwg_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите файл шаблона или схемы в AutoCAD... ", filter="DWG (*.dwg)")[0]  # выбираем исхоный файл
doc = acad.Documents.Open(dwg_file)
acad.Visible = True
ms = doc.ModelSpace
try:
    for i in doc.SelectionSets:
        i.Delete()
except:
    pass
time.sleep(2)
objSS = doc.SelectionSets.Add("toErase")

FilterType = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
FilterData = win32com.client.VARIANT(
    pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['INSERT'])
SELECT_ALL = 5

objSS.Select(SELECT_ALL, pythoncom.Empty,
             pythoncom.Empty, FilterType, FilterData)


for obj in objSS:
    if obj.EffectiveName in ["TOTAL", "KNF", "INCOMER", "AUTOMAT", "LINE"]:
        obj.Delete()
objSS.Delete()

time.sleep(1)

wb = opxl.load_workbook(excel_file, data_only=True)
sheet = wb["AutoCAD"]

#? get named range and convert them to squared massive of cells with rows and columns indices
def_range = wb.defined_names['DB_EXPORT']
dests = def_range.destinations
cells = []
for title, coord in dests:
    ws = wb[title]
    cells.append(ws[coord])
source_data = cells[0]  # ==> source_data[row][column]

# ! inserting TOTAL and KNF block in model space
insert_total()

# ! inserting INCOMER block in model space
insert_incomer()

# ! inserting AUTOMATs and LINEs blocks
rounded_data = round_data(source_data)
pt_x = 0.0
for column in rounded_data:
    pt1 = POINT(pt_x, 0.0, 0.0)
    automat_blk = ms.InsertBlock(pt1, "AUTOMAT", 1.0, 1.0, 1.0, 0)
    automat_visibility_props = automat_blk.GetDynamicBlockProperties()
    for prop in automat_visibility_props:
        if prop.PropertyName == "BreakerType":
            prop.Value = column[0]
    automat_blk_atts = automat_blk.GetAttributes()
    automat_blk_atts[0].TextString = column[16]
    automat_blk_atts[1].TextString = column[1]
    automat_blk_atts[2].TextString = column[2]
    automat_blk_atts[3].TextString = column[3]
    automat_blk_atts[4].TextString = column[4]
    automat_blk_atts[5].TextString = column[5]
    automat_blk_atts[6].TextString = column[8]
    automat_blk_atts[7].TextString = column[9]
    automat_blk_atts[8].TextString = column[10]
    automat_blk_atts[9].TextString = column[11]
    line_blk = ms.InsertBlock(pt1, "LINE", 1.0, 1.0, 1.0, 0)
    line_visibility_props = line_blk.GetDynamicBlockProperties()
    for prop in line_visibility_props:
        if prop.PropertyName == "Cunsumer":
            prop.Value = column[-1]
    line_blk_atts = line_blk.GetAttributes()
    line_blk_atts[0].TextString = column[17]
    line_blk_atts[1].TextString = column[18]
    line_blk_atts[2].TextString = column[19]
    line_blk_atts[3].TextString = column[20]
    line_blk_atts[4].TextString = column[21]
    line_blk_atts[5].TextString = column[22]
    line_blk_atts[6].TextString = column[23]
    line_blk_atts[7].TextString = column[24]
    line_blk_atts[8].TextString = column[25]
    line_blk_atts[9].TextString = column[26]
    line_blk_atts[10].TextString = column[27]
    line_blk_atts[11].TextString = column[28]
    line_blk_atts[12].TextString = column[29]
    pt_x += 25

wb.close()
