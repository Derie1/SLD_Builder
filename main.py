import win32com.client
import pythoncom
import time
import xlwings as xw
from PyQt5 import QtWidgets
from math import modf
import pyacadcom


# * functions
def POINT(x, y, z):
    return pyacadcom.AcadPoint(x, y, z).coordinates


# * read and store total data of panel from excel sheet named 'AutoCAD'
def get_total_data(excel_file):
    wb = xw.Book(excel_file)
    sheet = wb.sheets['AutoCAD']
    total_data = [sheet.range('A1').value,              # get Panel Name
                  round(sheet.range('B2').value, 2),    # get Py
                  round(sheet.range('B3').value, 2),    # get Kc
                  round(sheet.range('B4').value, 2),    # get Pp
                  round(sheet.range('B5').value, 2),    # get cos f
                  round(sheet.range('B6').value, 2),    # get Ip
                  round(sheet.range('B7').value, 2)     # get Isc(3)
                  ]
    return total_data


# * read and store balance data of panel from excel sheet named 'AutoCAD'
def get_balance_data(excel_file):
    wb = xw.Book(excel_file)
    sheet = wb.sheets['AutoCAD']
    balance_data = [sheet.range('A9').value,               # is balance nedd to be inserted (Yes/No)
                    round(sheet.range('B10').value, 1),    # get L1_Pp
                    round(sheet.range('B11').value, 1),    # get L2_Pp
                    round(sheet.range('B12').value, 1),    # get L3_Pp
                    round(sheet.range('C10').value, 1),    # get L1_Ip
                    round(sheet.range('C11').value, 1),    # get L2_Ip
                    round(sheet.range('C12').value, 1),    # get L3_Ip
                    round(sheet.range('D10').value),       # get L1_KNF
                    round(sheet.range('D11').value),       # get L2_KNF
                    round(sheet.range('D12').value)        # get L3_KNF
                    ]
    return balance_data


def get_incomer_data(excel_file):
    wb = xw.Book(excel_file)
    sheet = wb.sheets['AutoCAD']
    incomer_data = [sheet.range('A21').value,
                    sheet.range('A22').value,
                    sheet.range('A23').value,
                    sheet.range('A24').value,
                    sheet.range('C21').value,
                    sheet.range('C22').value,
                    sheet.range('C23').value,
                    sheet.range('C24').value,
                    sheet.range('F21').value,
                    sheet.range('F22').value,
                    sheet.range('B21').value,
                    sheet.range('B22').value,
                    sheet.range('B23').value,
                    sheet.range('B24').value,
                    sheet.range('D1').value
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
    columns = len(source_range[0, :])
    rounded_data = []
    while c < columns:
        column_c = source_range[:, c].value
        rounded_data.append(column_round(column_c))
        c += 1
    return rounded_data


def column_round(column):  # * part of previus function
    rounded_column = []
    for item in column:
        if type(item) == float:
            if modf(item)[0] == 0:
                rounded_column.append(round(item))
            else:
                rounded_column.append(round(item, 2))
        else:
            rounded_column.append(item)
    return rounded_column


# ! selecting source excel file and store data for groups
app = QtWidgets.QApplication([])
excel_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите исходный файл Excel... ",
                                                   filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем исхоный файл

with xw.App(visible=False) as xl_app:
    wb = xw.Book(excel_file)
    sheet = wb.sheets['AutoCAD']
    source_data = sheet.range('DB_EXPORT')
    wb.close()

# ! selecting source dwg file and deleting existing target blocks (if exist)
acad = pyacadcom.AutoCAD()
dwg_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите файл шаблона или схемы в AutoCAD... ",
                                                 filter="DWG (*.dwg)")[0]  # выбираем исхоный файл
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
# ! inserting TOTAL and KNF block in model space

with xw.App(visible=False) as xl_app:
    wb = xw.Book(excel_file)
    sheet = wb.sheets['AutoCAD']
    source_data = sheet.range('DB_EXPORT')

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
