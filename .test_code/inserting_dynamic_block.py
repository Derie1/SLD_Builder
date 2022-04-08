import win32com.client
import pythoncom
import xlwings as xw
from math import *
from PyQt5 import QtCore, QtWidgets


def POINT(x, y, z):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def round_data(source_range):
    c = 0
    columns = len(source_range[0, :])
    rounded_data = []
    while c < columns:
        column_c = source_range[:, c].value
        rounded_data.append(column_round(column_c))
        c += 1
    return rounded_data


def column_round(column):
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


app = QtWidgets.QApplication([])
excel_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите исходный файл Excel... ",
                                                   filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем исхоный файл
wb = xw.Book(excel_file)
sheet = wb.sheets['AutoCAD']
source_data = sheet.range('DB_EXPORT')

dwg_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите файл шаблона или схемы в AutoCAD... ",
                                                 filter="DWG (*.dwg)")[0]  # выбираем исхоный файл
acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.Documents.Open(dwg_file)
ms = doc.ModelSpace


rounded_data = round_data(source_data)
pt_x = 0.0
for column in rounded_data:
    print("\n")
    print(column[-1])
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
