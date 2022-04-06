import win32com.client
import pythoncom
import xlwings as xw
from PyQt5 import QtCore, QtWidgets


def POINT(x, y, z):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


app = QtWidgets.QApplication([])
# excel_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите исходный файл Excel... ",
#                                                   filter = "XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем исхоный файл

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument
ms = doc.ModelSpace

pt_x = 0.0
pt_y = 0.0
pt_z = 0.0
pt1 = POINT(pt_x, pt_y, pt_z)
line_block = ms.InsertBlock(pt1, "sw1", 1.0, 1.0, 1.0, 0)
visibility_props = line_block.GetDynamicBlockProperties()
for prop in visibility_props:
    if prop.PropertyName == "BreakerType":
        prop.Value = "SWITCH_RCD_3PH"

print(visibility_props)

# while i <= groups:
#     pt1 = POINT(pt_x, pt_y, pt_z)
#     line_block = ms.InsertBlock(pt1, "TOTAL", 1.0, 1.0, 1.0, 0)
#     for attr in line_block.GetAttributes():
#         attr.TextString = "PanelName"
#         attr.Update()
#     # line_block_atts = line_block.GetAttributes()
#     # line_block_atts(0).TextString = "PanelName"
#     # line_block_atts.Update()
#     pt_x += 25
#     i += 1
