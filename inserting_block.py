import win32com.client
import pythoncom
import xlwings as xw
from PyQt5 import QtCore, QtWidgets


def POINT(x, y, z):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


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


def insert_total():
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


app = QtWidgets.QApplication([])
excel_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите исходный файл Excel... ",
                                                   filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем исхоный файл

# total_data = get_total_data(excel_file)
# balance_data = get_balance_data(excel_file)
# print(total_data)
# print(balance_data)

acad = win32com.client.Dispatch("AutoCAD.Application")
# doc = acad.ActiveDocument
dwg_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите файл шаблона или схемы в AutoCAD... ",
                                                 filter="DWG (*.dwg)")[0]  # выбираем исхоный файл
# acad.Documents.Open(Application.GetOpenFilename("ACAD files(*.dwg),*.dwg*,All files(*.*),*.*", 1, "Select Autocad template file...", , False))
doc = acad.Documents.Open(dwg_file)
acad.Visible = True
ms = doc.ModelSpace

insert_total()


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
