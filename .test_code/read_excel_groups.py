import xlwings as xw
from math import *
from PyQt5 import QtCore, QtWidgets


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
                                                   filter="XLS (*.xls);XLSX (*.xlsx)")[0]

wb = xw.Book(excel_file)
sheet = wb.sheets['AutoCAD']
source_data = sheet.range('DB_EXPORT')

rounded_data = round_data(source_data)

print(rounded_data[0])


# print(group_data[:, 0].value, "\n")  # first column
# print(group_data[18:21, :].value, "\n")  # rows from 18 to 21
