import win32com.client
import time
import pythoncom
from PyQt5 import QtWidgets


app = QtWidgets.QApplication([])
acad = win32com.client.Dispatch("AutoCAD.Application")
dwg_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите файл шаблона или схемы в AutoCAD... ",
                                                 filter="DWG (*.dwg)")[0]  # выбираем исхоный файл
doc = acad.Documents.Open(dwg_file)

try:
    for i in doc.SelectionSets:
        i.Delete()
except:
    pass

time.sleep(1)
objSS = doc.SelectionSets.Add("toErase")

FilterType = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
FilterData = win32com.client.VARIANT(
    pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['INSERT'])
SELECT_ALL = 5

objSS.Select(SELECT_ALL, pythoncom.Empty,
             pythoncom.Empty, FilterType, FilterData)
print("\n")

for obj in objSS:
    if obj.EffectiveName == "TOTAL" or obj.EffectiveName == "KNF" or obj.EffectiveName == "INCOMER" or obj.EffectiveName == "AUTOMAT" or obj.EffectiveName == "LINE":
        obj.Delete()

objSS.Delete()
