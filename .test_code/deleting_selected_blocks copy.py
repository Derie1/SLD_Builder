import win32com.client
import pythoncom
from PyQt5 import QtWidgets
import pyacadcom


app = QtWidgets.QApplication([])
# acad = win32com.client.Dispatch("AutoCAD.Application")
acad = pyacadcom.AutoCAD()
dwg_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите файл шаблона или схемы в AutoCAD... ",
                                                 filter="DWG (*.dwg)")[0]  # выбираем исхоный файл
doc = acad.Documents.Open(dwg_file)
ms = doc.ModelSpace

for obj in ms:
    if obj.ObjectName == "AcDbBlockReference" and (obj.EffectiveName == "TOTAL" or obj.EffectiveName == "KNF" or obj.EffectiveName == "INCOMER" or obj.EffectiveName == "AUTOMAT" or obj.EffectiveName == "LINE"):
        obj.Delete()

print("\n")
