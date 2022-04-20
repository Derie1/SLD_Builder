import win32com.client
import pythoncom
from PyQt5 import QtWidgets
import pyacadcom


app = QtWidgets.QApplication([])
acad = pyacadcom.AutoCAD()
dwg_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите файл шаблона или схемы в AutoCAD... ",
                                                 filter="DWG (*.dwg)")[0]  # выбираем исхоный файл
try:
    doc = acad.Documents.Open(dwg_file)
except:
    print(
        f"error during file open, documents set type is: {type(acad.Documents)}")

try:
    for i in doc.SelectionSets:
        i.Delete()
except:
    print(
        f"error during selset delete, selset collection type is: {type(doc.SelectionSets)}, selset type is {type(i)}")
    pass
try:
    objSS = doc.SelectionSets.Add("blocks")
except:
    print(
        f"error during selset add, selset collection type is: {type(doc.SelectionSets)}")

FilterType = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
FilterData = win32com.client.VARIANT(
    pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['INSERT'])
SELECT_ALL = 5
try:
    objSS.Select(SELECT_ALL, pythoncom.Empty,
                 pythoncom.Empty, FilterType, FilterData)
except:
    print(f"error during filtering, selset type is: {type(objSS)}")
print("\n")

for obj in objSS:

    try:
        a = obj.EffectiveName
    except:
        print(f"error during getting filename, obj type is: {type(obj)}")

    try:
        if a in ["TOTAL", "KNF", "INCOMER", "AUTOMAT", "LINE"]:
            obj.Delete()
    except:
        print(f"error during deleting block, selset type is: {obj}")

try:
    objSS.Delete()
except:
    print(f"error during selset deleting, selset type is: {type(objSS)}")
