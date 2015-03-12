import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.IO")

from System.Windows.Forms import Application
from System.IO import Path

dll_dir = Path.GetDirectoryName(Application.ExecutablePath)
dll_path = Path.Combine(dll_dir, 'EcodistrictMessaging.dll')
clr.AddReferenceToFileAndPath(dll_path)

from Ecodistrict.Messaging import List, Number, InputSpecification, Outputs, Kpi

name = "SP PB ExcelTest"
description = "Excel test module"
moduleId = "SP_Excelmodule"
path = "L:\EcoDistr\EcoExcel.xlsx"
kpiList=["SP_Excelmodule-kpi-1"]

def input_specification():
    inputs = InputSpecification()
    inputs.Add("Testvalue_1",Number(label="Test varde 1", min=2, max=32))
    inputs.Add("Testvalue_2",Number(label="Test varde 2", min=33, max=53))
    inputs.Add("Testvalue_3",Number(label="Test varde 3", min=54, max=97))
    lst = List(label="lista")
    lst.Add(key="Listvalue_1",item=Number(label="Listvarde 1", min=0, max=100))
    lst.Add(key="Listvalue_2",item=Number(label="Listvarde 2", min=0, max=100))
    lst.Add(key="Listvalue_3",item=Number(label="Listvarde 3", min=0, max=100))
    lst.Add(key="Listvalue_4",item=Number(label="Listvarde 4", min=0, max=100))
    lst.Add(key="Listvalue_5",item=Number(label="Listvarde 5", min=0, max=100))
    inputs.Add("Listvarden", lst )
    return inputs

def run(indata, excel):
    excel.SetCellValue("EcoSheet1","B2",indata["Testvalue_1"])
    excel.SetCellValue("EcoSheet1","B3",indata["Testvalue_2"])
    excel.SetCellValue("EcoSheet1","B4",indata["Testvalue_3"])
    outputs = Outputs()
    kpi = excel.GetCellValue("EcoSheet1", "B11")
    outputs.Add(Kpi(kpi,"Average result","Number"))
    return outputs