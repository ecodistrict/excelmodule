import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.IO")

from System.Windows.Forms import Application
from System.IO import Path

dll_dir = Path.GetDirectoryName(Application.ExecutablePath)
dll_path = Path.Combine(dll_dir, 'EcodistrictMessaging.dll')
clr.AddReferenceToFileAndPath(dll_path)

from Ecodistrict.Messaging import List, Number, InputSpecification, Outputs, Kpi
# Dont change anything above

#Serverproperties
serverAdress="localhost"
port=4000
userId=123
userName="SP"
federation="TNOdemo"
subScribedEvent="dashboard"
publishedEvent="EcoDistrict"
#End Serverproperies

name = "SP PB ExcelTest"                 #Name of the module to be presented at the dashboard (Web frendly string)
description = "Excel test module"	     #Description of the module. (Web frendly string)
moduleId = "SP_Excelmodule"              #Should be unique as it identifies the module. (Web frendly string)
path = "L:\EcoDistr\EcoExcel.xlsx"       #Path to the Excel book
kpiList=["SP_Energy-1", "SP_Energy-2"]   #List of the kpi:s the mudule can calculate (Web frendly string)

#Function that detects the existence of a kpi.
#Do not change this function.
def kpiId_Exists(kpiId):
    for s in kpiList:
        if s==kpiId:
            return true
    return false

#Here should the input specification for the different kpi:s be defined
def input_specification(kpiId):
    inputs = InputSpecification()
    if kpiId=="SP_Energy-1":
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
    elif kpiId=="SP_Energy-2":
        inputs.Add("Testvalue_1",Number(label="Test varde 1", min=2, max=32))
        inputs.Add("Testvalue_2",Number(label="Test varde 2", min=33, max=53))
        lst = List(label="lista")
        lst.Add(key="Listvalue_1",item=Number(label="Listvarde 1", min=0, max=100))
        lst.Add(key="Listvalue_2",item=Number(label="Listvarde 2", min=0, max=100))
        lst.Add(key="Listvalue_5",item=Number(label="Listvarde 5", min=0, max=100))
        inputs.Add("Listvarden", lst )
    return inputs

#Here should the specification for the the Excel coupling for the different kpi:s be defined
def run(indata,kpiId,excel):
    outputs = Outputs()
    if kpiId=="SP_Energy-1":
        excel.SetCellValue("EcoSheet1","B2",indata["Testvalue_1"])
        excel.SetCellValue("EcoSheet1","B3",indata["Testvalue_2"])
        excel.SetCellValue("EcoSheet1","B4",indata["Testvalue_3"])
        kpi = excel.GetCellValue("EcoSheet1", "B11")
        outputs.Add(Kpi(kpi,"SP_Energy-1 result","Unit"))
    elif kpiId=="SP_Energy-2":
        excel.SetCellValue("EcoSheet1","B2",indata["Testvalue_1"])
        excel.SetCellValue("EcoSheet1","B3",indata["Testvalue_2"])
        kpi = excel.GetCellValue("EcoSheet1", "B11")
        outputs.Add(Kpi(kpi,"SP_Energy-2 result","Unit"))
    return outputs