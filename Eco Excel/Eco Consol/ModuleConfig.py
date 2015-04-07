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
#serverAdress="vps17642.public.cloudvps.com" #localhost"
#port=4000
#userId=123
#userName="SP"
#federation="EcoDistrict" #"TNOdemo"
#subScribedEvent="dashboard"
#publishedEvent="modules"
#End Serverproperies

name = "SP PB ExcelTest"                 #Name of the module to be presented at the dashboard (Web frendly string)
description = "Excel test module"	     #Description of the module. (Web frendly string)
moduleId = "SP_Excelmodule"               #Should be unique as it identifies the module.(Web frendly string)
#Path to the Excel book
path = "C:\Users\perbe\Documents\EcoDistr\EcoExcelSample.xlsx"
#List of the kpi:s the mudule can calculate (Web frendly string)  
kpiList=["SP_Average", "SP_Sum", "SP_Median"]   

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
    if kpiId=="SP_Average":
        inputs.Add("TestValue_1",Number(label="AverageValue1", min=2, max=32))
        inputs.Add("TestValue_2",Number(label="AverageValue2", min=33, max=53))
        inputs.Add("TestValue_3",Number(label="AverageValue3", min=54, max=97))
        inputs.Add("TestValue_4",Number(label="AverageValue4", min=54, max=97))
        inputs.Add("TestValue_5",Number(label="AverageValue5", min=54, max=97))
        inputs.Add("TestValue_6",Number(label="AverageValue6", min=54, max=97))
        inputs.Add("TestValue_7",Number(label="AverageValue7", min=54, max=97))
        inputs.Add("TestValue_8",Number(label="AverageValue8", min=54, max=97))
        inputs.Add("TestValue_9",Number(label="AverageValue9", min=54, max=97))
        inputs.Add("TestValue_10",Number(label="AverageValue10", min=54, max=97))
    elif kpiId=="SP_Sum":
        inputs.Add("Value_1",Number(label="SumValue1", min=2, max=32))
        inputs.Add("Value_2",Number(label="SumValue2", min=33, max=53))
        inputs.Add("Value_3",Number(label="SumValue3", min=54, max=97))
        inputs.Add("Value_4",Number(label="SumValue4", min=54, max=97))
        inputs.Add("Value_5",Number(label="SumValue5", min=54, max=97))
        inputs.Add("Value_6",Number(label="SumValue6", min=54, max=97))
        inputs.Add("Value_7",Number(label="SumValue7", min=54, max=97))
        inputs.Add("Value_8",Number(label="SumValue8", min=54, max=97))
        inputs.Add("Value_9",Number(label="SumValue9", min=54, max=97))
        inputs.Add("Value_10",Number(label="SumValue10", min=54, max=97))
    elif kpiId=="SP_Median":
        inputs.Add("MedValue_1",Number(label="MedianValue1", min=2, max=32))
        inputs.Add("MedValue_2",Number(label="MedianValue2", min=33, max=53))
        inputs.Add("MedValue_3",Number(label="MedianValue3", min=54, max=97))
        inputs.Add("MedValue_4",Number(label="MedianValue4", min=54, max=97))
        inputs.Add("MedValue_5",Number(label="MedianValue5", min=54, max=97))
        inputs.Add("MedValue_6",Number(label="MedianValue6", min=54, max=97))
        inputs.Add("MedValue_7",Number(label="MedianValue7", min=54, max=97))
        inputs.Add("MedValue_8",Number(label="MedianValue8", min=54, max=97))
        inputs.Add("MedValue_9",Number(label="MedianValue9", min=54, max=97))
        inputs.Add("MedValue_10",Number(label="MedianValue10", min=54, max=97))

    return inputs

#Here should the specification for the the Excel coupling for the different kpi:s be defined
def run(indata,kpiId,excel):
    outputs = Outputs()
    if kpiId=="SP_Average":
        excel.SetCellValue("Kpi_1","B3",indata["TestValue_1"])
        excel.SetCellValue("Kpi_1","B4",indata["TestValue_2"])
        excel.SetCellValue("Kpi_1","B5",indata["TestValue_3"])
        excel.SetCellValue("Kpi_1","B6",indata["TestValue_4"])
        excel.SetCellValue("Kpi_1","B7",indata["TestValue_5"])
        excel.SetCellValue("Kpi_1","B8",indata["TestValue_6"])
        excel.SetCellValue("Kpi_1","B9",indata["TestValue_7"])
        excel.SetCellValue("Kpi_1","B10",indata["TestValue_8"])
        excel.SetCellValue("Kpi_1","B11",indata["TestValue_9"])
        excel.SetCellValue("Kpi_1","B12",indata["TestValue_10"])
        kpi = excel.GetCellValue("Kpi_1", "B14")
        outputs.Add(Kpi(kpi,"SP_Average result","Unit"))
    elif kpiId=="SP_Sum":
        excel.SetCellValue("Kpi_2","B3",indata["Value_1"])
        excel.SetCellValue("Kpi_2","B4",indata["Value_2"])
        excel.SetCellValue("Kpi_2","B5",indata["Value_3"])
        excel.SetCellValue("Kpi_2","B6",indata["Value_4"])
        excel.SetCellValue("Kpi_2","B7",indata["Value_5"])
        excel.SetCellValue("Kpi_2","B8",indata["Value_6"])
        excel.SetCellValue("Kpi_2","B9",indata["Value_7"])
        excel.SetCellValue("Kpi_2","B10",indata["Value_8"])
        excel.SetCellValue("Kpi_2","B11",indata["Value_9"])
        excel.SetCellValue("Kpi_2","B12",indata["Value_10"])
        kpi = excel.GetCellValue("Kpi_2", "B14")
        outputs.Add(Kpi(kpi,"SP_Sum result","Unit"))
    elif kpiId=="SP_Median":
        excel.SetCellValue("Kpi_3","B3",indata["MedValue_1"])
        excel.SetCellValue("Kpi_3","B4",indata["MedValue_2"])
        excel.SetCellValue("Kpi_3","B5",indata["MedValue_3"])
        excel.SetCellValue("Kpi_3","B6",indata["MedValue_4"])
        excel.SetCellValue("Kpi_3","B7",indata["MedValue_5"])
        excel.SetCellValue("Kpi_3","B8",indata["MedValue_6"])
        excel.SetCellValue("Kpi_3","B9",indata["MedValue_7"])
        excel.SetCellValue("Kpi_3","B10",indata["MedValue_8"])
        excel.SetCellValue("Kpi_3","B11",indata["MedValue_9"])
        excel.SetCellValue("Kpi_3","B12",indata["MedValue_10"])
        kpi = excel.GetCellValue("Kpi_3", "B14")
        outputs.Add(Kpi(kpi,"SP_Median result","Unit"))
    return outputs