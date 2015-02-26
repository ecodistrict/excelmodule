import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.IO")
from System.Windows.Forms import Application
from System.IO import Path

dll_dir = Path.GetDirectoryName(Application.ExecutablePath)
dll_path = Path.Combine(dll_dir, 'EcodistrictMessaging.dll')
clr.AddReferenceToFileAndPath(dll_path)

from Ecodistrict.Messaging import List, Number, InputSpecification

name = "Renobuild"
description = "Description of the module"
moduleId = 123
path = '....xls'

def input_specification():
    inputs = InputSpecification()
    inputs.Add(Number(max=32))
    l = List()
    l.Add(Number(min=0))
    l.Add(Number(min=3))
    inputs.Add(l)
    return inputs

def run(indata, excel):
    return "test"