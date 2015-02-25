import clr
import System
import System.IO
dll_dir = System.IO.Path.GetDirectoryName(System.Application.ExecutablePath)
dll_path = System.IO.Path.Combine(dll_dir, 'DataTypes.dll')
clr.AddReferenceToFileAndPath(dll_path)
from DataTypes import List, Number, InputSpecification

name = "Renobuild"
description = "Description of the module"
moduleId = 123
path = '....xls'

def input_specification():
    inputs = InputSpecification()
    inputs.Add(Number(max=32))
    l = List()
    l.Add(Number(min=0))
    inputs.Add(l)
    return inputs

def run(indata, excel):
    return "test"