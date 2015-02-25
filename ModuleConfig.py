import clr
clr.AddReference("DataTypes.dll")
from IronPythonTest import List, Number, InputSpecification

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