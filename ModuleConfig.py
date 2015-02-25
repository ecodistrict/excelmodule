from IronPythonTest import List, Number, InputSpecification

name = "Renobuild"
description = "Description of the module"
moduleId = 123
path = '....xls'

def input_specification():
    inputs = InputSpecification()
    inputs.add(Number(max=32))
    l = List()
    l.add(Number(min=0))
    inputs.add(l)
    return inputs

def run(indata, excel):
    return "test"