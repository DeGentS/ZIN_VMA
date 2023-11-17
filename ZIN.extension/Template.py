# -*- coding: utf-8 -*-
__title__   = "Test_hostId"
__doc__     = """Hiermee kijken hoeveel elementen een waarde hebben onder de parameter "BERSnl_C_codering_systeem01" """

#-----------------------IMPORTS-------------------------------------------------------


import clr


clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import BuiltInCategory as Bic


#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#for testing
element_id = ElementId(18142629)
element = doc.GetElement(element_id)
#--------------------------------------------------------------------------------------


t               = Transaction(doc, __title__)   #dit zorgt ervoor dat bij de undo/redo button de naam van __title__ staat
                                                # Transaction = nodig wanneer we een handeling of aanpassing wensen door te voeren


print(element)