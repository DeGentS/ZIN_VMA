# -*- coding: utf-8 -*-
__title__ = "OM_Subcontractor\nDelete Value "                  # Name of the button displayed in Revit UI
__doc__ = """Version = 1.0
Date    = 26.10.2023
_____________________________________________________________________
Description:

Verwijder alle waarden aanwezig in de parameter OM_Subcontractor
_____________________________________________________________________
How-to:

-> Click on the button

_____________________________________________________________________
Last update:
- [26.10.2023] - 1.0 RELEASE

_____________________________________________________________________
Author: Sean De Gent"""                                                      # Button Description shown in Revit UI

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter, Transaction, BuiltInCategory

#----------------------VARIABELE--------------------------------------------------------

#Standaard VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------


genericmodels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType()

# Start a transaction
transaction = Transaction(doc, "Delete Subcontractor Value")
transaction.Start()



for element in genericmodels:
        if element.LookupParameter("OM_subcontractor"):
            element.LookupParameter("OM_subcontractor").Set("")


# Commit the transaction
transaction.Commit()
