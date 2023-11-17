# -*- coding: utf-8 -*-
__title__ = "D_link\nDelete Value "                  # Name of the button displayed in Revit UI
__doc__ = """Version = 1.0
Date    = 16.11.2023
_____________________________________________________________________
Description:

Delete de waarde in de parameter OMI_CTE_BimCollab

_____________________________________________________________________
How-to:

-> Click on the button

_____________________________________________________________________
Last update:
- [14.06.2023] - 1.0 RELEASE
_____________________________________________________________________
To-Do:
_____________________________________________________________________
Author: Sean De Gent"""                                                      # Button Description shown in Revit UI

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter, Transaction

# Importeren van .NET Windows Forms voor de OpenFileDialog
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox


# Get the active document
doc = __revit__.ActiveUIDocument.Document

# Start a transaction
transaction = Transaction(doc, "Delete D-link Value")
transaction.Start()

# Collect all elements in the document
collector = FilteredElementCollector(doc)
elements = collector.WhereElementIsNotElementType().ToElements()
element_count = collector.WhereElementIsNotElementType().GetElementCount()

# Loop through each element and delete value parameter
for element in elements:
    parameter = element.LookupParameter("OMI_CTE_BimCollab")
    if parameter is not None:
        if parameter and not parameter.IsReadOnly:
            element.LookupParameter("OMI_CTE_BimCollab").Set("")

print(element_count)

# Commit the transaction
transaction.Commit()

MessageBox.Show("Succes")
