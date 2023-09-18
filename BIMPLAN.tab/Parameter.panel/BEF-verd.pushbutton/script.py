# -*- coding: utf-8 -*-

__title__ = "BEF-verd"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 02-08-23
_____________________________________________________________________
Description:

Voorziet BEF-verdieping van de corresponderende naam 
van de level waarop het element gehost is.

_____________________________________________________________
Last update:

- [02-08-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""

#-----------------------IMPORTS-------------------------------------------------------

import clr
import sys
clr.AddReference("RevitAPI")
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction
from Autodesk.Revit.UI import TaskDialog


#----------------------VARIABLES--------------------------------------------------------
#Standaard VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------


def set_mark_as_level_name():
    # Get all non-nested elements in the model
    elements = FilteredElementCollector(doc).WhereElementIsNotElementType().WhereElementIsNotElementType().ToElements()

    count_unchanged = 0
    count_changed = 0

    # Dictionary to track the changes
    change_details = {
        'unchanged': [],
        'changed': [],
    }

    # Create a transaction to modify the document
    with Transaction(doc, 'Set BEF_verdieping as level name') as t:
        t.Start()

        for el in elements:
            try:
                # Check if the element has a level property
                if el.LevelId != ElementId.InvalidElementId:
                    # Get the level of the element
                    level = doc.GetElement(el.LevelId)

                    # Check if the element has a parameter named "BEF_verdieping"
                    bef_verdieping_param = el.LookupParameter('BEF_verdieping')

                    if bef_verdieping_param and level:
                        current_value = bef_verdieping_param.AsString()

                        # Set the value of the "BEF_verdieping" parameter to the first 4 characters of the element's level name
                        new_value = level.Name[:4]
                        bef_verdieping_param.Set(new_value)

                        if current_value == new_value:
                            count_unchanged += 1
                            change_details['unchanged'].append(el.Id)
                        else:
                            count_changed += 1
                            change_details['changed'].append(el.Id)
            except Exception as e:
                pass

        t.Commit()

    # Toon de resultaten in een TaskDialog
    dialog = TaskDialog("Set BEF_verdieping as Level Name Resultaat")
    dialog.MainInstruction = "Overzicht van wijzigingen in BEF_verdieping-parameter"

    # Gebruik str.format() om de waarden in te voegen
    dialog.MainContent = (
        "Aantal ongewijzigde elementen: {}\n"
        "Aantal gewijzigde elementen: {}".format(count_unchanged, count_changed)
    )

    # Voeg de ID's van de gewijzigde en ongewijzigde elementen toe aan de TaskDialog
    if count_unchanged > 0:
        dialog.ExpandedContent = "ID's van ongewijzigde elementen:\n" + "\n".join(str(id) for id in change_details['unchanged'])
    if count_changed > 0:
        dialog.ExpandedContent += "\n\nID's van gewijzigde elementen:\n" + "\n".join(str(id) for id in change_details['changed'])

    # Toon de TaskDialog
    dialog.Show()

# Roep de functie aan om het script uit te voeren en het resultaat te tonen
set_mark_as_level_name()
