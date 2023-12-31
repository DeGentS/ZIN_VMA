# -*- coding: utf-8 -*-

__title__ = "Import_Set\nSubcontractor"
__author__ = "Sean De Gent"
__doc__ = """Version = 0.0
Date    = 21-10-23
_____________________________________________________________________
Description:

Importeer excel subcontractor en wijzig;
        - OM-subcontractor
        - Family Type
_____________________________________________________________
Last update:

- [21-10-23] 0.0 START


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""

#-----------------------IMPORTS-------------------------------------------------------

import clr
import sys

clr.AddReference("RevitAPI")
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FilteredElementCollector, Transaction
from Autodesk.Revit.DB import BuiltInCategory as Bic

from Autodesk.Revit.UI import TaskDialog

#-----------------------CUSTOM_IMPORTS-------------------------------------------------------
import System
import shutil
import xlrd
import os

# Importeren van .NET Windows Forms voor de OpenFileDialog
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, DialogResult, MessageBox

# Importeren van de Microsoft.Office.Interop.Excel namespace
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass, XlFileFormat, XlCellType
from Microsoft.Office.Interop.Excel import XlDirection

#----------------------VARIABLES--------------------------------------------------------
#Standaard VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------

def read_excel_data(excel_file_path):
    """Leest de gegevens uit het Excel-bestand en retourneert een lijst van tuples."""
    # Een lijst om de gelezen gegevens van het Excel-bestand op te slaan
    data = []

    # Open het Excel-bestand en lees de gegevens
    excel_app = ApplicationClass()
    excel_app.Visible = False
    workbook = excel_app.Workbooks.Open(excel_file_path)
    worksheet = workbook.Sheets["Export"]

    # Zoek de laatste rij met gegevens in het werkblad
    last_row = worksheet.Cells(worksheet.Rows.Count, 1).End(XlDirection.xlUp).Row

    # Lees de gegevens van het werkblad en sla deze op in de 'data' lijst
    for row in range(2, last_row + 1):
        element_id = worksheet.Cells(row, 2).Value2
        new_typename = worksheet.Cells(row, 4).Value2
        subcontractor = worksheet.Cells(row, 6).Value2

        # # Controleer of de element-ID als geheel getal kan worden behandeld
        # if isinstance(element_id, (int, float)):
        #     element_id = int(element_id)

        data.append((element_id, new_typename, subcontractor))

    # Sluit het Excel-bestand
    workbook.Close()
    excel_app.Quit()

    return data

def update_element_subcontractor(data):
    """Bijwerken van de subcontractor parameter in Revit voor elk element in de data lijst."""

    # Start een transactie om alle updates in te groeperen
    with Transaction(doc, "Update Subcontractor") as transaction:
        transaction.Start()

        # Loop door de lijst met gegevens en pas het commentaarveld aan voor elk element
        for element_id, element_subcontractor, new_typename in data:
            try:
                # Zoek het element op basis van het element-ID
                element_id_int = int(element_id)
                element = doc.GetElement(ElementId(element_id_int))

                # Controleer of het element bestaat en of het een commentaarveld heeft dat kan worden gewijzigd
                if element and hasattr(element, "get_Parameter"):
                    parameter_name = "OM_subcontractor"
                    parameter_name02 = ""
                    parameter = element.LookupParameter(parameter_name)
                    if parameter and not parameter.IsReadOnly:
                        # Pas het commentaarveld aan met de bijgewerkte waarde
                        parameter.Set(str(element_subcontractor))
            except Exception as e:
                print("Fout bij het bijwerken van element " + str(element_id) + ": " + str(e))

        # Commit de transactie
        transaction.Commit()
