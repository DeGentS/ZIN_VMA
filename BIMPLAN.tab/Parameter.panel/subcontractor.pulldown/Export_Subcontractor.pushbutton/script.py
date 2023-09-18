# -*- coding: utf-8 -*-

__title__ = "Export subcontractor"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 18-09-23
_____________________________________________________________________
Description:

Hiermee zal men alle opeingen exporteren naar een excel-file ifv OM_Subcontractor

_____________________________________________________________
Last update:

- [18-09-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""

#-----------------------IMPORTS-------------------------------------------------------

import clr
import os
import sys
import System
import shutil

# Importeren van .NET Windows Forms voor de SaveFileDialog
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import SaveFileDialog, DialogResult, MessageBox

# Importeren van de Microsoft.Office.Interop.Excel namespace
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass, XlFileFormat
from Microsoft.Office.Interop.Excel import XlRgbColor

# Importeren van Revit API-elementen
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

def get_element_info(element, doc_title):
    element_guid = element.UniqueId
    element_id = element.Id.IntegerValue
    family_name = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
    type_name = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
    level_name = element.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM).AsValueString()
    subcontractor = element.LookupParameter("OM_subcontractor").AsString()
    host_model_name = element.LookupParameter("OMI_CTE_Host Model Name").AsString() if element.LookupParameter("OMI_CTE_Host Model Name") else "N/A"

    return element_guid, element_id, family_name, type_name, level_name, subcontractor, doc_title, host_model_name

def get_generic_models_with_text_in_family_name(doc, text):
    genericmodels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType()
    total_elements = 0
    matching_elements = []

    for gm in genericmodels:
        family_name = gm.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
        if text in family_name and gm.SuperComponent is None:
            matching_elements.append(gm)
            total_elements += 1

    return matching_elements, total_elements

def set_cell_format(cell, rgb_color, is_bold):
    cell.Interior.Color = rgb_color
    cell.Font.Bold = is_bold
    cell.Font.Color = 16777215  # Witte tekst (RGB-waarde voor wit)

def main():
    uidoc = __revit__.ActiveUIDocument
    doc = uidoc.Document
    doc_title = doc.Title  # Verkrijgen van de documenttitel (modelnaam)

    # Vraag de gebruiker om de opslaglocatie voor het Excel-bestand
    save_file_dialog = SaveFileDialog()
    save_file_dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    save_file_dialog.Title = "Selecteer een locatie om het Excel-bestand op te slaan"
    result = save_file_dialog.ShowDialog()

    if result == DialogResult.OK:
        excel_file_path = save_file_dialog.FileName
    else:
        # Toon een foutmelding als er geen locatie is geselecteerd
        MessageBox.Show("Geen locatie geselecteerd. Het script wordt afgebroken.")
        sys.exit()

    # Controleer of het Excel-bestand al bestaat en maak zo nodig een nieuw bestand aan
    if os.path.exists(excel_file_path):
        os.remove(excel_file_path)  # Verwijder het bestand als het al bestaat

    # Maak een nieuw Excel-bestand
    excel_app = ApplicationClass()
    workbook = excel_app.Workbooks.Add()
    worksheet = workbook.Sheets.Add()
    worksheet.Name = "Export"

    text_to_search = "Opening"
    genericmodels, total_elements = get_generic_models_with_text_in_family_name(doc, text_to_search)

    if total_elements == 0:
        # Toon een foutmelding als er geen elementen zijn gevonden
        message = "Geen elementen gevonden met '" + text_to_search + "' in de familienaam en SuperComponent op None. Het script wordt afgebroken."
        MessageBox.Show(message)
        sys.exit()

    # Voeg kolomkoppen toe aan het werkblad
    worksheet.Cells[1, 1].Value2 = "Element GUID"
    worksheet.Cells[1, 2].Value2 = "Element ID"
    worksheet.Cells[1, 3].Value2 = "Family Name"
    worksheet.Cells[1, 4].Value2 = "Type Name"
    worksheet.Cells[1, 5].Value2 = "Level Name"
    worksheet.Cells[1, 6].Value2 = "Subcontractor"
    worksheet.Cells[1, 7].Value2 = "Model Name"
    worksheet.Cells[1, 8].Value2 = "OMI_CTE_Host Model Name"

    # Voeg gegevens toe aan het werkblad
    for index, element in enumerate(genericmodels, start=2):
        element_info = get_element_info(element, doc_title)
        for col_num, info in enumerate(element_info, start=1):
            cell = worksheet.Cells[index, col_num]
            cell.Value2 = info
            if col_num in [1, 2]:  # Pas de opmaak alleen toe op de eerste twee kolommen
                set_cell_format(cell, 255, True)

    # Maak een tabel van de toegevoegde gegevens
    range_end = len(genericmodels) + 1
    table_range = worksheet.Range("A1", "H" + str(range_end))
    table = worksheet.ListObjects.Add(1, table_range, 0, 1, 1)
    table.Name = "GenericModelData"

    # Sla het Excel-bestand op en sluit Excel
    workbook.SaveAs(excel_file_path, XlFileFormat.xlOpenXMLWorkbook)
    workbook.Close()
    excel_app.Quit()

    # Toon een bevestigingsbericht
    MessageBox.Show("Gegevens succesvol geëxporteerd naar Excel.")

if __name__ == "__main__":
    main()
