# -*- coding: utf-8 -*-

__title__ = "Export subcontractor"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.1
Date    = 18-09-23
_____________________________________________________________________
Description:

Hiermee zal men alle openingen exporteren naar 
een excel-file ifv OM_Subcontractor

_____________________________________________________________
Last update:

- [18-09-23] 1.0 RELEASE
- [25-10-23] 1.1 Toevoegen; 
        - OMI_CTE_Element ID
        - Dimensies van de opening
        - communicatie "***_Comment" toegevoegd
        - toevoegen van Excel template

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

#-----------------------FUNCTIES-------------------------------------------------------

def get_element_info(element, doc_title):
    element_guid        = element.UniqueId
    element_id          = element.Id.IntegerValue
    OMI_CTE_Element     = element.LookupParameter("OMI_CTE_Element ID").AsString() if element.LookupParameter("OMI_CTE_Element ID") else "N/A" #Parameter toegevoegd om de elementen te kunnen
                                                                                                                                                # vanuit een containermodel/linked model
    family_name         = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
    type_name           = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
    level_name          = element.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM).AsValueString()
    subcontractor       = element.LookupParameter("OM_subcontractor").AsString()
    host_model_name     = element.LookupParameter("OMI_CTE_Host Model Name").AsString() if element.LookupParameter("OMI_CTE_Host Model Name") else "N/A"
    tc          = element.Symbol
    typecomment = tc.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString()

    #Afmetingen van het element
    hoogte      = element.LookupParameter("OMI_CLE_Height Total").AsValueString() if element.LookupParameter("OMI_CLE_Height Total") else "N/A"
    diepte     = element.LookupParameter("OMI_CLE_Width").AsValueString() if element.LookupParameter("OMI_CLE_Width") else "N/A"
    lengte      = element.LookupParameter("OMI_CLE_Length").AsValueString() if element.LookupParameter("OMI_CLE_Length") else "N/A"
    diameter    = element.LookupParameter("OMI_CLE_Diameter Total").AsValueString() if element.LookupParameter("OMI_CLE_Diameter Total") else "N/A"

    #communicatie van het element
    vma01       = element.LookupParameter("OMI_CTE_VMA - Comment 01").AsValueString() if element.LookupParameter("OMI_CTE_VMA - Comment 01") else "N/A"
    vma02       = element.LookupParameter("OMI_CTE_VMA - Comment 02").AsValueString() if element.LookupParameter("OMI_CTE_VMA - Comment 02") else "N/A"
    gr01        = element.LookupParameter("OMI_CTE_Greisch - Comment 01").AsValueString() if element.LookupParameter("OMI_CTE_Greisch - Comment 01") else "N/A"
    gr02        = element.LookupParameter("OMI_CTE_Greisch - Comment 02").AsValueString() if element.LookupParameter("OMI_CTE_Greisch - Comment 02") else "N/A"


    return element_guid, element_id, OMI_CTE_Element ,typecomment, subcontractor, level_name, family_name, type_name,hoogte, lengte, diameter,diepte, vma01, vma02, gr01, gr01, host_model_name, doc_title,

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

    # Verkrijg het pad naar het script
    script_path = os.path.abspath(__file__)

    # Bepaal het pad naar het Excel-sjabloon in dezelfde locatie als het script
    template_excel_path = os.path.join(os.path.dirname(script_path), "template.xlsx")

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

    # Controleer of het Excel-sjabloon bestaat
    if not os.path.exists(template_excel_path):
        # Toon een foutmelding als het sjabloon niet gevonden kan worden
        MessageBox.Show("Excel-sjabloon (template.xlsx) niet gevonden in dezelfde locatie als het script. Het script wordt afgebroken.")
        sys.exit()
    # # Kopieer het sjabloon naar de gewenste locatie voor het Excel-bestand
    shutil.copy(template_excel_path, excel_file_path)

    # Open het Excel-bestand
    excel_app = ApplicationClass()
    workbook = excel_app.Workbooks.Open(excel_file_path)
    worksheet = workbook.Sheets["Export"]

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
    worksheet.Cells[1, 3].Value2 = "OMI_CTE_Element"
    worksheet.Cells[1, 4].Value2 = "TypeComment"
    worksheet.Cells[1, 5].Value2 = "Subcontractor"
    worksheet.Cells[1, 6].Value2 = "Level Name"
    worksheet.Cells[1, 7].Value2 = "Family Name"
    worksheet.Cells[1, 8].Value2 = "Type Name"
    worksheet.Cells[1, 9].Value2 = "Hoogte"
    worksheet.Cells[1, 10].Value2 = "Lengte"
    worksheet.Cells[1, 11].Value2 = "Diameter"
    worksheet.Cells[1, 12].Value2 = "Diepte"
    worksheet.Cells[1, 13].Value2 = "VMA comment 01"
    worksheet.Cells[1, 14].Value2 = "VMA comment 02"
    worksheet.Cells[1, 15].Value2 = "GR Comment 01"
    worksheet.Cells[1, 16].Value2 = "GR Comment 02"
    worksheet.Cells[1, 17].Value2 = "OMI_CTE_Host Model Name"
    worksheet.Cells[1, 18].Value2 = "Discipline"


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
    table_range = worksheet.Range("A1", "r" + str(range_end))
    table = worksheet.ListObjects.Add(1, table_range, 0, 1, 1)
    table.Name = "GenericModelData"



    # Sla het Excel-bestand op en sluit Excel
    workbook.SaveAs(excel_file_path, XlFileFormat.xlOpenXMLWorkbook)
    workbook.Close()
    excel_app.Quit()

    # Toon een bevestigingsbericht
    MessageBox.Show("Gegevens succesvol geëxporteerd naar Excel.")

#-----------------------MAIN-------------------------------------------------------

if __name__ == "__main__":
    main()
