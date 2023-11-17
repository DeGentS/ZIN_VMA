# -*- coding: utf-8 -*-

__title__ = "D.Link"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 16-11-23
_____________________________________________________________________
Description:

Vult de Parameter OMI_CTE_BimCollab in obv geselecteerde excel-file.
Vervolgens kan men het model synchroniseren met Dalux.
In dalux kan men nu op de parameter OMI_CTE_BimCollab filteren 
waardoor men een overzicht verkrijgt van de issues.

_____________________________________________________________
Last update:

- [16-11-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""

#-----------------------IMPORTS-------------------------------------------------------
import clr
import os
import sys
import System

#Import for track and log
import datetime
import os

# Importeren van .NET Windows Forms voor de OpenFileDialog
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, DialogResult, MessageBox

# Importeren van de Microsoft.Office.Interop.Excel namespace
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass, XlFileFormat, XlCellType
from Microsoft.Office.Interop.Excel import XlDirection

# Importeren van Revit API-elementen
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

# #----------------------VARIABLES--------------------------------------------------------
# #VARIABLES
#
doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application


#-----------------------FUNCTION-------------------------------------------------------
#LOGBESTAND AANMAKEN
def write_log(start_time, duration):
    # Bepaal de locatie van het script
    script_path = os.path.dirname(os.path.realpath(__file__))
    log_file_path = os.path.join(script_path, "script_usage_log.txt")

    # Schrijf de log
    with open(log_file_path, "a") as log_file:
        log_file.write("{}: Script gestart om {}, duur {} seconden\n".format(datetime.datetime.now(),start_time, duration))

def read_excel_data(excel_file_path):
    """Leest de gegevens uit het Excel-bestand en retourneert een lijst van tuples."""
    # Een lijst om de gelezen gegevens van het Excel-bestand op te slaan
    data = []
    # Open het Excel-bestand en lees de gegevens
    excel_app = ApplicationClass()
    excel_app.Visible = False
    workbook = excel_app.Workbooks.Open(excel_file_path)
    worksheet = workbook.Sheets["BPL_ID to RVT"]  #het tabblad dient deze naam te hebben om te werken

    # Zoek de laatste rij met gegevens in het werkblad
    last_row = worksheet.Cells(worksheet.Rows.Count, 1).End(XlDirection.xlUp).Row

    # Lees de gegevens van het werkblad en sla deze op in de 'data' lijst
    for row in range(2, last_row + 1):
        element_id = worksheet.Cells(row, 3).Value2 #element id bevindt zich in de derde kolom van de excel
        value = worksheet.Cells(row, 4).Text #OMI_CTE_BimCollab waarde bevindt zich in de vierde kolom

        # Controleer of de element-ID als geheel getal kan worden behandeld
        if isinstance(element_id, (int, float)):
            element_id = int(element_id)

        data.append((element_id, value))

    # Sluit het Excel-bestand
    workbook.Close()
    excel_app.Quit()

    return data

def update_elements(data):
    """Bijwerken van de subcontractor parameter in Revit voor elk element in de data lijst."""
    # Markeer de starttijd van het script
    start_time = datetime.datetime.now()

    # Start een transactie om alle updates in te groeperen
    with Transaction(doc, "Update D-link") as transaction:
        transaction.Start()

        # Loop door de lijst met gegevens en pas het commentaarveld aan voor elk element
        for element_id, value in data:
            try:
                # Zoek het element op basis van het element-ID
                element_id_int = int(element_id)
                element = doc.GetElement(ElementId(element_id_int))

                # Controleer of het element bestaat en of het de parameter heeft dat kan worden gewijzigd
                if element and hasattr(element, "get_Parameter"):
                    parameter_name = "OMI_CTE_BimCollab"
                    parameter = element.LookupParameter(parameter_name)
                    if parameter and not parameter.IsReadOnly:
                        #controleer of de parameter reeds een waarde heeft, zoja mag men deze niet overschrijven
                        if parameter.HasValue:
                            current_value = parameter.AsValueString()
                            if current_value:
                                continue
                        # Pas aan met de bijgewerkte waarde indien de parameter nog geen waarde heeft
                        parameter.Set(str(value))


            except Exception as e:
                print("Fout bij het bijwerken van element " + str(element_id) + ": " + str(e))

        # Commit de transactie
        transaction.Commit()

    # Bereken de duur en schrijf naar het logbestand
    end_time = datetime.datetime.now()
    duration = (end_time - start_time).total_seconds()
    write_log(start_time, duration)

def main():
    """Hoofdfunctie om het importproces te starten."""
    # Vraag de gebruiker om het Excel-bestand te selecteren dat geïmporteerd moet worden
    open_file_dialog = OpenFileDialog()
    open_file_dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    open_file_dialog.Title = "Selecteer het Excel-bestand om te importeren"
    result = open_file_dialog.ShowDialog()

    if result == DialogResult.OK:
        excel_file_path = open_file_dialog.FileName
    else:
        # Toon een foutmelding als er geen bestand is geselecteerd
        MessageBox.Show("Geen bestand geselecteerd. Het script wordt afgebroken.")
        sys.exit()

    # Lees de gegevens uit het Excel-bestand
    data = read_excel_data(excel_file_path)

    # Werk de elementcommentaren bij in Revit
    update_elements(data)

    # Toon een bevestigingsbericht
    MessageBox.Show("Gegevens succesvol geïmporteerd naar Revit.")


if __name__ == "__main__":
    main()

