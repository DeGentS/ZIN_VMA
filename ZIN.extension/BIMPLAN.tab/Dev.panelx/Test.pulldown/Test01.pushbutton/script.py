# -*- coding: utf-8 -*-
__title__   = "Test_hostId"
__doc__     = """Hiermee kijken hoeveel elementen een waarde hebben onder de parameter "BERSnl_C_codering_systeem01" """

#-----------------------IMPORTS-------------------------------------------------------


import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit import DB

from Autodesk.Revit.DB import BuiltInCategory as Bic


#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#for testing
element_id = ElementId(23799111)
element = doc.GetElement(element_id)

# linked_room_id = ElementId(9880660)
# linked_room = doc.GetElement(linked_room_id)

# Create an empty list to store the linked documents
linked_doc = []

# Verkijg het linked model
collector = FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)

# Loop through each link instance in the collector
for link_instance in collector:
    # Get the linked document from the link instance
    link_document = link_instance.GetLinkDocument()
    # linked_doc.append(link_document)
    link_status = link_instance.LinkedFileStatus.GetLinkFileStatus()
    print(link_status,link_instance.Title)

    # Check if the linked document exists and if its title matches the desired title
    # if link_document and link_document.Title == "ZIN_JE5_ARC_EXE_Architectuur":
    #     linked_doc.append(link_document)

# Assuming the first linked document in the list is the one you're interested in
# (you should add checks for this), retrieve an element by its ID from that document:
# if linked_doc:
#     test_linked_document = linked_doc[0]  # take the first linked document
#     test_element_id = ElementId(9880660)  # replace '12345678' with your specific ID
#     test_element = test_linked_document.GetElement(test_element_id)

#--------------------------------------------------------------------------------------


t               = Transaction(doc, __title__)   #dit zorgt ervoor dat bij de undo/redo button de naam van __title__ staat
                                                # Transaction = nodig wanneer we een handeling of aanpassing wensen door te voeren
# L060.B.0B.02

# print(element.HostFace.ElementId)
# print(test_element.Id)
# linked_roomname = test_element.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
# linked_room_number = test_element.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
#
#
# zin_lokaal_element = element.LookupParameter("ZIN_lokaal").AsValueString() if element.LookupParameter("ZIN_lokaal").AsValueString() else "N/A"
#
# print(zin_lokaal_element)
#
# t.Start()
#
#
# element.LookupParameter("ZIN_lokaal").Set(linked_room_number)
#
# t.Commit()