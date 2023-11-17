# -*- coding: utf-8 -*-

__title__ = "Set Host ID "
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 26-10-23
_____________________________________________________________________
Description:

Hiermee zal de opening voorzien worden met de element ID van de Host
waarin de opening voorzien is.
_____________________________________________________________
Last update:

- [26-10-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""

#-----------------------IMPORTS-------------------------------------------------------

import clr
import os
import sys
import System


# Importeren van Revit API-elementen
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction

#----------------------VARIABLES--------------------------------------------------------

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
t               = Transaction(doc, __title__)

#----------------------MAIN--------------------------------------------------------


genericmodels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType()
Opening = 'Opening'
element_guid = None




for gm in genericmodels:
    if gm.SuperComponent is None:
        family_name = gm.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
        if Opening in family_name:
            omi_host_id_param = gm.LookupParameter("OMI_CTE_Host Element ID")
            host = gm.HostFace
            t.Start()
            if host is not None:
                linked_element_id = host.LinkedElementId
                # if linked_element_id != ElementId.InvalidElementId:
                omi_host_id_param.Set(linked_element_id)

            if host is None:
                omi_host_id_param.Set("N/A")

            t.Commit()
#
# except Exception as e:
#     print("Fout: {}".format(e))
#     t.RollBack()
