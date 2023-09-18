# -*- coding: utf-8 -*-

__title__ = "BEF-Zone"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 18-09-23
_____________________________________________________________________
Description:

A.d.h.v. het zoneringsmodel zal men de parameter BEF_zone invullen in alle elementen

_____________________________________________________________
Last update:

- [18-09-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""

#-----------------------IMPORTS-------------------------------------------------------
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, BoundingBoxIntersectsFilter, Outline, XYZ
from Autodesk.Revit.DB import ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Verkrijg alle massaelementen uit het gelinkte model
linked_doc = [x.GetLinkDocument() for x in FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance) if x.GetLinkDocument() and x.GetLinkDocument().Title == "ZIN_JE5_ARC_EXE_Zoneringsmodel"]

masses = FilteredElementCollector(linked_doc[0]).OfCategory(BuiltInCategory.OST_Mass).ToElements()

# Start een transactie om wijzigingen in het document aan te brengen
t = Transaction(doc, "Copy BEF_zone parameter")
t.Start()

try:
    for mass in masses:
        mass_param = mass.LookupParameter("BEF_zone")
        if mass_param:
            mass_zone_value = mass_param.AsString()  # Of gebruik .AsValueString(), afhankelijk van het parametertype

            # Creëer een filter op basis van de omtrek van de massa
            mass_bounding_box = mass.get_BoundingBox(None)
            outline = Outline(mass_bounding_box.Min, mass_bounding_box.Max)
            bb_filter = BoundingBoxIntersectsFilter(outline)

            # Verzamel alle elementen die de filtercriteria matchen (die zich binnen de mass bevinden)
            host_elements = FilteredElementCollector(doc).WherePasses(
                bb_filter).WhereElementIsNotElementType().ToElements()

            for host_element in host_elements:
                host_element_param = host_element.LookupParameter("BEF_zone")
                if host_element_param:
                    current_value = host_element_param.AsString()  # Of gebruik .AsValueString(), afhankelijk van het parametertype

                    # Controleer of de huidige waarde niet is ingesteld of niet overeenkomt met de mass_zone_value
                    if not host_element_param.IsReadOnly and (
                            current_value is None or current_value != mass_zone_value):
                        host_element_param.Set(mass_zone_value)

except Exception as e:
    print("Exception: {}".format(e))
    t.RollBack()
else:
    t.Commit()
