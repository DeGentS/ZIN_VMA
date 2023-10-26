# -*- coding: utf-8 -*-
__title__ = "BIM360"
__author__ = "Sean De Gent i.o.v. BimPlan"
__doc__ = """Version = 1.0
Date    = 02-08-23
_____________________________________________________________________
Description:

Druk op de knop om naar de BIM360 omgeving te gaan van ZIN
_____________________________________________________________________

Last update:

- [02.08.23] 1.0 RELEASE

author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________

"""

#-----------------------IMPORTS-------------------------------------------------------

import webbrowser
from pyrevit import script

#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------

# set up logger
logger = script.get_logger()

# URL to open
url = 'https://docs.b360.autodesk.com/projects/0e5a1474-c5d9-49cd-b8e8-7cb853cfdfe0/folders/'

try:
    # try to open the URL in the default web browser
    webbrowser.open(url)
    logger.info('Successfully opened ' + url)
except Exception as e:
    # if there's an error, log it
    logger.error('Failed to open ' + url, e)
