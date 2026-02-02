# -*- coding: utf-8 -*-
""" Trova elementi in base al GUID. """

__author__ = '2dto6d'
__title__ = 'Cerca GUID'

import codecs
import re
import unicodedata
import pyrevit
from pyrevit import *
from pyrevit import forms, script
import sys
import clr
import System
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType
import math
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

##############################################################
doc   = __revit__.ActiveUIDocument.Document  #type: Document
uidoc = __revit__.ActiveUIDocument						   
app   = __revit__.Application		
output = pyrevit.output.get_output()
aview = doc.ActiveView
##############################################################

def EstraiInfoOggetto(oggetto):
    # SE L'OGGETTO E' UNA FAMIGLIA CARICABILE MI PRENDE FAMIGLIA E TIPO
    if isinstance(oggetto,FamilyInstance):
        return [oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString(),oggetto.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()]
    # SE L'OGGETTO E' UNA FAMIGLIA DI SISTEMA MI PRENDE SOLO IL TIPO
    else:
        try:
            #return oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
            return [oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString(),oggetto.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()]
        except:
            try:
                return oggetto.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString()
            except:
                return "NON TROVATO"
            


Collector = FilteredElementCollector(doc,doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
Ricerca_GUID = forms.ask_for_string(    
    prompt='Inserisci GUID...',
    title='Cerca GUID'
)


newselection = List[ElementId]()
selection = uidoc.Selection

for element in Collector:
    
    Elemento_GUID = element.get_Parameter(BuiltInParameter.IFC_GUID).AsString()
        
    if Ricerca_GUID in Elemento_GUID:
        
        forms.toast(
            "Elemento trovato : {} - {}\nID:{}".format(element.Category.Name, EstraiInfoOggetto(element),element.Id),
             title="Cerca GUID",
            )
        newselection.Add(element.Id)
        selection.SetElementIds(newselection)
        break
        found = 1
        
if found != 1:
    forms.toast(
        "Elemento non presente",
            title="Cerca GUID",
        )