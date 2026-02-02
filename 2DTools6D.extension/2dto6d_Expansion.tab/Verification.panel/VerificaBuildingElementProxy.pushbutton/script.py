# -*- coding: utf-8 -*-

""" Verifica la presenza di elementi building element proxy """

__title__ = 'Check Building\nElement Proxy'
__author__ = '2dto6d'


######################################

import Autodesk.Revit
import Autodesk.Revit.DB
import Autodesk.Revit.DB.Architecture
import pyrevit
from pyrevit import forms, script
import codecs
import re
from re import *
from Autodesk.Revit.DB import BuiltInCategory, Category, CategoryType
import System
from System import Enum
from System.Collections.Generic import Dictionary
import unicodedata
import math
import clr
import os
import sys
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType
import time
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

#PER IL FLOATING POINT
from decimal import Decimal, ROUND_DOWN, getcontext

##############################################################
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument       
app   = __revit__.Application     
aview = doc.ActiveView
output = pyrevit.output.get_output()

t = Transaction(doc, "Colorazione IFCSaveAs")
##############################################################


#PREPARAZIONE OUTPUT
output = pyrevit.output.get_output()
# CREAZIONE LISTE DI OUTPUT DATA
BUILDINGELEMENTPROXY_CSV_OUTPUT = []
BUILDINGELEMENTPROXY_CSV_OUTPUT.append(["Categoria","Nome Famiglia","Nome Tipo","ID Elemento","IFC Class","Stato"])

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


# FILTRO ELEMENTI VALORIZZATI
Filtro_AssegnazioneEffettuata = ElementParameterFilter(HasValueFilterRule(ElementId(BuiltInParameter.IFC_EXPORT_ELEMENT_AS)))
Temp_Collector_IFC_Valorizzato = FilteredElementCollector(doc,aview.Id).WherePasses(Filtro_AssegnazioneEffettuata).WhereElementIsNotElementType().ToElements()
Collector_IFC_Valorizzato = []
# VERIFICA INCROCIATA PER EVITARE VALORI VUOTI 
for elemento in Temp_Collector_IFC_Valorizzato:
    if elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS).AsString() != "" and "CenterLine" not in str(elemento.Category.BuiltInCategory):
        Collector_IFC_Valorizzato.append(elemento)

DataTable = []
for Elemento in Collector_IFC_Valorizzato:
    if "CenterLine" in str(Elemento.Category.BuiltInCategory):
        continue
    else:
    # ESTRAGGO INFORMAZIONI
        Stato = 1
        Categoria = str(Elemento.Category.BuiltInCategory)
        ID_Elemento = Elemento.Id
        Nome_Famiglia = EstraiInfoOggetto(Elemento)[0]
        Nome_Tipo = EstraiInfoOggetto(Elemento)[1]

        param_ifc_class = Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS)
        param_ifc_predef = Elemento.get_Parameter(BuiltInParameter.IFC_EXPORT_PREDEFINEDTYPE)

        Current_IfcClass = param_ifc_class.AsValueString() if param_ifc_class and param_ifc_class.HasValue else ""
        Current_Predef = param_ifc_predef.AsValueString() if param_ifc_predef and param_ifc_predef.HasValue else ""


    # VERIFICA VALORIZZAZIONE EXPORT AS
    if Current_IfcClass != "IfcBuildingElementProxy":
        BUILDINGELEMENTPROXY_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,Current_IfcClass,Stato])
    else: #! VERIFICARE IN FUTURO SE LASCIANO LA CATEGORIA GENERIC MODEL O LA CAMBIANO CON ALTRO
        if Current_IfcClass == "IfcBuildingElementProxy" and Categoria == "OST_GenericModel":
            Stato = 1
            BUILDINGELEMENTPROXY_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,Current_IfcClass,Stato])
            DataTable.append([Nome_Famiglia, Nome_Tipo, Categoria, ID_Elemento, Current_IfcClass, ":heavy_white_check_mark:"])
        else:
            Stato = 0
            BUILDINGELEMENTPROXY_CSV_OUTPUT.append([Categoria,Nome_Famiglia,Nome_Tipo,ID_Elemento,Current_IfcClass,Stato])
            DataTable.append([Nome_Famiglia, Nome_Tipo, Categoria, ID_Elemento, Current_IfcClass, ":cross_mark:"])

output = pyrevit.output.get_output()
output.print_md("# Verifica Elementi 'Building Element Proxy'")
output.print_md("---")


if DataTable:
    output.freeze()
    output.print_table(table_data = DataTable, columns = ["Nome Famiglia", "Nome Tipo","Categoria", "ID Elemento","IfcClass","Verifica"])
    output.resize(1500, 900)
    output.unfreeze()
else:
    output.print_md("##:white_heavy_check_mark: Nessun elemento da verificare :white_heavy_check_mark:")



