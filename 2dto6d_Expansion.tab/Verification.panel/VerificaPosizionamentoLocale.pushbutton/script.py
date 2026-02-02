# -*- coding: utf-8 -*-
""" Verifica non siano presenti locali non delimitati nel progetto, ridondanti, non posizionati  """

__author__ = '2dto6d'
__title__ = 'Verifica Status\nLocali'

import codecs
import re
import unicodedata
import pyrevit
from pyrevit import forms, script
import math
import clr
import os
import sys
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import csv
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *
#PER IL FLOATING POINT
from decimal import Decimal, ROUND_DOWN, getcontext

##############################################################
doc   = __revit__.ActiveUIDocument.Document  #type: Document
uidoc = __revit__.ActiveUIDocument						   
app   = __revit__.Application		
aview = doc.ActiveView
output = pyrevit.output.get_output()

t = Transaction(doc, "Verifica Status Locali")

##############################################################


# CREAZIONE LISTE DI OUTPUT DATA
LOCAL_POSITION_CSV_OUTPUT = []
LOCAL_POSITION_CSV_OUTPUT.append(["Nome Locale","ID Elemento","Verifica","Stato"]) #DEFINIRE SE LOCALE RIDONTANTE O NON POSIZIONATO

######################################################################################
def EstraiFamigliaOggetto(oggetto):
	try:
		return oggetto.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
	except:
		try:
			return oggetto.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString()
		except:
			return "NON TROVATO"
		
LocaliProgetto = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()

NotEnclosed = BuiltInFailures.RoomFailures.RoomNotEnclosed.Guid
NotSingleRoom = BuiltInFailures.RoomFailures.RoomsInSameRegionRooms.Guid

LocaliNonRacchiusi = []
LocaliRidondanti = []
LocaliNonPosizionati = []
ListaErrori = []

for Warning in doc.GetWarnings():
	if Warning.GetFailureDefinitionId().Guid == NotEnclosed:
		ElementoErrorID = Warning.GetFailingElements()[0]
		ElementoError = doc.GetElement(ElementoErrorID)
		room_id = output.linkify(ElementoErrorID)
		ListaErrori.append(ElementoErrorID)
		room_name = ElementoError.get_Parameter(BuiltInParameter.ROOM_NAME).AsValueString()
		failure_message = "Locale non racchiuso"
		
		LocaliNonRacchiusi.append([room_name,room_id,failure_message])
		LOCAL_POSITION_CSV_OUTPUT.append([room_name,ElementoErrorID,failure_message,0])

	elif Warning.GetFailureDefinitionId().Guid == NotSingleRoom:
		ElementoErrorID = Warning.GetFailingElements()[1]
		ElementoError = doc.GetElement(ElementoErrorID)
		room_id = output.linkify(ElementoErrorID)
		ListaErrori.append(ElementoErrorID)
		room_name = ElementoError.get_Parameter(BuiltInParameter.ROOM_NAME).AsValueString()
		failure_message = "Locale ridondante"

		LocaliRidondanti.append([room_name,room_id,failure_message])
		LOCAL_POSITION_CSV_OUTPUT.append([room_name,ElementoErrorID,failure_message,0])

for Locale in LocaliProgetto:
	room_name = Locale.get_Parameter(BuiltInParameter.ROOM_NAME).AsValueString()
	room_id = output.linkify(Locale.Id)
	failure_message = "Locale non posizionato"
	if not Locale.Location:
		ListaErrori.append(Locale.Id)
		LocaliNonPosizionati.append([room_name,room_id,failure_message])
		LOCAL_POSITION_CSV_OUTPUT.append([room_name,Locale.Id,failure_message,0])

	if Locale.Id not in ListaErrori:
		LOCAL_POSITION_CSV_OUTPUT.append([room_name,Locale.Id,"Locale Posizionato",1])

		
output.print_md("# Verifica Status Locali")
output.print_md("---")

if LocaliNonPosizionati:
	output.print_table(table_data = LocaliNonPosizionati,title = "Locali Non Posizionati", columns = ["Nome Locale","ID Elemento","Verifica","Stato"],formats = ["","",""])
else:
	output.print_md("## Tutti i locali sono posizionati correttamente")
if LocaliRidondanti:
	output.print_table(table_data = LocaliRidondanti,title = "Locali Ridondanti", columns = ["Nome Locale","ID Elemento","Verifica","Stato"],formats = ["","",""])
else:
	output.print_md("## Non sono presenti locali ridondanti")
if LocaliNonRacchiusi:
	output.print_table(table_data = LocaliNonRacchiusi,title = "Locali Non Racchiusi", columns = ["Nome Locale","ID Elemento","Verifica","Stato"],formats = ["","",""])
else:
	output.print_md("## Non sono presenti locali non racchiusi")

output.print_md("---")