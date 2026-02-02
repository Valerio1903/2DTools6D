"""Gets element's Linked ID"""

__title__= "Get Linked\nElement ID"
__author__ = '2dto6d'

# Boilerplate text
import clr #permette di entrare in comunicazione tra le dll "common language runtime" ----Sempre necessario----

import sys
sys.path.append('C:\Program Files (x86)\IronPython 2.7\Lib')

import System   #permette di importare la libreria dei metodi interni del sistema windows ----Eliminare se non necessario----
from System import Array
from System.Collections.Generic import *

clr.AddReference('ProtoGeometry') #permette di importare la libreria delle geometrie di Dynamo
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitNodes") #permette di importare la libreria delle classi Revit di Dynamo
import Revit
clr.ImportExtensions(Revit.Elements) #Permette di importare il ToDSType
clr.ImportExtensions(Revit.GeometryConversion) #Permette di importare i metodi per convertire 

clr.AddReference("RevitServices") #permette di importare le librerie del Documento attivo e del Transaction necessarie per utilizzare le API
import RevitServices
from RevitServices.Persistence import DocumentManager #permette di importare il modulo del documento attivo
from RevitServices.Transactions import TransactionManager #permette di importare il modulo per eseguire modifiche

clr.AddReference("RevitAPI")  #permette di utilizzare la libreria delle APIdocs  ----Eliminare se non necessario----
clr.AddReference("RevitAPIUI") #permette di utilizzare l'User Interface e l'applicazione corrente di Revit ----Eliminare se non necessario----

import Autodesk 
from Autodesk.Revit.DB import *    #importare tutte le classi delle API  ----Eliminare se non necessario----
from Autodesk.Revit.UI import *    #importare tutte le classi delle APIUI  ----Eliminare se non necessario----

clr.AddReference('System.Windows.Forms')

from System.Windows.Forms import Clipboard

import subprocess

from pyrevit import script, forms

def copy2clip(txt):
    cmd='echo '+txt.strip()+'|clip'
    return subprocess.check_call(cmd, shell=True)


doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

uiviews = uidoc.GetOpenUIViews()
view = doc.ActiveView

outputWin = script.get_output()

sel1 = uidoc.Selection
ot = Selection.ObjectType.LinkedElement

try:
	el_ref = sel1.PickObjects(ot, "Pick a linked element.")
except:
	script.exit()

if not el_ref:
	script.exit()

outputInfo = []
allIds = []
allDocs = []

outputWin.print_md('# Linked Element ID')
outputWin.insert_divider()

for item in el_ref:
	
	linkInst = doc.GetElement(item.ElementId)
	linkDoc = linkInst.GetLinkDocument()
	linkEl = linkDoc.GetElement(item.LinkedElementId)
	
	allIds.append(str(item.LinkedElementId))
	allDocs.append(str(linkDoc.Title))

	if linkEl.GetType() == Autodesk.Revit.DB.FamilyInstance:
		outputWin.print_md("ID : {} of model : {}".format(linkEl.Id,linkDoc.Title))
		outputWin.print_md("Family Name : {}".format(linkEl.Symbol.FamilyName))
		outputWin.print_md("Type : {}".format(linkEl.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()))
		outputWin.print_md("\n")
	else:
		outputWin.print_md("ID : {} of model : {}".format(linkEl.Id,linkDoc.Title))
		outputWin.print_md("Family Type : {}".format(linkEl.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()))
		outputWin.print_md("\n")
outputWin.insert_divider()

forms.toast(
    "ID elementi copiati in clipboard",
    title="Linked element ID",
)

Clipboard.SetText(",".join(allIds))



