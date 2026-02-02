""" Align text to direction """
__title__= "Align Text\nTo Dir"
import clr
import System

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

t = Transaction(doc,"Align Text")

ActiveView = doc.ActiveView


if len(uidoc.Selection.GetElementIds()) != 0:
	SelezioneAttiva = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()]
	with forms.WarningBar(title='Seleziona i punti su cui allineare'):
		primoPunto = uidoc.Selection.PickPoint()
		secondoPunto = uidoc.Selection.PickPoint()
	AngoloRiferimento = secondoPunto.Subtract(primoPunto)
	t.Start()
	for Testo in SelezioneAttiva:
		DirTesto = Testo.BaseDirection
		Testo.Location.Rotate(Line.CreateBound(Testo.Coord,XYZ(Testo.Coord.X,Testo.Coord.Y,1)),DirTesto.AngleOnPlaneTo(AngoloRiferimento,ActiveView.ViewDirection))
	t.Commit()
