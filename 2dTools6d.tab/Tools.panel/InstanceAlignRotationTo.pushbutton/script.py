""" Align family instance to direction """
__title__= "Align Instance\nTo Dir"
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

t = Transaction(doc,"Align Istances")

ActiveView = doc.ActiveView


if len(uidoc.Selection.GetElementIds()) != 0:
	SelezioneAttiva = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()]
	with forms.WarningBar(title='Seleziona i punti su cui allineare'):
		primoPunto = uidoc.Selection.PickPoint()
		secondoPunto = uidoc.Selection.PickPoint()
	AngoloRiferimento = secondoPunto.Subtract(primoPunto)



	


	t.Start()
	
	for elemento in SelezioneAttiva:
		try:
			elemento.Location.Rotate(Line.CreateBound(elemento.Location.Point,XYZ(elemento.Location.Point.X,elemento.Location.Point.Y,1)),elemento.FacingOrientation.AngleOnPlaneTo(AngoloRiferimento,ActiveView.ViewDirection))
		except:
			AngoloIniziale = elemento.FacingOrientation.AngleOnPlaneTo(AngoloRiferimento,ActiveView.ViewDirection)
			AngoloUtile = AngoloIniziale + 1.5708
			elemento.Location.Rotate(Line.CreateBound(elemento.Location.Curve.Evaluate(0.5,True),XYZ(elemento.Location.Curve.Evaluate(0.5,True).X,elemento.Location.Curve.Evaluate(0.5,True).Y,1)),AngoloUtile)
	t.Commit()
else:
	pass