""" Fast grids quote"""
__title__= "Fast Grids \nQuote"

# Boilerplate text
import clr #permette di entrare in comunicazione tra le dll "common language runtime" ----Sempre necessario----


import sys
sys.path.append('C:/Program Files (x86)/IronPython 2.7/Lib')

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
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import *    #importare tutte le classi delle API  ----Eliminare se non necessario----
from Autodesk.Revit.UI import *    #importare tutte le classi delle APIUI  ----Eliminare se non necessario----

from pyrevit import forms

from rpw.ui.forms import TextInput

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

t = Transaction(doc,"Fast Grid Quote")

def transpose(matrix):
    rows = len(matrix)
    columns = len(matrix[0])

    matrix_T = []
    for j in range(columns):
        row = []
        for i in range(rows):
           row.append(matrix[i][j])
        matrix_T.append(row)

    return matrix_T


def reconstructPoint(point,newvalueX,newvalueY,newvalueZ):
	xcoord = point.X
	ycoord = point.Y
	zcoord = point.Z
	if newvalueX != "X":
		xcoord = newvalueX
	else:
		pass
	if newvalueY != "X":
		ycoord = newvalueY
	else:
		pass
	if newvalueZ != "X":
		zcoord = newvalueZ
	else:
		pass
	return XYZ(xcoord,ycoord,zcoord)

def getIntersection(line1,line2):
	array = clr.Reference[IntersectionResultArray]()
	result = line1.Intersect(line2, array)
	
	if result == SetComparisonResult.Overlap:
		intersection = array.Item[0].XYZPoint
		return intersection
		
def checkIntersection(line1,line2):
	array = clr.Reference[IntersectionResultArray]()
	result = line1.Intersect(line2, array)
	
	if result == SetComparisonResult.Overlap:
		check = True
		return check

t = Transaction(doc,"Fast Quote")

active_view = doc.ActiveView
active_selection = uidoc.Selection.GetElementIds()

multiple_selection = []


selected_grids = [doc.GetElement(x) for x in active_selection]


quote_reference = ReferenceArray()
for grid in selected_grids:
	quote_reference.Append(Reference(grid))

t.Start()
try:
	
	

	selected_grids_curve = [x.Curve.Direction for x in selected_grids]

	quote_direction = selected_grids_curve[0].CrossProduct(active_view.ViewDirection)
	
	quote_base_point = uidoc.Selection.PickPoint()
	quote_reference_line = Line.CreateUnbound(quote_base_point,quote_direction)
	doc.Create.NewDimension(active_view, quote_reference_line, quote_reference)
except:
	
	
	selected_grid_curve = [x.GetCurvesInView(DatumExtentType.ViewSpecific,active_view)[0].Direction for x in selected_grids] 
	quote_direction = active_view.ViewDirection.CrossProduct(selected_grid_curve[0])
	
	active_view.SketchPlane = SketchPlane.Create(doc,Plane.CreateByNormalAndOrigin(active_view.ViewDirection,active_view.Origin))

	quote_base_point = uidoc.Selection.PickPoint()
	
	quote_reference_line = Line.CreateUnbound(quote_base_point,quote_direction)
	doc.Create.NewDimension(active_view, quote_reference_line, quote_reference)
	
t.Commit()

