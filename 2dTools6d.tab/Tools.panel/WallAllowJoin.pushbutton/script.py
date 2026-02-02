
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
from System import Guid

import Autodesk 
from Autodesk.Revit.DB import *    #importare tutte le classi delle API  ----Eliminare se non necessario----
from Autodesk.Revit.UI import *    #importare tutte le classi delle APIUI  ----Eliminare se non necessario----
import Autodesk.Revit.DB as DB

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

t = Transaction(doc,"Allow Joins")

ids = uidoc.Selection.GetElementIds()

t.Start()
for i in ids:
    DB.WallUtils.AllowWallJoinAtEnd(doc.GetElement(i),1)
    DB.WallUtils.AllowWallJoinAtEnd(doc.GetElement(i),0)
t.Commit()