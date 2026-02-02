"""Creates a bounding box around linked element"""

__title__= "Section Box On\nLinked Element"
__author__ = '2dto6d'

import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from pyrevit import forms

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

uiviews = uidoc.GetOpenUIViews()
view = doc.ActiveView
uiview = [x for x in uiviews if x.ViewId == view.Id][0]

## Link List ##
links = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements()
linkString=[]
for link in links:
    try:
        linkString.append(link.GetLinkDocument().Title)
    except:
        linkString.append("<Unloaded Link>")

linkOfRoom = forms.SelectFromList.show(context = linkString,title = "Seleziona modello linkato", width  = 500, height = 500)
if not linkOfRoom:
	import script
	script.exit()

linkInstance = links[linkString.index(linkOfRoom)]
######
elemIdStr = forms.ask_for_string(prompt = "Insert Linked Element ID", title = "Linked Element Id")
if not elemIdStr:
	import script
	script.exit()

elemId = int(elemIdStr)
selection = linkInstance.GetLinkDocument().GetElement(ElementId(elemId))

# Funzione di filtro per selezionare solo viste 3D non template
def is_3d_view(v):
	return v.ViewType == ViewType.ThreeD and not v.IsTemplate

isoView = forms.select_views(
	title = "Select One View",
	button_name= "Select",
	width=500,
	multiple=False,
	filterfunc = is_3d_view
)

if not isoView:
	import script
	script.exit()

bb = selection.get_BoundingBox(None)

t = Transaction(doc,"Section Box On Id")

t.Start()
try:
	isoView.SetSectionBox(bb)
except:
	t.RollBack()
else:
	t.Commit()

uidoc.ActiveView = isoView
