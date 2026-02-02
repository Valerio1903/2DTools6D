"""Isolate Warning Elements in a New 3D View"""

__title__= 'Isolate\nWarnings'
__author__= '2dto6d'


import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('System')
from System.Collections.Generic import List

from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document


doc_warn = doc.GetWarnings()

if len(doc_warn) == 0:
	forms.alert('There are no Warnings to isolate ... Good job!', exitscript=True)

else:
		
	# Extract element IDs from warning messages
	warn_element_ids = []
	for warning in doc_warn:
		failing_elements = warning.GetFailingElements()
		for elem_id in failing_elements:
			if elem_id not in warn_element_ids:
				warn_element_ids.append(elem_id)

	doc_vtype = FilteredElementCollector(doc).OfClass(ViewFamilyType)

	present_views = FilteredElementCollector(doc).OfClass(View3D).WhereElementIsNotElementType().ToElements()

	for x in doc_vtype:
		if "3D" in x.FamilyName:
			vtid = x.Id
			break

	# Delete existing "Warnings View" if it exists
	for view in present_views:
		if view.Name == "Warnings View":
			with Transaction(doc, "Delete Old Warnings View") as t:
				t.Start()
				doc.Delete(view.Id)
				t.Commit()
			break

	# Create new warnings view
	t = Transaction(doc, "Create Isolate Warnings View")
	t.Start()
	try:
		isoView = View3D.CreateIsometric(doc, vtid)
		isoView.Name = "Warnings View"
		isoView.IsolateElementsTemporary(List[ElementId](warn_element_ids))

		doc.Regenerate()

		isoView.ConvertTemporaryHideIsolateToPermanent()
		t.Commit()
	except Exception as e:
		print(e)
		t.RollBack()

