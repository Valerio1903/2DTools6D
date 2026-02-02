"""Select Elements based on Rectangular selection and Category """

__title__= 'Rectangular\nBy Category'
__author__= '2dto6d'

import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import ElementId

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import *

from System.Collections.Generic import List

from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

categories = doc.Settings.Categories

model_cat = []
cat_dict = {}  # Dictionary to map category names to category objects


for c in categories:
	if c.CategoryType == CategoryType.Model:
		if "dwg"  not in c.Name and c.SubCategories.Size > 0 or c.CanAddSubcategory:
			model_cat.append(c.Name)
			cat_dict[c.Name] = c  # Store the category object


value = forms.SelectFromList.show(
        {'All Categories': model_cat},
        title='Select Categories',
        multiselect=False
    )


actual_selection = List[ElementId]()
selected_category = cat_dict[value]  # Get the actual category object


# Use the selected category's BuiltInCategory
# actual_selection.Add(selected_category.BuiltInCategory)

class MySelectionFilter(ISelectionFilter):
	def __init__(self):
		pass
	def AllowElement(self, element):
		if element.Category.Name == value:
			return True
		else:
			return False
	def AllowReference(self, element):
		return False

selection_filter = MySelectionFilter()

output = uidoc.Selection.PickElementsByRectangle(selection_filter)

outputID = []

for i in output:
	outputID.append(i.Id)

collection = List[ElementId](outputID)

select = uidoc.Selection.SetElementIds(collection)