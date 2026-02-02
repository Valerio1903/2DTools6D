"""Select Elements based on Category and Type, 1 - Select one or more Categories. 2 - Select Type"""

__title__= 'Multi-Category\n and Type'
__author__= '2dto6d'

import System
import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

from pyrevit import revit, forms

doc = __revit__.ActiveUIDocument.Document

categories =doc.Settings.Categories

model_cat = []
model_cate = []
for c in categories:
	if c.CategoryType == CategoryType.Model:
		if "dwg"  not in c.Name and c.SubCategories.Size > 0 or c.CanAddSubcategory:
			model_cat.append(c.Name)
			model_cate.append(c)


sortlist = sorted(model_cat)

value = forms.SelectFromList.show(
        {'All Categories': sortlist},
        title='Select Categories',
        multiselect=True
    )

category =[]
namer = []

for ci in model_cate:
	namer.append(ci.Name)

for n,c in zip(namer,model_cate):
	for r in value:
		if n == r:
			category.append(c)

categoriesId = []
bic = []

for c in category:
	categoriesId.append(c.Id)
	bic.append(System.Enum.ToObject(BuiltInCategory, c.Id.IntegerValue))

fam_typ = []
#fam_type = FilteredElementCollector(doc,doc.ActiveView.Id).OfCategory(category).WhereElementIsElementType().ToElements()
for c in bic:
	fam_typ.append(FilteredElementCollector(doc).OfCategory(c).WhereElementIsElementType().ToElements())

name_type = []

for t in fam_typ:
	for c in t:
		try:
			name_type.append(c.Category.Name+' | '+c.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
		except:
			name_type.append(c.Category.Name+' | '+c.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsString())
value_ty = forms.SelectFromList.show(
        {'All': name_type},
        title='Select Type',
        multiselect=True
    )

result_ty =[]
typ_name_result = []
for v in value_ty:
	result_ty.append(v.split(" | ",1))

for n in result_ty:
	typ_name_result.append(n[1])


selected_option = forms.CommandSwitchWindow.show(
		['YES','NO'],
		message='Select in Active View?',
)

collection = []
if selected_option == 'Yes':
	for c in bic:
		collection.append(FilteredElementCollector(doc,doc.ActiveView.Id).OfCategory(c).WhereElementIsNotElementType().ToElements())
else:
	for c in bic:
		collection.append(FilteredElementCollector(doc).OfCategory(c).WhereElementIsNotElementType().ToElements())

selectedID = []
for c in collection:
	for i in c:
		if i.Name in typ_name_result:
			selectedID.append(i.Id)

revit.get_selection().set_to(selectedID)


forms.toast(
	"{} elementi selezionati".format(len(selectedID)),
	title = "Selezione Multi Category"
)

