"""Rimuove il valore del "Contrassegno / Mark" per rimuovere il warning."""

__title__= 'Fix All\nMark Warnings'
__author__= '2dto6d'

import clr
import System

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document


# Recupera tutte le categorie disponibili nel documento
categories = doc.Settings.Categories

# Crea un dizionario per mappare nome categoria -> oggetto categoria (lookup O(1))
model_categories = {}

# Itera su tutte le categorie del documento
for c in categories:
	# Filtra solo le categorie di tipo Model (esclude Annotation, etc.)
	if c.CategoryType == CategoryType.Model:
		# Verifica che la categoria non sia legata a file DWG E che abbia sottocategorie o possa averne
		# Nota: le parentesi chiariscono la precedenza degli operatori logici
		if ("dwg" not in c.Name) and (c.SubCategories.Size > 0 or c.CanAddSubcategory):
			# Salva nel dizionario: chiave = nome, valore = oggetto categoria
			model_categories[c.Name] = c

# Ordina alfabeticamente i nomi delle categorie per una migliore UX
sorted_names = sorted(model_categories.keys())

# Mostra la finestra di selezione all'utente con le categorie ordinate
res = forms.SelectFromList.show(
	{'All Categories': sorted_names},  # Usa direttamente la lista ordinata
	title = 'Seleziona le categorie',
	multiselect = True
	)

# Converte i nomi selezionati negli oggetti categoria corrispondenti
# List comprehension con lookup O(1) sul dizionario (molto piu efficiente)
categoria_selezionata = [model_categories[sel] for sel in res]


"""
res = forms.SelectFromList.show(
        {'Mark Automatic Increment Categories': ['Doors','Curtain Panels','Mechanical Equipment','Plumbing Fixtures','Sprinklers','Windows','Electrical Equipment', 'Electrical Fixtures','Lighting Fixtures'],
		'All Categories':[x for x in model_cat]
		},
        title='Categories Selector',
        multiselect=True
    )
"""


categoriesId = []
bic = []
for c in categoria_selezionata:
	categoriesId.append(c.Id)
	bic.append(System.Enum.ToObject(BuiltInCategory, c.Id.IntegerValue))

def filcategories(document,Category): #Filtra tutti gli elementi in base a una lista di categorie
	if isinstance(Category, list):
		categoriesCollector = []
		for nId in categoriesId:
			categoriesCollector.append(FilteredElementCollector(document).OfCategoryId(nId).WhereElementIsNotElementType().ToElements())
		return categoriesCollector

outputelem = filcategories(doc,bic)

parametermark=[]

for o in outputelem:
	for x in o:
		parametermark.append(x.get_Parameter(BuiltInParameter.ALL_MODEL_MARK))


if not forms.alert('Stanno per essere cancellati i valori del campo "Contrassegno/Mark" per le categorie scelte, per proseguire premere "OK"'
                   , cancel=True):
    script.exit()

t = Transaction(doc,"Clean Mark")

t.Start()
try:
	for p in parametermark:
		p.Set('')

except:
	t.RollBack()
else:
	t.Commit()
