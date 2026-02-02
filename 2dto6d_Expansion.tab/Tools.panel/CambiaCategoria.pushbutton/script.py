""" Cambia la categoria degli oggetti selezionati e associa parametro materiale. """

__author__ = '2dto6d'
__title__ = 'Cambia Categoria'

from pyrevit import forms, revit, DB
from Autodesk.Revit.Exceptions import ArgumentException

# Revit document setup
doc = revit.doc
uidoc = revit.uidoc


# FamilyLoadOptions implementation
class FamilyLoadOptions(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        return True, True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        return True, False, True


def prendi_categoria():
    cats = doc.Settings.Categories
    ops = [c.Name for c in cats]
    scelta = forms.SelectFromList.show(ops, message='Scegli la categoria')

    if not scelta:
        return None

    for c in cats:
        if c.Name == scelta:
            return c.Id

def change_category(family, categoria):
    try:
        fam_doc = doc.EditFamily(family)
        transaction = DB.Transaction(fam_doc, "Change Category")
    except Exception as e:
        transaction.RollBack()
        return

    
    transaction.Start()

    owner_family = fam_doc.OwnerFamily
    owner_family.FamilyCategoryId = categoria

    transaction.Commit()

    load_options = FamilyLoadOptions()
    fam_doc.LoadFamily(doc, load_options)

    fam_doc.Close(False)

# Main execution
elementi_selezionati = uidoc.Selection.GetElementIds()
elements_in_view = [doc.GetElement(x) for x in elementi_selezionati]

if not elements_in_view:
    forms.alert("Nessun elemento selezionato. Seleziona almeno un elemento prima di eseguire lo script.", exitscript=True)

# Opzione per cambiare categoria (commentata)

cat_scelta = prendi_categoria()

try:
    for element in elements_in_view:
        change_category(element.Symbol.Family, cat_scelta)

except ArgumentException:
    
    alert = forms.alert("Non si possono assegnare categorie di sistema.", exitscript=True)