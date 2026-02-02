# -*- coding: utf-8 -*-

"""Verifica quali elementi NON hanno IFC Export As valorizzato"""

__title__ = 'Check Valorizzazione\nIFC Export As'
__author__ = '2dto6d'

from pyrevit import revit, script
from Autodesk.Revit.DB import *

doc = revit.doc
uidoc = revit.uidoc
aview = doc.ActiveView
output = script.get_output()

# Filtro elementi con parametro vuoto
filtro_vuoto = ElementParameterFilter(
    FilterStringRule(
        ParameterValueProvider(ElementId(BuiltInParameter.IFC_EXPORT_ELEMENT_AS)),
        FilterStringEquals(),
        ""
    )
)


# Collector elementi NON valorizzati
elementi_non_valorizzati = []
for el in FilteredElementCollector(doc, aview.Id)\
        .WherePasses(filtro_vuoto)\
        .WhereElementIsNotElementType():
    param = el.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS)
    if param is not None:
        elementi_non_valorizzati.append(el)

""" IMPLEMENTARE IN FUTURO 
# Collector elementi NON valorizzati, senza ifcName
elementi_non_valorizzati_name = []
for el in FilteredElementCollector(doc, aview.Id):
    try:
            param = el.LookupParameter("IfcName")
            if param is None:
                elementi_non_valorizzati_name.append([el,el.Id])
    except:
        pass
"""
# Output
output.print_md("# Verifica IFC Export As")
output.print_md("---")

# Tabella elementi non valorizzati
if elementi_non_valorizzati:
    output.print_md("## ✗ Elementi SENZA IFC Export As: {}".format(len(elementi_non_valorizzati)))
    table_no = []
    for el in elementi_non_valorizzati:
        cat = el.Category.Name if el.Category else "N/A"
        nome = [el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString(),el.Id]
        table_no.append([nome,cat, output.linkify(el.Id)])
    output.print_table(table_no, columns=["Tipo,""Categoria", "ID"])
else:
    output.print_md("## 🟢 Tutti gli elementi hanno IFC Export As valorizzato 🟢 ")