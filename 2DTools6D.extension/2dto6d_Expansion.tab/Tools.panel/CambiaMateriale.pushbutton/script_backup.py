import Autodesk.Revit
import Autodesk.Revit.DB
import Autodesk.Revit.DB.Architecture
import System
from System import Enum
from System.Collections.Generic import Dictionary
import unicodedata
import math
import clr
import os
import sys
import csv
import codecs
import re
from re import *

# Add RevitAPI and RevitAPIUI references
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

import pyrevit
from pyrevit import * # Assuming pyrevit is properly installed and accessible

# For floating point precision (though not directly used in the fix for this error)
from decimal import Decimal, ROUND_DOWN, getcontext

##############################################################
# Global Revit Document and UI Document objects
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
ActiveView = doc.ActiveView

##############################################################

# Definizione corretta della classe FamilyOption
class FamilyOption(object):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues[0] = True
        return True

    def OnSharedFamilyFound(
            self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues[0] = True
        return True

def edita_famiglia(famiglia):

    FamDoc = doc.EditFamily(famiglia)
    t_famiglia = None
    t_famiglia = Transaction(FamDoc, "Associa parametri materiale")

    # Collect all non-type elements within the family document
    collector = FilteredElementCollector(FamDoc).WhereElementIsNotElementType().ToElements()

    # Filter for FreeForm elements
    free_form_elements = [e for e in collector if isinstance(e, Autodesk.Revit.DB.FreeFormElement)]

    if not free_form_elements:
        FamDoc.Close(False)
        print("Nessun elemento FreeForm trovato nella famiglia: {0}. Chiudo senza salvare.".format(famiglia.Name))
        return

    # Find the "MATERIALE" family parameter.
    param_materiale = None
    for fam_param in FamDoc.FamilyManager.Parameters:
        if "MATERIALE" in fam_param.Definition.Name.upper():
            param_materiale = fam_param
            break

    if param_materiale is None:
        FamDoc.Close(False)
        print("Parametro 'MATERIALE' non trovato nella famiglia: {0}. Chiudo senza salvare.".format(famiglia.Name))
        return

    t_famiglia.Start()

    for element in free_form_elements:
        for p in element.Parameters:
            if "Material" in p.Definition.Name:
                # Associa il parametro del materiale dell'elemento FreeForm al parametro di famiglia "MATERIALE"
                FamDoc.FamilyManager.AssociateElementParameterToFamilyParameter(p, param_materiale)
                print("   Associato '{0}' (ID: {1}) su elemento FreeForm in '{2}'".format(p.Definition.Name, element.Id, famiglia.Name))


    t_famiglia.Commit() # Commit delle modifiche alla famiglia
    print("Modifiche salvate con successo per la famiglia: {0}".format(famiglia.Name))

    try:
        # USA LoadFamilyAndClose per caricare le modifiche nel documento host e chiudere la famiglia
        load_result = FamDoc.LoadFamilyAndClose(doc, FamilyOption())
        if load_result:
            print("Famiglia '{0}' ricaricata e chiusa con successo nel progetto.".format(famiglia.Name))
        else:
            print("Avviso: Famiglia '{0}' non e stata ricaricata correttamente (LoadFamilyAndClose ha restituito False).".format(famiglia.Name))

    except Exception as e:
        print("Errore critico durante il ricaricamento della famiglia '{0}' nel progetto: {1}".format(famiglia.Name, e))
        # Non usare 'pass' qui, e importante sapere se il ricaricamento fallisce!
        # Potresti voler riavvolgere la transazione o notificare l'utente.

    # Rimuovi questa riga, perche LoadFamilyAndClose gestisce la chiusura
    # FamDoc.Close(False)

collector_elementi_in_doc = FilteredElementCollector(doc, ActiveView.Id).WhereElementIsNotElementType().ToElements()

for elemento in collector_elementi_in_doc:
    if isinstance(elemento, Autodesk.Revit.DB.FamilyInstance):
        edita_famiglia(elemento.Symbol.Family)