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
from pyrevit import *

# For floating point precision
from decimal import Decimal, ROUND_DOWN, getcontext

##############################################################
# Global Revit Document and UI Document objects
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
ActiveView = doc.ActiveView

##############################################################

# Definizione corretta della classe FamilyOption che implementa IFamilyLoadOptions
class FamilyOption(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues = True
        return True

def edita_famiglia(famiglia):
    print("Inizio elaborazione famiglia: {0}".format(famiglia.Name))
    
    try:
        FamDoc = doc.EditFamily(famiglia)
    except Exception as e:
        print("Errore apertura famiglia {0}: {1}".format(famiglia.Name, e))
        return

    t_famiglia = None
    
    try:
        t_famiglia = Transaction(FamDoc, "Associa parametri materiale")

        # Collect all non-type elements within the family document
        collector = FilteredElementCollector(FamDoc).WhereElementIsNotElementType().ToElements()

        # Filter for FreeForm elements
        free_form_elements = [e for e in collector if isinstance(e, Autodesk.Revit.DB.FreeFormElement)]

        if not free_form_elements:
            FamDoc.Close(False)
            print("Nessun elemento FreeForm trovato nella famiglia: {0}. Chiudo senza salvare.".format(famiglia.Name))
            return

        # Find the "MATERIALE" family parameter
        param_materiale = None
        for fam_param in FamDoc.FamilyManager.Parameters:
            if "MATERIALE" in fam_param.Definition.Name.upper():
                param_materiale = fam_param
                break

        if param_materiale is None:
            FamDoc.Close(False)
            print("Parametro MATERIALE non trovato nella famiglia: {0}. Chiudo senza salvare.".format(famiglia.Name))
            return

        t_famiglia.Start()

        # Associa i parametri materiale
        for element in free_form_elements:
            for p in element.Parameters:
                if "Material" in p.Definition.Name:
                    # Associa il parametro del materiale elemento FreeForm al parametro di famiglia MATERIALE
                    FamDoc.FamilyManager.AssociateElementParameterToFamilyParameter(p, param_materiale)
                    print("   Associato {0} (ID: {1}) su elemento FreeForm in {2}".format(p.Definition.Name, element.Id, famiglia.Name))

        FamDoc.Regenerate()
        t_famiglia.Commit()
        print("Modifiche salvate con successo per la famiglia: {0}".format(famiglia.Name))

        # STRATEGIA NUOVA: Salva in file temporaneo e ricarica
        import tempfile
        temp_dir = tempfile.gettempdir()
        # Sostituisci caratteri problematici nel nome file
        safe_name = re.sub(r'[<>:"/\\|?*]', '_', famiglia.Name)
        temp_family_path = os.path.join(temp_dir, "{0}_temp.rfa".format(safe_name))
        
        print("Salvataggio famiglia temporaneo in: {0}".format(temp_family_path))
        
        # Salva la famiglia modificata come file temporaneo
        save_as_options = SaveAsOptions()
        save_as_options.OverwriteExistingFile = True
        FamDoc.SaveAs(temp_family_path, save_as_options)
        
        # Chiudi il documento famiglia
        FamDoc.Close(False)
        print("Famiglia salvata e documento chiuso: {0}".format(famiglia.Name))

        # Ricarica la famiglia nel documento principale
        t_main = Transaction(doc, "Ricarica famiglia modificata")
        t_main.Start()
        
        try:
            family_option = FamilyOption()
            
            # Prova prima con il metodo standard (due parametri)
            try:
                load_result = doc.LoadFamily(temp_family_path, family_option)
                if load_result:
                    print("Famiglia {0} ricaricata con successo nel progetto.".format(famiglia.Name))
                else:
                    print("Avviso: Famiglia {0} non e stata ricaricata correttamente.".format(famiglia.Name))
            except Exception as load_error:
                print("Errore durante il ricaricamento con metodo a 2 parametri: {0}".format(load_error))
                
                # Fallback: prova con metodo a 3 parametri
                try:
                    from System import Array
                    from clr import Reference
                    
                    # Crea una reference per il parametro di output
                    family_out = Reference[Family]()
                    load_result = doc.LoadFamily(temp_family_path, family_option, family_out)
                    
                    if load_result:
                        print("Famiglia {0} ricaricata con successo nel progetto (metodo alternativo).".format(famiglia.Name))
                    else:
                        print("Avviso: Famiglia {0} non e stata ricaricata correttamente (metodo alternativo).".format(famiglia.Name))
                        
                except Exception as alt_error:
                    print("Errore anche con metodo alternativo: {0}".format(alt_error))
            
            t_main.Commit()
            
        except Exception as load_error:
            print("Errore durante il ricaricamento della famiglia {0}: {1}".format(famiglia.Name, load_error))
            if t_main.HasStarted():
                t_main.RollBack()
        
        # Pulisci il file temporaneo
        try:
            if os.path.exists(temp_family_path):
                os.remove(temp_family_path)
                print("File temporaneo rimosso: {0}".format(temp_family_path))
        except Exception as cleanup_error:
            print("Avviso: Impossibile rimuovere file temporaneo {0}: {1}".format(temp_family_path, cleanup_error))

    except Exception as e:
        print("Errore critico durante elaborazione della famiglia {0}: {1}".format(famiglia.Name, e))
        if t_famiglia and t_famiglia.HasStarted():
            t_famiglia.RollBack()
        try:
            FamDoc.Close(False)
        except:
            pass

# Elabora tutte le famiglie nella vista attiva
print("Inizio elaborazione famiglie nella vista attiva...")

collector_elementi_in_doc = FilteredElementCollector(doc, ActiveView.Id).WhereElementIsNotElementType().ToElements()

famiglie_elaborate = set()  # Per evitare di elaborare la stessa famiglia piu volte

for elemento in collector_elementi_in_doc:
    if isinstance(elemento, Autodesk.Revit.DB.FamilyInstance):
        famiglia = elemento.Symbol.Family
        if famiglia.Name not in famiglie_elaborate:
            famiglie_elaborate.add(famiglia.Name)
            edita_famiglia(famiglia)

print("Elaborazione completata. Famiglie elaborate: {0}".format(len(famiglie_elaborate)))