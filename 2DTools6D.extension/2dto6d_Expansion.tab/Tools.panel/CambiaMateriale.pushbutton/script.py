""" Cambia la categoria degli oggetti selezionati e associa parametro materiale. """

__author__ = 'Roberto Dolfini'
__title__ = 'Assegna Materiale'

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType

import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

from System.Collections.Generic import List
from Autodesk.Revit.DB import IFamilyLoadOptions
from System import Guid

import pyrevit
from pyrevit import *
from pyrevit import forms, script

import os
import csv
from datetime import datetime

# Revit document setup
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
active_view = doc.ActiveView
output = pyrevit.output.get_output()

# GUID del parametro condiviso "Materiale"
MATERIAL_PARAM_GUID = Guid("238f411c-d82f-8006-832d-dab27cb3f782")

# Lista per raccogliere i dati del report
report_data = []

t = Transaction(doc, "Cambia Categoria")

# FamilyLoadOptions implementation
class FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        return True, True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        return True, False, True

# Function to get or create shared parameter
def get_or_create_shared_parameter(fam_doc, param_name, param_guid, family_name):
    """
    Ottiene o crea un parametro condiviso nella famiglia.
    Restituisce il FamilyParameter se trovato o creato, None altrimenti.
    """
    
    # Prima controlla se il parametro esiste gia' (per GUID)
    for param in fam_doc.FamilyManager.Parameters:
        if param.IsShared:
            # Ottieni il GUID del parametro condiviso
            if param.GUID == param_guid:
                status = "Parametro esistente"
                tipo = "ISTANZA" if param.IsInstance else "TIPO"
                report_data.append({
                    'Famiglia': family_name,
                    'Azione': 'Parametro trovato',
                    'Nome Parametro': param.Definition.Name,
                    'GUID': str(param_guid),
                    'Tipo': tipo,
                    'Stato': status
                })
                return param
    
    # Se non trovato, controlla anche per nome
    for param in fam_doc.FamilyManager.Parameters:
        if param.Definition.Name == param_name:
            if param.IsShared and param.GUID == param_guid:
                tipo = "ISTANZA" if param.IsInstance else "TIPO"
                report_data.append({
                    'Famiglia': family_name,
                    'Azione': 'Parametro trovato per nome',
                    'Nome Parametro': param_name,
                    'GUID': str(param_guid),
                    'Tipo': tipo,
                    'Stato': 'Trovato per nome e GUID'
                })
                return param
    
    # Se non esiste, proviamo a crearlo
    # Ottieni il file dei parametri condivisi
    shared_params_file = app.OpenSharedParameterFile()
    
    if shared_params_file is None:
        report_data.append({
            'Famiglia': family_name,
            'Azione': 'ERRORE',
            'Nome Parametro': param_name,
            'GUID': str(param_guid),
            'Tipo': 'N/A',
            'Stato': 'Nessun file parametri condivisi configurato'
        })
        return None
    
    # Cerca la definizione del parametro nel file condiviso
    shared_param_element = None
    
    # Cerca in tutti i gruppi del file condiviso
    for group in shared_params_file.Groups:
        for definition in group.Definitions:
            external_def = definition
            if external_def.GUID == param_guid:
                shared_param_element = external_def
                break
        if shared_param_element:
            break
    
    # Se non trovato nel file condiviso, proviamo a crearlo
    if shared_param_element is None:
        # Trova o crea un gruppo per il parametro
        group_name = "Materiali Custom"
        param_group = None
        
        for group in shared_params_file.Groups:
            if group.Name == group_name:
                param_group = group
                break
        
        if param_group is None:
            try:
                param_group = shared_params_file.Groups.Create(group_name)
            except Exception as e:
                report_data.append({
                    'Famiglia': family_name,
                    'Azione': 'ERRORE',
                    'Nome Parametro': param_name,
                    'GUID': str(param_guid),
                    'Tipo': 'N/A',
                    'Stato': 'Errore creazione gruppo: {}'.format(str(e))
                })
                return None
        
        # Crea la definizione del parametro nel file condiviso
        try:
            # Determina il tipo corretto per il parametro Material
            try:
                from Autodesk.Revit.DB import ForgeTypeId, SpecTypeId
                param_type = SpecTypeId.Reference.Material
            except:
                param_type = ParameterType.Material
            
            external_def_create_options = ExternalDefinitionCreationOptions(param_name, param_type)
            external_def_create_options.GUID = param_guid
            
            shared_param_element = param_group.Definitions.Create(external_def_create_options)
            
        except Exception as e:
            # Prova a recuperare se esiste gia' con nome diverso
            for definition in param_group.Definitions:
                if definition.GUID == param_guid:
                    shared_param_element = definition
                    break
    
    # Ora aggiungi il parametro alla famiglia
    if shared_param_element:
        try:
            # Aggiungi come parametro di istanza
            new_param = fam_doc.FamilyManager.AddParameter(
                shared_param_element,
                BuiltInParameterGroup.PG_MATERIALS,
                True  # True = parametro di istanza
            )
            report_data.append({
                'Famiglia': family_name,
                'Azione': 'Parametro creato',
                'Nome Parametro': param_name,
                'GUID': str(param_guid),
                'Tipo': 'ISTANZA',
                'Stato': 'Creato con successo'
            })
            return new_param
            
        except Exception as e:
            # Potrebbe gia' esistere, riprova a cercarlo
            for param in fam_doc.FamilyManager.Parameters:
                if param.IsShared and param.GUID == param_guid:
                    tipo = "ISTANZA" if param.IsInstance else "TIPO"
                    report_data.append({
                        'Famiglia': family_name,
                        'Azione': 'Parametro trovato dopo creazione',
                        'Nome Parametro': param.Definition.Name,
                        'GUID': str(param_guid),
                        'Tipo': tipo,
                        'Stato': 'Esistente'
                    })
                    return param
    
    return None

# Function to edit family and associate material parameter
def edit_family(family):
    family_name = family.Name
    elementi_info = []
    
    try:
        fam_doc = doc.EditFamily(family)
    except Exception as e:
        report_data.append({
            'Famiglia': family_name,
            'Azione': 'ERRORE apertura',
            'Nome Parametro': '',
            'GUID': '',
            'Tipo': '',
            'Stato': str(e)
        })
        return

    try:
        transaction = Transaction(fam_doc, "Associate Material Parameter")
        
        # Raccoglie tutti gli elementi nella famiglia
        collector = FilteredElementCollector(fam_doc).WhereElementIsNotElementType().ToElements()
        
        # Lista dei tipi di elementi che possono avere un parametro Material
        geometry_types = [
            FreeFormElement,
            Extrusion,
            Blend,
            Revolution,
            Sweep,
            SweptBlend,
            GenericForm
        ]
        
        # Trova tutti gli elementi geometrici che potrebbero avere un parametro Material
        geometric_elements = []
        for element in collector:
            # Verifica se l'elemento e' uno dei tipi geometrici
            is_geometry_type = False
            for geom_type in geometry_types:
                if isinstance(element, geom_type):
                    geometric_elements.append(element)
                    is_geometry_type = True
                    break
            
            # Se non e' un tipo geometrico noto, controlla se ha parametri Material
            if not is_geometry_type:
                if hasattr(element, 'Category') and element.Category is not None:
                    # Verifica se l'elemento ha parametri
                    if element.Parameters.Size > 0:
                        # Controlla se ha un parametro Material
                        for p in element.Parameters:
                            if p.Definition and "Material" in p.Definition.Name:
                                geometric_elements.append(element)
                                break

        if not geometric_elements:
            report_data.append({
                'Famiglia': family_name,
                'Azione': 'Nessun elemento geometrico',
                'Nome Parametro': '',
                'GUID': '',
                'Tipo': '',
                'Stato': 'Nessun elemento con parametro Material trovato'
            })
            fam_doc.Close(False)
            return

        # Ottieni o crea il parametro condiviso "Materiale"
        transaction.Start()
        
        material_param = get_or_create_shared_parameter(fam_doc, "Materiale", MATERIAL_PARAM_GUID, family_name)
        
        if material_param is None:
            transaction.RollBack()
            
            # Come fallback, prova a usare un parametro esistente
            for param in fam_doc.FamilyManager.Parameters:
                if "MATERIAL" in param.Definition.Name.upper() and param.IsInstance:
                    material_param = param
                    report_data.append({
                        'Famiglia': family_name,
                        'Azione': 'Fallback parametro',
                        'Nome Parametro': param.Definition.Name,
                        'GUID': '',
                        'Tipo': 'ISTANZA' if param.IsInstance else 'TIPO',
                        'Stato': 'Uso parametro esistente'
                    })
                    break
            
            if not material_param:
                fam_doc.Close(False)
                return
        
        transaction.Commit()

        # Ora associa il parametro materiale degli elementi al parametro di famiglia
        transaction2 = Transaction(fam_doc, "Associate Element Material to Family Parameter")
        transaction2.Start()

        elementi_associati = 0
        elementi_gia_associati = 0
        elementi_errore = 0
        
        for element in geometric_elements:
            element_id = element.Id.IntegerValue
            element_type = type(element).__name__
            
            for p in element.Parameters:
                if p.Definition is None:
                    continue
                    
                # Cerca il parametro Material dell'elemento
                if "Material" in p.Definition.Name:
                    # Verifica che sia del tipo corretto (ElementId per materiali)
                    if p.StorageType == StorageType.ElementId:
                        try:
                            # Prova ad associare il parametro
                            fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(p, material_param)
                            elementi_associati += 1
                            break
                        except Exception as e:
                            error_msg = str(e)
                            if "already associated" in error_msg.lower():
                                elementi_gia_associati += 1
                            else:
                                elementi_errore += 1
                            break

        transaction2.Commit()

        # Aggiungi il riepilogo al report
        report_data.append({
            'Famiglia': family_name,
            'Azione': 'Associazione completata',
            'Nome Parametro': material_param.Definition.Name if material_param else '',
            'GUID': str(MATERIAL_PARAM_GUID),
            'Tipo': 'ISTANZA' if material_param and material_param.IsInstance else 'TIPO',
            'Stato': 'Associati: {}, Gia associati: {}, Errori: {}'.format(
                elementi_associati, elementi_gia_associati, elementi_errore
            )
        })

        # Ricarica la famiglia nel progetto
        load_options = FamilyLoadOptions()
        try:
            result = fam_doc.LoadFamily(doc, load_options)
            if not result:
                report_data.append({
                    'Famiglia': family_name,
                    'Azione': 'ATTENZIONE',
                    'Nome Parametro': '',
                    'GUID': '',
                    'Tipo': '',
                    'Stato': 'Famiglia non ricaricata correttamente'
                })
        except Exception as e:
            report_data.append({
                'Famiglia': family_name,
                'Azione': 'ERRORE ricaricamento',
                'Nome Parametro': '',
                'GUID': '',
                'Tipo': '',
                'Stato': str(e)
            })

    except Exception as e:
        report_data.append({
            'Famiglia': family_name,
            'Azione': 'ERRORE inaspettato',
            'Nome Parametro': '',
            'GUID': '',
            'Tipo': '',
            'Stato': str(e)
        })
    finally:
        try:
            if fam_doc and fam_doc.IsModifiable:
                fam_doc.Close(False)
        except:
            pass

def save_report_to_csv():
    """Salva il report in un file CSV nella cartella del documento Revit"""
    
    # Ottieni il percorso del documento Revit
    doc_path = doc.PathName
    
    if not doc_path:
        # Se il documento non e' salvato, usa il desktop
        doc_dir = os.path.expanduser("~/Desktop")
        doc_name = "UnsavedDocument"
    else:
        doc_dir = os.path.dirname(doc_path)
        doc_name = os.path.splitext(os.path.basename(doc_path))[0]
    
    # Crea il nome del file con timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = "MaterialParameter_Report_{}_{}.csv".format(doc_name, timestamp)
    csv_path = os.path.join(doc_dir, csv_filename)
    
    # Scrivi il CSV
    try:
        with open(csv_path, 'wb') as csvfile:
            fieldnames = ['Famiglia', 'Azione', 'Nome Parametro', 'GUID', 'Tipo', 'Stato']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            # Scrivi l'intestazione
            writer.writeheader()
            
            # Scrivi i dati
            for row in report_data:
                writer.writerow(row)
        
        return csv_path
    except Exception as e:
        print("Errore nel salvataggio del CSV: {}".format(e))
        return None

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
    except Exception as e:
        return

    transaction = Transaction(fam_doc, "Change Category")
    transaction.Start()

    owner_family = fam_doc.OwnerFamily
    owner_family.FamilyCategoryId = categoria

    transaction.Commit()

    load_options = FamilyLoadOptions()
    result = fam_doc.LoadFamily(doc, load_options)

    fam_doc.Close(False)

# Main execution
elementi_selezionati = uidoc.Selection.GetElementIds()
elements_in_view = [doc.GetElement(x) for x in elementi_selezionati]

if not elements_in_view:
    forms.alert("Nessun elemento selezionato. Seleziona almeno un elemento prima di eseguire lo script.", exitscript=True)

# Opzione per cambiare categoria (commentata)
"""
cat_scelta = prendi_categoria()

for element in elements_in_view:
    change_category(element.Symbol.Family, cat_scelta)
"""
"""
# Esegui l'associazione del parametro materiale
print("Elaborazione in corso...")
"""
# Verifica preliminare del file parametri condivisi
if app.OpenSharedParameterFile() is None:
    report_data.append({
        'Famiglia': 'SISTEMA',
        'Azione': 'ATTENZIONE',
        'Nome Parametro': '',
        'GUID': '',
        'Tipo': '',
        'Stato': 'Nessun file parametri condivisi configurato'
    })

famiglie_processate = set()

for element in elements_in_view:
    try:
        famiglia = element.Symbol.Family
        if famiglia.Id not in famiglie_processate:
            edit_family(famiglia)
            famiglie_processate.add(famiglia.Id)
    except Exception as e:
        report_data.append({
            'Famiglia': 'ERRORE',
            'Azione': 'Errore elaborazione elemento',
            'Nome Parametro': '',
            'GUID': '',
            'Tipo': '',
            'Stato': str(e)
        })
"""
# Salva il report CSV
csv_path = save_report_to_csv()

# Mostra il riepilogo finale
print("\n" + "=" * 50)
print("PROCESSO COMPLETATO")
print("Famiglie processate: {}".format(len(famiglie_processate)))
if csv_path:
    print("Report salvato in: {}".format(csv_path))
else:
    print("Impossibile salvare il report CSV")
print("=" * 50)
"""