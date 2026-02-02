# -*- coding: utf-8 -*-
""" Disegna le linee della scope box selezionata.  """

__author__ = '2dto6d'
__title__ = 'Ricalca\nScopeBox'

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from pyrevit import script

##############################################################
doc   = __revit__.ActiveUIDocument.Document  #type: Document
uidoc = __revit__.ActiveUIDocument

t = Transaction(doc, "Ricalca ScopeBox")
##############################################################

Selezione = uidoc.Selection.GetElementIds()

if Selezione:

    ScopeBox = doc.GetElement(Selezione[0])

    #Estrazione Geometria ScopeBox
    GenericOption = Options()
    GeometriaScopeBox = ScopeBox.get_Geometry(GenericOption)
    Enum = GeometriaScopeBox.GetEnumerator()
    LineeScopeBox = []

    while Enum.MoveNext():
        if isinstance(Enum.Current, Line):
            if Enum.Current.GetEndPoint(0) > 1:
                pass
            else:
                LineeScopeBox.append(Enum.Current)
                
    NuoveLinee = []
    ZDaUsare = LineeScopeBox[0].GetEndPoint(0)
    linee_create = set()

    for Linea in LineeScopeBox:
        StartLinea = Linea.GetEndPoint(0)
        EndLinea = Linea.GetEndPoint(1)

        # Controlla se la linea è sul piano corretto
        if StartLinea.Z == ZDaUsare.Z and EndLinea.Z == ZDaUsare.Z:
            # Crea una chiave unica per evitare duplicati (ordina i punti)
            punti = sorted([(round(StartLinea.X, 6), round(StartLinea.Y, 6)),
                          (round(EndLinea.X, 6), round(EndLinea.Y, 6))])
            chiave = (punti[0], punti[1])

            if chiave not in linee_create:
                linee_create.add(chiave)
                nuovo_start = XYZ(StartLinea.X, StartLinea.Y, ZDaUsare.Z)
                nuovo_end = XYZ(EndLinea.X, EndLinea.Y, ZDaUsare.Z)

                try:
                    NuoveLinee.append(Line.CreateBound(nuovo_start, nuovo_end))
                except:
                    pass


    #Inizio la transazione
    t.Start()

    #Creazione di una vista di disegno per le linee
    active_view = doc.ActiveView
    plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, LineeScopeBox[0].GetEndPoint(0))
    sketch_plane = SketchPlane.Create(doc, plane)


    for line in NuoveLinee:
        DetailLine = doc.Create.NewDetailCurve(active_view, line)


    #Fine transazione
    t.Commit()


else:
    script.exit()

