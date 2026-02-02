# -*- coding: utf-8 -*-
"""Misura la distanza tra due punti, in clipboard viene salvato il valore"""
__title__ = 'Misura Distanza'
__author__ = '2dto6d'

import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import Clipboard
import pyrevit
from pyrevit import forms
from Autodesk.Revit.UI.Selection import ObjectSnapTypes

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

snap_types = (
    ObjectSnapTypes.Endpoints |
    ObjectSnapTypes.Intersections |
    ObjectSnapTypes.Midpoints
)

p1 = uidoc.Selection.PickPoint(snap_types, "Seleziona il primo vertice")
p2 = uidoc.Selection.PickPoint(snap_types, "Seleziona il secondo vertice")

distanza = p1.DistanceTo(p2) * 0.3048

Clipboard.SetText("{:.4f}".format(distanza))

forms.toast(
    "{} m".format(distanza),
    title="Fast Measure",
)