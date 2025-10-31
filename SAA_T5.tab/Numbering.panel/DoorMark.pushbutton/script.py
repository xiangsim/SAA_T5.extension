# -*- coding: utf-8 -*-
__title__ = "Door\nMark"
__author__ = "JK_Sim"
__doc__ = """Version = 1.4
Date    = 31.10.2025
_____________________________________________________________________
Description:

Fill in Door parameter "Mark" and "ELEMENT_ROOM ALLOCATION"
from Room "Number" and "Name", including linked rooms.

_____________________________________________________________________
How-to:

-> Select Doors & ONLY 1 Room (host or from link)
-> Run this script
-> Done
_____________________________________________________________________
"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, forms, script
from System import Guid

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ"

# ------------------------------------- INIT
sel_ids = uidoc.Selection.GetElementIds()
sel_refs = uidoc.Selection.GetReferences()

selected = [doc.GetElement(id) for id in sel_ids]

doors = []
room = None

# --- First pass: check host model
for el in selected:
    if el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_Doors):
        doors.append(el)

    elif el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms):
        room = el

# --- Second pass: check if a linked room is selected
if not room:
    for ref in sel_refs:
        if ref.LinkedElementId != ElementId.InvalidElementId:
            link_inst = doc.GetElement(ref.ElementId)
            if isinstance(link_inst, RevitLinkInstance):
                link_doc = link_inst.GetLinkDocument()
                linked_elem = link_doc.GetElement(ref.LinkedElementId)

                if linked_elem.Category and linked_elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms):
                    room = linked_elem
                    break

# ------------------------------------- VALIDATION
if not doors:
    forms.alert("No doors selected.")
    script.exit()

if room is None:
    forms.alert("No room found. Please select ONE room (host or linked).")
    script.exit()

p_num = room.LookupParameter("Number")
p_name = room.LookupParameter("Name")

if not p_num or not p_name:
    forms.alert("Room missing 'Number' or 'Name' parameter.")
    script.exit()

room_num = p_num.AsString() or ""
room_name = p_name.AsString() or ""

if not room_num or not room_name:
    forms.alert("Room Number or Name is empty.")
    script.exit()

# ------------------------------------- SORT DOORS
def get_center(elem):
    loc = elem.Location
    if isinstance(loc, LocationPoint):
        pt = loc.Point
        return pt.X, pt.Y
    return None

door_centers = [(d, get_center(d)) for d in doors]
door_centers = [dc for dc in door_centers if dc[1] is not None]

if len(door_centers) != len(doors):
    forms.alert("Some doors have no location point.")
    script.exit()

# Sort by Y (top to bottom), then X (left to right)
door_centers.sort(key=lambda x: (-x[1][1], x[1][0]))
sorted_doors = [dc[0] for dc in door_centers]

# ------------------------------------- ASSIGN VALUES
with Transaction(doc, "Door Mark from Room") as t:
    t.Start()

    for i, door in enumerate(sorted_doors):
        if i >= len(ALPHABET):
            output.print_md("Too many doors (max {}). Stopped at {}.".format(len(ALPHABET), i))
            break

        mark = room_num + ALPHABET[i]

        p_mark = door.LookupParameter("Mark")
        if p_mark and not p_mark.IsReadOnly:
            p_mark.Set(mark)

        p_alloc = door.LookupParameter("ELEMENT_ROOM ALLOCATION")
        if p_alloc and not p_alloc.IsReadOnly:
            p_alloc.Set(room_name)

    t.Commit()
