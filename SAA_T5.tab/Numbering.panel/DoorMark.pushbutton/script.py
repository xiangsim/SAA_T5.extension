# -*- coding: utf-8 -*-
__title__ = "Door\nMark"
__author__ = "JK_Sim"
__doc__ = """Version = 1.0
Date    = 28.10.2025
_____________________________________________________________________
Description:

Fill in Door parameter "Mark" and "ELEMENT_ROOM ALLOCATION" from Room "Number" and "Name"
_____________________________________________________________________
How-to:

-> Select doors & ONLY 1 room
-> Click the button
-> Done
_____________________________________________________________________
"""
from Autodesk.Revit.DB import (BuiltInCategory, LocationPoint, Transaction)
from pyrevit import revit, forms, script
from Snippets._selection import get_selected_elements

doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()
ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ"

# ------------------------------------------------- GET SELECTION
selected = get_selected_elements(uidoc)

# Filter doors and rooms
doors = [e for e in selected if e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_Doors)]
rooms = [e for e in selected if e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms)]

# ------------------------------------------------- VALIDATION
if not doors:
    forms.alert("No doors selected.")
    script.exit()

if len(rooms) != 1:
    forms.alert("Please select EXACTLY ONE room.\nSelected: {} rooms.".format(len(rooms)))
    script.exit()

# ------------------------------------------------- START TRANSACTION EARLY
with Transaction(doc, "Door Mark from Room") as t:
    t.Start()

    # --- NOW SAFE TO READ PARAMETERS ---
    room = rooms[0]
    p_num = room.LookupParameter("Number")
    p_name = room.LookupParameter("Name")

    if not p_num or not p_name:
        t.RollBack()
        forms.alert("Room missing 'Number' or 'Name' parameter.")
        script.exit()

    room_num = p_num.AsString() or ""
    room_name = p_name.AsString() or ""

    if not room_num or not room_name:
        t.RollBack()
        forms.alert("Room Number or Name is empty.")
        script.exit()

    # --- SORT DOORS ---
    def get_center(elem):
        loc = elem.Location
        if isinstance(loc, LocationPoint):
            pt = loc.Point
            return pt.X, pt.Y
        return None

    door_centers = [(d, get_center(d)) for d in doors]
    door_centers = [dc for dc in door_centers if dc[1] is not None]

    if len(door_centers) != len(doors):
        t.RollBack()
        forms.alert("Some doors have no location point.")
        script.exit()

    door_centers.sort(key=lambda x: (-x[1][1], x[1][0]))
    sorted_doors = [dc[0] for dc in door_centers]

    # --- ASSIGN MARKS ---
    for i, door in enumerate(sorted_doors):
        if i >= len(ALPHABET):
            output.print_md("Warning: Too many doors (max {}). Stopped at {}.".format(len(ALPHABET), i))
            break

        mark = room_num + ALPHABET[i]

        p_mark = door.LookupParameter("Mark")
        if p_mark and not p_mark.IsReadOnly:
            p_mark.Set(mark)

        p_alloc = door.LookupParameter("ELEMENT_ROOM ALLOCATION")
        if p_alloc and not p_alloc.IsReadOnly:
            p_alloc.Set(room_name)

    t.Commit()


