# -*- coding: utf-8 -*-
__title__ = "Total\nArea"
__author__ = "JK_Sim"
__doc__ = """Version = 6.0
Date    = 11.12.2025
Description:
Calculates total area of selected rooms.
- Seamlessly handles Native Rooms and Linked Rooms.
- No pop-ups or confirmations.
- Uses Reference selection to identify specific linked rooms.
"""

from Autodesk.Revit.DB import (
    BuiltInCategory, ElementId, RevitLinkInstance, SpatialElement
)
from pyrevit import revit, forms, script

# ---------------- INITIALIZATION ----------------
doc = revit.doc
uidoc = revit.uidoc

# ---------------- HELPERS ----------------
def get_internal_area(element):
    """Safely returns area if element is a placed room."""
    if hasattr(element, "Area") and element.Area > 0:
        return element.Area
    return 0.0

def to_sqm(sqft):
    return sqft * 0.09290304

# ---------------- MAIN LOGIC ----------------

# We use GetReferences() instead of GetElementIds()
# This is crucial because it contains the 'LinkedElementId'
sel_refs = uidoc.Selection.GetReferences()

if not sel_refs:
    forms.alert("Please select Room(s) first.", exitscript=True)

total_area_sqft = 0.0
room_count = 0

# Iterate through every picked object reference
for ref in sel_refs:
    element = None

    # --- CASE A: Linked Element (Tab-Selected) ---
    if ref.LinkedElementId != ElementId.InvalidElementId:
        link_inst = doc.GetElement(ref.ElementId)
        if isinstance(link_inst, RevitLinkInstance):
            link_doc = link_inst.GetLinkDocument()
            if link_doc:
                element = link_doc.GetElement(ref.LinkedElementId)

    # --- CASE B: Native Element (Direct Select) ---
    else:
        element = doc.GetElement(ref.ElementId)

    # --- PROCESS THE ELEMENT ---
    # We treat both exactly the same now
    if element and isinstance(element, SpatialElement):
        if element.Category and element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms):
            area = get_internal_area(element)
            
            # Add to total if valid
            if area > 0:
                total_area_sqft += area
                room_count += 1

# ---------------- OUTPUT ----------------
if room_count == 0:
    forms.alert("No placed rooms found in selection.", warn_icon=True)
    script.exit()

val_sqm = to_sqm(total_area_sqft)
val_sqft = total_area_sqft

msg = "Total Area Calculation\n" \
      "---------------------------\n" \
      "Rooms Selected: {}\n" \
      "---------------------------\n" \
      "{:.2f} m²\n" \
      "{:.2f} sqft".format(room_count, val_sqm, val_sqft)

forms.alert(msg, title="Total Area")