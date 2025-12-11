# -*- coding: utf-8 -*-
__title__ = "Total\nArea"
__author__ = "JK_Sim"
__doc__ = """Version = 1.0
Date    = 11.12.2025
Description:
Calculates total area of selected rooms.
Supports:
1. Native Rooms (Pre-selected).
2. Linked Rooms (Calculates all rooms in selected Link Instance).
"""

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, 
    SpatialElement, RevitLinkInstance, ElementId
)
import sys

# ---------------- Helpers ----------------
def to_sqm(sqft_value):
    """Converts internal sqft to sqm."""
    return sqft_value * 0.09290304

def get_room_area(element):
    """Safely gets area from a Room element."""
    # Check if element is a Room (SpatialElement)
    if hasattr(element, "Area"):
        return element.Area
    return 0.0

# ---------------- MAIN ----------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# 1. Get Selection
sel_ids = uidoc.Selection.GetElementIds()

if not sel_ids:
    forms.alert("Please select Room(s) or Link(s) first.", exitscript=True)

valid_rooms = []
total_area_internal = 0.0
processed_count = 0

# 2. Iterate Selection
for i in sel_ids:
    e = doc.GetElement(i)
    
    # CASE A: Native Room
    if isinstance(e, DB.Architecture.Room):
        if e.Area > 0:
            total_area_internal += e.Area
            processed_count += 1
            
    # CASE B: Revit Link Instance
    # (If a link is selected, we calculate rooms inside it)
    elif isinstance(e, RevitLinkInstance):
        link_doc = e.GetLinkDocument()
        if link_doc:
            # Collect all rooms in the linked document
            # Note: We collect 'SpatialElement' to catch Rooms/Areas
            collector = FilteredElementCollector(link_doc)\
                        .OfCategory(BuiltInCategory.OST_Rooms)\
                        .WhereElementIsNotElementType()\
                        .ToElements()
            
            for linked_room in collector:
                if linked_room.Area > 0:
                    total_area_internal += linked_room.Area
                    processed_count += 1

# 3. Validation & Output
if processed_count == 0:
    forms.alert("No placed rooms found in selection.\n(Note: Unplaced rooms have 0 Area)", title="Result")
    sys.exit()

# 4. Conversion & Display
# Revit internal area is always Square Feet
val_sqft = total_area_internal
val_sqm = to_sqm(total_area_internal)

res_msg = "Total Area Calculation:\n\n" \
          "Count: {} Room(s)\n" \
          "--------------------------------\n" \
          "{:.2f} m²\n" \
          "{:.2f} sqft".format(processed_count, val_sqm, val_sqft)

forms.alert(res_msg, title="Total Room Area")