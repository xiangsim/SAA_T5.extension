# -*- coding: utf-8 -*-
__title__ = "Split Wall\nAt Column"
__author__ = "JK_Sim"
__doc__ = """Version = 2.7
Date    = 16.12.2025
_____________________________________________________________________
Description:

Splits selected Walls at Linked Columns.

Fixes in v2.7:
  • CORNER JOINS: Fixed issue where walls ending in a column (corner condition)
    would still try to auto-join.
  • LOGIC UPDATE: Now checks if a wall end is an 'Original End' or a 'Cut Face'.
    Only 'Original Ends' are allowed to join. All 'Cut Faces' are disallowed.
_____________________________________________________________________
"""

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
from System.Collections.Generic import List
import math

doc = revit.doc
uidoc = revit.uidoc

# ------------------------------------- INIT
sel_ids = uidoc.Selection.GetElementIds()
sel_refs = uidoc.Selection.GetReferences()

walls = []
linked_cols = []

# --- Pass 1: Walls
for id in sel_ids:
    el = doc.GetElement(id)
    if isinstance(el, Wall):
        walls.append(el)

# --- Pass 2: Linked Columns
for ref in sel_refs:
    if ref.LinkedElementId != ElementId.InvalidElementId:
        link_inst = doc.GetElement(ref.ElementId)
        if isinstance(link_inst, RevitLinkInstance):
            link_doc = link_inst.GetLinkDocument()
            try:
                linked_elem = link_doc.GetElement(ref.LinkedElementId)
                if linked_elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_StructuralColumns):
                    trans = link_inst.GetTotalTransform()
                    opt = Options()
                    opt.DetailLevel = ViewDetailLevel.Fine
                    opt.ComputeReferences = True
                    geom_elem = linked_elem.get_Geometry(opt)
                    
                    solids = []
                    def get_solids(geom_col):
                        for g in geom_col:
                            if isinstance(g, Solid) and g.Volume > 0:
                                solids.append(g)
                            elif isinstance(g, GeometryInstance):
                                get_solids(g.GetInstanceGeometry())
                    get_solids(geom_elem)
                    world_solids = [SolidUtils.CreateTransformed(s, trans) for s in solids]
                    if world_solids:
                        linked_cols.append(world_solids)
            except:
                pass

if not walls or not linked_cols:
    forms.alert("Select Walls AND Tab-select Linked Columns.")
    script.exit()

# ------------------------------------- HELPER FUNCTIONS

def manage_wall_joins_geometric(wall, p_current_start, p_current_end, p_orig_start, p_orig_end):
    """
    Decides joins based on geometry, not index.
    """
    # Tolerance for point comparison
    tol = 0.001

    # --- START JOIN ---
    # Allow join ONLY if this start point is the Original Wall's start point
    if p_current_start.DistanceTo(p_orig_start) < tol:
        if not WallUtils.IsWallJoinAllowedAtEnd(wall, 0):
            try: WallUtils.AllowWallJoinAtEnd(wall, 0)
            except: pass
    else:
        # It's a cut point (or gap edge) -> Disallow
        if WallUtils.IsWallJoinAllowedAtEnd(wall, 0):
            try: WallUtils.DisallowWallJoinAtEnd(wall, 0)
            except: pass

    # --- END JOIN ---
    # Allow join ONLY if this end point is the Original Wall's end point
    if p_current_end.DistanceTo(p_orig_end) < tol:
        if not WallUtils.IsWallJoinAllowedAtEnd(wall, 1):
            try: WallUtils.AllowWallJoinAtEnd(wall, 1)
            except: pass
    else:
        # It's a cut point -> Disallow
        if WallUtils.IsWallJoinAllowedAtEnd(wall, 1):
            try: WallUtils.DisallowWallJoinAtEnd(wall, 1)
            except: pass

def get_view_tags_map(doc, view_id):
    tag_map = {}
    collector = FilteredElementCollector(doc, view_id).OfClass(IndependentTag)
    for tag in collector:
        host_id_val = None
        try:
            if hasattr(tag, "GetTaggedLocalElementIds"): 
                ids = tag.GetTaggedLocalElementIds()
                if ids: host_id_val = ids[0].IntegerValue
            elif hasattr(tag, "TaggedLocalElementId"): 
                host_id_val = tag.TaggedLocalElementId.IntegerValue
        except:
            pass 
        if host_id_val:
            if host_id_val not in tag_map: tag_map[host_id_val] = []
            tag_map[host_id_val].append(tag)
    return tag_map

def get_insert_snapshot(wall, tag_map):
    snapshot = {}
    ids = wall.FindInserts(True, False, False, False)
    for id in ids:
        el = doc.GetElement(id)
        if el and isinstance(el.Location, LocationPoint):
            pt = el.Location.Point
            key = "{:.4f},{:.4f},{:.4f}".format(pt.X, pt.Y, pt.Z)
            p_mark = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            mark_val = p_mark.AsString() if p_mark else None
            tags_data = []
            if id.IntegerValue in tag_map:
                for t in tag_map[id.IntegerValue]:
                    t_info = { 'TypeId': t.GetTypeId(), 'HeadPos': t.TagHeadPosition, 
                               'Orientation': t.TagOrientation, 'HasLeader': t.HasLeader }
                    tags_data.append(t_info)
            snapshot[key] = { 'Mark': mark_val, 'Tags': tags_data }
    return snapshot

def restore_inserts_data(doc, wall, snapshot, view):
    ids = wall.FindInserts(True, False, False, False)
    wall_curve = wall.Location.Curve
    inserts_to_delete = []

    for id in ids:
        insert = doc.GetElement(id)
        if not insert: continue
        loc = insert.Location
        if not isinstance(loc, LocationPoint): continue
        pt = loc.Point
        result = wall_curve.Project(pt)
        
        is_on_segment = False
        if result and pt.DistanceTo(result.XYZPoint) < 0.00328:
            is_on_segment = True
        
        if is_on_segment:
            key = "{:.4f},{:.4f},{:.4f}".format(pt.X, pt.Y, pt.Z)
            if key in snapshot:
                data = snapshot[key]
                # Restore Mark
                if data['Mark']:
                    p_mark = insert.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
                    if p_mark:
                        try:
                            if p_mark.AsString() != data['Mark']: p_mark.Set(data['Mark'])
                        except: pass
                # Restore Tags
                if data['Tags'] and insert.Id != data.get('OrigId', ElementId.InvalidElementId):
                    for t_info in data['Tags']:
                        try:
                            ref = Reference(insert)
                            IndependentTag.Create(doc, t_info['TypeId'], view.Id, ref, 
                                                  t_info['HasLeader'], t_info['Orientation'], t_info['HeadPos'])
                        except: pass
        else:
            inserts_to_delete.append(id)
            
    if inserts_to_delete:
        try: doc.Delete(List[ElementId](inserts_to_delete))
        except: pass

# ------------------------------------- EXECUTION

with Transaction(doc, "Split Wall At Column") as t:
    t.Start()
    
    view = doc.ActiveView
    sketch_plane = view.SketchPlane
    global_tag_map = get_view_tags_map(doc, view.Id)
    walls_to_process = list(walls)

    for wall in walls_to_process:
        if not wall.IsValidObject: continue

        w_curve_orig = wall.Location.Curve
        p_start_anchor = w_curve_orig.GetEndPoint(0)
        p_end_anchor   = w_curve_orig.GetEndPoint(1)
        
        # 1. SNAPSHOT
        master_snapshot = get_insert_snapshot(wall, global_tag_map)
        
        # 2. FIND GAPS
        gap_intervals = []
        for col_solids in linked_cols:
            col_pts = []
            for solid in col_solids:
                opt = SolidCurveIntersectionOptions()
                try:
                    inter = solid.IntersectWithCurve(w_curve_orig, opt)
                    for k in range(inter.SegmentCount):
                        col_pts.append(inter.GetCurveSegment(k).GetEndPoint(0))
                        col_pts.append(inter.GetCurveSegment(k).GetEndPoint(1))
                except: pass
            
            if col_pts:
                col_pts.sort(key=lambda p: p_start_anchor.DistanceTo(p))
                if len(col_pts) >= 2:
                    gap_intervals.append((col_pts[0], col_pts[-1], p_start_anchor.DistanceTo(col_pts[0])))

        if not gap_intervals: continue

        # 3. MERGE OVERLAPS
        gap_intervals.sort(key=lambda x: x[2])
        merged_gaps = []
        if gap_intervals:
            curr_s, curr_e, _ = gap_intervals[0]
            for i in range(1, len(gap_intervals)):
                next_s, next_e, _ = gap_intervals[i]
                if p_start_anchor.DistanceTo(next_s) < p_start_anchor.DistanceTo(curr_e):
                    if p_start_anchor.DistanceTo(next_e) > p_start_anchor.DistanceTo(curr_e):
                        curr_e = next_e
                else:
                    merged_gaps.append((curr_s, curr_e))
                    curr_s, curr_e = next_s, next_e
            merged_gaps.append((curr_s, curr_e))

        # 4. BUILD COORDINATES
        points_chain = [p_start_anchor]
        gap_flags = []
        for g_s, g_e in merged_gaps:
            points_chain.append(g_s)
            gap_flags.append(False) # Wall
            points_chain.append(g_e)
            gap_flags.append(True)  # Gap
        points_chain.append(p_end_anchor)
        gap_flags.append(False)

        wall_segments = []
        gap_segments = []
        for i in range(len(points_chain) - 1):
            p1, p2 = points_chain[i], points_chain[i+1]
            if p1.DistanceTo(p2) < 0.003: continue
            if i < len(gap_flags) and gap_flags[i]: gap_segments.append((p1, p2))
            else: wall_segments.append((p1, p2))

        if len(wall_segments) < 1: continue

        # 5. COPIES
        copies_needed = len(wall_segments) - 1
        created_copies = []
        if copies_needed > 0:
            try:
                for _ in range(copies_needed):
                    c_ids = ElementTransformUtils.CopyElement(doc, wall.Id, XYZ.Zero)
                    created_copies.append(doc.GetElement(c_ids[0]))
            except: pass
        
        # 6. ASSIGN
        walls_to_update = []
        walls_to_update.append((wall, wall_segments[0]))
        for i, cp in enumerate(created_copies):
            if i+1 < len(wall_segments):
                walls_to_update.append((cp, wall_segments[i+1]))

        # 7. UPDATE GEOMETRY & JOINS
        for w_obj, (p1, p2) in walls_to_update:
            try:
                # Geometry
                new_curve = Line.CreateBound(p1, p2)
                w_obj.Location.Curve = new_curve
                
                # JOINS (New Geometric Logic)
                # We check: Is p1 the Original Start? Is p2 the Original End?
                manage_wall_joins_geometric(w_obj, p1, p2, p_start_anchor, p_end_anchor)

                # Data
                restore_inserts_data(doc, w_obj, master_snapshot, view)
            except Exception as e:
                print("Error: {}".format(e))

        # 8. ROOM SEPARATION
        for gs, ge in gap_segments:
            try:
                gap_line = Line.CreateBound(gs, ge)
                c_array = CurveArray()
                c_array.Append(gap_line)
                doc.Create.NewRoomBoundaryLines(sketch_plane, c_array, view)
            except: pass

    t.Commit()