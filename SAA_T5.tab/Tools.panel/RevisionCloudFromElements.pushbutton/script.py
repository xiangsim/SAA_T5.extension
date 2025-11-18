# -*- coding: utf-8 -*-
__title__ = "RevisionCloud\nFromElements"
__author__ = "JK_Sim"
__doc__ = """Version 2.8
Date: 18.11.2025
_____________________________________________________________________
Description:
Generate Revision Cloud from selected elements:
_____________________________________________________________________
How-to:

-> Select element(s)
-> Run this script
-> Done
_________
_____________________________________________________________________
"""

import clr, math
from Autodesk.Revit.DB import *
from pyrevit import forms
from System.Collections.Generic import List
from Autodesk.Revit.DB import Transaction

doc  = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

MM_TO_FT    = 1.0 / 304.8
OFFSET_DIST = 300 * MM_TO_FT
THICKNESS   = 10  * MM_TO_FT

# --------------------------- basic utils ----------------------------
def distinct_xy(points, prec=6):
    """Remove duplicate XY points (Z flattened)."""
    seen = {}
    for p in points:
        k = (round(p.X, prec), round(p.Y, prec))
        seen[k] = XYZ(p.X, p.Y, 0)
    return list(seen.values())

def convex_hull(points):
    """2D convex hull in XY; returns list of points in order."""
    if len(points) < 3:
        return points[:]
    pts = sorted(points, key=lambda p: (p.X, p.Y))
    def cross(o, a, b):
        return (a.X-o.X)*(b.Y-o.Y) - (a.Y-o.Y)*(b.X-o.X)
    lower, upper = [], []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    return lower[:-1] + upper[:-1]

def curveloop_from_points(points):
    cl = CurveLoop()
    if len(points) < 2:
        return cl
    tol = doc.Application.ShortCurveTolerance
    for i in range(len(points)):
        p1, p2 = points[i], points[(i + 1) % len(points)]
        if p1.DistanceTo(p2) > tol:
            cl.Append(Line.CreateBound(p1, p2))
    return cl

def project_model_points_to_view(pt_list, model_view, viewport):
    """
    Convert model-space points directly into sheet-space points
    BEFORE CurveLoop creation.
    """
    crop = model_view.CropBox
    A1 = (crop.Min + crop.Max) * 0.5    # model center

    box = viewport.GetBoxOutline()
    A2 = (box.MinimumPoint + box.MaximumPoint) * 0.5   # sheet center

    S = 1.0 / float(model_view.Scale if model_view.Scale else 1)

    rot_map = {0: 0.0, 1: math.pi/2, 2: math.pi, 3: 3*math.pi/2}
    ang = rot_map.get(int(viewport.Rotation), 0.0)
    trot = Transform.CreateRotation(XYZ.BasisZ, ang)

    out = []
    for p in pt_list:
        d = XYZ((p.X - A1.X) * S, (p.Y - A1.Y) * S, 0)
        v = trot.OfVector(d)
        out.append(XYZ(A2.X + v.X, A2.Y + v.Y, 0))

    return out

# -------------------- scope box clipping helpers --------------------
def get_scopebox_rect(view):
    """
    Returns the active view's Scope Box rectangle in MODEL XY:
    (xmin, ymin, xmax, ymax), or None if no scope box is assigned.
    """
    param = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
    if not param:
        return None

    sb_id = param.AsElementId()
    if not sb_id or sb_id.IntegerValue < 1:
        return None

    sb = doc.GetElement(sb_id)
    if not sb:
        return None

    bb = sb.get_BoundingBox(None)
    if not bb:
        return None

    return (bb.Min.X, bb.Min.Y, bb.Max.X, bb.Max.Y)

def loop_to_xy_points(loop):
    """
    Convert a CurveLoop into a list of XY points (Z=0), preserving order.
    """
    pts = []
    it = loop.GetCurveLoopIterator()
    first = None
    last = None
    tol = doc.Application.ShortCurveTolerance

    while it.MoveNext():
        c = it.Current
        p0 = c.GetEndPoint(0)
        if first is None:
            first = p0
        pts.append(XYZ(p0.X, p0.Y, 0))
        last = c.GetEndPoint(1)

    if last is not None and first is not None:
        if last.DistanceTo(first) > tol:
            pts.append(XYZ(last.X, last.Y, 0))

    return pts

def clip_polygon_to_rect(points, xmin, ymin, xmax, ymax):
    """
    Sutherland–Hodgman clipping of a polygon (list[XYZ]) against
    axis-aligned rectangle [xmin,xmax] x [ymin,ymax] in model XY.
    Returns list[XYZ] (possibly empty).
    """
    if not points:
        return []

    def clip_edge(pts, inside, intersect):
        if not pts:
            return []
        out = []
        prev = pts[-1]
        prev_in = inside(prev)
        for curr in pts:
            curr_in = inside(curr)
            if curr_in:
                if not prev_in:
                    out.append(intersect(prev, curr))
                out.append(curr)
            elif prev_in:
                out.append(intersect(prev, curr))
            prev, prev_in = curr, curr_in
        return out

    def inside_left(p):   return p.X >= xmin
    def inside_right(p):  return p.X <= xmax
    def inside_bottom(p): return p.Y >= ymin
    def inside_top(p):    return p.Y <= ymax

    def intersect_x(p1, p2, x_clip):
        dx = (p2.X - p1.X)
        if abs(dx) < 1e-9:
            return XYZ(x_clip, p1.Y, 0)
        t = (x_clip - p1.X) / dx
        y = p1.Y + t * (p2.Y - p1.Y)
        return XYZ(x_clip, y, 0)

    def intersect_y(p1, p2, y_clip):
        dy = (p2.Y - p1.Y)
        if abs(dy) < 1e-9:
            return XYZ(p1.X, y_clip, 0)
        t = (y_clip - p1.Y) / dy
        x = p1.X + t * (p2.X - p1.X)
        return XYZ(x, y_clip, 0)

    pts_out = clip_edge(points, inside_left,  lambda a, b: intersect_x(a, b, xmin))
    pts_out = clip_edge(pts_out, inside_right, lambda a, b: intersect_x(a, b, xmax))
    pts_out = clip_edge(pts_out, inside_bottom, lambda a, b: intersect_y(a, b, ymin))
    pts_out = clip_edge(pts_out, inside_top,   lambda a, b: intersect_y(a, b, ymax))
    return pts_out

# --------------------- oriented frame helpers -----------------------
def horizontalize(v):
    return XYZ(v.X, v.Y, 0)

def safe_norm(v):
    L = (v.X*v.X + v.Y*v.Y + v.Z*v.Z) ** 0.5
    return XYZ(v.X/L, v.Y/L, v.Z/L) if L > 1e-9 else None

def get_element_frame(el):
    """
    Returns (origin, x_axis, y_axis) for an element.
    - FamilyInstance: use instance transform basis.
    - Wall: use wall curve direction.
    - Fallback: bbox center, world X/Y.
    """
    if isinstance(el, FamilyInstance):
        try:
            tf = el.GetTransform()
            origin = tf.Origin
            x = safe_norm(horizontalize(tf.BasisX)) or XYZ(1, 0, 0)
            y = safe_norm(XYZ(-x.Y, x.X, 0)) or XYZ(0, 1, 0)
            return origin, x, y
        except:
            pass

    if isinstance(el, Wall):
        loc = el.Location
        if isinstance(loc, LocationCurve):
            c  = loc.Curve
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            dir2d = horizontalize(XYZ(p1.X - p0.X, p1.Y - p0.Y, 0))
            x = safe_norm(dir2d) or XYZ(1, 0, 0)
            y = safe_norm(XYZ(-x.Y, x.X, 0)) or XYZ(0, 1, 0)
            bb = el.get_BoundingBox(None)
            origin = (bb.Min + bb.Max) * 0.5 if bb else p0
            return origin, x, y

    bb = el.get_BoundingBox(None)
    origin = (bb.Min + bb.Max) * 0.5 if bb else XYZ(0, 0, 0)
    return origin, XYZ(1, 0, 0), XYZ(0, 1, 0)

# ------------------ per-type loop extractors ------------------------
def get_room_loops(room):
    """Original room boundary logic (unchanged)."""
    loops = []
    try:
        opts = SpatialElementBoundaryOptions()
        segs = room.GetBoundarySegments(opts)
        if not segs:
            return loops
        for loop in segs:
            cl = CurveLoop()
            for seg in loop:
                c = seg.GetCurve()
                p0, p1 = c.GetEndPoint(0), c.GetEndPoint(1)
                cl.Append(Line.CreateBound(XYZ(p0.X, p0.Y, 0), XYZ(p1.X, p1.Y, 0)))
            if cl.NumberOfCurves() > 1:
                loops.append(cl)
    except:
        pass
    return loops

def wall_convex_outline_points(el, view):
    """Walls: convex hull of solid edges (unchanged)."""
    opt = Options()
    opt.IncludeNonVisibleObjects = False
    opt.View = view

    geo = el.get_Geometry(opt)
    pts = []
    if geo:
        for g in geo:
            if isinstance(g, Solid) and g.Volume > 1e-9:
                for e in g.Edges:
                    c = e.AsCurve()
                    pts.extend([
                        XYZ(c.GetEndPoint(0).X, c.GetEndPoint(0).Y, 0),
                        XYZ(c.GetEndPoint(1).X, c.GetEndPoint(1).Y, 0)
                    ])
            elif isinstance(g, GeometryInstance):
                for sg in g.GetInstanceGeometry():
                    if isinstance(sg, Solid) and sg.Volume > 1e-9:
                        for e in sg.Edges:
                            c = e.AsCurve()
                            pts.extend([
                                XYZ(c.GetEndPoint(0).X, c.GetEndPoint(0).Y, 0),
                                XYZ(c.GetEndPoint(1).X, c.GetEndPoint(1).Y, 0)
                            ])

    if len(pts) < 3:
        bb = el.get_BoundingBox(None)
        if bb:
            pts = [
                XYZ(bb.Min.X, bb.Min.Y, 0),
                XYZ(bb.Max.X, bb.Min.Y, 0),
                XYZ(bb.Max.X, bb.Max.Y, 0),
                XYZ(bb.Min.X, bb.Max.Y, 0)
            ]

    return convex_hull(distinct_xy(pts, 6))

# -------- model elements: oriented bbox (4 pts) --------
def model_oriented_bbox_points(el, view):
    """
    Oriented bounding box from:
    1) FamilySymbol BoundingBoxXYZ (using active view)
    2) Transformed using instance transform into MODEL space
    Returns 4 pts in model XY (flattened Z) for downstream processing.
    """

    # Try to get FamilySymbol bbox
    try:
        sym = el.Symbol
    except:
        sym = None

    if sym:
        bb = sym.get_BoundingBox(view)
    else:
        bb = None

    # Fallback to instance bbox
    if not bb:
        bb = el.get_BoundingBox(view) or el.get_BoundingBox(None)
        if not bb:
            return []

    minX, minY, minZ = bb.Min.X, bb.Min.Y, bb.Min.Z
    maxX, maxY, maxZ = bb.Max.X, bb.Max.Y, bb.Max.Z

    # Corners in symbol/element local space
    box_local = [
        XYZ(minX, minY, 0),
        XYZ(maxX, minY, 0),
        XYZ(maxX, maxY, 0),
        XYZ(minX, maxY, 0)
    ]

    # Instance transform → model coordinates
    try:
        T_inst = el.GetTransform()
    except:
        T_inst = Transform.Identity

    pts_model = [T_inst.OfPoint(p) for p in box_local]
    pts_model = [XYZ(p.X, p.Y, 0) for p in pts_model]

    return pts_model

# --------------------------------------------------------------------
def oriented_bbox_points(el, view):
    """For annotation elements: view-oriented bbox in sheet plane."""
    bbox = el.get_BoundingBox(view)
    if not bbox:
        return []
    n = view.ViewDirection.Normalize()
    try:
        vx = view.RightDirection.Normalize()
        vy = view.UpDirection.Normalize()
    except:
        ref = XYZ.BasisZ if abs(n.Z) < 0.9 else XYZ.BasisX
        vx = n.CrossProduct(ref).Normalize()
        vy = n.CrossProduct(vx).Normalize()
    x_axis = vx
    y_axis = n.CrossProduct(x_axis).Normalize()
    center = (bbox.Min + bbox.Max) * 0.5
    base_pts = [
        XYZ(bbox.Min.X, bbox.Min.Y, 0),
        XYZ(bbox.Max.X, bbox.Min.Y, 0),
        XYZ(bbox.Max.X, bbox.Max.Y, 0),
        XYZ(bbox.Min.X, bbox.Max.Y, 0)
    ]
    def to_local(p):
        v = p - center
        return (v.DotProduct(x_axis), v.DotProduct(y_axis))
    xs, ys = zip(*[to_local(p) for p in base_pts])
    xmin, xmax = min(xs), max(xs)
    ymin, ymax = min(ys), max(ys)
    p1 = center + x_axis * xmin + y_axis * ymin
    p2 = center + x_axis * xmax + y_axis * ymin
    p3 = center + x_axis * xmax + y_axis * ymax
    p4 = center + x_axis * xmin + y_axis * ymax
    return [
        XYZ(p1.X, p1.Y, 0),
        XYZ(p2.X, p2.Y, 0),
        XYZ(p3.X, p3.Y, 0),
        XYZ(p4.X, p4.Y, 0)
    ]

def get_annotation_loops(el, view):
    pts = oriented_bbox_points(el, view)
    return [curveloop_from_points(pts)] if pts else []

def get_model_loops(el, view):
    """
    Model loops:
      - Walls: convex hull (wall_convex_outline_points).
      - Other model elements: oriented bbox (model_oriented_bbox_points).
    """
    if isinstance(el, Wall):
        pts = wall_convex_outline_points(el, view)
    else:
        pts = model_oriented_bbox_points(el, view)
    return [curveloop_from_points(pts)] if pts else []

# ----------------------------- ops -----------------------------------
def offset_loop(loop, dist):
    try:
        return CurveLoop.CreateViaOffset(loop, dist, XYZ.BasisZ)
    except:
        return loop

def make_wafer(loop):
    """Extrude wafer geometry with loop flattened to Z=0 plane."""
    try:
        flat = CurveLoop()
        it = loop.GetCurveLoopIterator()
        while it.MoveNext():
            c = it.Current
            s, e = c.GetEndPoint(0), c.GetEndPoint(1)
            flat.Append(Line.CreateBound(XYZ(s.X, s.Y, 0), XYZ(e.X, e.Y, 0)))
        return GeometryCreationUtilities.CreateExtrusionGeometry([flat], XYZ.BasisZ, THICKNESS)
    except:
        return None

# --- multi-group union (preserve disjoint solids) ---
def _bbox2d(bb):
    return (XYZ(bb.Min.X, bb.Min.Y, 0), XYZ(bb.Max.X, bb.Max.Y, 0))

def _bb_overlap2d(bbA, bbB, tol):
    amin, amax = _bbox2d(bbA)
    bmin, bmax = _bbox2d(bbB)
    return not (amax.X < bmin.X - tol or bmax.X < amin.X - tol or
                amax.Y < bmin.Y - tol or bmax.Y < amin.Y - tol)

def _solid_bbox(s):
    try:
        return s.GetBoundingBox()
    except:
        return None

def _try_union(a, b):
    try:
        return BooleanOperationsUtils.ExecuteBooleanOperation(a, b, BooleanOperationsType.Union)
    except:
        try:
            return BooleanOperationsUtils.ExecuteBooleanOperation(b, a, BooleanOperationsType.Union)
        except:
            return None

def merge_solids_groups(solids):
    if not solids:
        return []
    tol = doc.Application.ShortCurveTolerance
    groups = []
    for s in solids:
        placed = False
        sb = _solid_bbox(s)
        if sb is None:
            groups.append(s)
            continue
        for i, g in enumerate(groups):
            gb = _solid_bbox(g)
            if gb and _bb_overlap2d(sb, gb, tol * 10.0):
                u = _try_union(g, s)
                if u:
                    groups[i] = u
                    placed = True
                    break
        if not placed:
            groups.append(s)
    changed = True
    while changed and len(groups) > 1:
        changed = False
        new_groups = []
        while groups:
            base = groups.pop()
            j = 0
            while j < len(groups):
                cand = groups[j]
                bb1, bb2 = _solid_bbox(base), _solid_bbox(cand)
                if bb1 and bb2 and _bb_overlap2d(bb1, bb2, tol * 10.0):
                    u = _try_union(base, cand)
                    if u:
                        base = u
                        groups.pop(j)
                        changed = True
                        continue
                j += 1
            new_groups.append(base)
        groups = new_groups
    return groups

def get_top_face(solid):
    faces = [f for f in solid.Faces if isinstance(f, PlanarFace) and f.FaceNormal.Z > 0.9]
    if not faces:
        return None
    faces.sort(key=lambda f: f.Area, reverse=True)
    return faces[0]

def map_curve_loop_to_sheet(loop, model_view, viewport):
    tol = doc.Application.ShortCurveTolerance
    segs = []
    try:
        trs = list(model_view.GetModelToProjectionTransforms())
        chosen = trs[0] if trs else None
        if not chosen:
            raise Exception("No transform found.")
        t_model_to_proj  = chosen.GetModelToProjectionTransform()
        t_proj_to_sheet  = viewport.GetProjectionToSheetTransform()
        t_model_to_sheet = t_proj_to_sheet.Multiply(t_model_to_proj)
        it = loop.GetCurveLoopIterator()
        while it.MoveNext():
            c_s = it.Current.CreateTransformed(t_model_to_sheet)
            if c_s and c_s.ApproximateLength > tol:
                segs.append(c_s)
    except:
        crop = model_view.CropBox
        A1   = (crop.Min + crop.Max) * 0.5
        box  = viewport.GetBoxOutline()
        A2   = (box.MinimumPoint + box.MaximumPoint) * 0.5
        inv_scale = 1.0 / float(model_view.Scale if model_view.Scale else 1.0)
        rot_map = {0: 0.0, 1: math.pi/2, 2: math.pi, 3: 3*math.pi/2}
        ang = rot_map.get(int(viewport.Rotation), 0.0)
        trot = Transform.CreateRotation(XYZ.BasisZ, ang)

        def map_pt(p):
            d = XYZ((p.X - A1.X) * inv_scale, (p.Y - A1.Y) * inv_scale, 0)
            v = trot.OfVector(d)
            return XYZ(A2.X + v.X, A2.Y + v.Y, 0)

        it = loop.GetCurveLoopIterator()
        while it.MoveNext():
            c = it.Current
            p1, p2 = map_pt(c.GetEndPoint(0)), map_pt(c.GetEndPoint(1))
            if p1.DistanceTo(p2) > tol:
                segs.append(Line.CreateBound(p1, p2))

    cl_sheet = CurveLoop()
    if not segs:
        return cl_sheet

    sorted_curves = [segs[0]]
    segs_left = segs[1:]
    while segs_left:
        last_end = sorted_curves[-1].GetEndPoint(1)
        next_idx = None
        for i, c in enumerate(segs_left):
            s, e = c.GetEndPoint(0), c.GetEndPoint(1)
            if last_end.DistanceTo(s) < tol * 10:
                next_idx = i
                break
            elif last_end.DistanceTo(e) < tol * 10:
                next_idx = i
                segs_left[i] = c.CreateReversed()
                break
        if next_idx is not None:
            sorted_curves.append(segs_left.pop(next_idx))
        else:
            break

    if len(sorted_curves) > 1:
        first_start = sorted_curves[0].GetEndPoint(0)
        last_end    = sorted_curves[-1].GetEndPoint(1)
        if first_start.DistanceTo(last_end) > tol * 10:
            try:
                sorted_curves.append(Line.CreateBound(last_end, first_start))
            except:
                pass

    for c in sorted_curves:
        try:
            cl_sheet.Append(c)
        except:
            pass

    return cl_sheet

def reverse_if_needed(loop):
    pts = []
    it = loop.GetCurveLoopIterator()
    while it.MoveNext():
        pts.append(it.Current.GetEndPoint(0))
    if len(pts) < 3:
        return loop
    a = sum(
        pts[i].X * pts[(i+1) % len(pts)].Y -
        pts[(i+1) % len(pts)].X * pts[i].Y
        for i in range(len(pts))
    ) * 0.5
    if a <= 0:
        return loop
    rev = CurveLoop()
    curves = []
    it = loop.GetCurveLoopIterator()
    while it.MoveNext():
        curves.insert(0, it.Current.CreateReversed())
    for c in curves:
        rev.Append(c)
    return rev

# ----------------------------- main ----------------------------------
model_view = doc.ActiveView
if isinstance(model_view, ViewSheet):
    forms.alert("Run this from a MODEL view, not a sheet.", title="Error")
    raise Exception("Wrong view type.")

# ------------------------------------- SELECTION (supports links)
sel_ids  = uidoc.Selection.GetElementIds()
sel_refs = uidoc.Selection.GetReferences()

if not sel_ids and not sel_refs:
    forms.alert("No elements selected.", title="Error")
    raise Exception("No elements selected.")

elements = []  # host elements + (linked_element, link_instance) tuples

# --- 1) Host model elements
for eid in sel_ids:
    el = doc.GetElement(eid)
    if el:
        elements.append(el)

# --- 2) Linked elements picked via tab-select
for ref in sel_refs:
    if ref.LinkedElementId != ElementId.InvalidElementId:
        link_inst = doc.GetElement(ref.ElementId)
        if isinstance(link_inst, RevitLinkInstance):
            link_doc = link_inst.GetLinkDocument()
            if not link_doc:
                continue
            linked_el = link_doc.GetElement(ref.LinkedElementId)
            if linked_el:
                # store as a tuple so we know this is from a link
                elements.append((linked_el, link_inst))

sheet_view, viewport = None, None
for vp in FilteredElementCollector(doc).OfClass(Viewport):
    if vp.ViewId == model_view.Id:
        viewport = vp
        sheet_view = doc.GetElement(vp.SheetId)
        break
if not sheet_view or not viewport:
    forms.alert("This view is not placed on any sheet.", title="Error")
    raise Exception("No viewport found.")

revisions = list(FilteredElementCollector(doc).OfClass(Revision))
if not revisions:
    raise Exception("No revisions found.")
rev_labels = [
    "Revision {} | {}".format(r.SequenceNumber, r.Description or "No Description")
    for r in sorted(revisions, key=lambda x: x.SequenceNumber, reverse=True)
]
sel = forms.SelectFromList.show(rev_labels, title="Select Revision")
if not sel:
    raise Exception("User cancelled.")
seq = int(sel.split(" | ")[0].replace("Revision ", ""))
revision = {r.SequenceNumber: r for r in revisions}[seq]

# STEP 1–3: Collect & flatten loops (per element, incl. linked)
all_loops = []

for el in elements:
    try:
        # --------------------------------------------------------
        # --------------- LINKED ELEMENTS (tuple) -----------------
        # --------------------------------------------------------
        if isinstance(el, tuple):
            linked_el, link_inst = el
            T = link_inst.GetTotalTransform()

            # 1) extract geometry in host view (VERY IMPORTANT)
            if isinstance(linked_el, SpatialElement):
                loops = get_room_loops(linked_el)

            else:
                loops = get_model_loops(linked_el, model_view)   # <--- FIXED HERE

            if not loops:
                continue

            # 2) apply link transform
            for cl in loops:
                new_cl = CurveLoop()
                it = cl.GetCurveLoopIterator()
                while it.MoveNext():
                    c = it.Current
                    new_cl.Append(c.CreateTransformed(T))
                all_loops.append(new_cl)

            continue   # done with linked element

        # --------------------------------------------------------
        # --------------------- HOST ELEMENTS --------------------
        # --------------------------------------------------------
        if isinstance(el, SpatialElement):
            loops = get_room_loops(el)
        else:
            cat = el.Category
            if cat and cat.CategoryType.ToString() == "Annotation":
                loops = get_annotation_loops(el, model_view)
            else:
                loops = get_model_loops(el, model_view)

        if loops:
            all_loops.extend(loops)

    except:
        continue

# Optional: clip loops against Scope Box (if any)
scope_rect = get_scopebox_rect(model_view)
if scope_rect:
    xmin, ymin, xmax, ymax = scope_rect
    clipped = []
    for cl in all_loops:
        pts = loop_to_xy_points(cl)
        pts_clip = clip_polygon_to_rect(pts, xmin, ymin, xmax, ymax)
        if len(pts_clip) >= 3:
            clipped.append(curveloop_from_points(pts_clip))
    all_loops = clipped

if not all_loops:
    raise Exception("No valid curve loops generated (inside scope box).")

# STEP 4: Offset, extrude wafers, and group-merge (preserve disjoint)
all_solids = []
for cl in all_loops:
    if not cl or cl.NumberOfCurves() < 1:
        continue
    off = offset_loop(cl, OFFSET_DIST)
    solid = make_wafer(off)
    if solid:
        all_solids.append(solid)

if not all_solids:
    raise Exception("No geometry created.")

solid_groups = merge_solids_groups(all_solids)
if not solid_groups:
    raise Exception("Failed to assemble solids.")

# STEP 5: For each group, take top face -> outer loop -> map to sheet -> stitch -> cloud
def pick_outer_loop(face):
    loops = list(face.GetEdgesAsCurveLoops())
    if not loops:
        return None
    if len(loops) == 1:
        return loops[0]
    def loop_len(cl):
        total = 0.0
        it = cl.GetCurveLoopIterator()
        while it.MoveNext():
            total += it.Current.Length
        return total
    return max(loops, key=loop_len)

tol = doc.Application.ShortCurveTolerance
max_gap = tol * 100.0
cloud_curve_sets = []

for s in solid_groups:
    face = get_top_face(s)
    if not face:
        continue
    outer_loop = pick_outer_loop(face)
    if not outer_loop:
        continue

    sheet_loop = map_curve_loop_to_sheet(outer_loop, model_view, viewport)
    sheet_loop = reverse_if_needed(sheet_loop)

    curves = List[Curve]()
    first_start, prev_end = None, None
    it = sheet_loop.GetCurveLoopIterator()
    while it.MoveNext():
        c = it.Current
        if not c or c.Length < tol:
            continue
        spt, ept = c.GetEndPoint(0), c.GetEndPoint(1)
        if first_start is None:
            first_start = spt
        if prev_end is not None:
            gap = prev_end.DistanceTo(spt)
            if gap > tol and gap < max_gap:
                try:
                    curves.Add(Line.CreateBound(prev_end, spt))
                except:
                    pass
        curves.Add(c)
        prev_end = ept

    if first_start and prev_end:
        gap = prev_end.DistanceTo(first_start)
        if gap > tol and gap < max_gap:
            try:
                curves.Add(Line.CreateBound(prev_end, first_start))
            except:
                pass

    if curves.Count > 0:
        cloud_curve_sets.append(curves)

if not cloud_curve_sets:
    raise Exception("No curve sets for revision cloud.")

with Transaction(doc, "Create Revision Cloud (Selection)") as t:
    t.Start()
    for curves in cloud_curve_sets:
        cloud = RevisionCloud.Create(doc, sheet_view, revision.Id, curves)
        try:
            param = cloud.LookupParameter("Comments")
            if param and not param.IsReadOnly:
                param.Set("SAA")
            else:
                cloud.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("SAA")
        except:
            pass
    t.Commit()
