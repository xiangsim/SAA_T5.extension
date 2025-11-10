# -*- coding: utf-8 -*-
__title__ = "RevisionCloud\nFromSelection"
__author__ = "JK_Sim"
__doc__ = """Version 2.3
Date: 10.11.2025
_____________________________________________________________________
Description:
Generate Revision Cloud from selected elements:
Rooms, model elements, or annotation elements.

Pipeline:
1) Collect selected elements
2) Build per-element CurveLoop(s) on a unified plane (Z=0):
   - Rooms: boundary loops projected to Z=0
   - Model: convex hull from geometry edges (auto-oriented), Z=0
   - Annotations: oriented bbox rectangle, Z=0
3) Offset each loop by 300 mm
4) Extrude wafers and boolean-union into one
5) Take merged top face → map to sheet → **stitch gaps & close** → cloud
_____________________________________________________________________
"""

import clr, math
from Autodesk.Revit.DB import *
from pyrevit import forms
from System.Collections.Generic import List
from Autodesk.Revit.DB import Transaction

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

MM_TO_FT = 1.0 / 304.8
OFFSET_DIST = 300 * MM_TO_FT
THICKNESS   = 10 * MM_TO_FT

# --------------------------- basic utils ----------------------------
def distinct_xy(points, prec=6):
    seen = {}
    for p in points:
        k = (round(p.X, prec), round(p.Y, prec))
        seen[k] = XYZ(p.X, p.Y, 0.0)
    return list(seen.values())

def convex_hull(points):
    if len(points) < 3:
        return points[:]
    pts = sorted(points, key=lambda p: (p.X, p.Y))
    def cross(o, a, b):
        return (a.X - o.X) * (b.Y - o.Y) - (a.Y - o.Y) * (b.X - o.X)
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
    n = len(points)
    if n < 2:
        return cl
    tol = doc.Application.ShortCurveTolerance
    for i in range(n):
        p1 = XYZ(points[i].X, points[i].Y, 0.0)
        p2 = XYZ(points[(i+1)%n].X, points[(i+1)%n].Y, 0.0)
        if p1.DistanceTo(p2) > tol:
            cl.Append(Line.CreateBound(p1, p2))
    return cl

# --------------- per-type loop extractors (flatten to Z=0) ----------
def get_room_loops(room):
    loops = []
    try:
        opts = SpatialElementBoundaryOptions()
        segs = room.GetBoundarySegments(opts)
        if not segs: return loops
        for loop in segs:
            cl = CurveLoop()
            for seg in loop:
                c = seg.GetCurve()
                p0, p1 = c.GetEndPoint(0), c.GetEndPoint(1)
                cl.Append(Line.CreateBound(XYZ(p0.X, p0.Y, 0.0), XYZ(p1.X, p1.Y, 0.0)))
            if cl.NumberOfCurves() > 1:
                loops.append(cl)
    except:
        pass
    return loops

def model_convex_outline_points(el, view):
    opt = Options()
    opt.DetailLevel = ViewDetailLevel.Fine
    opt.IncludeNonVisibleObjects = False
    geo = el.get_Geometry(opt)
    pts = []
    if geo:
        for g in geo:
            if isinstance(g, Solid) and g.Volume > 1e-9:
                for e in g.Edges:
                    c = e.AsCurve()
                    pts.append(XYZ(c.GetEndPoint(0).X, c.GetEndPoint(0).Y, 0.0))
                    pts.append(XYZ(c.GetEndPoint(1).X, c.GetEndPoint(1).Y, 0.0))
            elif isinstance(g, GeometryInstance):
                inst_geo = g.GetInstanceGeometry()
                for sg in inst_geo:
                    if isinstance(sg, Solid) and sg.Volume > 1e-9:
                        for e in sg.Edges:
                            c = e.AsCurve()
                            pts.append(XYZ(c.GetEndPoint(0).X, c.GetEndPoint(0).Y, 0.0))
                            pts.append(XYZ(c.GetEndPoint(1).X, c.GetEndPoint(1).Y, 0.0))
    if len(pts) < 3:
        bbox = el.get_BoundingBox(None)
        if bbox:
            pts.extend([
                XYZ(bbox.Min.X, bbox.Min.Y, 0.0),
                XYZ(bbox.Max.X, bbox.Min.Y, 0.0),
                XYZ(bbox.Max.X, bbox.Max.Y, 0.0),
                XYZ(bbox.Min.X, bbox.Max.Y, 0.0),
            ])
    pts = distinct_xy(pts, 5)
    return convex_hull(pts)

def oriented_bbox_points(el, view):
    bbox = el.get_BoundingBox(view)
    if not bbox: return []
    n = view.ViewDirection.Normalize()
    try:
        vx = view.RightDirection.Normalize()
        vy = view.UpDirection.Normalize()
    except:
        ref = XYZ.BasisZ if abs(n.Z) < 0.9 else XYZ.BasisX
        vx  = n.CrossProduct(ref).Normalize()
        vy  = n.CrossProduct(vx).Normalize()
    x_axis = None
    if isinstance(el, Dimension) and el.Curve is not None:
        v = el.Curve.GetEndPoint(1) - el.Curve.GetEndPoint(0)
        v = v - n * v.DotProduct(n)
        if v.GetLength() > 1e-9:
            x_axis = v.Normalize()
    if x_axis is None: x_axis = vx
    y_axis = n.CrossProduct(x_axis).Normalize()
    center = (bbox.Min + bbox.Max) * 0.5
    base_pts = [
        XYZ(bbox.Min.X, bbox.Min.Y, 0.0),
        XYZ(bbox.Max.X, bbox.Min.Y, 0.0),
        XYZ(bbox.Max.X, bbox.Max.Y, 0.0),
        XYZ(bbox.Min.X, bbox.Max.Y, 0.0),
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
    return [XYZ(p1.X, p1.Y, 0.0), XYZ(p2.X, p2.Y, 0.0),
            XYZ(p3.X, p3.Y, 0.0), XYZ(p4.X, p4.Y, 0.0)]

def get_annotation_loops(el, view):
    pts = oriented_bbox_points(el, view)
    return [curveloop_from_points(pts)] if pts else []

def get_model_loops(el, view):
    pts = model_convex_outline_points(el, view)
    return [curveloop_from_points(pts)] if pts else []

# ----------------------------- ops -----------------------------------
def offset_loop(loop, dist):
    try:    return CurveLoop.CreateViaOffset(loop, dist, XYZ.BasisZ)
    except: return loop

def make_wafer(loop):
    try:    return GeometryCreationUtilities.CreateExtrusionGeometry([loop], XYZ.BasisZ, THICKNESS)
    except: return None

def merge_solids(solids):
    if not solids: return None
    merged = solids[0]
    for s in solids[1:]:
        try:
            merged = BooleanOperationsUtils.ExecuteBooleanOperation(merged, s, BooleanOperationsType.Union)
        except:
            pass
    return merged

def get_top_face(solid):
    faces = [f for f in solid.Faces if isinstance(f, PlanarFace) and f.FaceNormal.Z > 0.9]
    if not faces: return None
    faces.sort(key=lambda f: f.Area, reverse=True)
    return faces[0]

def map_curve_loop_to_sheet(loop, model_view, viewport):
    """
    Map a model-space CurveLoop to sheet-space.
    Handles discontinuities safely by rebuilding the CurveLoop from transformed segments.
    """
    tol = doc.Application.ShortCurveTolerance
    segs = []  # store transformed segments first

    # --- Transform step ---
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
        # Fallback manual transform
        crop = model_view.CropBox
        A1   = (crop.Min + crop.Max) * 0.5
        box  = viewport.GetBoxOutline()
        A2   = (box.MinimumPoint + box.MaximumPoint) * 0.5
        inv_scale = 1.0 / float(model_view.Scale if model_view.Scale else 1.0)
        rot_map = {0:0.0,1:math.pi/2,2:math.pi,3:3*math.pi/2}
        ang   = rot_map.get(int(viewport.Rotation), 0.0)
        trot  = Transform.CreateRotation(XYZ.BasisZ, ang)
        def map_pt(p):
            d = XYZ((p.X - A1.X)*inv_scale, (p.Y - A1.Y)*inv_scale, 0.0)
            v = trot.OfVector(d)
            return XYZ(A2.X+v.X, A2.Y+v.Y, 0.0)
        it = loop.GetCurveLoopIterator()
        while it.MoveNext():
            c = it.Current
            p1, p2 = map_pt(c.GetEndPoint(0)), map_pt(c.GetEndPoint(1))
            if p1.DistanceTo(p2) > tol:
                segs.append(Line.CreateBound(p1, p2))

    # --- Rebuild a safe, stitched CurveLoop ---
    cl_sheet = CurveLoop()
    if not segs:
        return cl_sheet

    # sort curves by connectivity
    sorted_curves = [segs[0]]
    segs_left = segs[1:]
    while segs_left:
        last_end = sorted_curves[-1].GetEndPoint(1)
        next_idx = None
        for i, c in enumerate(segs_left):
            s, e = c.GetEndPoint(0), c.GetEndPoint(1)
            if last_end.DistanceTo(s) < tol*10:
                next_idx = i
                break
            elif last_end.DistanceTo(e) < tol*10:
                next_idx = i
                segs_left[i] = c.CreateReversed()
                break
        if next_idx is not None:
            sorted_curves.append(segs_left.pop(next_idx))
        else:
            # no connection found, break to prevent infinite loop
            break

    # close gap if any
    if len(sorted_curves) > 1:
        first_start = sorted_curves[0].GetEndPoint(0)
        last_end    = sorted_curves[-1].GetEndPoint(1)
        if first_start.DistanceTo(last_end) > tol*10:
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
    if len(pts) < 3: return loop
    a = sum(pts[i].X*pts[(i+1)%len(pts)].Y - pts[(i+1)%len(pts)].X*pts[i].Y for i in range(len(pts)))*0.5
    if a <= 0: return loop
    rev = CurveLoop()
    curves = []
    it = loop.GetCurveLoopIterator()
    while it.MoveNext():
        curves.insert(0, it.Current.CreateReversed())
    for c in curves: rev.Append(c)
    return rev

# ----------------------------- main ----------------------------------
model_view = doc.ActiveView
if isinstance(model_view, ViewSheet):
    forms.alert("Run this from a MODEL view, not a sheet.", title="Error")
    raise Exception("Wrong view type.")

sel_ids = uidoc.Selection.GetElementIds()
if not sel_ids:
    forms.alert("No elements selected.", title="Error")
    raise Exception("No elements selected.")
elements = [doc.GetElement(eid) for eid in sel_ids]

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
rev_labels = ["Revision {} | {}".format(r.SequenceNumber, r.Description or "No Description")
              for r in sorted(revisions, key=lambda x: x.SequenceNumber, reverse=True)]
sel = forms.SelectFromList.show(rev_labels, title="Select Revision")
if not sel: raise Exception("User cancelled.")
seq = int(sel.split(" | ")[0].replace("Revision ", ""))
revision = {r.SequenceNumber: r for r in revisions}[seq]

# STEP 1–3: loops → offset → wafers
all_solids = []
for el in elements:
    loops = []
    try:
        if isinstance(el, SpatialElement):
            loops = get_room_loops(el)
        else:
            cat = el.Category
            if cat and cat.CategoryType.ToString() == "Annotation":
                loops = get_annotation_loops(el, model_view)
            else:
                loops = get_model_loops(el, model_view)
    except:
        continue

    for cl in loops or []:
        if not cl or cl.NumberOfCurves() < 1:
            continue
        off   = offset_loop(cl, OFFSET_DIST)
        solid = make_wafer(off)
        if solid:
            all_solids.append(solid)

if not all_solids:
    raise Exception("No geometry created.")

# STEP 4: union + top face
merged = merge_solids(all_solids)
if not merged: raise Exception("Failed to merge solids.")
top_face = get_top_face(merged)
if not top_face: raise Exception("No top face found.")

# STEP 5: map to sheet, STITCH & CLOSE, then cloud
loops = list(top_face.GetEdgesAsCurveLoops())
outer_loop = loops[0] if loops else None
if len(loops) > 1:
    def loop_len(cl):
        total = 0.0
        it = cl.GetCurveLoopIterator()
        while it.MoveNext(): total += it.Current.Length
        return total
    outer_loop = max(loops, key=loop_len)

sheet_loop = map_curve_loop_to_sheet(outer_loop, model_view, viewport)
sheet_loop = reverse_if_needed(sheet_loop)

tol = doc.Application.ShortCurveTolerance
max_gap = tol * 100.0  # stitch only reasonable gaps

curves = List[Curve]()
first_start = None
prev_end    = None

it = sheet_loop.GetCurveLoopIterator()
while it.MoveNext():
    c = it.Current
    if not c or c.Length < tol:
        continue
    s = c.GetEndPoint(0)
    e = c.GetEndPoint(1)
    if first_start is None:
        first_start = s
    if prev_end is not None:
        gap = prev_end.DistanceTo(s)
        if gap > tol and gap < max_gap:
            try:
                curves.Add(Line.CreateBound(prev_end, s))
            except:
                pass
    curves.Add(c)
    prev_end = e

# close last → first if needed
if first_start and prev_end:
    gap = prev_end.DistanceTo(first_start)
    if gap > tol and gap < max_gap:
        try:
            curves.Add(Line.CreateBound(prev_end, first_start))
        except:
            pass

with Transaction(doc, "Create Revision Cloud (Selection)") as t:
    t.Start()
    RevisionCloud.Create(doc, sheet_view, revision.Id, curves)
    t.Commit()
