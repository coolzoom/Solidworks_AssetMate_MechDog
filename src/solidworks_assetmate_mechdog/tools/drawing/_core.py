"""Shared drawing internals: constants, caches, COM helpers.

Everything the feature modules need lives here; they do
`from ._core import *` (explicit __all__ includes the private names).
"""

from __future__ import annotations

import glob

import os

import time

from ...sw_session import (
    tool_guard,
    SW_DOC_DRAWING,
    SW_DOC_PART,
    call_method,
    error,
    ok,
    open_document,
    variant,
)

drawing_ok = ok

_call_method = call_method

VIEW_ALIASES = {
    "front": ("*前视", "*Front"),
    "back": ("*后视", "*Back"),
    "left": ("*左视", "*Left"),
    "right": ("*右视", "*Right"),
    "top": ("*上视", "*Top"),
    "bottom": ("*下视", "*Bottom"),
    "isometric": ("*等轴测", "*Isometric"),
}

POSITIONS = {
    "front": (0.105, 0.150),
    "top": (0.105, 0.250),
    "right": (0.295, 0.150),
    "left": (0.295, 0.150),
    "back": (0.105, 0.060),
    "bottom": (0.105, 0.060),
    "isometric": (0.300, 0.235),
}

SUPPORTED_FORMATS = {".pdf", ".dwg", ".dxf"}

def _activate(sw, path):
    import pythoncom

    errors = variant(0, pythoncom.VT_BYREF | pythoncom.VT_I4)
    try:
        return sw.ActivateDoc3(path, False, 0, errors)
    except Exception:
        return None

def _find_template(paper):
    """Locate a drawing template without relying on user preferences."""
    want = paper.lower()
    roots = [
        r"C:\ProgramData\SOLIDWORKS",
        r"C:\ProgramData\SolidWorks",
    ]
    found = []
    for root in roots:
        found += glob.glob(os.path.join(root, "**", "*.drwdot"), recursive=True)
    if not found:
        return None
    for path in found:
        base = os.path.basename(path).lower()
        if base in ("gb_%s.drwdot" % want, "%s.drwdot" % want):
            return path
    return sorted(found)[0]

def _new_document(sw, template, attempts=8, delay=0.5):
    doc = None
    for _ in range(attempts):
        doc = sw.NewDocument(template, 0, 0, 0)
        if doc is not None:
            return doc
        time.sleep(delay)
    return None

def _iter_views(drawing):
    """Flatten GetViews (nested tuple of raw dispatch pointers) into IView objects."""
    import win32com.client as win

    flat = []

    def collect(obj):
        if isinstance(obj, (tuple, list)):
            for item in obj:
                collect(item)
        elif obj is not None:
            flat.append(win.Dispatch(obj))

    try:
        collect(drawing.GetViews)
    except Exception:
        return []
    return flat

def _place_views(views, sheet_w, sheet_h, gap=0.022, safety=None, pad=0.0):
    """Repack views using their REAL projected outlines so nothing overlaps.

    The bounding-box estimate used before creation is only good enough to pick a
    scale; an isometric view projects wider than the box, which is how views ended
    up overlapping. Here we read each view's actual outline and pack them into one
    or two columns, then move them with IView.Position (needs a VT_ARRAY|VT_R8
    VARIANT — a plain tuple is silently misread as (0, y)).
    """
    import pythoncom
    from win32com.client import VARIANT

    if safety:
        left = float(safety.get("left", sheet_w * 0.08))
        right = float(safety.get("right", sheet_w * 0.03))
        top = float(safety.get("top", sheet_h * 0.03))
        bottom = float(safety.get("bottom", sheet_h * 0.20))
    else:
        left, right = sheet_w * 0.08, sheet_w * 0.03
        top, bottom = sheet_h * 0.03, sheet_h * 0.20

    # Reserve room for dimension text: the projected outline covers geometry only, so a
    # view placed flush against the parts list has its dimensions drawn over the border.
    # P2-13: deflate GetOutline first — it reports every view 11.9 mm wider and taller
    # than the ink, which packs views further apart than they need to be.
    boxes = []
    for view in views:
        try:
            o = precise_outline(view) or list(view.GetOutline or [])
            boxes.append([view, (o[2] - o[0]) + pad, (o[3] - o[1]) + pad])
        except Exception:
            continue
    if not boxes:
        return {"fits": False, "columns": 0}

    avail_w = sheet_w - left - right
    avail_h = sheet_h - top - bottom
    columns = [[]]
    for cols in (2, 1):
        trial = [[] for _ in range(cols)]
        heights = [0.0] * cols
        for box in sorted(boxes, key=lambda b: -b[2]):
            index = heights.index(min(heights))
            trial[index].append(box)
            heights[index] += box[2] + gap
        total_w = sum(max((b[1] for b in col), default=0) for col in trial) + gap * (cols - 1)
        total_h = max(heights) if heights else 0.0
        columns = trial
        if total_w <= avail_w and total_h <= avail_h:
            break

    # GB templates keep a parts-list column on the left and a title block bottom-right;
    # staying inside the printable area is not enough — views must avoid those tables.
    x = left + pad / 2
    for col in columns:
        col_w = max((b[1] for b in col), default=0.0)
        y = sheet_h - top - pad / 2
        for view, width, height in col:
            center_x = x + width / 2
            center_y = y - height / 2
            try:
                view.Position = VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_R8, [center_x, center_y]
                )
            except Exception:
                pass
            y -= height + gap
        x += col_w + gap

    fits = True
    for view, _, _ in boxes:
        try:
            o = list(view.GetOutline or [])
            if o[0] < 0 or o[1] < 0 or o[2] > sheet_w or o[3] > sheet_h:
                fits = False
        except Exception:
            fits = False
    return {"fits": fits, "columns": len(columns)}

def _relayout(doc):
    """Repack views after annotation: dimension text extends a view's occupancy.

    A view whose outline fitted before auto_dimension can clip its dimensions
    against the sheet border afterwards, so annotation steps re-run the packer.
    """
    views, sheets = [], []
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            sheets.append(view)
        else:
            views.append(view)
    if not sheets or not views:
        return None
    try:
        o = list(sheets[0].GetOutline)
        w, h = o[2] - o[0], o[3] - o[1]
        result = _place_views(views, w, h, safety=template_safe_area(sheets[0], w, h))
        try:
            doc.GraphicsRedraw2()
        except Exception:
            pass
        return result
    except Exception:
        return None

def _model_view_name(model, semantic):
    try:
        names = list(model.GetModelViewNames or ())
    except Exception:
        names = []
    for alias in VIEW_ALIASES.get(semantic, ()):
        if alias in names:
            return alias
    return None

def _model_box(model):
    """Model bounding box in metres, or None.

    IPartDoc.GetPartBox works for parts; on an assembly it can raise or return zeros, and
    the old code then silently used scale 1.0 for a 1.69 m machine (reported as D-3).
    Falls back to the union of component boxes so auto_fit has something to work with.
    """
    def _valid(value):
        try:
            values = list(value or ())
            if len(values) < 6:
                return None
            box = (abs(values[3] - values[0]), abs(values[4] - values[1]),
                   abs(values[5] - values[2]))
            return box if box[0] > 0.001 and box[1] > 0.001 else None
        except Exception:
            return None

    try:
        found = _valid(model.GetPartBox(True))
        if found:
            return found
    except Exception:
        pass
    # Assembly path: union of visible component boxes.
    try:
        components = list(model.GetComponents(False) or ())
    except Exception:
        components = []
    low, high = None, None
    for component in components:
        try:
            import win32com.client as win

            comp = win.Dispatch(component)
            box = _valid(comp.GetBox(False, False))
            if box is None:
                continue
            values = list(comp.GetBox(False, False))
            if low is None:
                low = [values[0], values[1], values[2]]
                high = [values[3], values[4], values[5]]
            else:
                low = [min(low[i], values[i]) for i in range(3)]
                high = [max(high[i], values[i + 3]) for i in range(3)]
        except Exception:
            continue
    if low is not None:
        box = (abs(high[0] - low[0]), abs(high[1] - low[1]), abs(high[2] - low[2]))
        if box[0] > 0.001 and box[1] > 0.001:
            return box
    return None

def _standard_scale(fit):
    """Nearest standard drawing scale not larger than the available fit."""
    for candidate in (1.0, 0.5, 0.2, 0.1, 0.05, 0.02, 0.01, 0.005, 0.002):
        if candidate <= fit + 1e-9:
            return candidate
    return 0.002

OUTLINE_MARGIN_M = 0.0119  # measured constant: GetOutline is 11.9 mm larger per side

def precise_outline(view):
    """Projected outline in SHEET metres with the fixed GetOutline inflation removed.

    P2-13: `IView.GetOutline` inflates width AND height by a constant 11.9 mm on every
    view at every scale (measured on four views at 1:1 and 2:1 — front 59x28 reads
    70.9x39.9, top 59x80 reads 70.9x91.9, etc.). It is a fixed offset, not a projection
    error, so deflating by that constant recovers the real ink box at O(1) cost.

    This is the single source of truth for packing / overlap / out-of-sheet / hit-testing.
    A per-vertex "exact" variant is deliberately NOT provided: every route to real
    geometry points was measured dead on 2026-10-31 (a 31-face valve body):
      * `IView.GetVisibleEntities2(comp, 1)` returns SILHOUETTE edges — GetStartVertex
        returns None, only GetCurve has a value;
      * `IBody2.GetVertices` returns None for a body whose edges are all closed circles;
      * `IEdge.GetStartVertex` returns None through call_method as well;
      * `ICurve` on this build has no GetEndPoint / GetStartPoint / GetTessPoints.
    So deflation is not a shortcut, it is the only workable route. Accuracy measured
    against the rendered ink box: 58.98 mm vs 59.2 mm true (0.4%).

    Returns [x0, y0, x1, y1] in the same sheet frame as GetOutline, or None when it
    cannot be deflated safely (callers then fall back to the inflated but valid box).
    """
    try:
        o = list(view.GetOutline or [])
    except Exception:
        return None
    if len(o) < 4:
        return None
    m = OUTLINE_MARGIN_M / 2.0
    w = o[2] - o[0] - OUTLINE_MARGIN_M
    h = o[3] - o[1] - OUTLINE_MARGIN_M
    if w <= 0 or h <= 0:
        # The view is smaller than the inflation margin itself (a tiny detail view or
        # a blank one). Deflating would invert or collapse the box, which would make
        # _place_views stack views on top of each other — worse than the phantom margin.
        # Signal "cannot deflate" so every caller keeps the original GetOutline box.
        return None
    return [o[0] + m, o[1] + m, o[2] - m, o[3] - m]

def _dispatch_specific(annotation):
    """Resolve an annotation to its IDisplayDimension (or None)."""
    import win32com.client as win
    try:
        return win.Dispatch(annotation.GetSpecificAnnotation)
    except Exception:
        return None

def _save_drawing(doc):
    """Persist the drawing. Annotation steps otherwise live only in memory and the
    .SLDDRW handed to the user has views but none of the dimensions."""
    try:
        doc.SaveSilent
        return True
    except Exception:
        return False

def _open_drawing(path):
    from ...sw_session import get_sw

    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError("file not found: %s" % path)
    if os.path.splitext(path)[1].lower() != ".slddrw":
        raise ValueError("expected a .SLDDRW file, got %s" % path)
    doc = open_document(path, SW_DOC_DRAWING)
    _activate(get_sw(), doc.GetPathName or doc.GetTitle)
    return doc

def _annotation_dimension_full_name(annotation, win):
    """Model-level identity of a dimension annotation (e.g. D1@草图1@灌装阀阀体.Part).

    `IAnnotation.GetName` is only a sheet-unique label (细节项目272), so grouping by it can
    never detect the same driving dimension inserted into four different views. IDimension
    .FullName is the identity P1-6 needs; `Name` is the fallback on older builds.
    """
    try:
        specific = win.Dispatch(annotation.GetSpecificAnnotation)
    except Exception:
        return None
    try:
        dim = win.Dispatch(specific.GetDimension2(0))
    except Exception:
        return None
    for attr in ("FullName", "Name"):
        try:
            value = str(getattr(dim, attr))
        except Exception:
            continue
        if value:
            return value
    return None

def _dedupe_inserted_dimensions(doc, before_names, win):
    """Delete duplicate driving dimensions among the annotations JUST inserted (P1-6).

    SolidWorks' own `eliminate_duplicates` flag on InsertModelAnnotations3 did not stop the
    same D1@草图1 from landing once per view (56 dims across 4 views). Keep the FIRST
    occurrence of each model full name and delete the rest. Only annotations whose sheet
    name was absent before the insert are eligible, so dimensions the user already placed
    are never touched. Returns (duplicate_groups, removed_names, failures).
    """
    groups = {}
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        for raw in (view.GetAnnotations or ()):
            try:
                annotation = win.Dispatch(raw)
                if int(annotation.GetType) != 4:  # 4 = dimension annotation
                    continue
                sheet_name = str(annotation.GetName)
            except Exception:
                continue
            if sheet_name in before_names:
                continue
            full = _annotation_dimension_full_name(annotation, win)
            if full:
                groups.setdefault(full, []).append((sheet_name, annotation))
    duplicates = {k: v for k, v in groups.items() if len(v) > 1}
    removed, failed = [], []
    for full, entries in duplicates.items():
        for sheet_name, annotation in entries[1:]:
            try:
                doc.ClearSelection2(True)
                if not annotation.Select(False):
                    raise RuntimeError("Select(False) returned False")
                if doc.SelectionManager.GetSelectedObjectCount2(-1) < 1:
                    raise RuntimeError("nothing selected")
                _call_method(doc, "EditDelete")
                removed.append(sheet_name)
            except Exception as exc:
                failed.append({"name": sheet_name, "full_name": full,
                               "reason": str(exc)[:80]})
    if removed:
        try:
            _call_method(doc, "EditRebuild3")
        except Exception:
            pass
        try:
            doc.ClearSelection2(True)
        except Exception:
            pass
    return duplicates, removed, failed

def _filter_inserted_dimensions(doc, before_names, win, keep):
    """Delete inserted dimensions the plan did not keep (N3: plan ↔ auto_dimension).

    `keep` is `planner.keep_dimensions`, e.g.
    `[{"feature": "外形", "dimension": "总长", "value_mm": 1550.12}, ...]`.

    A dimension survives when either
      * its MODEL feature name equals an entry's `feature` (full names look like
        `D1@凸台-拉伸1@PART.Part`, so the middle segment is the feature), or
      * its value matches an entry's `value_mm` within `tol` — the planner's 外形 bucket
        describes bounding-box values, not feature names, so value matching is the only
        way those entries can ever hit.

    Safety valve: if NOTHING matches, nothing is deleted and `aborted` says why. A filter
    that recognises zero dimensions is far more likely to be wrong than the sheet is.
    Returns (kept, removed, failures, aborted_reason).
    """
    features = set()
    values = []
    for entry in keep or []:
        if not isinstance(entry, dict):
            continue
        feature = str(entry.get("feature") or "").strip()
        if feature and feature not in ("外形", "总长", "总宽", "总高"):
            features.add(feature)
        try:
            values.append(float(entry.get("value_mm")))
        except Exception:
            continue

    def _value_of(annotation):
        try:
            dim = win.Dispatch(win.Dispatch(annotation.GetSpecificAnnotation).GetDimension2(0))
            return float(dim.SystemValue) * 1000.0
        except Exception:
            return None

    candidates = []
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        for raw in (view.GetAnnotations or ()):
            try:
                annotation = win.Dispatch(raw)
                if int(annotation.GetType) != 4:
                    continue
                sheet_name = str(annotation.GetName)
            except Exception:
                continue
            if sheet_name in before_names:      # never touch what the user placed
                continue
            full = _annotation_dimension_full_name(annotation, win) or ""
            parts = full.split("@")
            feature = parts[1] if len(parts) >= 3 else ""
            value = _value_of(annotation)
            candidates.append((sheet_name, full, feature, value, annotation))

    if not candidates:
        return 0, [], [], "no dimensions were inserted by this call"

    def _keep(candidate):
        _, _, feature, value, _ = candidate
        if feature and feature in features:
            return True
        if value is None:
            return False
        for entry_value in values:
            tol = max(0.05, abs(entry_value) * 1e-4)
            if abs(value - entry_value) <= tol:
                return True
        return False

    kept = [c for c in candidates if _keep(c)]
    if not kept:
        return 0, [], [], ("none of the %d inserted dimensions matched the plan "
                           "(%d features / %d values) — deleted nothing"
                           % (len(candidates), len(features), len(values)))

    removed, failed = [], []
    for sheet_name, full, _, _, annotation in candidates:
        if _keep((sheet_name, full, _, _, annotation)):
            continue
        try:
            doc.ClearSelection2(True)
            if not annotation.Select(False):
                raise RuntimeError("Select(False) returned False")
            if doc.SelectionManager.GetSelectedObjectCount2(-1) < 1:
                raise RuntimeError("nothing selected")
            _call_method(doc, "EditDelete")
            removed.append(sheet_name)
        except Exception as exc:
            failed.append({"name": sheet_name, "full_name": full, "reason": str(exc)[:80]})
    if removed:
        try:
            _call_method(doc, "EditRebuild3")
        except Exception:
            pass
        try:
            doc.ClearSelection2(True)
        except Exception:
            pass
    return len(kept), removed, failed, None

def _push_dimensions_outside(doc, win, before_names, gap_m=0.012, step_m=0.008):
    """Move the dimensions this call inserted to just OUTSIDE their view (P0-1).

    SolidWorks inserts model dimensions at a fixed ~10 mm offset whatever the part size, so
    on a Ø46 disc 25 dimensions land inside the outline and overprint each other. The
    report's workaround was to scale the view up to 4:1 — that changes what the drawing
    states instead of fixing the layout.

    Each dimension is assigned to the NEAREST outline side of the view that contains it and
    stacked outward: distance = gap + slot * step. Within one side the stacking order follows
    the other axis, so two dimensions on the same side never share a distance.

    Only annotations NOT in `before_names` are touched — a dimension the user placed by hand
    is never moved.

    Returns (placed_count, failures, sides_used).
    """
    views = [v for v in _iter_views(doc) if getattr(v, "Type", None) != 1]
    boxes = []
    for view in views:
        box = precise_outline(view)
        if box and len(box) >= 4:
            boxes.append((view, box))
    if not boxes:
        return 0, [{"reason": "no view outlines readable"}], {}

    buckets = {}
    failures = []
    for view, box in boxes:
        try:
            raws = list(view.GetAnnotations or ())
        except Exception:
            continue
        for raw in raws:
            try:
                annotation = win.Dispatch(raw)
                if int(annotation.GetType) != 4:
                    continue
                name = str(annotation.GetName)
            except Exception:
                continue
            if name in before_names:
                continue                     # placed by the user, not by this call
            try:
                position = [float(v) for v in annotation.GetPosition]
            except Exception:
                continue

            # Owning view: the box that contains it, else the one with the nearest centre.
            owner_box = None
            for _candidate, candidate_box in boxes:
                if (candidate_box[0] <= position[0] <= candidate_box[2]
                        and candidate_box[1] <= position[1] <= candidate_box[3]):
                    owner_box = candidate_box
                    break
            if owner_box is None:
                best = None
                for _candidate, candidate_box in boxes:
                    cx = (candidate_box[0] + candidate_box[2]) / 2.0
                    cy = (candidate_box[1] + candidate_box[3]) / 2.0
                    distance = (position[0] - cx) ** 2 + (position[1] - cy) ** 2
                    if best is None or distance < best[0]:
                        best = (distance, candidate_box)
                if best is None:
                    continue
                owner_box = best[1]

            edges = {
                "left": abs(position[0] - owner_box[0]),
                "right": abs(owner_box[2] - position[0]),
                "bottom": abs(position[1] - owner_box[1]),
                "top": abs(owner_box[3] - position[1]),
            }
            side = min(edges, key=edges.get)
            order = position[1] if side in ("left", "right") else position[0]
            buckets.setdefault((str(view.Name), side), []).append(
                (order, annotation, position, owner_box))

    placed = 0
    sides_used = {}
    for (view_name, side), items in buckets.items():
        # Sort so the stacking follows the view's own reading order; the slot decides the
        # distance from the outline, so equal-order items still separate.
        items.sort(key=lambda entry: entry[0], reverse=(side in ("right", "top")))
        sides_used["%s:%s" % (view_name, side)] = len(items)
        for slot, (_order, annotation, position, box) in enumerate(items):
            offset = gap_m + slot * step_m
            if side == "left":
                target = (box[0] - offset, position[1])
            elif side == "right":
                target = (box[2] + offset, position[1])
            elif side == "bottom":
                target = (position[0], box[1] - offset)
            else:
                target = (position[0], box[3] + offset)
            try:
                annotation.SetPosition(float(target[0]), float(target[1]), 0)
                placed += 1
            except Exception as exc:
                failures.append({"view": view_name, "side": side,
                                 "reason": str(exc)[:70]})
    return placed, failures, sides_used

def _nearest_circle_edge(drawing_path, view_name, x, y, tolerance_m=0.005):
    """Circular edge of `view` nearest to sheet point (x, y), with its sheet centre.

    `doc.Extension.SelectByID2("", "EDGE", x, y)` cannot hit drawing edges by point (a
    v0.8.23 finding), so every single-point call on a perfectly valid hole returned
    `edge_not_found` — the report's 3/3 failures — while the batch tool succeeded on the
    same sheet, because that one ENUMERATES edges. Do the same: enumerate the view's
    circles, project each centre onto the sheet, take the nearest. x/y then only chooses
    WHICH hole instead of having to land exactly on the curve.

    Returns (match, None) or (None, reason).
    """
    import win32com.client as win

    doc = _open_drawing(drawing_path)
    view = None
    for candidate in _iter_views(doc):
        if getattr(candidate, "Type", None) == 1:
            continue
        if view_name and candidate.Name != view_name:
            continue
        view = candidate
        break
    if view is None:
        return None, "view_not_found"
    axis_map, reason = _view_axis_map(view)
    # `_view_axis_map` reads the part box through open_document(model_path), which switches
    # the ACTIVE document to the model — the view object and every edge taken from it are
    # stale afterwards, and `edge.Select2` on a stale handle silently returns False (that is
    # why the first attempt at this fix still ended in callout_failed). Come back to the
    # drawing and re-acquire both the view and its edges.
    doc = _open_drawing(drawing_path)
    view = None
    for candidate in _iter_views(doc):
        if getattr(candidate, "Type", None) == 1:
            continue
        if view_name and candidate.Name != view_name:
            continue
        view = candidate
        break
    if view is None:
        return None, "view went away after re-opening the drawing"
    try:
        doc.ActivateView(view.Name)
    except Exception:
        pass
    doc.ClearSelection2(True)
    try:
        position = list(view.Position)
        scale = float(view.ScaleDecimal)
    except Exception:
        return None, "view position unreadable"
    try:
        edges = [win.Dispatch(e) for e in (view.GetVisibleEntities2(view, 1) or ())]
    except Exception:
        return None, "view has no visible entities"

    best = None
    for edge in edges:
        try:
            curve = win.Dispatch(edge.GetCurve)
            if curve is None or not curve.IsCircle:
                continue
            params = list(curve.CircleParams)
            centre_model = (float(params[0]), float(params[1]), float(params[2]))
            radius_m = float(params[6])
        except Exception:
            continue
        if not axis_map:
            continue
        axes = axis_map["axes"]
        centre_m = axis_map["centre_m"]
        cx = position[0] + (centre_model[axes[0]] - centre_m[axes[0]]) * scale
        cy = position[1] + (centre_model[axes[1]] - centre_m[axes[1]]) * scale
        distance = ((cx - x) ** 2 + (cy - y) ** 2) ** 0.5
        if best is None or distance < best[0]:
            best = (distance, edge, [round(cx, 5), round(cy, 5)],
                    round(radius_m * 2000, 2),
                    [round(centre_model[0], 6), round(centre_model[1], 6),
                     round(centre_model[2], 6)])
    if best is None:
        if not axis_map:
            return None, "axes_mismatch: %s" % (reason or "no model->sheet mapping")
        return None, "no circular edge in this view"
    if best[0] > tolerance_m:
        return None, ("nearest circular edge is %.1f mm away (allowed %.0f mm)"
                      % (best[0] * 1000, tolerance_m * 1000))
    # Deliberately returns NO edge: reading the part box switched the active document, so
    # every COM handle taken here belongs to a dead one. `Select2` on such a handle reports
    # success while `AddHoleCallout2` then returns None — the symptom that made this tool
    # look impossible to fix. The caller re-finds the circle by diameter on its own handle.
    return {"centre": best[2], "diameter_mm": best[3],
            "centre_model": best[4],
            "distance_m": round(best[0], 5)}, None

def ann_position(note):
    """Position of a note's annotation, best effort (None when unavailable)."""
    try:
        return list(note.GetAnnotation.GetPosition)
    except Exception:
        return []

def _sheet_notes(sheet):
    """All INote objects on the sheet (annotations whose type is 6)."""
    import win32com.client as win

    flat = []

    def collect(obj):
        if isinstance(obj, (tuple, list)):
            for item in obj:
                collect(item)
        elif obj is not None:
            flat.append(win.Dispatch(obj))

    try:
        collect(sheet.GetAnnotations)
    except Exception:
        return []
    notes = []
    for ann in flat:
        try:
            if ann.GetType != 6:  # swAnnotationType_Note
                continue
            notes.append(ann.GetSpecificAnnotation)
        except Exception:
            continue
    return notes

TITLE_SLOTS = {
    "name": (0.391, 0.036),
    "code": (0.391, 0.022),
    "material": (0.344, 0.021),
}

def _note_at(sheet, target, tolerance=0.006):
    """Sheet note whose anchor is nearest `target` (metres), within tolerance."""
    import math
    import win32com.client as win

    try:
        raw = sheet.GetAnnotations
    except Exception:
        return None
    items = list(raw) if isinstance(raw, (tuple, list)) else ([raw] if raw else [])
    best = None
    for item in items:
        if item is None:
            continue
        try:
            ann = win.Dispatch(item)
            if ann.GetType != 6:  # notes only
                continue
            pos = list(ann.GetPosition)
            note = ann.GetSpecificAnnotation
        except Exception:
            continue
        distance = math.hypot(pos[0] - target[0], pos[1] - target[1])
        if distance <= tolerance and (best is None or distance < best[0]):
            best = (distance, note)
    return best[1] if best else None

def _note_name(note):
    """Reported name of a title-block cell, so the caller can verify what was written.

    P3-4: the field locator falls back to coordinates, and on a GB template those cells are
    easy to confuse — one report found the `material` slot occupied by the drawing code
    (GZ1000/8.06-02) until it was corrected. Echoing the cell's own annotation name plus its
    position makes that kind of mix-up visible in the receipt instead of only in the PDF.
    """
    if note is None:
        return None
    try:
        return str(_call_method(note, "GetAnnotation").GetName)
    except Exception:
        return None

def _projection_axes(view, edges):
    """Which two model axes the view projects (xy=front, xz=top, zy=right).

    Chosen by matching the view's outline aspect ratio against the model extents.
    Dimension values are PROJECTED lengths, so this must be known before comparing
    a new dimension against existing ones (a 1253 mm edge can dimension as 37.6 mm).

    P2-13: the aspect ratio uses the DEFLATED outline. GetOutline inflates width and height
    by the same 11.9 mm, which does not preserve the ratio — on a 70.9x39.9 view the raw
    ratio is 1.78 while the real 59x28 view is 2.11, enough to pick the wrong axis pair.
    """
    try:
        box = precise_outline(view) or list(view.GetOutline)
        width, height = box[2] - box[0], box[3] - box[1]
        if width <= 0 or height <= 0:
            return None
        ratio = width / height
    except Exception:
        return None

    pts = []
    for edge in edges:
        try:
            pts.append((edge.GetStartVertex.GetPoint, edge.GetEndVertex.GetPoint))
        except Exception:
            continue
    if not pts:
        return None
    lo = [min(min(a[i], b[i]) for a, b in pts) for i in range(3)]
    hi = [max(max(a[i], b[i]) for a, b in pts) for i in range(3)]
    size = [hi[i] - lo[i] for i in range(3)]

    best, best_err = None, None
    for axes in ((0, 1), (0, 2), (2, 1)):
        a, b = axes
        if size[a] <= 1e-6 or size[b] <= 1e-6:
            continue
        err = abs(size[a] / size[b] - ratio)
        if best_err is None or err < best_err:
            best, best_err = axes, err
    return best

def _edge_model_length(edge, axes):
    """Projected length along the two axes the view shows (this equals the dimension value)."""
    try:
        p1 = edge.GetStartVertex.GetPoint
        p2 = edge.GetEndVertex.GetPoint
    except Exception:
        return None
    if axes is None:
        return None
    a, b = axes
    return ((p1[a] - p2[a]) ** 2 + (p1[b] - p2[b]) ** 2) ** 0.5

def drawing_dimension_values(doc):
    """All dimension values already on the sheet, rounded to 0.1 mm.

    IView.GetDisplayDimensions only returns a subset (13 of 56 measured), so read the
    view's annotation list instead: dimension annotations have GetType() == 4, and the
    value comes from GetSpecificAnnotation().GetDimension2(0).SystemValue.
    """
    import win32com.client as win

    values = set()
    for view in _iter_views(doc):
        try:
            raw = view.GetAnnotations
        except Exception:
            continue
        items = list(raw) if isinstance(raw, (tuple, list)) else ([raw] if raw else [])
        for item in items:
            try:
                ann = win.Dispatch(item)
                if ann.GetType != 4:  # 4 = dimension annotation
                    continue
                dim = ann.GetSpecificAnnotation.GetDimension2(0)
                if dim is not None:
                    values.add(round(float(dim.SystemValue), 4))
            except Exception:
                continue
    return values

def _nearest_edge_at_sheet_point(view, x, y, tolerance):
    """Pick the visible edge whose midpoint projects nearest to sheet point (x, y).

    Drawing edges report MODEL coordinates, so a sheet point cannot be used directly.
    The view's projection plane is inferred from its outline aspect ratio against the
    model extents (xy = front, xz = top, zy = right), then each edge midpoint is mapped
    with paper = view.Position + (model - centre) * scale.

    Returns (distance_m, (p1, p2), projection) for the NEAREST edge even when that edge
    is OUTSIDE `tolerance`, so the caller can emit a self-calibration diagnostic instead
    of a bare "not found" (P1-8). `projection` carries the plane/pos/scale/centre needed
    to report both model and sheet coordinates. Returning None means only "no visible
    edges could be enumerated" — the tolerance check belongs to the caller.
    """
    import math
    import win32com.client as win

    try:
        edges = [win.Dispatch(e) for e in (view.GetVisibleEntities2(view, 1) or ())]
        outline = list(view.GetOutline)
        scale = float(view.ScaleDecimal)
        pos = list(view.Position)
    except Exception:
        return None
    if not edges or scale == 0:
        return None

    pts = []
    for edge in edges:
        try:
            p1 = edge.GetStartVertex.GetPoint
            p2 = edge.GetEndVertex.GetPoint
            pts.append((p1, p2))
        except Exception:
            continue
    if not pts:
        return None

    lo = [min(min(a[i], b[i]) for a, b in pts) for i in range(3)]
    hi = [max(max(a[i], b[i]) for a, b in pts) for i in range(3)]
    size = [hi[i] - lo[i] for i in range(3)]
    centre = [(hi[i] + lo[i]) / 2 for i in range(3)]

    width = outline[2] - outline[0]
    height = outline[3] - outline[1]
    if width <= 0 or height <= 0:
        return None
    ratio = width / height
    planes = {
        "xy": (0, 1),
        "xz": (0, 2),
        "zy": (2, 1),
    }
    best_plane, best_err = None, None
    for name, (a, b) in planes.items():
        if size[a] <= 1e-6 or size[b] <= 1e-6:
            continue
        err = abs(size[a] / size[b] - ratio)
        if best_err is None or err < best_err:
            best_plane, best_err = (a, b), err
    if best_plane is None:
        return None
    ai, bi = best_plane

    best = None
    for (p1, p2) in pts:
        mx = (p1[ai] + p2[ai]) / 2.0
        my = (p1[bi] + p2[bi]) / 2.0
        px = pos[0] + (mx - centre[ai]) * scale
        py = pos[1] + (my - centre[bi]) * scale
        distance = math.hypot(px - x, py - y)
        if best is None or distance < best[0]:
            best = (distance, (p1, p2))
    if best is None:
        return None
    plane_name = {(0, 1): "xy", (0, 2): "xz", (2, 1): "zy"}.get((ai, bi), "?")
    projection = {"plane": plane_name, "ai": ai, "bi": bi,
                  "pos": pos, "scale": scale, "centre": centre}
    return (best[0], best[1], projection)

def _view_at_point(doc, x, y):
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        try:
            # P2-13: hit-test the deflated box. GetOutline's phantom 11.9 mm margin made
            # points that sit just outside a view answer as "inside" it.
            o = precise_outline(view) or list(view.GetOutline)
        except Exception:
            continue
        if o[0] <= x <= o[2] and o[1] <= y <= o[3]:
            return view
    return None

def _view_outline_hint(doc):
    """'视图名 [l, b, r, t]' per view, for the point_outside_views error (P3-1)."""
    parts = []
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        try:
            outline = precise_outline(view) or list(view.GetOutline)
            name = str(view.Name)
        except Exception:
            continue
        if outline and len(outline) >= 4:
            parts.append("%s [%.3f, %.3f, %.3f, %.3f]" % (
                name, outline[0], outline[1], outline[2], outline[3]))
    return "；".join(parts) if parts else "（未能读取任何视图 outline）"

def _free_spot_for_view(doc, half_w, half_h, prefer_x, prefer_y, margin=0.006, exclude=None):
    """Lowest-clash sheet position for a view of ±half_w/±half_h (P1-2).

    Detail views were dropped at a fixed height (`sheet_h * 0.30`), so the second one landed
    at y = -0.0487 m — off the sheet — and overlapped its neighbour. Placement is decided
    BEFORE the view exists, from an estimate of its drawn size: scan a 10 mm grid over the
    sheet inner box, reject any cell that overlaps an existing view outline, and take the
    free cell closest to the requested spot so repeated calls stay predictable.

    Writing `IView.Position` on a detail view afterwards only moves it vertically (x snaps
    back), so x has to be right at CreateDetailViewAt4 time.
    """
    try:
        box = list(doc.GetFirstView.GetOutline)
        sheet_w, sheet_h = box[2] - box[0], box[3] - box[1]
    except Exception:
        return None
    occupied = []
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        try:
            name = str(view.Name)
            if exclude and name == exclude:
                continue
            o = precise_outline(view) or list(view.GetOutline)
        except Exception:
            continue
        if o and len(o) >= 4:
            occupied.append((o[0] - margin, o[1] - margin, o[2] + margin, o[3] + margin))

    def _clash(x, y):
        left, right = x - half_w, x + half_w
        bottom, top = y - half_h, y + half_h
        if left < 0.008 or bottom < 0.008 or right > sheet_w - 0.008 or top > sheet_h - 0.008:
            return True
        for o in occupied:
            if not (right <= o[0] or left >= o[2] or top <= o[1] or bottom >= o[3]):
                return True
        return False

    step = 0.01
    best, best_d2 = None, None
    y = half_h + 0.008
    while y <= sheet_h - half_h - 0.008 + 1e-9:
        x = half_w + 0.008
        while x <= sheet_w - half_w - 0.008 + 1e-9:
            if not _clash(x, y):
                d2 = (x - prefer_x) ** 2 + (y - prefer_y) ** 2
                if best_d2 is None or d2 < best_d2:
                    best, best_d2 = [round(x, 4), round(y, 4)], d2
            x += step
        y += step
    return best

_EDGE_CACHE = {}

EDGE_CACHE_LIMIT = 32

def _template_lines(sheet, sheet_w=None, sheet_h=None):
    """Template geometry as (x1, y1, x2, y2) sheet-space segments.

    GetLines3 returns 12 numbers per line: 6 of metadata then x1,y1,z1,x2,y2,z2.
    (Read naively with a stride of 6 it looks like garbage every other record.)
    """
    try:
        raw = sheet.GetLines3
    except Exception:
        return []
    if not raw:
        return []
    values = list(raw)
    limit_x = (sheet_w or 0.6) * 1.05
    limit_y = (sheet_h or 0.45) * 1.05
    segments = []
    for base in range(0, len(values) - 5, 6):
        try:
            x1, y1, x2, y2 = values[base], values[base + 1], values[base + 3], values[base + 4]
        except Exception:
            continue
        # Each record carries 6 metadata values before the coordinates, so a naive
        # stride-6 read also yields junk like (0, 2, 1, 1); drop anything outside
        # the sheet instead of guessing the record layout.
        if not (-1e-6 <= x1 <= limit_x and -1e-6 <= y1 <= limit_y):
            continue
        if not (-1e-6 <= x2 <= limit_x and -1e-6 <= y2 <= limit_y):
            continue
        segments.append((x1, y1, x2, y2))
    return segments

def template_safe_area(sheet, sheet_w, sheet_h):
    """Usable region of a sheet: inside the frame, clear of the title block / parts list.

    GB templates vary with paper size, so the margins are measured from the template's
    own lines instead of being hard-coded for A3 (which is why A2 layouts overflowed).
    """
    segments = _template_lines(sheet, sheet_w, sheet_h)
    fallback = {
        "left": sheet_w * 0.14, "right": sheet_w * 0.01,
        "top": sheet_h * 0.012, "bottom": sheet_h * 0.22,
        "source": "ratio",
    }
    if len(segments) < 8:
        return fallback

    horizontals = [s for s in segments if abs(s[1] - s[3]) < 1e-5 and abs(s[2] - s[0]) > 0.004]
    verticals = [s for s in segments if abs(s[0] - s[2]) < 1e-5 and abs(s[3] - s[1]) > 0.004]
    if not horizontals or not verticals:
        return fallback

    # Use the outermost extent of ALL segments: larger sheets carry centre lines that
    # would otherwise be mistaken for the frame edge (A2 measured right=0.30 instead of 0.59).
    frame_left = min(min(s[0], s[2]) for s in segments)
    frame_right = max(max(s[0], s[2]) for s in segments)
    frame_bottom = min(min(s[1], s[3]) for s in segments)
    frame_top = max(max(s[1], s[3]) for s in segments)

    # Title block: the highest horizontal line sitting in the bottom-right corner.
    title_candidates = [
        s for s in horizontals
        if min(s[0], s[2]) > frame_left + (frame_right - frame_left) * 0.45
        and s[1] < frame_bottom + (frame_top - frame_bottom) * 0.30
    ]
    title_top = max((s[1] for s in title_candidates), default=frame_bottom)

    # Parts list / addendum column: the right-most vertical line in the left margin.
    # Its right edge is drawn as many short row separators, so the length threshold
    # must stay small or the column looks absent.
    list_candidates = [
        s for s in verticals
        if s[0] < frame_left + (frame_right - frame_left) * 0.30
        and max(s[1], s[3]) > frame_bottom + (frame_top - frame_bottom) * 0.30
    ]
    list_right = max((s[0] for s in list_candidates), default=frame_left)

    margin = 0.004
    # The measured left edge also catches the frame's row/column index strip, which is
    # ~30% of the sheet width on A3 — leaving that much empty would starve the views.
    # Cap the reserved band by ratio so the template still adapts across paper sizes.
    left_band = min(list_right - frame_left + margin, sheet_w * 0.14)
    bottom_band = min(title_top - frame_bottom + margin, sheet_h * 0.24)
    measured = {
        "left": frame_left + left_band,
        "right": sheet_w - frame_right + margin,
        "top": sheet_h - frame_top + margin,
        "bottom": bottom_band,
        "source": "template",
        "title_block_top": round(title_top, 4),
        "parts_list_right": round(list_right, 4),
    }

    # Sanity check before trusting the measurement. On A2 the frame scan missed the right
    # border entirely and reported a usable band of 212 mm out of 594 mm, which then drove
    # the scale down to 1:10 on a part that fits at 1:2. If the reserved margins eat more
    # than half the sheet, the scan is wrong — fall back to ratios, which are size-relative
    # and cannot collapse like that.
    reserved_x = measured["left"] + measured["right"]
    reserved_y = measured["bottom"] + measured["top"]
    if reserved_x > sheet_w * 0.5 or reserved_y > sheet_h * 0.5:
        fallback_used = dict(fallback)
        fallback_used["source"] = "ratio (template scan rejected: reserved %.0f%% x %.0f%%)" % (
            reserved_x / sheet_w * 100, reserved_y / sheet_h * 100)
        return fallback_used
    return measured

def _view_axis_map(view):
    """Which two MODEL axes does this view display? (P1-4)

    A cut plane given in model coordinates must be translated into the view's sketch, so
    the tool has to know which model axis each sheet axis shows. There is no reliable API
    for this on SW2024/2026 (measured 2026-10-31):
      - `IView.GetViewPlaneNormal` does not exist (property and method both fail);
      - `IView.ModelToViewTransform` returns 16 doubles that are NOT a plain row-major 4x4,
        and `IMathPoint.MultiplyTransform` does not exist either.

    So the mapping is matched by SIZE: the view's deflated outline (GetOutline minus the
    measured 11.9 mm constant, P2-13) must equal the part box extent along the two
    displayed axes. Returns (info, reason); info carries axes / centre / extent / path.
    An ambiguous match (e.g. a cube) is refused rather than guessed.

    The caller must re-activate the DRAWING afterwards: reading the part box goes through
    open_document(model_path), which switches SolidWorks' active document.
    """
    from ...sw_session import call_method

    try:
        model_path = str(call_method(view, "GetReferencedModelName") or "")
    except Exception:
        try:
            model_path = str(view.GetReferencedModelName or "")
        except Exception:
            model_path = ""
    if not model_path or not os.path.exists(model_path):
        return None, "no_referenced_model"

    # Read the view's own outline BEFORE switching documents.
    box = precise_outline(view)
    if not box:
        return None, "outline_too_small"
    width = box[2] - box[0]
    height = box[3] - box[1]
    # The outline is in SHEET metres and the part box is in MODEL metres, so the extents
    # have to be scaled before they can be compared. Without this factor every view that
    # is not 1:1 failed to match and the tool answered `axes_mismatch` — which is exactly
    # what a report saw on 30 of 30 circles, and what made single-point hole callouts
    # impossible (they need a circle centre in sheet coordinates).
    try:
        view_scale = float(view.ScaleDecimal)
    except Exception:
        view_scale = 1.0
    if view_scale <= 1e-9:
        view_scale = 1.0

    try:
        model = open_document(model_path)
        raw = list(model.GetPartBox(True) or ())
    except Exception:
        return None, "model_box_failed"
    if len(raw) < 6:
        return None, "model_box_missing"
    lo = [float(raw[i]) for i in range(3)]
    hi = [float(raw[i]) for i in range(3, 6)]
    extent = [hi[i] - lo[i] for i in range(3)]
    if min(extent) <= 1e-6:
        return None, "model_box_degenerate"
    centre = [(lo[i] + hi[i]) / 2 for i in range(3)]

    pairs = ((0, 1), (0, 2), (2, 1), (1, 0), (2, 0), (1, 2))
    scored = []
    for a, b in pairs:
        if extent[a] <= 1e-6 or extent[b] <= 1e-6:
            continue
        scored.append((abs(extent[a] * view_scale - width)
                       + abs(extent[b] * view_scale - height), (a, b)))
    if not scored:
        return None, "axes_not_found"
    scored.sort(key=lambda item: item[0])
    best_err, best = scored[0]
    if best_err > max(width, height) * 0.05:
        return None, "axes_mismatch"
    if len(scored) > 1 and abs(scored[1][0] - best_err) < 1e-6:
        return None, "axes_ambiguous"
    return {
        "axes": [int(best[0]), int(best[1])],
        "centre_m": centre,
        "extent_m": extent,
        "model_path": model_path,
        "outline_m": [width, height],
    }, None

SECTION_KEEP_TOLERANCE = 0.10

def _section_sanity(source_view, section, orientation, depth):
    """Did the cut land on a real cross-section? (P1-4)

    A section plane perpendicular to the view's horizontal axis cannot change the extent
    along that axis, so the section's WIDTH must match the source view's width; same idea
    for a vertical cut and the height. Comparing exactly that dimension is geometry, not a
    heuristic.

    Measured 2026-10-31 on 灌装阀阀体 (source 59x28 mm after deflation):
      horizontal, offset -0.012 -> section 59x80, kept width 59.0 -> ratio 1.000 (correct)
      horizontal, offset +0.012 -> section 14x50, kept width 14.4 -> ratio 0.244 (wrong)
    A 10% tolerance separates those without flagging ordinary cross-sections. Partial
    sections (depth>0) legitimately shrink, so they are reported as not_checked.
    """
    src = precise_outline(source_view)
    sec = precise_outline(section)
    if not src or not sec:
        return {"checked": False,
                "reason": "outline too small for a deflated comparison"}
    src_w, src_h = src[2] - src[0], src[3] - src[1]
    sec_w, sec_h = sec[2] - sec[0], sec[3] - sec[1]
    info = {
        "checked": True,
        "method": "kept_dimension",
        "orientation": str(orientation),
        "source_size_mm": [round(src_w * 1000, 2), round(src_h * 1000, 2)],
        "section_size_mm": [round(sec_w * 1000, 2), round(sec_h * 1000, 2)],
        "tolerance": SECTION_KEEP_TOLERANCE,
    }
    if depth and float(depth) > 0:
        info["checked"] = False
        info["reason"] = ("partial section (depth>0) shrinks the view by design, so the "
                          "kept dimension is not comparable")
        return info
    if str(orientation) == "vertical":
        keep_src, keep_sec, which = src_h, sec_h, "height"
    else:
        keep_src, keep_sec, which = src_w, sec_w, "width"
    ratio = (keep_sec / keep_src) if keep_src else 0.0
    info["kept_dimension"] = which
    info["source_kept_mm"] = round(keep_src * 1000, 2)
    info["section_kept_mm"] = round(keep_sec * 1000, 2)
    info["ratio"] = round(ratio, 3)
    info["verdict"] = "ok" if abs(ratio - 1.0) <= SECTION_KEEP_TOLERANCE else "suspicious"
    return info

def _view_edge_count(view):
    """Count a view's visible edges without re-opening the document (P2-9).

    A DETAIL view whose boundary circle missed the parent geometry reports NO visible
    components, so this returns 0 — that is a genuinely empty view, not an API quirk.
    Measured 2026-10-31 on door.SLDPRT: boundary inside the parent outline -> 1 edge,
    boundary outside -> 0 edges. Used to reject empty detail views instead of returning
    ok:true for a view that shows nothing.
    """
    try:
        comps = list(view.GetVisibleComponents or ())
    except Exception:
        comps = []
    total = 0
    for comp in comps:
        try:
            total += len(list(view.GetVisibleEntities2(comp, 1) or ()))
        except Exception:
            continue
    return total

def _apply_detail_scale(doc, detail, scale):
    """Set a detail view's scale, clearing the parent-scale lock first. Returns readback."""
    applied = None
    try:
        doc.ActivateView(detail.Name)
    except Exception:
        pass
    for _ in range(4):
        try:
            detail.UseParentScale = False
        except Exception:
            pass
        try:
            detail.ScaleDecimal = float(scale)
        except Exception:
            pass
        try:
            applied = float(detail.ScaleDecimal)
        except Exception:
            applied = None
        if applied is not None and abs(applied - float(scale)) < 0.01:
            return applied
        try:
            doc.GraphicsRedraw2()
        except Exception:
            pass
    return applied

TOLERANCE_TYPES = {
    "none": 0, "basic": 1, "bilateral": 2, "limit": 3, "symmetric": 4,
    "min": 5, "max": 6, "fit": 7, "fit_with_tolerance": 8,
}

def _find_dimension(doc, value_mm, view_name=None):
    """Dimension on the sheet whose value matches `value_mm` (whole millimetres)."""
    import win32com.client as win

    target = round(float(value_mm))
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        if view_name and view.Name != view_name:
            continue
        try:
            raw = view.GetAnnotations
        except Exception:
            continue
        items = list(raw) if isinstance(raw, (tuple, list)) else ([raw] if raw else [])
        for item in items:
            if item is None:
                continue
            try:
                ann = win.Dispatch(item)
                if ann.GetType != 4:
                    continue
                dim = ann.GetSpecificAnnotation.GetDimension2(0)
                if dim is None:
                    continue
                if round(float(dim.SystemValue) * 1000) == target:
                    return view, dim
            except Exception:
                continue
    return None, None

ANNOTATION_KINDS = {
    2: "centerline", 4: "dimension", 5: "geometric_tolerance", 6: "note",
    7: "surface_finish", 8: "datum_tag", 9: "datum_target", 11: "balloon",
    12: "weld_symbol", 13: "center_mark", 14: "hole_callout",
}

TEMPLATE_TEXT = {
    "标记", "处数", "分区", "签名", "日期", "年 月 日", "阶段标记", "阶 段 标 记",
    "更改文件号", "设计", "审核", "校核", "批准", "工艺", "标准化", "重量", "比例",
    "共 张", "第 张", "图样名称", "图样代号", "材料", "数量", "更改单号",
}

DISPLAY_CODE_MAP = {
    "<MOD-DIAM>": "Ø",
    "<MOD-RAD>": "R",
    "<MOD-DEG>": "°",
    "<MOD-PM>": "±",
    "<MOD-SQ>": "□",
    "<MOD-SPH>": "S",
}

def _resolve_display_text(raw_text, value_mm=None):
    """Turn display format codes into what the sheet actually prints (P3-3 / P1-1).

    Returns the human-readable string. Callers keep the original as `raw_text` so a
    mismatch between code and rendering stays visible instead of being papered over.
    `<DIM>` is the value placeholder; when `value_mm` is known it is filled in, which is
    how "Ø20" is reconstructed from "<MOD-DIAM>" + SystemValue.
    """
    text = str(raw_text or "")
    if not text:
        return None
    upper = text.upper()
    for code, symbol in DISPLAY_CODE_MAP.items():
        if code in upper:
            start = 0
            while True:
                index = upper.find(code, start)
                if index < 0:
                    break
                text = text[:index] + symbol + text[index + len(code):]
                upper = text.upper()
                start = index + len(symbol)
    if "<DIM>" in upper and value_mm is not None:
        number = ("%.3f" % float(value_mm)).rstrip("0").rstrip(".")
        start = 0
        while True:
            upper = text.upper()
            index = upper.find("<DIM>")
            if index < 0:
                break
            text = text[:index] + number + text[index + 5:]
    return text or None

SELECT_TYPE_NAMES = {
    10: "vertex", 11: "edge", 12: "face", 13: "sketch_segment",
    14: "sketch_point", 20: "component", 2: "body",
}

def _view_kind(view):
    """What sort of view this is (IView.Type: 1=sheet, 2=section, 3=detail, 4=projected).

    Only used to label a circle's annotatability: report a view kind rather than making the
    caller infer it from the view name, which is a localised string.
    """
    try:
        view_type = int(view.Type)
    except Exception:
        return "unknown"
    # Measured on SW2024: a drawing SHEET is 1, a standard model view is 7, a section 2 and
    # a detail 3. (The first version of this map guessed and printed "type_7" for every view.)
    return {1: "sheet", 2: "section", 3: "detail", 4: "projected", 7: "model"}.get(
        view_type, "type_%d" % view_type)

def _callout_diameter_matches(text, expected_mm, tol=0.08):
    """True/False when the callout prints an explicit Ø; None when it prints none.

    A report showed callouts reading "Ø6 贯穿 / Ø99 ↧3" on a part with no Ø99 anywhere.
    The text PARSED and `IsHoleCallout` was true, so the tool signed off on a dimension that
    contradicts the geometry. Comparing the printed values with the measured edge diameter
    is the only cheap way to catch that — a wrong number is worse than a missing callout,
    because the drawing looks authoritative.
    """
    import re

    values = [float(v) for v in re.findall(r"[Ø⌀]\s*([0-9]+(?:\.[0-9]+)?)", str(text or ""))]
    if not values:
        return None
    expected = float(expected_mm)
    tol_abs = max(0.2, abs(expected) * float(tol))
    return any(abs(v - expected) <= tol_abs for v in values)

SURFACE_FINISH_TYPES = {"machining": 1, "prohibited": 2, "none": 0}

LEADER_STYLES = {"none": 0, "straight": 1, "bent": 2}

def _sf_render_page(doc, out_pdf, work_dir):
    """Export the CURRENT in-memory drawing to PDF and render page 1 to a PIL image.

    P2-11: used twice (before and after inserting the symbol) so a pixel diff can prove
    the symbol really rendered. Raises on any missing dependency so the caller can
    degrade to visibility.checked=false instead of failing the whole tool.
    """
    from PIL import Image

    doc.SaveAsSilent(out_pdf, False)
    if not os.path.exists(out_pdf):
        raise RuntimeError("SaveAsSilent wrote no PDF")
    rendered = render_drawing(out_pdf, work_dir, scale=2.0)
    if not rendered.get("ok"):
        raise RuntimeError(rendered.get("error") or "render failed")
    page = (rendered.get("pages") or [None])[0]
    if not page:
        raise RuntimeError("no page rendered")
    return Image.open(page["path"]).convert("L")

def _sf_visibility_check(doc, x, y, before_img, after_img):
    """Pixel-diff two rendered pages in a window around the requested sheet point.

    SolidWorks sheet origin is bottom-left while a rendered image origin is top-left,
    so y must be flipped. Measured 2026-10-31 on door.SLDPRT: a correctly placed
    surface finish changed 265 pixels in a +/-40 px window at (x, y).
    """
    from PIL import ImageChops

    sheet = list(doc.GetFirstView.GetOutline)
    sw_m, sh_m = sheet[2] - sheet[0], sheet[3] - sheet[1]
    if sw_m <= 0 or sh_m <= 0:
        raise RuntimeError("bad sheet outline")
    px = int(float(x) / sw_m * before_img.width)
    py = int((1.0 - float(y) / sh_m) * before_img.height)
    half = 50
    box = (max(0, px - half), max(0, py - half),
           min(before_img.width, px + half), min(before_img.height, py + half))
    diff = ImageChops.difference(before_img.crop(box), after_img.crop(box))
    changed = sum(1 for v in diff.getdata() if v > 25)
    return {
        "checked": True,
        "method": "pdf_render_diff",
        "pixel": [px, py],
        "window_px": [box[2] - box[0], box[3] - box[1]],
        "changed_pixels": changed,
        "visible": changed >= 20,
        "sheet_m": [round(sw_m, 4), round(sh_m, 4)],
    }

GTOL_SYMBOLS = {
    "position": "位置度",
    "concentricity": "同轴度",
    "perpendicularity": "垂直度",
    "parallelism": "平行度",
    "flatness": "平面度",
    "straightness": "直线度",
    "circularity": "圆度",
    "cylindricity": "圆柱度",
    "profile_line": "线轮廓度",
    "profile_surface": "面轮廓度",
    "angularity": "倾斜度",
    "symmetry": "对称度",
    "runout": "圆跳动",
    "total_runout": "全跳动",
    "none": "",
}

def _delete_gtol(doc, gtol):
    """Best-effort removal of a frame whose symbol could not be written (P2-10).

    Leaving a frame whose first cell is empty or a raw format code would be worse than
    failing: the caller would believe a geometric symbol is on the sheet when it is not.
    """
    from ...sw_session import call_method

    try:
        annotation = gtol.GetAnnotation
        if annotation is None:
            return False
        doc.ClearSelection2(True)
        if not annotation.Select(False):
            return False
        if doc.SelectionManager.GetSelectedObjectCount2(-1) < 1:
            return False
        call_method(doc, "EditDelete")
        call_method(doc, "EditRebuild3")
        doc.ClearSelection2(True)
        return True
    except Exception:
        return False

def _invoke_as_method(obj, name, *args):
    """Kept as a thin alias; the single implementation lives in sw_session.call_method."""
    from ...sw_session import call_method

    return call_method(obj, name, *args)


def render_drawing(pdf_path, output_dir=None, scale=2.0):
    """Render every PDF page to PNG so the result can actually be looked at."""
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.exists(pdf_path):
        raise FileNotFoundError("file not found: %s" % pdf_path)
    try:
        import pypdfium2 as pdfium
    except ImportError:
        import sys as _sys

        # A report noted the obvious-looking fix (install into the system Python) does
        # nothing, because the MCP host runs this server from its own venv. Naming the
        # exact interpreter removes the guesswork.
        return error(
            "missing_dependency", "pypdfium2 is required to render PDF pages",
            "Install it into the interpreter that RUNS this server, not the system "
            "Python: \"%s\" -m pip install pypdfium2 . Since 0.8.23 it is a hard "
            "dependency, so reinstalling this package also fixes it."
            % _sys.executable)
    out_dir = os.path.abspath(output_dir) if output_dir else os.path.dirname(pdf_path)
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    pages = []
    doc = pdfium.PdfDocument(pdf_path)
    for index in range(len(doc)):
        page = doc[index]
        image = page.render(scale=float(scale)).to_pil()
        target = os.path.join(out_dir, "%s_p%d.png" % (stem, index + 1))
        image.save(target)
        pages.append({"page": index + 1, "path": target, "width": image.width,
                      "height": image.height, "bytes": os.path.getsize(target)})
    return ok(pdf=pdf_path, page_count=len(pages), pages=pages)


__all__ = ['render_drawing', 'ANNOTATION_KINDS', 'DISPLAY_CODE_MAP', 'EDGE_CACHE_LIMIT', 'GTOL_SYMBOLS', 'LEADER_STYLES', 'OUTLINE_MARGIN_M', 'POSITIONS', 'SECTION_KEEP_TOLERANCE', 'SELECT_TYPE_NAMES', 'SUPPORTED_FORMATS', 'SURFACE_FINISH_TYPES', 'SW_DOC_DRAWING', 'SW_DOC_PART', 'TEMPLATE_TEXT', 'TITLE_SLOTS', 'TOLERANCE_TYPES', 'VIEW_ALIASES', '_EDGE_CACHE', '_activate', '_annotation_dimension_full_name', '_apply_detail_scale', '_call_method', '_callout_diameter_matches', '_dedupe_inserted_dimensions', '_delete_gtol', '_dispatch_specific', '_edge_model_length', '_filter_inserted_dimensions', '_find_dimension', '_find_template', '_free_spot_for_view', '_invoke_as_method', '_iter_views', '_model_box', '_model_view_name', '_nearest_circle_edge', '_nearest_edge_at_sheet_point', '_new_document', '_note_at', '_note_name', '_open_drawing', '_place_views', '_projection_axes', '_push_dimensions_outside', '_relayout', '_resolve_display_text', '_save_drawing', '_section_sanity', '_sf_render_page', '_sf_visibility_check', '_sheet_notes', '_standard_scale', '_template_lines', '_view_at_point', '_view_axis_map', '_view_edge_count', '_view_kind', '_view_outline_hint', 'ann_position', 'annotations', 'call_method', 'drawing_dimension_values', 'drawing_ok', 'error', 'glob', 'ok', 'open_document', 'os', 'precise_outline', 'template_safe_area', 'time', 'tool_guard', 'variant']
