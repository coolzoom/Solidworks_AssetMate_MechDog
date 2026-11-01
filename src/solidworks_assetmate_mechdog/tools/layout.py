"""Sheet-layout operations: place views, edit dimensions, add notes/balloons/BOM.

Every one of these was previously "not available" and had to be written as raw COM by the
caller, which cost hours: `IView.Position` silently mis-writes unless dispatched as
PROPERTYPUT with a VARIANT array, `InsertNote` needs GetAnnotation called as a method, and
deleting a dimension needs SelectByID2 by position plus EditDelete.

Kept in a separate module because these are layout concerns, not geometry.
"""
from __future__ import annotations

import os

from ..sw_session import byref_dispatch_null, error, get_sw, tool_guard
from .bom_native import default_template, populate_top_level_bom
from .drawing import _open_drawing, _iter_views, drawing_ok, _save_drawing, GTOL_SYMBOLS  # noqa: F401


def _prop_put(obj, name, value):
    """Set a property through IDispatch::Invoke with DISPATCH_PROPERTYPUT.

    Plain `obj.X = value` goes through pywin32's own marshalling, which silently writes
    the wrong thing for array-valued properties: setting Position to (0.12, 0.20) produced
    (0, 0.12). Explicit PROPERTYPUT with a VARIANT array is the only reliable form.
    """
    import pythoncom

    ole = obj._oleobj_
    dispid = ole.GetIDsOfNames(name)
    return ole.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, True, value)


# call_method lives in sw_session; this module used to carry its own copy, which risked
# drifting from the real one. Keep the single import and drop the duplicate.
from ..sw_session import call_method as _call_method


def _dodge_view_label(doc, view, gap_m=0.006, max_steps=12):
    """Move this view's dimensions off its own scale / label note (P1-9).

    Changing a section view's scale makes SolidWorks regenerate the "A-A" / "2.5 : 1" label
    at a position of its own choosing, and that position ignores dimensions already sitting
    there — a report had 2.80 and the A-A tag printed on top of each other the moment a view
    went 4:1 -> 2.5:1. `layout_dimensions` only treats the label as an obstacle when it
    happens to run afterwards, which is not when the damage occurs.

    The label is any non-empty note on the view: `IView.GetLabel` is not reachable under
    late binding, and the ratio is not always a separate note. Overlapping dimensions
    are pushed up one line height at a time until they clear every label box.
    """
    import win32com.client as win

    try:
        raws = list(view.GetAnnotations or ())
    except Exception:
        return None

    # A view's own notes ARE its label ("A", "A-A", "2.5 : 1") — the ratio is not always a
    # separate note, and `IView.GetLabel` is unreachable under late binding, so take every
    # non-empty note on this view as a label candidate.
    labels = []
    for raw in raws:
        try:
            annotation = win.Dispatch(raw)
            if int(annotation.GetType) != 6:
                continue
            note = win.Dispatch(annotation.GetSpecificAnnotation)
            text = str(note.GetText or "")
            position = [float(v) for v in annotation.GetPosition]
        except Exception:
            continue
        if not text.strip():
            continue
        labels.append({"text": text[:20], "position": position})
    if not labels:
        return None

    char_h = 0.0035
    boxes = []
    for label in labels:
        # `<VLLABEL>/<VLSCALEV>` is what a view label's GetText really returns (confirmed
        # 2026-11-01 — v0.8.24 concluded it was unreproducible because
        # list_drawing_annotations happened not to surface it). It renders as "A—A" over
        # "2.5 : 1", so 19 code characters are about 8 visible ones; using the raw length
        # made the keep-out 40 mm wide and shoved dimensions away for no reason.
        shown = label["text"]
        if "<VL" in shown.upper():
            shown = (shown.upper().replace("<VLLABEL>", "A—A")
                     .replace("<VLSCALEV>", "2.5 : 1"))
        half_w = max(char_h, len(shown) * char_h * 0.6)
        boxes.append([label["position"][0] - half_w, label["position"][1] - char_h,
                      label["position"][0] + half_w, label["position"][1] + char_h])

    def _hits(candidate):
        for box in boxes:
            if (box[0] - gap_m < candidate[2] and candidate[0] < box[2] + gap_m
                    and box[1] - gap_m < candidate[3] and candidate[1] < box[3] + gap_m):
                return True
        return False

    moved = 0
    for raw in raws:
        try:
            annotation = win.Dispatch(raw)
            if int(annotation.GetType) != 4:
                continue
            position = [float(v) for v in annotation.GetPosition]
        except Exception:
            continue
        half_text = 0.006
        candidate = [position[0] - half_text, position[1] - char_h * 0.5,
                     position[0] + half_text, position[1] + char_h * 0.5]
        if not _hits(candidate):
            continue
        for _step in range(max_steps):
            position[1] += char_h * 1.8
            candidate = [position[0] - half_text, position[1] - char_h * 0.5,
                         position[0] + half_text, position[1] + char_h * 0.5]
            if not _hits(candidate):
                break
        try:
            annotation.SetPosition(float(position[0]), float(position[1]), 0)
            moved += 1
        except Exception:
            continue
    return {"labels": [{"text": entry["text"],
                        "position_m": [round(v, 4) for v in entry["position"]]}
                       for entry in labels],
            "dimensions_moved": moved}


@tool_guard("set_view_failed")
def set_view(drawing_path, view_name, x_m=None, y_m=None, scale=None, save=True,
             break_alignment=False):
    """Move and/or rescale one view. Coordinates are sheet metres.

    Position is written with DISPATCH_PROPERTYPUT + VARIANT(VT_ARRAY|VT_R8); a plain
    assignment silently lands the view somewhere else.

    Section and detail views keep projection alignment with their parent, so the applied
    position may differ from the requested one. The result reports `requested`,
    `position_after_m` and `aligned_to_parent` separately — previously it just said ok:true
    while quietly moving y, and a reporter needed two rounds to notice.
    """
    import pythoncom
    from win32com.client import VARIANT

    doc = _open_drawing(drawing_path)
    target = None
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        if view.Name == view_name:
            target = view
            break
    if target is None:
        names = [v.Name for v in _iter_views(doc) if getattr(v, "Type", None) != 1]
        return error("view_not_found", "no view named %r" % view_name,
                     "Available: %s" % ", ".join(names))

    # P1-7: a section view is locked to its parent along one axis, so the requested
    # coordinate comes back "corrected" and multi-view layout gets stuck in a narrow
    # strip beside the parent. IView.RemoveAlignment is the UI's 解除对齐; note it only
    # clears THIS view's alignment, not what its children inherit.
    alignment_broken = False
    if break_alignment:
        try:
            _call_method(target, "RemoveAlignment")
            alignment_broken = True
        except Exception:
            alignment_broken = False

    requested = None
    if x_m is not None and y_m is not None:
        requested = [round(float(x_m), 4), round(float(y_m), 4)]
        _prop_put(target, "Position",
                  VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                          [float(x_m), float(y_m)]))
    if scale is not None:
        _prop_put(target, "ScaleDecimal", float(scale))
    try:
        _call_method(doc, "EditRebuild3")
    except Exception:
        pass
    try:
        doc.GraphicsRedraw2
    except Exception:
        pass

    # Read back AFTER the rebuild. Section/detail views keep projection alignment
    # with their parent, so SolidWorks may have moved the view away from the
    # requested position. P2-12: report requested vs applied separately instead of
    # only saying ok:true while quietly changing y.
    applied = {"view": view_name, "position_m": None, "scale": None}
    position_after = None
    try:
        position_after = [round(v, 4) for v in target.Position]
    except Exception:
        pass
    if x_m is not None and y_m is not None:
        applied["position_m"] = position_after
    if scale is not None:
        try:
            applied["scale"] = float(target.ScaleDecimal)
        except Exception:
            pass
    # P1-9: a scale change regenerates the view label at a SW-chosen spot; move this
    # view's dimensions off it before the caller sees the result.
    label_dodge = None
    if scale is not None:
        try:
            label_dodge = _dodge_view_label(doc, target)
            if label_dodge and label_dodge.get("dimensions_moved"):
                _call_method(doc, "EditRebuild3")
                try:
                    doc.GraphicsRedraw2
                except Exception:
                    pass
        except Exception:
            label_dodge = None
    aligned = bool(
        requested is not None and position_after is not None
        and (abs(position_after[0] - requested[0]) > 1e-4
             or abs(position_after[1] - requested[1]) > 1e-4)
    )
    if save:
        _save_drawing(doc)
    # P0-3: moving a PARENT view does not carry its section cut lines with it — the
    # cut line lives in the parent's sketch, and SolidWorks leaves it at the old sheet
    # coordinates, so the section loses its cutting basis while still looking valid.
    # The report paid for this twice (rebuild the section -> lose that view's
    # dimensions). We cannot move the sketch from here, so say it before saving.
    children = []
    # NOTE: this used to reuse `aligned`, which means "the applied position does NOT
    # match the request" — so a view that moved exactly where asked skipped the check
    # and reported no derived views. Ask instead: did the caller request a position?
    if requested is not None and position_after is not None:
        import win32com.client as _win

        for other in _iter_views(doc):
            if getattr(other, "Type", None) == 1:
                continue
            try:
                if str(other.Name) == str(view_name):
                    continue
                # `GetBaseView` is a PROPERTY that returns a raw PyIDispatch: reading
                # `.Name` straight off it fails and the exception was being swallowed,
                # so this check silently always reported "no derived views" (measured
                # 2026-10-31: a section view does answer 工程图视图1 once dispatched).
                base = other.GetBaseView
                if base is None:
                    continue
                if str(_win.Dispatch(base).Name) == str(view_name):
                    children.append(str(other.Name))
            except Exception:
                continue
    return drawing_ok(
        drawing=doc.GetTitle,
        requested=requested,
        position_after_m=position_after,
        aligned_to_parent=aligned,
        alignment_reason=("section/detail views keep projection alignment with the "
                          "parent view" if aligned else None),
        # P1-9: what the label did, and how many dimensions had to move off it.
        view_label=label_dodge,
        alignment_broken=alignment_broken,
        # P0-3: empty list means no derived views (detail/section) reference this one.
        derived_views=children,
        cut_lines_detached_warning=(
            "本视图有派生视图 %r：切割线/边界草图不会随父视图平移，"
            "请用 export_drawing + render_drawing 复核剖面位置，"
            "或重建剖面（重建会丢该视图上已插入的尺寸）。" % children
            if children else None),
        **applied,
    )


def place_views(drawing_path, placements, save=True):
    """Place several views at once: [{"view": name, "x": , "y": , "scale": }] (metres).

    Accepts `set_view`'s key names too (`view_name` / `x_m` / `y_m`), because passing one
    tool's vocabulary to the other used to drop every field and report "no view named
    None" (N5). An entry with neither key is reported as `missing_view_key` instead of
    being forwarded as None.
    """
    results = []
    for index, item in enumerate(placements or []):
        if not isinstance(item, dict):
            results.append({"ok": False, "index": index, "error": "placement_not_object"})
            continue
        name = item.get("view", item.get("view_name"))
        if name is None:
            results.append({
                "ok": False, "index": index, "error": "missing_view_key",
                "message": "这是第 %d 条 placement，缺视图名" % index,
                "hint": ("键名用 view 或 view_name 都可以；坐标用 x/y 或 x_m/y_m（单位：米）。"
                         "示例：[{\"view\": \"工程图视图1\", \"x\": 0.1, \"y\": 0.2}]"),
            })
            continue
        x = item.get("x", item.get("x_m"))
        y = item.get("y", item.get("y_m"))
        results.append(set_view(drawing_path, name,
                                x_m=x, y_m=y,
                                scale=item.get("scale"), save=False))
    if save:
        try:
            _save_drawing(_open_drawing(drawing_path))
        except Exception:
            pass
    ok_count = sum(1 for r in results if r.get("ok"))
    return drawing_ok(placed=ok_count, total=len(results), results=results)


@tool_guard("delete_dimensions_failed")
def delete_dimensions(drawing_path, keep_values=None, view_name=None,
                      keep_tolerance_mm=0.05, save=True):
    """Delete dimensions whose value is not in keep_values.

    SolidWorks has no bulk delete: each dimension is selected by position and removed with
    EditDelete. Measured 2026-10-29: 14 dimensions in ONE call took < 5 s. An earlier
    timing note overestimated this by ~100x and made callers budget ten minutes for a
    job that takes seconds. This batches the loop and reports what was removed, so a
    49-dimension sheet can be reduced to the planned shortlist without hand-written COM.
    """
    import win32com.client as win

    doc = _open_drawing(drawing_path)
    keep = sorted(float(v) for v in (keep_values or []))
    removed, kept, failed = [], [], []
    rows = []
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        if view_name and view.Name != view_name:
            continue
        try:
            items = list(view.GetAnnotations or ())
        except Exception:
            continue
        for item in items:
            try:
                ann = win.Dispatch(item)
                if int(ann.GetType) != 4:
                    continue
                specific = ann.GetSpecificAnnotation
                dim = specific.GetDimension2(0)
                if dim is None:
                    continue
                value = round(float(dim.SystemValue) * 1000, 3)
                rows.append((view, ann, value))
            except Exception:
                continue

    for view, ann, value in rows:
        if keep and any(abs(value - target) <= keep_tolerance_mm for target in keep):
            kept.append({"value_mm": value, "name": str(ann.GetName)})
            continue
        try:
            position = list(ann.GetPosition)
            name = str(ann.GetName)
            doc.ClearSelection2(True)
            try:
                doc.ActivateView(view.Name)
            except Exception:
                pass
            # SelectByID2 with an empty Name returns False for dimensions; the annotation
            # name is what actually selects it. Position is kept as a fallback for builds
            # where the name differs from the selection name.
            selected = False
            # The Callout argument must be a byref null VARIANT; passing None raises
            # 'type mismatch' and SelectByID2 returns False (quirk section 3).
            callout = byref_dispatch_null()
            for candidate, px, py, pz in ((name, 0.0, 0.0, 0.0),
                                          (name + "@" + view.Name, 0.0, 0.0, 0.0),
                                          ("", position[0], position[1], position[2])):
                try:
                    if doc.Extension.SelectByID2(candidate, "DIMENSION", px, py, pz,
                                                 False, 0, callout, 0):
                        selected = True
                        break
                except Exception:
                    continue
            if not selected:
                failed.append({"value_mm": value, "reason": "select failed",
                               "name": name})
                continue
            doc.EditDelete()
            removed.append({"value_mm": value, "name": name})
        except Exception as exc:
            failed.append({"value_mm": value, "reason": str(exc)[:80]})

    doc.ClearSelection2(True)
    try:
        _call_method(doc, "EditRebuild3")
    except Exception:
        pass
    try:
        doc.GraphicsRedraw2
    except Exception:
        pass
    if save:
        _save_drawing(doc)
    return drawing_ok(
        drawing=doc.GetTitle,
        removed=len(removed), kept=len(kept), failed=len(failed),
        removed_items=removed[:40], kept_values=[k["value_mm"] for k in kept][:40],
        failed_items=failed[:10],
        note="按值保留（默认容差 0.05 mm）。SolidWorks 无批量删除接口；实测 14 条一次调用 < 5 s。",
    )


def _note_bbox(text, x, y, height_m):
    """Conservative bounding box for a note anchored at (x, y) with its TOP-LEFT there.

    Chinese glyphs are roughly square, Latin roughly 0.55 em; 0.75 is the safe middle and
    over-estimating only costs a slightly larger keep-out, never a missed collision.
    """
    lines = str(text).splitlines() or [""]
    line_h = height_m * 1.6
    width = max(1, max(len(line) for line in lines)) * height_m * 0.75
    return [float(x), float(y) - line_h * len(lines), float(x) + width, float(y)]


def _boxes_overlap(a, b, pad=0.0):
    return (a[0] - pad < b[2] and b[0] - pad < a[2]
            and a[1] - pad < b[3] and b[1] - pad < a[3])


@tool_guard("note_failed")
def add_note(drawing_path, text, x_m, y_m, height_mm=None, bold=False, save=True,
             avoid_collisions=False):
    """Add a drawing note (技术要求 lines). x/y are sheet metres.

    `InsertNote` returns an INote; positioning it needs GetAnnotation invoked as a METHOD —
    pywin32 otherwise dispatches it as PROPERTYGET and the call fails.

    `avoid_collisions=True` (P3-2) walks the note DOWN until its estimated box clears every
    model view. A report showed the first line of a 技术要求 block being overprinted by a
    section view's arrow, and `move_annotation` was the only cure — which meant a rebuild if
    it failed. The note box is estimated from line count and height; views are taken from
    `precise_outline`, so the 11.9 mm phantom margin does not push notes away for no reason.
    """
    doc = _open_drawing(drawing_path)
    doc.ClearSelection2(True)

    collisions_avoided = 0
    position_requested = [round(float(x_m), 4), round(float(y_m), 4)]
    if avoid_collisions:
        from .drawing import precise_outline

        height_m = (float(height_mm) / 1000.0) if height_mm else 0.0035
        obstacles = []
        for view in _iter_views(doc):
            if getattr(view, "Type", None) == 1:
                continue
            box = precise_outline(view)
            if box:
                obstacles.append(box)
        while collisions_avoided < 24:
            box = _note_bbox(text, x_m, y_m, height_m)
            if not any(_boxes_overlap(box, obstacle, pad=0.002)
                       for obstacle in obstacles):
                break
            step = (box[3] - box[1]) + 0.004
            if float(y_m) - step < 0.02:
                break
            y_m = float(y_m) - step
            collisions_avoided += 1

    created = doc.InsertNote(str(text))
    if created is None:
        return error("note_failed", "InsertNote returned null",
                     "Some templates refuse notes while a view is active; activate the sheet first.")
    note = created
    try:
        annotation = _call_method(note, "GetAnnotation")
        import win32com.client as win

        annotation = win.Dispatch(annotation)
        _call_method(annotation, "SetPosition2", float(x_m), float(y_m), 0.0)
    except Exception:
        pass
    if height_mm:
        try:
            note.SetHeight(float(height_mm) / 1000.0)
        except Exception:
            try:
                _call_method(note, "SetHeight", float(height_mm) / 1000.0)
            except Exception:
                pass
    if bold:
        try:
            note.SetBold(True)
        except Exception:
            pass
    try:
        _call_method(doc, "EditRebuild3")
    except Exception:
        pass
    try:
        doc.GraphicsRedraw2
    except Exception:
        pass
    if save:
        _save_drawing(doc)
    name = None
    try:
        name = str(_call_method(note, "GetAnnotation").GetName)
    except Exception:
        pass
    # P1-5: report the box the note actually occupies. A multi-line note shifted by
    # 49 mm looked identical to one shifted by 4 mm in the old receipt — only
    # `adjusted: true` — so the caller could not tell whether it now sat on a view or
    # the title block.
    final_bbox = _note_bbox(str(text), float(x_m), float(y_m),
                            (float(height_mm) / 1000.0) if height_mm else 0.0035)
    return drawing_ok(drawing=doc.GetTitle, note=name, text=str(text),
                      position_m=[round(float(x_m), 4), round(float(y_m), 4)],
                      position_requested_m=position_requested,
                      adjusted=bool(collisions_avoided),
                      collisions_avoided=collisions_avoided,
                      bbox_final=[round(v, 4) for v in final_bbox],
                      shifted_m=([round(float(x_m) - position_requested[0], 4),
                                  round(float(y_m) - position_requested[1], 4)]
                                 if collisions_avoided else [0.0, 0.0]))


@tool_guard("balloon_failed")
def add_balloon(drawing_path, view_name, x_m, y_m, text=None, item_number=None,
                leader_x_m=None, leader_y_m=None, save=True):
    """Add a BOM balloon to a view.

    Uses InsertBOMBalloon2 when an item number is known (attaches to the component under
    the point); otherwise inserts a plain balloon with the given text.
    """
    doc = _open_drawing(drawing_path)
    target = None
    for view in _iter_views(doc):
        if getattr(view, "Type", None) != 1 and view.Name == view_name:
            target = view
            break
    if target is None:
        return error("view_not_found", "no view named %r" % view_name)
    doc.ActivateView(target.Name)
    doc.ClearSelection2(True)
    # Both entry points go through call_method: direct attribute access on this binding
    # raises for some of them, and one user report on SW2026 shows `InsertBalloon`
    # resolving to nothing at all (DISP_E_UNKNOWNNAME) — i.e. renamed or removed in
    # that version. Reported honestly rather than hidden behind a bare except.
    balloon_error = None
    ballon = None
    if item_number is not None:
        for attempt in (5, 4, 3):
            try:
                args = [item_number] + [0] * (attempt - 1)
                if attempt == 5:
                    args[-1] = 0.0
                ballon = _call_method(doc, "InsertBOMBalloon2", *args)
                if ballon is not None:
                    break
            except Exception as exc:
                balloon_error = "%s(%d args): %s" % ("InsertBOMBalloon2", attempt,
                                                     str(exc)[:60])
    if ballon is None:
        # `InsertBalloon` does not exist on this binding at all (GetIDsOfNames: False on
        # SW2024 SP4, DISP_E_UNKNOWNNAME on SW2026 per a user report). `InsertBOMBalloon`
        # and `InsertBOMBalloon2` do. Try the older entry point before giving up.
        for api in ("InsertBOMBalloon", "InsertBalloon"):
            try:
                ballon = _call_method(doc, api, 0, 0, 0,
                                      float(x_m), float(y_m),
                                      float(leader_x_m or x_m),
                                      float(leader_y_m or y_m))
                if ballon is not None:
                    balloon_error = None
                    break
            except Exception as exc:
                balloon_error = "%s; %s: %s" % (balloon_error or "no callout",
                                                api, str(exc)[:60])
    if ballon is None:
        return error("balloon_failed",
                     "no balloon could be created",
                     "%s | Verified on SW2024 SP4 only; a SW2026 report shows "
                     "InsertBalloon as DISP_E_UNKNOWNNAME (renamed or removed there)."
                     % (balloon_error or "both entry points returned null"))
    applied_text = None
    if text:
        try:
            ballon.SetText(str(text))
            applied_text = str(text)
        except Exception:
            try:
                _call_method(ballon, "SetText", str(text))
                applied_text = str(text)
            except Exception:
                pass
    try:
        _call_method(doc, "EditRebuild3")
    except Exception:
        pass
    try:
        doc.GraphicsRedraw2
    except Exception:
        pass
    if save:
        _save_drawing(doc)
    return drawing_ok(drawing=doc.GetTitle, view=target.Name, text=applied_text,
                      item_number=item_number)


@tool_guard("sheet_line_failed")
def add_sheet_line(drawing_path, x1_m, y1_m, x2_m, y2_m, save=True):
    """Draw a line in TRUE sheet space.

    `SketchManager.CreateLine` on a drawing interprets coordinates as model space and
    scales them by the view scale: asking for a line at (236, 62) produced a fan of
    diagonals across the sheet. This activates the sheet view first, then converts the
    requested sheet point through the sheet's own transform, which is the only way to get
    the line where it was asked for.
    """
    doc = _open_drawing(drawing_path)
    sheet = doc.GetFirstView
    try:
        doc.ActivateSheet(doc.GetCurrentSheet().GetName)
    except Exception:
        pass
    doc.ClearSelection2(True)
    try:
        doc.InsertSketch()
    except Exception:
        pass
    sketch = doc.SketchManager
    # Sheet space -> sketch space: a drawing sketch is placed at the sheet origin, so
    # sheet metres are already the right unit. Passing millimetres is what produced the
    # scaled fan of lines.
    sketch.CreateLine(float(x1_m), float(y1_m), 0.0,
                      float(x2_m), float(y2_m), 0.0)
    sketch.InsertSketch(True)
    try:
        _call_method(doc, "EditRebuild3")
    except Exception:
        pass
    try:
        doc.GraphicsRedraw2
    except Exception:
        pass
    if save:
        _save_drawing(doc)
    return drawing_ok(drawing=doc.GetTitle,
                      line_m=[[round(float(x1_m), 4), round(float(y1_m), 4)],
                              [round(float(x2_m), 4), round(float(y2_m), 4)]],
                      note="坐标为图纸空间米；图纸草图不接受毫米。")


@tool_guard("bom_table_failed")
def add_bom_table(drawing_path, view_name=None, configuration="Default",
                  template=None, save=True):
    """Insert a POPULATED native TopLevelOnly BOM into an assembly drawing.

    Verified working 2026-10-29 on SW2024 SP4: 10 rows x 4 columns with a quantity
    column, stable across save/reopen.

    Why the two-step sequence is required (and why everything simpler failed):
      * `InsertBomTable4(BomType=TopLevelOnly=2)` creates ONLY the header here, and
        `BomFeature.SetConfigurations` does not repair it (every marshalling form returns
        True or False but the column count never moves off 3).
      * `InsertBomTable4(BomType=Indented=3)` with a valid configuration DOES populate
        rows (120 rows, 4 columns) — proving assembly, template and model reference are all
        fine; only TopLevelOnly's configuration linkage is broken.
      * Converting the populated table with `BomFeature.TableType = 2` yields the 9 module
        rows plus ONE duplicate quantity column (5 columns); hiding that last column gives
        the target shape.

    Do NOT call SetConfigurations in this flow — it clears the table back to its header.
    Do NOT try to add a quantity column via SetColumnType — SOLIDWORKS forbids promoting a
    normal column to Quantity (SetColumnType2(..., 203, ...) returns False).
    """
    from ..sw_session import error, get_sw

    doc = _open_drawing(drawing_path)
    sw = get_sw()
    template = template or default_template(sw)
    if not os.path.exists(template):
        return error("template_missing", "BOM template not found: %s" % template)

    views = [v for v in _iter_views(doc) if getattr(v, "Type", None) != 1]
    if not views:
        return error("no_views", "drawing has no model views")
    target = views[0]
    if view_name:
        target = next((v for v in views if v.Name == view_name), None)
        if target is None:
            return error("view_not_found", "no view named %r" % view_name)

    configuration = configuration or getattr(target, "ReferencedConfiguration", "Default")

    result = populate_top_level_bom(
        doc, target, template,
        configuration=configuration or "Default",
    )
    if save:
        _save_drawing(doc)

    return drawing_ok(
        drawing=doc.GetTitle,
        created=True,
        native_bom=True,
        populated=True,
        rows=result["rows"],
        columns=result["columns"],
        data_rows=result["data_rows"],
        headers=result["headers"],
        table_type=result["table_type"],
        referenced_model=result["referenced_model"],
        sample_data=result["data"][:3],
        saved=bool(save),
    )




# --- delete / move: the missing half of the API surface -------------------------------
#
# Every create-only design funnels into the same dead end: one wrong parameter means
# rebuilding the whole sheet, because nothing on it can be removed. A report measured this
# precisely — four full rebuilds for one drawing, three of them solely because the sheet
# carried leftovers (`delete_view` / `delete_annotation` absent). `name` was already exposed
# by list_drawing_annotations, so deletion was the cheapest possible fix.

# SelectByID2 does not accept a generic "ANNOTATION" — notes, GT frames and surface symbols
# each have their own type string, and the reported `name` ("细节项目256") is not resolvable
# by ID at all. Selecting by the annotation's own position is what actually works.
DELETABLE_SELECT_TYPES = ("NOTE", "GTOL", "SFSYMBOL", "DIMENSION", "ANNOTATION")


def _annotation_total(doc):
    """Annotation count summed over EVERY view.

    The v0.8.14 report's P0 was partly a measurement bug: `doc.GetFirstView.GetAnnotations`
    is the SHEET, and its count does not change when a DIMENSION is deleted, because
    dimensions belong to the model views. Comparing only the sheet made a successful
    delete answer `delete_failed`. Measured 2026-10-31 on a real sheet: the sheet stayed at
    83 while 工程图视图1 went 8 -> 7 and the total went 122 -> 121.
    """
    total = 0
    counted = False
    for view in _iter_views(doc):
        try:
            total += len(list(view.GetAnnotations or ()))
            counted = True
        except Exception:
            continue
    return total if counted else None


def _find_annotation_by_name(doc, name, position_m=None):
    """Return (annotation, position_m) for an annotation with this reported name.

    Iterates EVERY view, not just the first one. `doc.GetFirstView` is the sheet, and its
    GetAnnotations returns sheet furniture only — but in practice every annotation, including
    model-view ones, hangs off the sheet view too (the name reported by
    list_drawing_annotations, e.g. "细节项目256", does appear here). Sweeping all views is
    still the correct approach: it costs one extra loop and covers the case where an
    annotation belongs to a model view.

    `position_m` disambiguates repeated names: SolidWorks names dimensions D1/D2/D3 per
    sketch, so one sheet can hold eight annotations called "D1". Passing the position that
    list_drawing_annotations reported selects that specific one instead of the first hit.
    """
    import win32com.client
    try:
        candidates = []
        for view in _iter_views(doc):
            try:
                # _call_method is this module's alias for sw_session.call_method. Writing
                # `call_method` here raised NameError, which the surrounding except turned
                # into a silent "no candidates" — the finder then reported not_found for
                # names the listing tool had just printed.
                candidates.extend(list(_call_method(view, "GetAnnotations") or ()))
            except Exception:
                continue
        for raw in candidates:
            annotation = win32com.client.Dispatch(raw)
            try:
                if str(annotation.GetName) != name:
                    continue
                position = [float(v) for v in annotation.GetPosition]
                if position_m is not None:
                    try:
                        if not all(abs(position[i] - float(position_m[i])) <= 1e-3
                                   for i in range(2)):
                            continue
                    except Exception:
                        continue
                return annotation, position
            except Exception:
                continue
    except Exception:
        pass
    return None, None


def _select_by_name(doc, name, select_type, byref_dispatch_null):
    """Select one object by its reported name. Returns True when the selection took."""
    try:
        ok_selected = doc.Extension.SelectByID2(name, select_type, 0, 0, 0,
                                               False, 0, byref_dispatch_null(), 0)
    except Exception:
        return False
    if not ok_selected:
        return False
    try:
        return doc.SelectionManager.GetSelectedObjectCount2(-1) > 0
    except Exception:
        return ok_selected


@tool_guard("delete_view_failed")
def delete_view(drawing_path, view_name, save=True):
    """Delete a drawing view (wrong section, empty detail, stray projection).

    Without this, a single wrong `add_section_view` forces rebuilding the entire sheet.
    """
    from ..sw_session import byref_dispatch_null, error, get_sw

    doc = _open_drawing(drawing_path)
    before = [str(v.Name) for v in _iter_views(doc)
              if getattr(v, "Type", None) != 1]
    if view_name not in before:
        return error("view_not_found",
                     "no model view named %r" % view_name,
                     "available: %s" % ", ".join(before))
    doc.ClearSelection2(True)
    if not _select_by_name(doc, view_name, "DRAWINGVIEW", byref_dispatch_null):
        return error("select_failed",
                     "SelectByID2 could not select %r as DRAWINGVIEW" % view_name)
    _call_method(doc, "EditDelete")
    try:
        _call_method(doc, "EditRebuild3")
    except Exception:
        pass
    after = [str(v.Name) for v in _iter_views(doc)
             if getattr(v, "Type", None) != 1]
    if len(after) >= len(before):
        return error("delete_failed",
                     "EditDelete did not remove the view",
                     "before=%s after=%s" % (before, after),
                     before=before, after=after)
    if save:
        _save_drawing(doc)
    return drawing_ok(deleted=view_name, views_before=before, views_after=after,
                      saved=bool(save))


@tool_guard("delete_annotation_failed")
def delete_annotation(drawing_path, name, save=True, position_m=None):
    """Delete one annotation by name (broken hole callout, unwanted GT, stray note).

    Uses `IAnnotation.Select(False)` then EditDelete. Measured on a real sheet:
    Select3 returned False and selected nothing; Select(False) selected exactly one object
    and EditDelete removed it. SelectByID2 with the reported name never worked either —
    `<unknown>`-style names like "细节项目34" are not resolvable ids.

    All four kinds were re-verified on 2026-10-31 against a report that claimed every type
    failed: note, geometric tolerance, surface finish and dimension all delete. Dimensions
    are the reason the success test now counts EVERY view instead of the sheet: deleting a
    dimension leaves the sheet's own annotation count untouched, so the old check answered
    `delete_failed` for a delete that had worked (sheet stayed 83, 工程图视图1 went 8 -> 7).

    `position_m` selects among repeated names — SolidWorks reuses D1/D2/D3 per sketch, so
    one sheet can hold eight annotations called "D1". Pass the `position_m` that
    `list_drawing_annotations` reported for the one you mean.
    """
    from ..sw_session import error

    doc = _open_drawing(drawing_path)
    doc.ClearSelection2(True)
    annotation, position = _find_annotation_by_name(doc, name, position_m)
    if annotation is None:
        return error("annotation_not_found",
                     "no annotation named %r%s on this sheet"
                     % (name, " at %r" % (position_m,) if position_m else ""),
                     "list names first with list_drawing_annotations")
    try:
        selected = annotation.Select(False)
    except Exception as exc:
        return error("select_failed", "IAnnotation.Select raised: %s" % str(exc)[:80])
    try:
        count = doc.SelectionManager.GetSelectedObjectCount2(-1)
    except Exception:
        count = None
    if not selected or count == 0:
        return error("select_failed",
                     "Select(False) did not select %r" % name,
                     "selected count=%r" % count)

    before = _annotation_total(doc)
    _call_method(doc, "EditDelete")
    try:
        _call_method(doc, "EditRebuild3")
    except Exception:
        pass
    after = _annotation_total(doc)
    doc.ClearSelection2(True)
    removed = bool(before is not None and after is not None and after < before)
    if not removed and position is not None:
        # Fallback: a view can hand back a cached annotation list, leaving the total
        # unchanged. Re-find the exact object instead of trusting the count.
        again, again_position = _find_annotation_by_name(doc, name, position)
        if again is None or again_position is None:
            removed = True
        else:
            removed = not all(abs(again_position[i] - position[i]) <= 1e-3
                              for i in range(2))
    if not removed and before is not None and after is not None:
        return error("delete_failed",
                     "EditDelete did not remove %r" % name,
                     "annotation count %s -> %s (all views)" % (before, after))
    if save:
        _save_drawing(doc)
    return drawing_ok(deleted=name, removed=removed,
                      annotations_before=before, annotations_after=after,
                      counted_scope="all_views",
                      position_m=[round(v, 4) for v in position] if position else None,
                      saved=bool(save))


@tool_guard("move_annotation_failed")
def move_annotation(drawing_path, name, x_m, y_m, save=True, position_m=None):
    """Move an annotation or dimension to a sheet position (metres).

    Dimensions pile up on the centre line when auto_dimension runs on a section view; there
    was no way to pull them apart, which made the output unreadable.

    `position_m` selects among repeated names (SolidWorks reuses D1/D2/D3 per sketch) —
    pass the position `list_drawing_annotations` reported. Verified on 2026-10-31 for a
    note: `move_annotation` returned the new position, which a v0.8.14 report said was
    impossible.
    """
    from ..sw_session import byref_dispatch_null, error

    doc = _open_drawing(drawing_path)
    doc.ClearSelection2(True)
    annotation, before = _find_annotation_by_name(doc, name, position_m)
    if annotation is None:
        return error("annotation_not_found",
                     "no annotation named %r on this sheet" % name,
                     "list names first with list_drawing_annotations")
    try:
        before = [round(float(v), 4) for v in annotation.GetPosition]
    except Exception:
        before = None
    try:
        annotation.SetPosition(float(x_m), float(y_m), 0)
    except Exception as exc:
        return error("move_failed", "SetPosition raised: %s" % str(exc)[:80],
                     "section/detail views keep projection alignment and refuse some moves")
    try:
        _call_method(doc, "EditRebuild3")
    except Exception:
        pass
    after = None
    try:
        after = [round(float(v), 4) for v in annotation.GetPosition]
    except Exception:
        pass
    if save:
        _save_drawing(doc)
    return drawing_ok(moved=name, position_before_m=before, position_after_m=after,
                      requested=[round(float(x_m), 4), round(float(y_m), 4)],
                      applied=after != before, saved=bool(save))


@tool_guard("set_dimension_position_failed")
def set_dimension_position(drawing_path, name, x_m, y_m, save=True, position_m=None):
    """Move ONE dimension to a sheet position (metres), by its reported name.

    P0-2: `list_drawing_annotations` already gives every dimension a stable `name`; this is
    the single-dimension counterpart to `layout_dimensions`' bulk auto-placement. Unlike
    `move_annotation` it verifies the target really is a dimension (GetType 4), reports the
    dimension value, and refuses notes / GTOL / surface-finish symbols so a typo cannot
    silently move the wrong kind of annotation.

    `position_m` disambiguates repeated names — SolidWorks reuses D1/D2/D3 per sketch, and a
    40-dimension sheet can hold 19 of them called "D1". Pass the `position_m` that
    `list_drawing_annotations` reported. Without it the first match wins, which is exactly
    the wrong-dimension trap this tool otherwise exists to avoid.
    """
    import win32com.client as win

    doc = _open_drawing(drawing_path)
    doc.ClearSelection2(True)
    annotation, matched = _find_annotation_by_name(doc, name, position_m)
    if annotation is None:
        return error("annotation_not_found",
                     "no annotation named %r%s on this sheet"
                     % (name, " at %r" % (position_m,) if position_m else ""),
                     "List names first with list_drawing_annotations; a repeated name "
                     "needs position_m to tell the copies apart.")
    try:
        kind = int(annotation.GetType)
    except Exception:
        kind = None
    if kind != 4:
        return error("not_a_dimension",
                     "%r is annotation type %r, not a dimension (4)" % (name, kind),
                     "Use move_annotation for notes / GTOL / surface-finish symbols.")
    value_mm = None
    try:
        specific = win.Dispatch(annotation.GetSpecificAnnotation)
        dim = win.Dispatch(specific.GetDimension2(0))
        value_mm = round(float(dim.SystemValue) * 1000.0, 3)
    except Exception:
        pass
    try:
        before = [round(float(v), 4) for v in annotation.GetPosition]
    except Exception:
        before = None
    try:
        annotation.SetPosition(float(x_m), float(y_m), 0)
    except Exception as exc:
        return error("move_failed", "SetPosition raised: %s" % str(exc)[:80],
                     "Section/detail views keep projection alignment and refuse some moves.")
    try:
        _call_method(doc, "EditRebuild3")
    except Exception:
        pass
    try:
        after = [round(float(v), 4) for v in annotation.GetPosition]
    except Exception:
        after = None
    if save:
        _save_drawing(doc)
    return drawing_ok(
        dimension=name, dimension_value_mm=value_mm,
        position_before_m=before, position_after_m=after,
        requested=[round(float(x_m), 4), round(float(y_m), 4)],
        matched_position_m=([round(v, 4) for v in matched] if matched else None),
        disambiguated_by_position=bool(position_m),
        applied=after != before, saved=bool(save))
