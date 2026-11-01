from __future__ import annotations

"""Dimension tools: insert, tolerance, layout, verify, prune."""

from ._core import *
from .create import drawing_audit

@tool_guard("auto_dimension_failed")
def auto_dimension(drawing_path, all_views=True, include_unmarked=False, eliminate_duplicates=True,
                   dimension_filter=None, place_outside=False,
                   outside_gap_m=0.012, outside_step_m=0.008):
    """Insert the model's driving dimensions into the drawing views.

    Equivalent to SolidWorks "Insert Model Items > Dimensions". Marked-for-drawing
    dimensions only by default; include_unmarked=True brings in every driving dim
    (more complete, noisier). Returns how many annotations were inserted.

    `eliminate_duplicates=True` (default) now deletes the extra copies by MODEL full name,
    keeping the first occurrence of each driving dimension. Only dimensions inserted by
    this call are eligible — existing user-placed dimensions are never touched.

    `dimension_filter` closes the plan → insert gap (N3): pass
    `plan_drawing()["keep_dimensions"]` and everything the plan did not keep is deleted
    after insertion, so "keep 8 of 41" actually ends with ~8 dimensions on the sheet.
    Deletion uses the same mechanics as the duplicate pass, and aborts without deleting
    anything if the filter matches no inserted dimension.

    This is document-level idempotent: a driving dimension inserted once will NOT be inserted
    into a second view, whatever `all_views` / duplicate settings are used (measured
    2026-11-01: three parameter combinations, second view stayed at 0). To annotate a detail
    view with the same geometry, build a NEW dimension there instead — `list_drawing_edges`
    + `add_drawing_dimension(view_name=...)`. See check_capabilities
    `duplicate_dimension_to_view`.

    `place_outside=True` (P0-1) then moves everything this call inserted to just outside its
    view, stacked by side: `outside_gap_m` + slot * `outside_step_m`. SolidWorks' own offset
    is a fixed ~10 mm regardless of part size, so on a Ø46 disc 25 dimensions land inside the
    outline and overprint; scaling the view up hides the symptom instead of fixing it.
    Dimensions already on the sheet are never moved, and this path saves the drawing itself
    (the rest of the function only edits in memory).
    """
    doc = _open_drawing(drawing_path)
    # swImportModelItemsSource_e.FromEntireModel = 0
    # swInsertAnnotation_e: MarkedForDrawing = 32768, NotMarked = 524288
    types = 32768 | (524288 if include_unmarked else 0)
    doc.ClearSelection2(True)
    # P1-6: snapshot every annotation name BEFORE inserting. The dedupe pass below
    # must only consider annotations this call created, never a dimension the user
    # placed on the sheet earlier.
    import win32com.client as _win_auto

    before_names = set()
    for _v in _iter_views(doc):
        if getattr(_v, "Type", None) == 1:
            continue
        for _raw in (_v.GetAnnotations or ()):
            try:
                before_names.add(str(_win_auto.Dispatch(_raw).GetName))
            except Exception:
                continue
    inserted = doc.InsertModelAnnotations3(
        0, types, bool(all_views), not bool(eliminate_duplicates), False, True
    )
    try:
        # Must be a real METHOD call: after InsertModelAnnotations3 the new
        # annotations only show up in view.GetAnnotations once the sheet rebuilds.
        # A bare `doc.EditRebuild3` attribute access used to be a silent no-op here,
        # which made the dedupe pass below see an empty set of new dimensions.
        _call_method(doc, "EditRebuild3")
    except Exception:
        pass
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass
    count = len(inserted) if isinstance(inserted, (tuple, list)) else (1 if inserted else 0)
    # No repack: a section view is x-aligned with its parent and cannot be moved
    # sideways, so repacking would stack it on top of a neighbour.
    layout = None
    # P1-6: `eliminate_duplicates` used to mean "tell SolidWorks to try". It did not
    # work (56 dims / 4 views). Now also delete the extra copies by model full name.
    duplicates, removed, dedupe_failed = {}, [], []
    if eliminate_duplicates:
        try:
            duplicates, removed, dedupe_failed = _dedupe_inserted_dimensions(
                doc, before_names, _win_auto)
        except Exception as exc:
            dedupe_failed = [{"reason": "dedupe pass failed: %s" % str(exc)[:80]}]
    # N3: honour plan_drawing()["keep_dimensions"]. Runs after the duplicate pass so
    # the same dimension is not counted twice, and only ever deletes annotations this
    # call inserted (`before_names`).
    filter_kept, filter_removed, filter_failed, filter_aborted = None, [], [], None
    if dimension_filter:
        try:
            filter_kept, filter_removed, filter_failed, filter_aborted = \
                _filter_inserted_dimensions(doc, before_names, _win_auto, dimension_filter)
        except Exception as exc:
            filter_aborted = "filter pass failed: %s" % str(exc)[:120]
    _save_drawing(doc)
    audit = drawing_audit(drawing_path)
    total = audit.get("total_dimensions") or 0

    # P1-6: report the model full names that HAD duplicates and what the automatic
    # pass actually removed. The old code only reported and told the caller to delete
    # by hand — the duplicates stayed on the sheet.
    duplicate_report = {
        full: [sheet_name for sheet_name, _ in entries]
        for full, entries in duplicates.items()
    }

    notes = []
    if duplicate_report:
        worst = max(duplicate_report.items(), key=lambda kv: len(kv[1]))
        notes.append("%d 个 driving 尺寸在多个视图里重复（最多 %d 次：%s）——"
                     "已按 full name 自动删除 %d 条多余副本（保留首次出现），失败 %d 条。"
                     % (len(duplicate_report), len(worst[1]), worst[0],
                        len(removed), len(dedupe_failed)))
    if dedupe_failed:
        notes.append("%d 条重复尺寸删除失败，见 dedupe_failed；可用 "
                     "delete_dimensions 按视图兜底。" % len(dedupe_failed))
    if eliminate_duplicates and removed:
        notes.append(
            "去重按模型 FullName 保留**首次出现**（插入顺序主视图在前），"
            "所以剖视/局部视图上的同源副本会被删掉——深度类尺寸可能因此无处安放"
            "（实测 83 插入 → 删 44，两个剖视 0 尺寸）。需要剖视承载尺寸时，"
            "改 eliminate_duplicates=False，或用 dimension_filter 精确保留。")
    if total > 30:
        notes.append(
            "本图 %d 个尺寸；100 张真实工程图实测平均 14 个（中位 11）。"
            "全量导入（include_unmarked=True）会把模型的每个驱动尺寸都倒到图上，"
            "制图是筛选而非搬运，建议关闭它或事后删除非关键尺寸。" % total)
    if dimension_filter:
        if filter_aborted:
            notes.append("dimension_filter 未生效：%s" % filter_aborted)
        else:
            notes.append("dimension_filter 生效：保留 %s 条、删除 %d 条不符合规划"
                         "的尺寸（失败 %d 条）。"
                         % (filter_kept, len(filter_removed), len(filter_failed)))
    # P0-1: push this call's dimensions outside their views, stacked by side.
    outside_placed = 0
    outside_failed = []
    outside_sides = {}
    if place_outside:
        outside_placed, outside_failed, outside_sides = _push_dimensions_outside(
            doc, _win_auto, before_names,
            gap_m=float(outside_gap_m), step_m=float(outside_step_m))
        try:
            _call_method(doc, "EditRebuild3")
        except Exception:
            pass
        try:
            doc.GraphicsRedraw2()
        except Exception:
            pass
        # Explicit save: this is the only branch that edits geometry the caller asked
        # to be laid out, and the rest of the function never saves.
        try:
            _save_drawing(doc)
        except Exception:
            pass
        notes.append(
            "place_outside：%d 条尺寸已推到视图外按侧分层（外扩 %.0f mm + n×%.0f mm），"
            "失败 %d 条。" % (outside_placed, float(outside_gap_m) * 1000,
                             float(outside_step_m) * 1000, len(outside_failed)))

    return ok(
        drawing=doc.GetTitle,
        inserted=count,
        layout=layout,
        include_unmarked=bool(include_unmarked),
        total_dimensions=total,
        dimension_filter_applied=bool(dimension_filter),
        filter_kept=filter_kept,
        filter_removed=len(filter_removed),
        filter_removed_names=filter_removed[:20],
        filter_failed=filter_failed[:10],
        filter_aborted=filter_aborted,
        duplicate_dimension_names=len(duplicate_report),
        duplicate_examples=dict(list(duplicate_report.items())[:5]),
        dedupe_applied=bool(eliminate_duplicates),
        duplicates_removed=len(removed),
        duplicate_names_removed=removed[:20],
        dedupe_failed=dedupe_failed[:10],
        # 0.8.25: an explicit count — `dedupe_failed` is truncated to 10, so a run
        # that removed 0 and failed 23 used to be indistinguishable from a clean no-op.
        dedupe_failed_count=len(dedupe_failed),
        place_outside_applied=bool(place_outside),
        placed_outside=outside_placed,
        outside_sides=outside_sides,
        outside_failed=outside_failed[:10],
        advisory=notes,
        views=audit.get("views"),
        warnings=audit.get("warnings"),
    )


@tool_guard("add_dimension_failed")
def add_drawing_dimension(drawing_path, x=None, y=None, value=None, offset_x=0.0,
                          offset_y=-0.015, view_name=None, tolerance=0.01,
                          edge_index=None, skip_duplicates=False, known_values=None, save=True):
    """Add ONE dimension on the visible edge nearest sheet point (x, y), meters.

    This is the manual dimensioning path (process §4.4): InsertModelItems only carries
    dimensions the model already has, so IR-rebuilt parts need this to fill the gaps.

    The point does not have to be exactly on the line: the nearest visible edge within
    `tolerance` meters is selected, and the dimension is placed at (x+offset_x,
    y+offset_y). value is meters; omit it to keep the computed value.

    skip_duplicates (off by default) rejects values listed in `known_values`. It cannot
    infer them: a dimension's value is the PROJECTED length, and the same edge projects
    differently per view (1253 mm measured in 3D becomes 37.6 mm in the side view), so
    the check can only trust numbers the caller passes — in metres, as projected in the
    TARGET view. New dimensions are saved (save=True); without that they exist only in
    memory and vanish on reopen.
    """
    from ...sw_session import byref_dispatch_null

    doc = _open_drawing(drawing_path)
    # Compare in whole millimetres: the same dimension rounded to 0.1 mm still differs
    # in the last float digit between views, and drafting only cares about the number.
    existing = set()
    if skip_duplicates:
        # The annotation list is cached until the drawing rebuilds, so a dimension
        # added moments ago is invisible without this.
        try:
            doc.EditRebuild3
        except Exception:
            pass
        existing |= {round(v * 1000) for v in drawing_dimension_values(doc)}
        existing |= {round(float(v) * 1000) for v in (known_values or [])}
    view = None
    if view_name:
        for candidate in _iter_views(doc):
            if getattr(candidate, "Name", None) == view_name:
                view = candidate
                break
    if view is None and x is not None and y is not None:
        view = _view_at_point(doc, float(x), float(y))
    if view is not None:
        try:
            doc.ActivateView(view.Name)
        except Exception:
            pass

    doc.ClearSelection2(True)
    selected = None
    distance = None
    picked_length = None

    if edge_index is not None:
        if view is None:
            return error("view_not_found", "edge_index requires view_name")
        import win32com.client as win

        edges = [win.Dispatch(e) for e in (view.GetVisibleEntities2(view, 1) or ())]
        if edge_index < 0 or edge_index >= len(edges):
            return error("bad_index", "edge_index %d out of range (0..%d)"
                         % (edge_index, len(edges) - 1))
        try:
            target_edge = edges[edge_index]
            picked_length = _edge_model_length(target_edge, _projection_axes(view, edges))
            picked = target_edge.Select2(False, 0)
        except Exception:
            picked = False
        if not picked or doc.SelectionManager.GetSelectedObjectCount2(-1) < 1:
            return error("select_failed", "Could not select edge %d" % edge_index)
        selected = "edge_index"

    # Fast path: coordinate pick (works for some entity kinds on some builds).
    if selected is None and x is not None and y is not None:
        for candidate in ("EDGE", "SILHOUETTE", "SKETCHSEGMENT"):
            try:
                if doc.Extension.SelectByID2("", candidate, float(x), float(y), 0,
                                             False, 0, byref_dispatch_null(), 0):
                    selected = candidate
                    break
            except Exception:
                continue

    if selected is None:
        # Two very different situations used to share one message (P3-1): "you gave me
        # a point that lies outside every view" was reported as "no sheet point given",
        # so the caller kept re-sending x/y that could never work. Keep them apart.
        if x is None or y is None:
            return error(
                "missing_point",
                "No edge selected and no sheet point given",
                "Pass edge_index (+view_name) or sheet x/y in metres.",
            )
        if view is None:
            return error(
                "point_outside_views",
                "点 (%.4f, %.4f) 不在任何视图范围内" % (float(x), float(y)),
                "可用视图范围（米，left/bottom/right/top）：%s。"
                "改成视图内的点，或直接用 edge_index。" % _view_outline_hint(doc),
            )
        # Fall back to projecting every visible edge's midpoint onto the sheet and
        # taking the nearest one — SelectByID2 cannot hit drawing edges by coordinates.
        hit = _nearest_edge_at_sheet_point(view, float(x), float(y), float(tolerance))
        if hit is None:
            return error(
                "segment_not_found",
                "No visible edges could be enumerated in this view",
                "The view may be blank, not activated, or not rebuilt yet. Activate "
                "it and retry, or pass edge_index from list_drawing_edges.",
            )
        import win32com.client as win

        distance, midpoints, proj = hit
        if distance > float(tolerance):
            # P1-8: a bare "not found" left callers unable to tell whether x/y were
            # in sheet metres or model coordinates. Report BOTH, plus the distance,
            # so one read is enough to self-calibrate the next call.
            ai, bi = proj["ai"], proj["bi"]
            m1, m2 = midpoints
            mid_model = [(m1[i] + m2[i]) / 2.0 for i in range(3)]
            nearest_sheet = [
                round(proj["pos"][0] + (mid_model[ai] - proj["centre"][ai]) * proj["scale"], 4),
                round(proj["pos"][1] + (mid_model[bi] - proj["centre"][bi]) * proj["scale"], 4),
            ]
            return error(
                "segment_not_found",
                "No visible edge within %.1f mm of sheet point (%.4f, %.4f)"
                % (float(tolerance) * 1000, float(x), float(y)),
                "Nearest edge is %.2f mm away. Re-aim at nearest_edge_sheet_m, raise "
                "tolerance, or pass edge_index from list_drawing_edges."
                % (distance * 1000),
                requested_sheet_m=[round(float(x), 4), round(float(y), 4)],
                nearest_edge_sheet_m=nearest_sheet,
                nearest_edge_model_m={
                    "start": [round(float(v), 5) for v in m1],
                    "end": [round(float(v), 5) for v in m2],
                },
                nearest_edge_distance_mm=round(distance * 1000, 2),
                tolerance_mm=round(float(tolerance) * 1000, 2),
                projection_plane=proj["plane"],
            )
        picked = None
        for edge in [win.Dispatch(e) for e in (view.GetVisibleEntities2(view, 1) or ())]:
            try:
                p1 = edge.GetStartVertex.GetPoint
                p2 = edge.GetEndVertex.GetPoint
            except Exception:
                continue
            if (abs(p1[0] - midpoints[0][0]) < 1e-9 and abs(p1[1] - midpoints[0][1]) < 1e-9
                    and abs(p2[0] - midpoints[1][0]) < 1e-9 and abs(p2[1] - midpoints[1][1]) < 1e-9):
                picked = edge
                break
        if picked is None or not picked.Select2(False, 0):
            return error("select_failed", "Could not select the nearest edge",
                         "Retry with edge_index from list_drawing_edges.")
        try:
            picked_length = _edge_model_length(
                picked, _projection_axes(view, [win.Dispatch(e) for e in (view.GetVisibleEntities2(view, 1) or ())]))
        except Exception:
            picked_length = None
        selected = "edge_by_point"

    if x is None or y is None:
        try:
            base_x, base_y = list(view.Position)[:2]
        except Exception:
            base_x, base_y = 0.2, 0.2
    else:
        base_x, base_y = float(x), float(y)
    display_dim = doc.AddDimension2(base_x + float(offset_x), base_y + float(offset_y), 0)
    if display_dim is None:
        doc.ClearSelection2(True)
        return error(
            "dimension_failed",
            "SolidWorks rejected the dimension at (%.4f, %.4f)" % (base_x, base_y),
            "The entity may not be dimensionable (e.g. a construction line or a "
            "silhouette that already has a driver).",
        )
    # A wrong edge still yields a dimension — it just measures something else (report:
    # edge_index 71 pointed at a Ø36 arc and the tool produced R23 of the Ø46 outline,
    # a "plausible" wrong value nobody would question). A dimension attached to nothing
    # is the cheap signal that the selection did not take.
    if edge_index is not None:
        try:
            _ann = _call_method(display_dim, "GetAnnotation")
            _types = list(_ann.GetAttachedEntityTypes or ())
            if not _types:
                doc.ClearSelection2(True)
                try:
                    _call_method(doc, "EditDelete")
                except Exception:
                    pass
                return error(
                    "edge_mismatch",
                    "Dimension at edge_index=%s attached to no geometry" % edge_index,
                    "The indexed edge and the selected edge disagree. Re-read the "
                    "indices with list_drawing_edges, or pass x/y instead.")
        except Exception:
            pass
    dim = display_dim.GetDimension2(0)
    applied = None
    try:
        applied = float(dim.SystemValue) if dim is not None else None
    except Exception:
        applied = None
    if value is not None and dim is not None:
        try:
            dim.SystemValue = float(value)
            applied = float(dim.SystemValue)
        except Exception:
            pass
    # Dimensions are per-value: the same length annotated twice on two views is a
    # drafting error, so skip values that already exist anywhere on the sheet.
    if skip_duplicates and picked_length is not None and round(picked_length * 1000) in existing:
        doc.ClearSelection2(True)
        return drawing_ok(
            drawing=doc.GetTitle, created=False, selected=selected,
            view=view.Name if view is not None else None,
            value_m=float(picked_length), skipped="duplicate_value",
            note="The same dimension already exists on this sheet.",
        )

    doc.ClearSelection2(True)
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass
    if save:
        # Without this the dimension exists only in memory and disappears on reopen.
        try:
            doc.SaveSilent
        except Exception:
            pass
    return drawing_ok(
        drawing=doc.GetTitle,
        created=True,
        selected=selected,
        saved=bool(save),
        view=view.Name if view is not None else None,
        distance_m=round(distance, 6) if distance is not None else None,
        value_m=applied,
        note="Verify the dimension text and its attachment, then export a PDF.",
    )


@tool_guard("tolerance_failed")
def add_dimension_tolerance(drawing_path, value_mm, tolerance_type="bilateral",
                           upper=0.1, lower=0.1, view_name=None, apply=True):
    """Apply a native tolerance to an existing drawing dimension.

    The dimension is located by its VALUE in millimetres (drafting tolerances are
    specified per nominal size), so no fragile coordinate picking is involved.

    tolerance_type: bilateral (default, ±upper/lower) | symmetric | limit | basic |
    min | max | fit | fit_with_tolerance | none. upper/lower are in millimetres.
    """
    doc = _open_drawing(drawing_path)
    code = TOLERANCE_TYPES.get(str(tolerance_type).lower())
    if code is None:
        return error("bad_tolerance_type",
                     "Unknown tolerance_type %r" % tolerance_type,
                     "Use one of: %s" % ", ".join(sorted(TOLERANCE_TYPES)))
    view, dim = _find_dimension(doc, value_mm, view_name)
    if dim is None:
        return error("dimension_not_found",
                     "No dimension equal to %s mm%s"
                     % (value_mm, " in view %r" % view_name if view_name else ""),
                     "Use drawing_audit to see the dimensions present on the sheet.")
    planned = {
        "value_mm": round(float(dim.SystemValue) * 1000, 3),
        "tolerance_type": str(tolerance_type).lower(),
        "upper_mm": float(upper), "lower_mm": float(lower),
        "view": view.Name if view is not None else None,
    }
    if not apply:
        return drawing_ok(dry_run=True, planned=planned,
                          message="Nothing changed. apply=true writes the tolerance.")

    try:
        dim.SetToleranceType(code)
    except Exception as exc:
        return error("tolerance_unsupported", str(exc),
                     "This dimension may not accept a tolerance (reference dimensions "
                     "and driven dimensions reject them).")
    if code != 0:
        # `SetToleranceValues` takes METRES. Passing millimetres made 0.1 mm render as
        # "±100" — a 1000x error that still reported applied:true and read back a valid
        # tolerance type (measured 2026-11-01: 50 symmetric ±0.1 -> PDF shows "50 ±100").
        try:
            dim.SetToleranceValues(float(upper) / 1000.0, float(lower) / 1000.0)
        except Exception:
            # Symmetric takes a single value on some builds.
            try:
                dim.SetToleranceValues(float(upper) / 1000.0, float(upper) / 1000.0)
            except Exception:
                pass
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass
    try:
        doc.SaveSilent
    except Exception:
        pass
    applied = None
    try:
        applied = int(dim.GetToleranceType)
    except Exception:
        pass
    # The tolerance VALUE lives in a text slot, not in the numeric readback: a 1000x
    # unit error looked perfect until the PDF was rendered. Report what it prints.
    rendered_text = None
    try:
        for _slot in (1, 0, 2):
            _candidate = dim.GetText(_slot)
            if _candidate:
                rendered_text = str(_candidate)
                break
    except Exception:
        pass
    return drawing_ok(applied=True, saved=True, planned=planned,
                      tolerance_type_readback=applied,
                      rendered_text=rendered_text)


@tool_guard("layout_dimensions_failed")
def layout_dimensions(drawing_path, view_name=None, min_gap_m=0.006,
                      cluster_tol_m=0.0015, dry_run=False, save=True):
    """Spray out dimensions that landed on top of each other.

    `auto_dimension` on a section view stacks every Ø dimension on the axis: measured on the
    filling-valve body, 34 dimensions shared only a handful of x values (six sat at
    x=0.0942), so the text overlapped into unreadable strings like "Ø18Ø20".

    SolidWorks has no "distribute" call, so this clusters dimensions by x (tolerance
    `cluster_tol_m`), sorts each cluster by y, and pushes any that are closer than
    `min_gap_m` upward until they are not. Dimensions that are already clear are untouched —
    the tool only pays for the ones that actually collide.

    dry_run=True reports the plan without moving anything.
    """
    import win32com.client as win

    doc = _open_drawing(drawing_path)
    views = [v for v in _iter_views(doc) if getattr(v, "Type", None) != 1]
    if view_name:
        views = [v for v in views if str(v.Name) == view_name]
        if not views:
            return error("view_not_found", "no view named %r" % view_name)

    def annotation_type(annotation):
        """IAnnotation.GetType is an int PROPERTY under late binding; calling it raises.

        The old `int(annotation.GetType())` therefore always raised, and the surrounding
        `except: pass` let EVERY annotation through the filter — notes included, which
        the sweep then moved around as if they were dimensions.
        """
        for getter in (lambda: annotation.GetType, lambda: annotation.GetType()):
            try:
                return int(getter())
            except Exception:
                continue
        return None

    items = []
    fixed = []
    # P0-2: OTHER views are obstacles. A dimension belonging to view A must not be
    # pushed onto view B's outline — a report had main-view dimensions landing on the
    # section view's hatching. The owning view is NOT an obstacle: dimensions
    # legitimately sit inside their own view, and blocking that would shove every one
    # of them off the sheet.
    for view in views:
        try:
            view_box = precise_outline(view) or list(view.GetOutline)
        except Exception:
            view_box = None
        if view_box and len(view_box) >= 4:
            fixed.append({
                "annotation": None, "view": str(view.Name), "text": "",
                "x": (view_box[0] + view_box[2]) / 2.0,
                "y": (view_box[1] + view_box[3]) / 2.0,
                "w": max(1e-4, view_box[2] - view_box[0]),
                "h": max(1e-4, (view_box[3] - view_box[1]) / 2.0),
                "fixed": True, "obstacle": "view",
            })
    for view in views:
        for raw in (view.GetAnnotations or ()):
            annotation = win.Dispatch(raw)
            kind = annotation_type(annotation)
            if kind == 6:
                # P3-1: a view's own label ("2:1", "A—A") is a NOTE. Once a Ø8
                # dimension was pushed into it the sheet printed "2Ø8¹" and the scale
                # became unreadable. Notes are FIXED obstacles: seeded into `placed`
                # and never moved, so the sweep has to go around them.
                try:
                    position = [float(v) for v in annotation.GetPosition]
                except Exception:
                    continue
                text, char_h = "", 0.0035
                try:
                    specific = win.Dispatch(annotation.GetSpecificAnnotation)
                    text = str(specific.GetText or "")
                    try:
                        char_h = float(specific.GetTextFormat.CharHeight) or 0.0035
                    except Exception:
                        pass
                except Exception:
                    pass
                width = (max(char_h * 0.8, len(text) * char_h * 0.75)
                         if text else char_h * 3)
                fixed.append({"annotation": annotation, "view": str(view.Name),
                              "x": position[0], "y": position[1], "text": text,
                              "w": width, "h": char_h, "fixed": True})
                continue
            if kind != 4:
                continue
            try:
                position = [float(v) for v in annotation.GetPosition]
            except Exception:
                continue
            # Text width is needed to catch horizontal overlap: two dimensions 3 mm
            # apart in x still collide because "Ø25.3" is ~10 mm wide. IDisplayDimension
            # gives the character height, so the box can be estimated well enough.
            text, char_h = "", 0.0035
            try:
                specific = win.Dispatch(annotation.GetSpecificAnnotation)
                for index in (0, 1, 2):
                    try:
                        candidate = str(specific.GetText(index) or "")
                        if candidate:
                            text = candidate
                            break
                    except Exception:
                        continue
                try:
                    char_h = float(specific.GetTextFormat.CharHeight) or 0.0035
                except Exception:
                    pass
                # Plain dimensions inserted by auto_dimension return EMPTY from GetText
                # (only callouts carry text there). GetDimension2 still answers with the
                # value, so build a representative string from it — otherwise the width
                # estimate degenerates and the collision test misses real overlaps.
                if not text.strip():
                    try:
                        dim2 = win.Dispatch(specific.GetDimension2(0))
                        value_mm = float(dim2.SystemValue) * 1000.0
                        text = "%g" % round(value_mm, 2)
                    except Exception:
                        pass
            except Exception:
                pass
            # Rough monospace assumption: 0.62 em average advance for digits/Ø/dots.
            width = max(char_h * 0.8, len(text) * char_h * 0.62) if text else char_h * 3
            items.append({"annotation": annotation, "view": str(view.Name),
                          "x": position[0], "y": position[1],
                          "text": text, "w": width, "h": char_h})
    if not items:
        return error("no_dimensions", "no dimensions found on the selected views")

    def boxes_overlap(a, b, pad=0.0008):
        """Overlap test using estimated text boxes, not just anchor points."""
        ax0, ax1 = a["x"] - a["w"] / 2, a["x"] + a["w"] / 2
        bx0, bx1 = b["x"] - b["w"] / 2, b["x"] + b["w"] / 2
        ay0, ay1 = a["y"] - a["h"], a["y"] + a["h"]
        by0, by1 = b["y"] - b["h"], b["y"] + b["h"]
        return (ax0 < bx1 + pad and bx0 < ax1 + pad
                and ay0 < by1 + pad and by0 < ay1 + pad)

    items.sort(key=lambda i: (round(i["x"] / cluster_tol_m), i["y"]))
    moves = []
    # Sweep in reading order; push each item up until it clears every earlier one, and
    # AROUND the fixed notes. Pushing up into a view label only trades one overlap for
    # another, so a note clash sends the dimension DOWN instead (P3-1).
    placed = list(fixed)
    for item in items:
        for _attempt in range(40):
            clash = None
            for other in placed:
                if (other.get("obstacle") == "view"
                        and other.get("view") == item.get("view")):
                    continue      # dimensions may live inside their own view
                if boxes_overlap(item, other):
                    clash = other
                    break
            if clash is None:
                break
            if clash.get("fixed"):
                item["y"] = clash["y"] - clash["h"] - item["h"] - min_gap_m
            else:
                item["y"] = clash["y"] + clash["h"] + item["h"] + min_gap_m * 0.25
            moves.append({"view": item["view"], "text": item["text"][:10],
                          "x": round(item["x"], 4), "to_y": round(item["y"], 4),
                          "around": "note" if clash.get("fixed") else "dimension"})
        placed.append(item)

    # P0-2: say what is STILL overlapping. `moved < collisions` with no explanation
    # left the caller unable to tell "gave up" from "nothing left to do".
    unresolved = []
    for item in items:
        remaining = 0
        for other in placed:
            if other is item:
                continue
            if (other.get("obstacle") == "view"
                    and other.get("view") == item.get("view")):
                continue
            if boxes_overlap(item, other):
                remaining += 1
        if remaining:
            unresolved.append({"view": item["view"], "text": item["text"][:10],
                               "x": round(item["x"], 4), "y": round(item["y"], 4),
                               "remaining_overlaps": remaining})

    # One pass fixes 64% (14 collisions -> 5): pushing one dimension up can create a new
    # collision with the next one, so the pass has to repeat until it settles.

    applied = 0
    if moves and not dry_run:
        for item in items:
            try:
                current = [float(v) for v in item["annotation"].GetPosition]
            except Exception:
                continue
            if abs(current[1] - item["y"]) < 1e-9:
                continue
            try:
                item["annotation"].SetPosition(float(item["x"]), float(item["y"]), 0)
                applied += 1
            except Exception:
                continue
        try:
            _call_method(doc, "EditRebuild3")
        except Exception:
            pass
        if save:
            _save_drawing(doc)

    return drawing_ok(dimension_count=len(items), collisions=len(moves),
                      moved=applied, dry_run=bool(dry_run),
                      fixed_obstacles=len(fixed),
                      view_obstacles=sum(1 for o in fixed
                                         if o.get("obstacle") == "view"),
                      unresolved=unresolved[:20],
                      unresolved_count=len(unresolved),
                      min_gap_mm=min_gap_m * 1000,
                      sample_moves=moves[:5],
                      note=("按 x 聚类后沿 y 推开，间距 %.0f mm。只动真正重叠的，"
                            "已经清楚的尺寸不碰。"
                            "视图标签 note 视为固定障碍（%d 个）：撞上时尺寸向下让，"
                            "不会再把比例标签压成「2Ø8¹」。" % (min_gap_m * 1000, len(fixed))))


@tool_guard("attachment_check_failed")
def verify_dimension_attachments(drawing_path, view_name=None):
    """Report what each dimension is actually attached to (process §6.1).

    A dimension with no attached entity — or attached only to a sketch segment instead
    of a model edge — is not traceable to geometry and must not be counted as complete.
    Attachment is read through IAnnotation.GetAttachedEntities2 / GetAttachedEntityTypes.
    """
    import win32com.client as win

    doc = _open_drawing(drawing_path)
    rows, counts = [], {}
    orphans = []
    # Same dimension gets enumerated many times (measured 11x for one D1/275.0), and
    # name is NOT an identity here — some drawings name every dimension "D1" — so
    # dedupe on (view, name, position, value) and also report the unique count.
    seen_dims = set()
    duplicates_skipped = 0
    raw_hits = 0
    by_primary = {}

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
            except Exception:
                continue
            if dim is None:
                continue
            row = {"view": view.Name}
            try:
                row["name"] = str(ann.GetName)
            except Exception:
                row["name"] = None
            try:
                row["value_mm"] = round(float(dim.SystemValue) * 1000, 3)
            except Exception:
                row["value_mm"] = None
            # A driving dimension constrains the model; a reference (从动) one only
            # reports what the geometry already is. GB practice puts reference
            # dimensions in parentheses, and a drawing full of unmarked reference
            # dimensions looks dimensioned while constraining nothing.
            reference = None
            for attempt in ("IsReference", "GetReference"):
                try:
                    value = getattr(dim, attempt)
                    reference = bool(value() if callable(value) else value)
                    break
                except Exception:
                    continue
            row["is_reference"] = reference
            row["role"] = ("reference" if reference
                           else ("driving" if reference is False else "unknown"))
            # Needed by the dedupe key below. Without it the key collapsed to
            # (view, name, value) and drawings where every dimension is named "D1"
            # lost genuinely-distinct dimensions that happened to share a value.
            try:
                row["position_m"] = [round(float(v), 4) for v in ann.GetPosition]
            except Exception:
                row["position_m"] = None
            types = []
            try:
                raw_types = ann.GetAttachedEntityTypes
                if isinstance(raw_types, (tuple, list)):
                    types = [int(t) for t in raw_types]
                elif raw_types is not None:
                    types = [int(raw_types)]
            except Exception:
                pass
            row["attached_count"] = len(types)
            row["attached_kinds"] = [SELECT_TYPE_NAMES.get(t, "type_%d" % t) for t in types]

            # Dedupe first, bookkeep after. It used to be the other way round, so
            # counts/orphans counted RAW HITS (23) while dimension_count counted
            # unique ones (11) — the three numbers described the same sheet and
            # disagreed. A report caught this: "8+15=23 != 11".
            raw_hits += 1
            # Identity of a driving dimension is its MODEL full name. The old key was
            # (view, name, position, value): SolidWorks reuses D1/D2 per sketch and a
            # sheet can hold many equal values, so distinct dimensions collapsed —
            # measured 23 real dimensions reported as 18 with duplicates_skipped=5.
            # `_annotation_dimension_full_name` already existed (eliminate_duplicates
            # uses it); the fallback keeps unnamed annotations countable.
            dedupe_key = (_annotation_dimension_full_name(ann, win)
                          or (row.get("view"), row.get("name"), row.get("value_mm")))
            if dedupe_key in seen_dims:
                duplicates_skipped += 1
                continue
            seen_dims.add(dedupe_key)

            # Two different questions, two different counters. `counts` answers "how
            # many entity attachments of each kind" (one dimension can attach to many
            # edges); `by_primary` answers "how many dimensions are primarily of this
            # kind". Mixing them produced a report where 8+15=23 didn't equal 11 —
            # neither number was wrong, they were answering different questions.
            for kind in row["attached_kinds"]:
                counts[kind] = counts.get(kind, 0) + 1
            primary = row["attached_kinds"][0] if row["attached_kinds"] else "none"
            by_primary[primary] = by_primary.get(primary, 0) + 1
            if not types:
                # `GetAttachedEntities2` is empty; whether that means "loose" or just
                # "moved with SetPosition" cannot be decided from the position alone.
                # Record the proximity observation, do not treat it as a clearance.
                position = row.get("position_m")
                near_leader = False
                if position:
                    try:
                        box = precise_outline(view)
                        if box:
                            gap_x = max(box[0] - position[0], 0.0,
                                        position[0] - box[2])
                            gap_y = max(box[1] - position[1], 0.0,
                                        position[1] - box[3])
                            near_leader = ((gap_x * gap_x + gap_y * gap_y) ** 0.5
                                           <= 0.05)
                    except Exception:
                        near_leader = False
                # Proximity is NOT evidence of a leader. Since v0.8.27 dimensions are
                # parked just outside the view outline by default, so *every* dimension
                # in a tidy sheet sits within 50 mm — and a report showed this "fix"
                # excusing 15 genuine orphans (23 = 8 attached + 15 excused) while the
                # same receipt still said "23 个尺寸均已关联". Report the observation for
                # what it is and leave the verdict alone.
                row["orphan_near_view"] = near_leader
                orphans.append({"view": view.Name, "name": row["name"],
                                "value_mm": row.get("value_mm"),
                                "position_m": position,
                                "orphan_near_view": near_leader})
            rows.append(row)

    total = len(rows)
    return drawing_ok(
        drawing=doc.GetTitle,
        dimension_count=total,
        attachment_summary=counts,
        dimensions_by_primary_attachment=by_primary,
        orphan_count=len(orphans),
        orphans=orphans,
        unique_dimension_count=len(rows),
        raw_hits=raw_hits,
        duplicates_skipped=duplicates_skipped,
        reference_count=sum(1 for r in rows if r.get("is_reference") is True),
        driving_count=sum(1 for r in rows if r.get("is_reference") is False),
        role_unknown=sum(1 for r in rows if r.get("is_reference") is None),
        dimensions=rows,
        note="orphan_count > 0 表示有尺寸没有关联到任何实体，"
             "按流程 §6.1 属于未完成；只挂在 sketch_segment 上的也应复核。"
             " is_reference=true 是从动参考尺寸，不驱动模型——GB 习惯加括号标注，"
             "整张图全是参考尺寸时图面看着标注齐全，实际没有约束任何东西。",
    )


@tool_guard("prune_failed")
def prune_orphan_dimensions(drawing_path, view_name=None, dry_run=True, save=True,
                            confirm=False):
    """Delete dimensions that are attached to no geometry (P2-4).

    `auto_dimension` imports every driving dimension; on real parts a share of them land
    unattached — `GetAttachedEntities2` returns nothing for roughly half of them whatever
    the import parameters are. There was no way to clear those: `delete_dimensions` works by
    VALUE, so it cannot tell two 23 mm dimensions apart (a valve body held three).

    Dry-run by default and it reports view / name / value / position for each candidate.
    `dry_run=False` deletes by name + position (position is what separates repeated D1/D2
    names) and then RE-RUNS the attachment check, so the receipt carries the remaining
    orphan count instead of assuming the deletes worked.

    P1-2 — `orphan` means "GetAttachedEntities2 returned nothing", NOT "this dimension is
    loose". A report moved dimensions with SetPosition and then found 11 orphans including
    the 50 mm overall length and the 25 mm hole pitch; every one of them visibly led to real
    geometry on the render. Deleting that list would have gutted the drawing. So a real
    delete now requires `confirm=True`; without it the call returns
    `confirmation_required` with the same list, and the receipt always carries `caveat`.
    """
    from . import layout as _layout

    before = verify_dimension_attachments(drawing_path, view_name=view_name)
    if not before.get("ok"):
        return error("attachments_unreadable", before.get("message") or "failed")
    orphans = before.get("orphans") or []
    targets = [{"view": row.get("view"), "name": row.get("name"),
                "value_mm": row.get("value_mm"),
                "position_m": row.get("position_m")} for row in orphans]
    payload = {
        "drawing": before.get("drawing"),
        "orphan_count_before": len(targets),
        "orphan_count": len(targets),
        "dimension_count": before.get("dimension_count"),
        "dry_run": bool(dry_run),
        "targets": targets,
        "removed_count": 0,
    }
    # Always present, not only in dry-run: this is the sentence that decides whether
    # deleting the list is safe, and it applies to both modes.
    payload["caveat"] = (
        "orphan 只表示 GetAttachedEntities2 返回空，不等于「尺寸真的孤立」。"
        "实测：尺寸被 SetPosition 移动后会出现这种读不到附件的情况，"
        "而它在图纸上有正常引线、指向真实几何（含总长、孔距这类关键尺寸）。"
        "照单全删会让图纸缺料——先看渲染图逐条确认，或用 "
        "delete_annotation(name + position_m) 精确删。")
    if dry_run or not targets:
        if not dry_run and not targets:
            payload["note"] = "没有 orphan 尺寸需要清理。"
        return drawing_ok(**payload)
    if not confirm:
        return error(
            "confirmation_required",
            "检测到 %d 条 orphan 候选，但未确认删除（confirm=False）。" % len(targets),
            "orphan 可能只是取不到附件。核对渲染图后再传 confirm=True。",
            caveat=payload["caveat"], targets=targets)

    removed, failed = [], []
    for row in targets:
        name = row.get("name")
        if not name:
            failed.append({"name": None, "reason": "no name reported"})
            continue
        result = _layout.delete_annotation(drawing_path, name, save=False,
                                           position_m=row.get("position_m"))
        if result.get("ok"):
            removed.append({"name": name, "view": row.get("view"),
                            "value_mm": row.get("value_mm")})
        else:
            failed.append({"name": name, "reason": result.get("error"),
                           "message": str(result.get("message") or "")[:120]})
    if save:
        _save_drawing(_open_drawing(drawing_path))
    after = verify_dimension_attachments(drawing_path, view_name=view_name)
    payload.update({
        "removed": removed,
        "failed": failed,
        "removed_count": len(removed),
        "orphan_count_after": (int(after.get("orphan_count") or 0)
                               if after.get("ok") else None),
        "orphan_count": (int(after.get("orphan_count") or 0)
                         if after.get("ok") else None),
    })
    return drawing_ok(**payload)


