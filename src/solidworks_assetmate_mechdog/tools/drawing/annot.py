from __future__ import annotations

"""Annotation tools: marks, callouts, title block, gtol, finish."""

from ._core import *

@tool_guard("center_marks_failed")
def add_center_marks(drawing_path, include_slots=True, extended_lines=True):
    """Auto-insert center marks and centerlines on every hole/slot in every view.

    Requires activating each view first — drawing annotation ops act on the active
    view. GetCenterMarkCount under-reports on this binding, so `views_marked` (the
    API's own success count) is the reliable signal, not the count.
    """
    doc = _open_drawing(drawing_path)
    insert_type = 1 | (4 if include_slots else 0)  # Hole = 1, Slots = 4
    insert_option = 2  # swCenterMarkStyle_e.swCenterMark_Single
    processed = marked = 0
    results = []
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        processed += 1
        try:
            doc.ActivateView(view.Name)
        except Exception:
            pass
        try:
            ok_view = view.AutoInsertCenterMarks2(
                insert_type, insert_option, bool(include_slots), bool(include_slots),
                True, 0.0025, 0.0, bool(extended_lines), True, 0.0,
            )
        except Exception as exc:
            results.append({"view": view.Name, "ok": False, "error": str(exc)})
            continue
        if ok_view:
            marked += 1
        results.append({"view": view.Name, "ok": bool(ok_view)})
    try:
        doc.EditRebuild3
    except Exception:
        pass
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass
    layout = None
    _save_drawing(doc)
    return ok(
        drawing=doc.GetTitle,
        layout=layout,
        views_processed=processed,
        views_marked=marked,
        views=results,
        note="GetCenterMarkCount under-reports on this binding; views_marked is the "
             "reliable signal. Verify visually with render_drawing.",
    )


@tool_guard("hole_callout_failed")
def add_hole_callout(drawing_path, x, y, view_name=None, model_path=None,
                     save=True):
    """Insert a NATIVE hole callout at a sheet point (meters) on a hole edge, then read
    it back (process §4.5: the tool's success flag is not evidence).

    Read-back returns the callout text and the dimensioned value. When model_path is
    given, the value is matched against the model's cylinder inventory so a wrong edge
    (callout on a Ø6 hole that should be Ø30) is reported instead of silently accepted.
    """
    doc = _open_drawing(drawing_path)
    target = None
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        if view_name and view.Name != view_name:
            continue
        target = view
        break
    if target is not None:
        try:
            doc.ActivateView(target.Name)
        except Exception:
            pass
    before = 0
    if target is not None:
        try:
            before = int(target.GetAnnotationCount)
        except Exception:
            before = 0
    doc.ClearSelection2(True)
    # AddHoleCallout2 needs the hole's projected EDGE selected first: the selection is
    # what carries the owning view and the circular edge into the callout. Without it
    # SolidWorks returns null even at a correct point.
    from ...sw_session import byref_dispatch_null

    # Match by ENUMERATING circles first (see _nearest_circle_edge); coordinate
    # hit-testing stays only as a fallback for views whose axis mapping is unavailable.
    circle_match, why = _nearest_circle_edge(
        drawing_path, view_name or (str(target.Name) if target is not None else None),
        float(x), float(y))
    # The helper re-opens the drawing internally, so `doc` and `target` here are stale.
    doc = _open_drawing(drawing_path)
    target = None
    for _view in _iter_views(doc):
        if getattr(_view, "Type", None) == 1:
            continue
        if view_name and _view.Name != view_name:
            continue
        target = _view
        break
    if target is not None:
        try:
            doc.ActivateView(target.Name)
        except Exception:
            pass
    # x/y chooses WHICH hole (via the nearest circle); the leader anchor is placed just
    # outside the view, exactly like the batch tool. Passing the circle centre here made
    # AddHoleCallout2 return None even with the edge correctly selected — the text
    # anchor and the edge selection compete for the same point.
    callout_point = [float(x), float(y)]
    try:
        outline = list(target.GetOutline or ())
        if len(outline) >= 4:
            callout_point = [outline[2] + 0.015, (outline[1] + outline[3]) / 2.0]
    except Exception:
        pass
    circle_matched = None
    if circle_match is not None:
        circle_matched = {"diameter_mm": circle_match["diameter_mm"],
                          "distance_m": circle_match["distance_m"],
                          "centre_m": list(circle_match["centre"])}
        callout_point = list(circle_match["centre"])
        # Re-find and select the same circle on THIS handle. Matching by diameter keeps
        # it independent of the projection that the helper already did.
        import win32com.client as _win_pick

        picked = False
        _cands = []
        # Match the EXACT circle: several holes can share a diameter, and picking by
        # diameter alone annotates a different hole than the caller pointed at (that is
        # why the first attempt selected an edge, reported count 1, and still got null).
        _want = circle_match.get("centre_model") or []
        try:
            for _edge in [_win_pick.Dispatch(e)
                          for e in (target.GetVisibleEntities2(target, 1) or ())]:
                try:
                    _curve = _win_pick.Dispatch(_edge.GetCurve)
                    if _curve is None or not _curve.IsCircle:
                        continue
                    _params = list(_curve.CircleParams)
                    if len(_want) == 3:
                        _d = sum(abs(float(_params[i]) - float(_want[i]))
                                 for i in range(3))
                        if _d > 1e-6:
                            continue
                    elif abs(float(_params[6]) * 2000
                             - float(circle_match["diameter_mm"])) > 0.2:
                        continue
                    _cands.append((float(_params[2]), _edge))
                except Exception:
                    continue
        except Exception:
            _cands = []
        # A hole has two concentric circles (near face and far face) and only the near
        # one is annotatable. Same rule as the batch tool: take the smallest z.
        for _z, _edge in sorted(_cands, key=lambda entry: entry[0]):
            try:
                if _edge.Select2(False, 0):
                    picked = True
                    break
            except Exception:
                continue
        if not picked:
            return error("select_failed",
                         "Found the Ø%g circle but could not select it on this document"
                         % circle_match["diameter_mm"],
                         "Use add_hole_callouts (batch) for this view instead.")
    else:
        selected = False
        for sel_type in ("EDGE", "SILHOUETTE"):
            try:
                if doc.Extension.SelectByID2("", sel_type, float(x), float(y), 0,
                                             False, 0, byref_dispatch_null(), 0):
                    selected = True
                    break
            except Exception:
                continue
        if not selected:
            return error("edge_not_found",
                         "No hole edge near (%.4f, %.4f)" % (x, y),
                         "Pass the hole CENTRE from list_circular_edges.",
                         nearest=why)
    callout = doc.AddHoleCallout2(callout_point[0], callout_point[1], 0.0)
    if callout is None:
        # Keep the evidence: with the circle matched and its edge selected, a null here
        # means AddHoleCallout2 itself refused — not that the point was wrong. Reporting
        # the match lets the next round tell those apart without re-running blind.
        _sel_count = None
        try:
            _sel_count = int(doc.SelectionManager.GetSelectedObjectCount2(-1))
        except Exception:
            pass
        return error(
            "callout_failed",
            "AddHoleCallout2 returned None at (%.4f, %.4f)" % tuple(callout_point),
            "Circle {dia} matched at distance {dist} m; selection before the call: "
            "{sel}. With a match and a selection, SolidWorks still refused — this "
            "circle is probably not an annotatable hole (an outer arc, or the far face). "
            "The batch tool has the same limit: on door.SLDPRT it annotates Ø8/Ø15 and "
            "never the four Ø6 circles. Pick another circle from list_circular_edges."
            .format(
                dia=(circle_matched or {}).get("diameter_mm"),
                dist=(circle_matched or {}).get("distance_m"),
                sel=_sel_count),
            matched_circle=circle_matched, selected_before_call=_sel_count)

    readback = {}
    # GetText is a method on some builds and a property on others: call it, and only
    # fall back to a bare attribute read when calling raises.
    try:
        text_value = callout.GetText()
        readback["text"] = str(text_value) if not callable(text_value) else None
    except Exception:
        try:
            text_value = callout.GetText
            readback["text"] = str(text_value) if not callable(text_value) else None
        except Exception:
            readback["text"] = None
    try:
        dim = callout.GetDimension2(0)
        if dim is not None:
            readback["value_mm"] = round(float(dim.SystemValue) * 1000, 3)
    except Exception:
        pass
    if not readback.get("value_mm"):
        # A callout may hang off a display dimension rather than a plain dimension.
        try:
            readback["value_mm"] = round(float(callout.SystemValue) * 1000, 3)
        except Exception:
            pass

    match = None
    if model_path:
        try:
            from . import ledger as _ledger

            led = _ledger.build_feature_ledger(model_path, include_faces=True)
            values = sorted({round(h["diameter_mm"], 1) for h in (led.get("holes") or [])})
            got = readback.get("value_mm")
            if got is not None and values:
                nearest = min(values, key=lambda v: abs(v - got))
                match = {
                    "read_mm": got,
                    "nearest_model_diameter_mm": nearest,
                    "delta_mm": round(got - nearest, 3),
                    "ok": abs(got - nearest) <= max(0.2, nearest * 0.02),
                    "model_diameters_mm": values[:20],
                }
        except Exception as exc:
            match = {"error": str(exc)[:120]}

    after = 0
    if target is not None:
        try:
            after = int(target.GetAnnotationCount)
        except Exception:
            after = 0
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass
    if save:
        _save_drawing(doc)
    return drawing_ok(
        drawing=doc.GetTitle,
        created=True,
        view=target.Name if target is not None else None,
        annotations_before=before,
        annotations_after=after,
        readback=readback,
        model_match=match,
        # Which circle the point resolved to. Without it a caller cannot tell "the hole
        # I meant" from "the nearest hole", and this tool used to fail outright instead.
        matched_circle=circle_matched,
        saved=bool(save),
        note="比对 model_match.ok；false 表示标注落在别处或模型无对应孔径。",
    )


@tool_guard("hole_callouts_failed")
def add_hole_callouts(drawing_path, view_name=None, min_diameter_mm=0.0,
                      max_callouts=40, budget_s=25.0, leader_x_m=None,
                      leader_y_m=None, save=True):
    """Native hole callouts on the holes VISIBLE in a view, one callout per hole.

    Verified on SW2026 (2026-10-28):
      * `AddHoleCallout2` does NOT show the modal dialog the 2001-era Remarks warns about;
        it returns in ~0.03 s.
      * Success must be judged by `IDisplayDimension.IsHoleCallout` — a hole callout is a
        DisplayDimension (GetType 4), there is no separate annotation type.
      * **The edge must face the viewer.** On the panel-box door (6 mm sheet) the Ø6 edges
        at z=6.0 (the far face) all returned null, while Ø8 and Ø15 edges at z=0.0 succeeded.
        So a through hole contributes two circular edges and only the near one is annotatable.

    Hence: collect edges per view, group by projected centre, and keep ONE edge per hole —
    preferring a complete circle, else the one nearest the viewer along the sheet normal.

    `leader_x_m` / `leader_y_m` (P1-7) pin the callout text to a sheet point in metres.
    By default the text is stacked just outside the view's outline, which on a compact part
    threw the leader 100 mm past the frame border. The per-callout 12 mm vertical stack
    still applies when a custom anchor is given, so callouts do not overlap each other.
    Each created item reports `anchor_m` (requested) and `position_m` (read back).

    Performance: only this view's visible edges are enumerated. A whole-model scan took
    >400 s; one view has 9-40 circular edges and the whole call takes ~19 s.
    """
    import time

    import win32com.client as win

    doc = _open_drawing(drawing_path)
    targets = []
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        if view_name and view.Name != view_name:
            continue
        targets.append(view)
    if not targets:
        return error("no_views", "no model views to annotate")

    created, failed = [], []
    started = time.time()
    budget_exceeded = False
    floor_m = max(float(min_diameter_mm) * 0.0005, 1e-6)
    scanned = 0
    edges_scanned = 0
    rejected = 0
    groups_total = 0

    for view in targets:
        try:
            doc.ActivateView(view.Name)
        except Exception:
            pass
        try:
            components = list(view.GetVisibleComponents or ())
        except Exception:
            components = []

        # --- collect, then group by projected centre -------------------------------
        # A circle wider than 40% of the view is an outer contour, not a hole. The
        # track-wheel outline (Ø220-290 mm) came back as IsHoleCallout=true with text
        # "R"; rejecting it after the call still left a radius callout on the sheet and
        # still incremented annotation.hole_callout. Judge before inserting.
        view_box = list(view.GetOutline or [])
        view_width = (view_box[2] - view_box[0]) if len(view_box) >= 4 else None
        candidates = []
        for component in components:
            try:
                edges = list(view.GetVisibleEntities2(component, 1) or ())
            except Exception:
                continue
            for raw in edges:
                try:
                    edge = win.Dispatch(raw)
                    curve = win.Dispatch(edge.GetCurve)
                    if curve is None or not curve.IsCircle:
                        continue
                    params = list(curve.CircleParams)
                    radius_m = float(params[6])
                    if radius_m < floor_m:
                        continue
                    # Count every circular edge the scan met, not just the ones that
                    # survive the filters: "scanned: 0" next to "failed: 20" reads as
                    # "this view had nothing", when 20 edges were seen and rejected.
                    edges_scanned += 1
                    if view_width and radius_m * 2 > view_width * 0.4:
                        rejected += 1
                        failed.append({"view": view.Name,
                                       "diameter_mm": round(radius_m * 2000, 2),
                                       "reason": "skipped: circle is >40% of the view "
                                                 "width, treated as an outer contour"})
                        continue
                    candidates.append({
                        "edge": edge,
                        "radius_m": radius_m,
                        "centre": (round(float(params[0]), 4), round(float(params[1]), 4)),
                        "z": float(params[2]),
                        "complete": edge.GetStartVertex is None,
                    })
                except Exception:
                    continue
        scanned += len(candidates)

        grouped = {}
        for item in candidates:
            # Same centre + same radius = same hole seen from both faces.
            key = (item["centre"], round(item["radius_m"], 4))
            grouped.setdefault(key, []).append(item)
        groups_total += len(grouped)

        outline = list(view.GetOutline or [])
        default_x = (outline[2] + 0.015) if len(outline) >= 4 else 0.30
        default_y = ((outline[1] + outline[3]) / 2.0) if len(outline) >= 4 else 0.20
        # P1-7: the old default anchor sat just outside the view outline, which on a
        # compact part threw the leader well past the sheet border. Callers can now
        # pin the text to any sheet point; the vertical stack below still separates
        # consecutive callouts.
        anchor_x = float(leader_x_m) if leader_x_m is not None else default_x
        anchor_y = float(leader_y_m) if leader_y_m is not None else default_y

        # Consecutive-failure breaker. On a louvered sheet every one of 90 circular
        # edges returned null after a full Select2 + AddHoleCallout2 round trip, and
        # the run burned 82 s to produce nothing. Ten straight nulls means this view's
        # circles are not holes; stop and say so instead of grinding through the rest.
        consecutive_nulls = 0
        for key, members in grouped.items():
            if len(created) + len(failed) >= int(max_callouts):
                break
            if consecutive_nulls >= 10:
                failed.append({"view": view.Name, "radius_mm": None,
                               "reason": "stopped: 10 consecutive callouts returned null "
                                         "(this view's circles are not holes)"})
                break
            if time.time() - started > float(budget_s):
                budget_exceeded = True
                break
            # Prefer a complete circle (the near face usually reads as one), then the
            # edge nearest the viewer along z.
            chosen = sorted(members, key=lambda m: (not m["complete"], m["z"]))[0]
            doc.ClearSelection2(True)
            try:
                if not chosen["edge"].Select2(False, 0):
                    failed.append({"view": view.Name,
                                   "diameter_mm": round(chosen["radius_m"] * 2000, 2),
                                   "reason": "Select2 returned False", "z": chosen["z"]})
                    continue
            except Exception as exc:
                failed.append({"view": view.Name,
                               "diameter_mm": round(chosen["radius_m"] * 2000, 2),
                               "reason": "select: %s" % str(exc)[:50], "z": chosen["z"]})
                continue

            slot = len(created) + len(failed)
            try:
                display = doc.AddHoleCallout2(anchor_x, anchor_y - slot * 0.012, 0.0)
            except Exception as exc:
                failed.append({"view": view.Name,
                               "diameter_mm": round(chosen["radius_m"] * 2000, 2),
                               "reason": "callout: %s" % str(exc)[:50], "z": chosen["z"]})
                continue
            if display is None:
                failed.append({"view": view.Name,
                               "diameter_mm": round(chosen["radius_m"] * 2000, 2),
                               "reason": "returned None (edge not annotatable from this view)",
                               "z": chosen["z"], "hidden_side": chosen["z"] != 0.0})
                continue
            dim = win.Dispatch(display)
            try:
                is_hole = bool(dim.IsHoleCallout)
            except Exception:
                is_hole = None
            if is_hole is False:
                failed.append({"view": view.Name,
                               "diameter_mm": round(chosen["radius_m"] * 2000, 2),
                               "reason": "not a hole callout", "z": chosen["z"]})
                continue
            text = None
            for index in (0, 1, 2):
                try:
                    candidate = getattr(dim, "GetText")(index)
                    if candidate:
                        text = str(candidate)
                        break
                except Exception:
                    continue
            # IsHoleCallout is also True for a plain radius callout on an outer arc:
            # a 220 mm track-wheel outline came back as is_hole_callout=true with
            # text "R". Accepting that let "R" stand in for hole information and
            # suppressed the missing-hole blocker — the tool fooling its own check.
            def _discard(display_dim):
                """Delete a callout we are about to reject.

                Rejecting AFTER `AddHoleCallout2` still left the annotation on the sheet
                (and still counted in annotation.hole_callout), so the inventory and the
                pixels disagreed. Remove it before reporting failure.
                """
                try:
                    ann = win.Dispatch(_call_method(display_dim, "GetAnnotation"))
                    doc.ClearSelection2(True)
                    if ann.Select(False):
                        _call_method(doc, "EditDelete")
                except Exception:
                    pass

            text_upper = str(text or "").upper()
            has_diameter = ("Ø" in str(text or "")) or ("MOD-DIAM" in text_upper) \
                or text_upper.startswith("M") or " X " in text_upper
            radius_only = text_upper.startswith("R") and not has_diameter
            # v0.8.33: an UNREADABLE text is NOT evidence of a bad callout. On SW2026
            # `GetText(0/1/2)` returns None for every callout, so treating "no Ø found"
            # as failure deleted all of them (measured: 8 created -> 0 on the sheet).
            # Only reject on POSITIVE evidence — text that explicitly starts with R.
            text_unreadable = not str(text or "").strip()
            if radius_only:
                _discard(dim)
                failed.append({"view": view.Name,
                               "diameter_mm": round(chosen["radius_m"] * 2000, 2),
                               "reason": "not a hole callout (text=%r, likely an outer arc)"
                                         % (text,),
                               "z": chosen["z"], "radius_only": True})
                continue
            if not has_diameter and not text_unreadable:
                _discard(dim)
                failed.append({"view": view.Name,
                               "diameter_mm": round(chosen["radius_m"] * 2000, 2),
                               "reason": "callout text carries no diameter (text=%r)"
                                         % (text,),
                               "z": chosen["z"]})
                continue
            # P1-1: the printed number must agree with the measured edge.
            expected_d = round(chosen["radius_m"] * 2000, 2)
            if _callout_diameter_matches(text, expected_d) is False:
                _discard(dim)
                failed.append({"view": view.Name,
                               "diameter_mm": expected_d,
                               "reason": ("text_value_mismatch: callout says %r but the "
                                          "edge measures Ø%.2f" % (str(text)[:40], expected_d)),
                               "z": chosen["z"]})
                continue
            callout_point = [anchor_x, anchor_y - slot * 0.012]
            # `GetAnnotation` must be invoked as a METHOD under late binding (see
            # layout.add_note). Read the real text position back so the caller can
            # verify SolidWorks honoured the requested leader anchor (P1-7).
            position_m = None
            try:
                ann = win.Dispatch(_call_method(dim, "GetAnnotation"))
                position_m = [round(float(v), 4) for v in ann.GetPosition]
            except Exception:
                pass
            consecutive_nulls = 0
            created.append({
                "view": view.Name,
                "diameter_mm": round(chosen["radius_m"] * 2000, 2),
                "is_hole_callout": is_hole,
                "text": text,
                # Same signal the earlier rejection path needs: keep the callout but say
                # the text could not be read, so the caller decides on the geometry.
                "text_unreadable": text_unreadable,
                "z": chosen["z"],
                "anchor_m": [round(callout_point[0], 4), round(callout_point[1], 4)],
                "position_m": position_m,
                # Kept so the SetText rewrite below can reach the IDisplayDimension.
                # Stripped before returning — COM objects must not leak into JSON.
                "_annotation": dim,
            })

        if budget_exceeded or len(created) + len(failed) >= int(max_callouts):
            break

    doc.ClearSelection2(True)
    try:
        doc.EditRebuild3()
    except Exception:
        pass
    if save:
        _save_drawing(doc)
    # --- rewrite the raw format codes into readable text -----------------------
    # AddHoleCallout2 leaves GetText as "<MOD-DIAM>" which renders as those literal
    # characters. IDisplayDimension.SetText replaces it (value lands in slot 1, not 0).
    for item in created:
        dim_obj = item.pop("_annotation", None)
        if dim_obj is None:
            continue
        # Only rewrite a RAW FORMAT CODE. Writing "Ø15" into slot 1 replaced the
        # diameter but left the depth sitting in another slot, so a Ø15 hole rendered as
        # "Ø1515" — the report's "Ø9 -> Θ99". A Ø8 through-hole has no depth and looked
        # fine, which is why this hid for so long. When the callout already reads
        # correctly, SolidWorks' own hole-table text is the right answer.
        # NO SetText. Writing "Ø15" into one slot replaced the diameter while the
        # depth stayed in another, so the sheet printed "Ø1515" — the report's
        # "Ø9 -> Θ99". A Ø8 through hole has no depth, which is why it looked fine.
        # SolidWorks' own hole-table text is correct; only the RECEIPT needs a value,
        # and only when the text cannot be read (SW2026 returns None for every slot).
        _native = ""
        for _probe in (1, 0, 2):
            try:
                _candidate = dim_obj.GetText(_probe)
            except Exception:
                continue
            if _candidate and str(_candidate).strip():
                _native = str(_candidate)
                break
        if _native and "<MOD-" not in _native.upper():
            item["text"] = _native
            item["text_source"] = "native"
            continue
        # The raw "<MOD-DIAM>" must be replaced or it prints literally, but writing only
        # the diameter leaves the DEPTH in another slot: a Ø15 hole then reads "Ø1515"
        # (the report's "Ø9 -> Θ99"). So write the diameter, then blank the other slots.
        diameter = item.get("diameter_mm")
        try:
            target = chr(216) + "%g" % round(float(diameter), 2)
        except Exception:
            continue
        for slot in (1, 0, 2):
            try:
                dim_obj.SetText(slot, target)
            except Exception:
                continue
            break
        # KNOWN LIMIT (0.8.36): when the hole also carries a depth, that depth lives in
        # another text slot and SetText cannot clear it without erasing the annotation —
        # blanking the other slots made the callout disappear entirely (measured). So a
        # Ø15 hole WITH depth 15 still prints "Ø1515". A through hole (Ø8 here) is fine.
        item["text"] = target
        item["text_source"] = "SetText_depth_may_trail"
        texts = []
        for probe in range(3):
            try:
                texts.append(str(dim_obj.GetText(probe) or ""))
            except Exception:
                texts.append("")
        filled = [t for t in texts if t.strip()]
        if filled:
            item["text"] = filled[0]

    unresolved = [c for c in created
                  if "<MOD-" in str(c.get("text") or "").upper()]
    for item in unresolved:
        item["text_unresolved"] = True

    return drawing_ok(
        drawing=doc.GetTitle,
        circular_edges=scanned,
        edges_scanned=edges_scanned,
        rejected_outer_contour=rejected,
        holes_found=groups_total,
        created_count=len(created),
        failed_count=len(failed),
        created=created,
        failed=failed[:20],
        leader_anchor_requested_m=(
            [round(float(leader_x_m), 4), round(float(leader_y_m), 4)]
            if leader_x_m is not None and leader_y_m is not None else None),
        unresolved_count=len(unresolved),
        unresolved_note=(("%d 条标注的文本仍是未解析格式码（<MOD-DIAM> 等），渲染出来是"
                          "字面量——SetText 改写没成功。建议删掉改用 Ø 尺寸标注。")
                         % len(unresolved)) if unresolved else None,
        # Judge the budget with the same clock that produced elapsed_s: a run whose
        # candidates were all rejected never entered the insert loop, so it reported
        # budget_exceeded=False while taking 55 s of a 45 s budget.
        budget_exceeded=bool(budget_exceeded or (time.time() - started) > float(budget_s)),
        elapsed_s=round(time.time() - started, 2),
        saved=bool(save),
        note=("孔标注用 IsHoleCallout 判定；同一孔的上下两条圆边只标一条（优先完整圆、"
              "再取最靠观察者的一面）。hidden_side=true 表示该边在背面，SolidWorks 不给标注。"),
        # P1-2: "0 created / 8 failed" used to read as an ordinary partial result, so
        # callers retried with different offsets and burned whole rounds. Measured
        # 2026-10-31 on a valve body: the edge IS selectable, its adjacent face IS
        # selectable, the callout point sits exactly on the circle, and AddHoleCallout2
        # still returns null for every edge — while the same call succeeds on sheet-metal
        # cut holes. That is a geometry/build limit, not a parameter problem, so name it.
        # P1-1: a run that produced callouts AND rejected some is `partial`, not `ok`.
        # "budget_exceeded: true + 5 failed + verdict ok" was the old behaviour.
        verdict=("ok" if created and not failed else
                 ("partial" if created else
                  ("no_native_callout" if groups_total else "nothing_to_annotate"))),
        native_callout_note=(
            None if not failed else
            ("%d 条成功、%d 条被拒（原因见 failed；被拒的标注已从图上撤销，不留残留）。"
             "实测：边与相邻面都能选中、锚点就在圆上时 AddHoleCallout2 仍可能返回 null，"
             "属几何/构建限制，换 offset 或换视图不会改变结果。"
             % (len(created), len(failed)))),
    )


@tool_guard("title_block_failed")
def set_title_block(drawing_path, name=None, code=None, material=None, apply=False,
                   include_reversed=True, scale=None):
    """Fill title-block placeholders such as the literal text “图样名称” / “图样代号”.

    GB templates ship these as plain note text with quotes, so nothing is linked to
    the model and every drawing shows the placeholder. Dry-run by default; only the
    fields you pass are touched, unknown fields are reported, never invented.

    `图样代号` appears twice: in the title block (0 deg) and at the sheet's top-left
    corner rotated 180 deg — the GB convention that makes the number readable after
    folding the sheet. Both are filled so the drawing stays consistent; pass
    include_reversed=False to leave the rotated marker alone.
    """
    doc = _open_drawing(drawing_path)
    sheet = doc.GetFirstView
    notes = _sheet_notes(sheet)
    planned, unmatched, skipped_reversed, skipped_extra = [], [], [], []

    wants = []
    if name is not None:
        wants.append(("name", "图样名称", str(name)))
    if code is not None:
        wants.append(("code", "图样代号", str(code)))
    if material is not None:
        wants.append(("material", "材质", str(material)))

    # The scale cell holds a bare value like "1:5", not a labelled field, so the
    # marker-based loop below never finds it. Match the value pattern instead — and
    # only inside the title-block band, or a note that happens to read "1:5" elsewhere
    # would be overwritten.
    if scale is not None:
        import re as _re

        scale_pattern = _re.compile(r"^\s*\d+(?:\.\d+)?\s*:\s*\d+(?:\.\d+)?\s*$")
        try:
            sheet_box = list(sheet.GetOutline or [])
            sheet_h = (sheet_box[3] - sheet_box[1]) if len(sheet_box) >= 4 else 0.297
        except Exception:
            sheet_h = 0.297
        for note in notes:
            try:
                text = note.GetText
            except Exception:
                continue
            if not isinstance(text, str) or not scale_pattern.match(text):
                continue
            try:
                position = [round(v, 4) for v in ann_position(note)]
            except Exception:
                position = None
            if position and position[1] > sheet_h * 0.15:
                continue  # outside the title-block band
            planned.append({
                "field": "scale", "before": text, "after": str(scale), "note": note,
                "angle_deg": 0.0, "position": position,
            })
        if not any(p["field"] == "scale" for p in planned):
            unmatched.append({"field": "scale", "marker": "n:m",
                              "reason": "title block has no scale value in the bottom band"})

    for field, marker, new_value in wants:
        candidates = []
        for note in notes:
            try:
                text = note.GetText
            except Exception:
                continue
            if not isinstance(text, str) or marker not in text:
                continue
            try:
                position = [round(v, 4) for v in ann_position(note)]
            except Exception:
                position = None
            try:
                angle = float(note.Angle)
            except Exception:
                angle = 0.0
            reversed_note = abs(angle) > 0.01
            if reversed_note and not include_reversed:
                skipped_reversed.append({
                    "field": field, "text": text,
                    "angle_deg": round(angle * 180 / 3.141592653589793, 1)})
                continue
            if field == "material":
                before = text
                after = text.replace("<未指定>", new_value).replace("<未指定 >", new_value)
                if after == before:
                    after = ("材质 %s" % new_value if "<" not in text
                             else text.replace(
                                 text[text.find("<"):text.find(">") + 1], new_value))
            else:
                before = text
                after = new_value
            candidates.append({
                "field": field, "before": before, "after": after, "note": note,
                "angle_deg": round(angle * 180 / 3.141592653589793, 1),
                "position": position, "_reversed": reversed_note,
            })

        if candidates:
            # P1-4: several cells can carry the marker. A GB frame repeats 图样代号 in a
            # 180-degree note OUTSIDE the frame, and writing both put a second code in
            # the corner (measured: code at [0.053, 0.285] and [0.391, 0.022]). Keep
            # ONE — upright first, then nearest the configured slot.
            slot = TITLE_SLOTS.get(field)

            def _rank(candidate, _slot=slot):
                distance = 0.0
                if _slot and candidate.get("position"):
                    dx = candidate["position"][0] - _slot[0]
                    dy = candidate["position"][1] - _slot[1]
                    distance = (dx * dx + dy * dy) ** 0.5
                return (1 if candidate["_reversed"] else 0, distance)

            candidates.sort(key=_rank)
            chosen = candidates[0]
            chosen.pop("_reversed", None)
            planned.append(chosen)
            for entry in candidates[1:]:
                skipped_extra.append({
                    "field": field, "text": entry.get("before"),
                    "position": entry.get("position"),
                    "angle_deg": entry.get("angle_deg"),
                    "reason": "another cell carries the same marker; kept the upright "
                              "one nearest the configured slot"})
        else:
            slot = TITLE_SLOTS.get(field)
            note = _note_at(sheet, slot) if slot else None
            if note is not None:
                try:
                    before = note.GetText
                except Exception:
                    before = ""
                planned.append({
                    "field": field, "before": before if isinstance(before, str) else "",
                    "after": new_value, "note": note, "via": "position",
                })
            elif slot:
                # P1-4 third tier: the cell can be an EMPTY table cell with no note at
                # all — the report on gb_a3.drwdot came back with planned: [] for every
                # field. Create the note at the slot instead of declaring it unmatched.
                created_note = None
                try:
                    created_note = doc.InsertNote(str(new_value))
                except Exception:
                    created_note = None
                if created_note is not None:
                    try:
                        import win32com.client as _win_tb

                        _created_ann = _win_tb.Dispatch(
                            _call_method(created_note, "GetAnnotation"))
                        # SetPosition2 MUST go through call_method: pywin32 dispatches it
                        # as a property and the call raises MEMBERNOTFOUND.
                        _call_method(_created_ann, "SetPosition2",
                                     float(slot[0]), float(slot[1]), 0.0)
                    except Exception:
                        pass
                    try:
                        created_note.SetHeight(0.0025)
                    except Exception:
                        pass
                    planned.append({
                        "field": field, "before": "", "after": new_value,
                        "note": created_note, "via": "created",
                    })
                else:
                    unmatched.append({
                        "field": field, "marker": marker,
                        "reason": "no matching text and InsertNote failed"})
            else:
                unmatched.append({"field": field, "marker": marker})

    result = {
        "drawing": doc.GetTitle,
        "planned": [dict({k: v for k, v in p.items() if k != "note"},
                         note_name=_note_name(p.get("note")),
                         cell_position_m=([round(v, 4) for v in ann_position(p["note"])]
                                          if p.get("note") is not None else None))
                    for p in planned],
        "unmatched": unmatched,
        "skipped_reversed": skipped_reversed,
        "skipped_extra_matches": skipped_extra,
        "applied": False,
    }
    if not apply:
        result["dry_run"] = True
        result["message"] = "Nothing changed. apply=true writes these title-block fields."
        return ok(**result)

    written = []
    for item in planned:
        cell = {
            "note_name": _note_name(item.get("note")),
            "cell_position_m": ([round(v, 4) for v in ann_position(item["note"])]
                                if item.get("note") is not None else None),
        }
        try:
            item["note"].SetText(item["after"])
            written.append(dict({"field": item["field"], "before": item["before"],
                                 "after": item["after"], "ok": True}, **cell))
        except Exception as exc:
            written.append(dict({"field": item["field"], "ok": False,
                                 "error": str(exc)}, **cell))
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass
    doc.SaveSilent
    result.update(applied=True, written=written, saved=True)
    return ok(**result)


@tool_guard("annotations_failed")
def list_drawing_annotations(drawing_path, view_name=None):
    """Inventory every annotation on the sheet, classified by kind.

    Process §6.1 requires checking that datums, geometric tolerances and surface
    finishes exist and say what they should. SolidWorks can create these only through
    early binding (a C# add-in), but READING them works here — that is enough for a
    self-check to state "this drawing has no geometric tolerance at all".

    Geometric tolerance frames are read through IGtol.GetFrameValues.
    """
    import win32com.client as win

    doc = _open_drawing(drawing_path)
    rows, counts = [], {}
    seen = set()
    try:
        outline = list(doc.GetFirstView.GetOutline or [])
        sheet_w = (outline[2] - outline[0]) if len(outline) >= 4 else 0.42
        sheet_h = (outline[3] - outline[1]) if len(outline) >= 4 else 0.297
    except Exception:
        sheet_w, sheet_h = 0.42, 0.297

    def scan(view, scope):
        try:
            raw = view.GetAnnotations
        except Exception:
            return
        items = list(raw) if isinstance(raw, (tuple, list)) else ([raw] if raw else [])
        for item in items:
            if item is None:
                continue
            try:
                ann = win.Dispatch(item)
                kind_code = int(ann.GetType)
            except Exception:
                continue
            kind = ANNOTATION_KINDS.get(kind_code, "other_%d" % kind_code)
            # Reset per iteration: leaving the previous value in place made a callout
            # inherit the PREVIOUS annotation's text when its own read failed.
            callout_text = ""
            # A hole callout is a DisplayDimension, not its own annotation type — the
            # only way to tell it apart from an ordinary dimension is IsHoleCallout.
            if kind == "dimension":
                try:
                    specific = ann.GetSpecificAnnotation
                    if specific is not None and bool(specific.IsHoleCallout):
                        kind = "hole_callout"
                except Exception:
                    pass
            name = None
            try:
                name = str(ann.GetName)
            except Exception:
                name = None
            # IsHoleCallout is ALSO true for a radius dimension on an outer arc, and
            # `GetText` does not exist on those objects — so a text-based test silently
            # passes on unreadable text and lets "R" stand in for hole information.
            # The name prefix is reliable and costs nothing: RD* = radius dimension.
            if kind == "hole_callout" and name:
                upper_name = name.upper()
                callout_text = ""
                # GetText lives on IDisplayDimension, NOT on IAnnotation — calling it on
                # the annotation raises and the swallowed exception made every callout
                # look like it had no text, so all of them fell through to the RD-prefix
                # rule and even real Ø callouts were filed as radius dimensions.
                try:
                    specific = ann.GetSpecificAnnotation
                except Exception:
                    specific = None
                for index in (0, 1, 2):
                    try:
                        callout_text = str(specific.GetText(index) or "")
                        if callout_text:
                            break
                    except Exception:
                        continue
                upper_text = callout_text.upper()
                looks_like_hole = (
                    "Ø" in callout_text or "MOD-DIAM" in upper_text
                    or upper_text.startswith("M") or " X " in upper_text
                )
                # Order matters: readable text that names a diameter wins over the
                # name prefix, because a genuine hole callout can also be named RD*.
                # Only when the text is unreadable do we fall back to the prefix.
                if looks_like_hole:
                    pass  # keep hole_callout
                elif upper_name.startswith("RD"):
                    kind = "radius_callout"
                elif not callout_text:
                    # No text and no RD prefix: do not claim it is a hole callout.
                    # "Unreadable" must not read as "verified".
                    kind = "hole_callout_unverified"
                else:
                    kind = "radius_callout"
            position = None
            try:
                position = [round(v, 4) for v in ann.GetPosition]
            except Exception:
                position = None
            # Dedupe on scope+name+position: the same name (D1) can be hit from several
            # scans, but a genuine duplicate has the same position too.
            key = (scope, name, tuple(position) if position else None)
            if name and key in seen:
                continue
            if name:
                seen.add(key)
            row = {"kind": kind, "name": name, "scope": scope}
            # `text` must be written AFTER row exists. It used to be written before
            # this line, which meant: first scanned annotation -> UnboundLocalError
            # (the tool returned annotations_failed); any later one -> the text was
            # appended to the PREVIOUS row, silently shifting data between lines.
            if kind in ("hole_callout", "radius_callout", "hole_callout_unverified"):
                # GetText can return raw format codes like "<MOD-DIAM>" instead of "Ø".
                # The classification above already treats MOD-DIAM as a diameter, but
                # handing the code straight to the caller is unreadable — flag it and
                # convert the one code we know for sure.
                row["text"] = callout_text or None
                if callout_text and "<MOD-" in callout_text.upper():
                    row["text_is_format_code"] = True
                    row["text"] = (callout_text.upper()
                                   .replace("<MOD-DIAM>", "Ø")
                                   .replace("<MOD-DEG>", "°"))
            if position:
                row["position_m"] = position
            if kind == "geometric_tolerance":
                try:
                    gtol = ann.GetSpecificAnnotation
                    row["frames"] = int(gtol.GetFrameCount)
                    row["frame1"] = [str(v) for v in gtol.GetFrameValues(1)]
                    row["text"] = str(gtol.GetText(1))
                except Exception:
                    pass
            if kind in ("dimension", "hole_callout", "radius_callout",
                        "hole_callout_unverified"):
                dimension_value = None
                specific = None
                try:
                    specific = ann.GetSpecificAnnotation
                    dim = specific.GetDimension2(0)
                    if dim is not None:
                        dimension_value = round(float(dim.SystemValue) * 1000, 3)
                        row["value_mm"] = dimension_value
                        # A callout whose text is just "Ø" / "R" (or None) tells the
                        # reader nothing, and the value is already on this row. Reports
                        # hit this four times: radius_callout text=None, hole_callout
                        # text="Ø" with no number.
                        if kind in ("hole_callout", "radius_callout",
                                    "hole_callout_unverified"):
                            # `SystemValue` is NOT consistently radius or diameter for
                            # callouts: measured on one sheet, a printed Ø8 answered 4.0
                            # (radius) while Ø15 answered 15.0 (already diameter). So
                            # trust the printed text when it carries a number, and only
                            # fall back to SystemValue — flagged — when it does not.
                            import re as _re

                            _bare = str(row.get("text") or "").strip()
                            _printed = _re.search(r"([0-9]+(?:\.[0-9]+)?)", _bare)
                            if _printed:
                                row["diameter_mm"] = float(_printed.group(1))
                                row["diameter_source"] = "text"
                            elif _bare in ("", "Ø", "⌀", "R"):
                                row["diameter_mm"] = round(float(dimension_value), 3)
                                row["diameter_source"] = "value_unverified"
                                _prefix = ("R" if (_bare == "R"
                                                   or kind == "radius_callout")
                                           else "Ø")
                                row["text"] = "%s%g" % (
                                    _prefix, round(float(dimension_value), 3))
                                row["text_was_bare_symbol"] = True
                        try:
                            row["tolerance_type"] = int(dim.GetToleranceType)
                        except Exception:
                            pass
                except Exception:
                    specific = None
                if specific is not None:
                    # P1-1 / P3-3: `GetText()` with NO slot returns '' for a dimension —
                    # the value and its display codes live in slot 1 (measured
                    # 2026-10-31: a Ø20 dimension answers GetText(0)='' and
                    # GetText(1)='<MOD-DIAM>'). Reading no slot meant every plain
                    # dimension looked text-less, so delivery_check counted 0 Ø
                    # dimensions on a sheet that visibly printed Ø28/Ø23/Ø12/Ø18/Ø13.
                    try:
                        row["dimension_type"] = int(specific.GetType)
                    except Exception:
                        pass
                    if kind == "dimension":
                        raw_value_text = ""
                        for index in (1, 0, 2):
                            try:
                                probe = str(specific.GetText(index) or "")
                            except Exception:
                                probe = ""
                            if probe:
                                raw_value_text = probe
                                break
                        if raw_value_text:
                            row["raw_text"] = raw_value_text
                            resolved = _resolve_display_text(
                                raw_value_text, dimension_value)
                            if resolved:
                                row["text"] = resolved
                            if "<MOD-" in raw_value_text.upper():
                                row["text_is_format_code"] = True
            if kind == "note":
                try:
                    text = ann.GetSpecificAnnotation.GetText
                    if isinstance(text, str):
                        row["text"] = text
                except Exception:
                    pass
                # Title-block and frame text is template furniture, not annotation the
                # drafter added: a freshly created A3 sheet reports 82 "notes" whose
                # content is 审核/校核/设计/标记. Counting them hides real notes and
                # inflates the note tally that drawing_audit and delivery_check read.
                try:
                    position = row.get("position_m")
                    if position and sheet_w and sheet_h:
                        if position[1] < sheet_h * 0.13 and position[0] > sheet_w * 0.5:
                            kind = "title_block_text"
                except Exception:
                    pass
                # Position alone missed the frame zone labels (1 2 3 4 / A B C D /
                # 分区 / 更改文件号 / 年 月 日 / 阶 段 标 记): an empty A3 sheet still
                # reported 40 "notes", so annotation_count read 91 against 24 real
                # annotations. Match on content too — a one- or two-character token, or
                # a known template phrase, is never a drafter's note.
                template_phrase = str(row.get("text") or "").strip()
                is_furniture = template_phrase in TEMPLATE_TEXT or (
                    0 < len(template_phrase) <= 2
                    and all(ch.isalnum() for ch in template_phrase)
                )
                if kind in ("note", "title_block_text") and is_furniture:
                    kind = "template_text"
            row["kind"] = kind
            rows.append(row)

    for view in _iter_views(doc):
        is_sheet = getattr(view, "Type", None) == 1
        if view_name and not is_sheet and view.Name != view_name:
            continue
        scan(view, "sheet" if is_sheet else view.Name)

    # Totals are derived FROM the rows. Incrementing a counter during the scan and
    # appending the row after a dedupe `continue` made them disagree: a repeated name
    # bumped the count but produced no row, so totals said 15 dimensions while only 1
    # row existed. One source of truth instead of two.
    from collections import Counter

    counts = dict(Counter(row.get("kind") for row in rows))

    return drawing_ok(
        drawing=doc.GetTitle,
        totals=counts,
        # Count only annotation the drafter (or this toolkit) put there. Template
        # furniture inflated this to 91 against 24 real annotations on an empty sheet,
        # which made the "real drawings average 46.5 notes" baseline incomparable.
        annotation_count=sum(
            1 for row in rows
            if row.get("kind") not in ("title_block_text", "template_text")
        ),
        title_block_text_count=counts.get("title_block_text", 0),
        template_text_count=counts.get("template_text", 0),
        annotations=rows,
        notes=[
            "geometric_tolerance 的 frame1 依次为 公差值/基准1/基准2/基准3/符号位。",
            "surface_finish 只能报存在，内容读取需早绑定。",
        ],
    )


@tool_guard("surface_finish_failed")
def add_surface_finish(drawing_path, x, y, roughness="3.2", sym_type="machining",
                       leader_type="none", lay_symbol=0, arrow_type=0,
                       prod_method="", sample_len="", min_roughness="",
                       roughness_spacing="", save=True, verify_render=False,
                       verify_dwg=False):
    """Insert a NATIVE surface finish symbol (Ra value) at a sheet point.

    `IModelDocExtension.InsertSurfaceFinishSymbol3` takes FOURTEEN arguments — that count
    is why every earlier attempt failed with "non-optional parameter". Order:
      (SymType, LeaderType, X, Y, Z, LaySymbol, ArrowType, MachAllowance, OtherVals,
       ProdMethod, SampleLen, MaxRoughness, MinRoughness, RoughnessSpacing)

    x/y are sheet metres. roughness is the Ra value as text (e.g. "3.2").

    P2-11 (0.8.20): InsertSurfaceFinishSymbol3 IGNORES the x/y arguments — the symbol
    actually lands at the sheet origin (0,0), so echoing position_m=[x,y] was a lie. The
    annotation is now moved with SetPosition and read back as `position_actual_m`.
    `verify_render=True` exports a PDF before and after, renders both to PNG and pixel-
    diffs a window at (x,y) to prove the symbol is visible there (`visibility.visible`).
    `verify_dwg=True` additionally exports a DWG and checks it is non-empty; DWG is
    vector data with no pixel channel, so visibility is proven through the PDF path.
    """
    import shutil as _shutil
    import tempfile as _tempfile

    doc = _open_drawing(drawing_path)
    sym = SURFACE_FINISH_TYPES.get(str(sym_type).lower(), sym_type)
    leader = LEADER_STYLES.get(str(leader_type).lower(), leader_type)
    work_dir = None
    try:
        before_img = None
        visibility = None
        dwg_check = None
        if verify_render:
            try:
                work_dir = _tempfile.mkdtemp(prefix="assetmate_sf_")
                before_img = _sf_render_page(
                    doc, os.path.join(work_dir, "before.pdf"), work_dir)
            except Exception as exc:
                visibility = {"checked": False,
                              "reason": "PDF render unavailable: %s" % str(exc)[:120]}
        try:
            symbol = doc.Extension.InsertSurfaceFinishSymbol3(
                int(sym), int(leader), float(x), float(y), 0.0,
                int(lay_symbol), int(arrow_type),
                "", "", str(prod_method), str(sample_len),
                str(roughness), str(min_roughness), str(roughness_spacing),
            )
        except Exception as exc:
            return error("surface_finish_failed", str(exc),
                         "Check the 14-argument order; this signature is easy to mismatch.")
        if symbol is None:
            return error("surface_finish_failed",
                         "SolidWorks returned null for the surface finish symbol",
                         "Verify the point is on the sheet and the symbol type is valid.")

        # P2-11: the API put the symbol at (0,0). Move it and read the REAL position.
        position_actual_m = None
        position_applied = False
        try:
            annotation = symbol.GetAnnotation
            if annotation is not None:
                position_applied = bool(annotation.SetPosition(float(x), float(y), 0.0))
                position_actual_m = [round(float(v), 4) for v in annotation.GetPosition]
        except Exception:
            pass
        from ...sw_session import call_method as _call_method_sf
        try:
            _call_method_sf(doc, "EditRebuild3")
        except Exception:
            pass
        try:
            doc.GraphicsRedraw2()
        except Exception:
            pass
        readback = {}
        try:
            readback["annotation_name"] = str(symbol.GetAnnotation.GetName)
        except Exception:
            pass
        if save:
            _save_drawing(doc)

        # P2-11: render AFTER and diff against the before image.
        if verify_render and before_img is not None and visibility is None:
            try:
                after_img = _sf_render_page(
                    doc, os.path.join(work_dir, "after.pdf"), work_dir)
                visibility = _sf_visibility_check(doc, x, y, before_img, after_img)
            except Exception as exc:
                visibility = {"checked": False,
                              "reason": "PDF render unavailable: %s" % str(exc)[:120]}
        if verify_dwg:
            try:
                if work_dir is None:
                    work_dir = _tempfile.mkdtemp(prefix="assetmate_sf_")
                dwg_path = os.path.join(work_dir, "surface_finish.dwg")
                doc.SaveAsSilent(dwg_path, False)
                size = os.path.getsize(dwg_path) if os.path.exists(dwg_path) else 0
                dwg_check = {
                    "checked": True, "path": dwg_path, "bytes": size,
                    "nonempty": size > 1000,
                    "note": ("DWG is vector data — only file size is checked here; "
                             "pixel visibility is proven through the PDF path "
                             "(verify_render)."),
                }
            except Exception as exc:
                dwg_check = {"checked": False, "reason": str(exc)[:120]}
        return drawing_ok(
            created=True, saved=bool(save), roughness=str(roughness),
            sym_type=str(sym_type), leader_type=str(leader_type),
            position_m=[round(float(x), 4), round(float(y), 4)],
            position_actual_m=position_actual_m,
            position_applied=position_applied,
            visibility=visibility,
            dwg_check=dwg_check,
            readback=readback,
            note=("0.8.20：InsertSurfaceFinishSymbol3 会忽略传入的 x/y，符号实际落在图纸原点；"
                  "创建后用 IAnnotation.SetPosition 移到请求点并回读 position_actual_m。"
                  "verify_render=True 时前后各导出 PDF 渲染 PNG 做像素差分，证明符号确实"
                  "出现在请求位置（visible=true）；verify_dwg=True 只检查 DWG 文件非空"
                  "（矢量无像素通道，可见性以 PDF 为准）。"),
        )
    finally:
        if work_dir:
            _shutil.rmtree(work_dir, ignore_errors=True)


@tool_guard("gtol_failed")
def add_geometric_tolerance(drawing_path, x, y, tolerance="0.05", datum="A",
                            symbol="none", view_name=None, edge_index=None,
                            second_datum="", third_datum="", save=True):
    """Insert a NATIVE geometric tolerance frame (形位公差) into a drawing.

    Verified 2026-10-28 on SW2026: the API exists and works, but only when the call is
    dispatched as a METHOD. `doc.InsertGtol()` goes through PROPERTYGET and raises
    DISP_E_MEMBERNOTFOUND — which is why four earlier attempts concluded "impossible".

    x/y are sheet metres. Attach to geometry by passing view_name + edge_index (the same
    index `list_drawing_edges` returns); without it the frame is free-standing.
    """
    import win32com.client as win

    doc = _open_drawing(drawing_path)

    attached = None
    if view_name and edge_index is not None:
        for view in _iter_views(doc):
            if getattr(view, "Type", None) == 1 or view.Name != view_name:
                continue
            try:
                doc.ActivateView(view.Name)
                doc.ClearSelection2(True)
                edges = []
                try:
                    for component in (list(view.GetVisibleComponents or ())):
                        try:
                            edges.extend(list(view.GetVisibleEntities2(component, 1) or ()))
                        except Exception:
                            continue
                except Exception:
                    edges = []
                if edge_index < 0 or edge_index >= len(edges):
                    return error("bad_edge_index", "edge_index %s out of range (%d edges)"
                                 % (edge_index, len(edges)))
                entity = win.Dispatch(edges[edge_index])
                if entity.Select2(False, 0):
                    attached = "%s[%d]" % (view.Name, edge_index)
            except Exception:
                attached = None
            break

    raw = _invoke_as_method(doc, "InsertGtol")
    if raw is None:
        return error("gtol_failed", "InsertGtol returned null",
                     "Select an edge/face first; without a selection the frame is free-standing.")
    gtol = win.Dispatch(raw)

    code = GTOL_SYMBOLS.get(str(symbol).lower(), "")
    symbol_written = ""
    try:
        gtol.SetFrameSymbols2(1, code, False, "", False, "", "", "", "")
        # P2-10: SetFrameSymbols2 returns None even on success — the ONLY reliable
        # proof is a read-back. A "<MOD-*" code stored as literal text (or an ignored
        # write) must not be reported as a real geometric symbol.
        symbol_written = str(gtol.GetFrameSymbols2(1)[0] or "")
    except Exception as exc:
        return error("gtol_symbol_failed", str(exc)[:120])
    if code and symbol_written != code:
        # Refuse to leave a half-built frame that claims a symbol it does not show.
        removed = _delete_gtol(doc, gtol)
        return error(
            "gtol_symbol_not_applied",
            "Requested symbol %r but the first frame cell reads back %r — the symbol "
            "was not applied on this SolidWorks build." % (code, symbol_written),
            "Use symbol='none' to create a values-only frame, or add the geometric "
            "symbol manually in the UI. Do not treat this frame as a complete GD&T "
            "callout.",
            symbol_requested=code, symbol_readback=symbol_written,
            symbol_is_text=True, frame_removed=removed)
    try:
        wrote = bool(gtol.SetFrameValues2(1, str(tolerance), "", str(datum),
                                          str(second_datum), str(third_datum)))
    except Exception as exc:
        return error("gtol_value_failed", str(exc)[:120])

    position_ok = None
    try:
        annotation = gtol.GetAnnotation
        position_ok = bool(annotation.SetPosition(float(x), float(y), 0.0))
    except Exception:
        annotation = None

    try:
        doc.EditRebuild3()
    except Exception:
        pass
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass

    # Read back field by field: one unsupported accessor must not blank the whole block.
    readback = {"symbol_written": symbol_written}
    if annotation is not None:
        try:
            readback["name"] = str(annotation.GetName)
        except Exception:
            pass
        try:
            readback["position_m"] = [round(v, 4) for v in annotation.GetPosition]
        except Exception:
            pass
    for getter, key in (("GetFrameCount", "frame_count"),):
        try:
            readback[key] = int(getattr(gtol, getter))
        except Exception:
            pass
    for getter in ("GetFrameValues2", "GetFrameValues"):
        try:
            values = getattr(gtol, getter)(1)
            readback["frame1"] = [str(v) for v in values]
            break
        except Exception:
            continue

    if save:
        _save_drawing(doc)
    return drawing_ok(
        created=True, saved=bool(save), attached_to=attached,
        tolerance=str(tolerance), datum=str(datum), symbol=str(symbol),
        values_written=wrote, position_set=position_ok, readback=readback,
        symbol_written=symbol_written, symbol_is_text=bool(code),
        note=("形位公差是原生对象，可用 list_drawing_annotations 读出（kind="
              "geometric_tolerance）。0.8.20：符号格以可读文字写入第一格并回读校验"
              "（<MOD-*> 格式码在本环境只会被当字面文本，不解析成几何符号）；"
              "写入不生效会删除框格并返回 gtol_symbol_not_applied。"
              "symbol=none 时不写符号，只有公差值与基准。"),
    )


