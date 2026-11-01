from __future__ import annotations

"""View tools: flat pattern, section, detail, edge/circle listing."""

from ._core import *

@tool_guard("flat_pattern_failed")
def add_flat_pattern_view(drawing_path, model_path, config_name=None, position=None,
                         scale=None, hide_bend_lines=False, flip=False):
    """Add a sheet-metal FLAT PATTERN view (the unfolded blank with bend lines).

    Sheet metal is dimensioned on its flat pattern — blank size, hole positions, bend
    directions — not on the folded view. The part must have a Flat-Pattern feature and
    be saved to disk.

    CreateFlatPatternViewFromModelView3(ModelName, ConfigName, LocX, LocY, LocZ,
    HideBendLines, FlipView) — 7 arguments, no config guessing: pass config_name or the
    part's active configuration is used.
    """
    from ...sw_session import get_sw, open_document, SW_DOC_PART

    model_path = os.path.abspath(model_path)
    if not os.path.exists(model_path):
        raise FileNotFoundError("file not found: %s" % model_path)
    if not model_path.lower().endswith(".sldprt"):
        return error("bad_input", "Flat pattern needs a sheet-metal .SLDPRT",
                     "Assemblies cannot be unfolded; pass the sheet-metal part.")
    doc = _open_drawing(drawing_path)
    sheet = doc.GetFirstView
    outline = list(sheet.GetOutline or [])
    sheet_w = (outline[2] - outline[0]) if outline else 0.42
    sheet_h = (outline[3] - outline[1]) if outline else 0.297

    if not config_name:
        try:
            model = open_document(model_path, SW_DOC_PART)
            config_name = model.ConfigurationManager.ActiveConfiguration.Name
        except Exception:
            config_name = "Default"
    config_name = config_name or "Default"

    if position:
        px, py = float(position[0]), float(position[1])
    else:
        safety = template_safe_area(sheet, sheet_w, sheet_h)
        px = safety["left"] + (sheet_w - safety["left"] - safety["right"]) / 2
        py = safety["bottom"] + (sheet_h - safety["bottom"] - safety["top"]) / 2

    view = doc.CreateFlatPatternViewFromModelView3(
        model_path, config_name, px, py, 0.0, bool(hide_bend_lines), bool(flip)
    )
    if view is None:
        return error(
            "flat_pattern_failed",
            "SolidWorks returned null for '%s' (config '%s')" % (
                os.path.basename(model_path), config_name),
            "The part must be SHEET METAL (have a Flat-Pattern feature) and saved to disk.",
        )
    if scale:
        try:
            view.ScaleDecimal = float(scale)
        except Exception:
            pass
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass
    layout = _relayout(doc)
    _save_drawing(doc)
    try:
        o = list(view.GetOutline or [])
        box = {"width_m": round(o[2] - o[0], 4), "height_m": round(o[3] - o[1], 4)}
    except Exception:
        box = None
    return drawing_ok(
        drawing=doc.GetTitle, view=view.Name, config=config_name,
        scale=float(view.ScaleDecimal), outline=box, layout=layout,
        hide_bend_lines=bool(hide_bend_lines), flip=bool(flip),
    )


@tool_guard("section_failed")
def add_section_view(drawing_path, view_name=None, orientation="horizontal",
                     offset_m=0.0, position=None, label="A", flip=False, depth=0.0,
                     cut_plane_normal=None, cut_plane_origin=None, verify_cut=True):
    """Create a SECTION VIEW through a model view.

    Sign convention (measured on the 灌装阀阀体 top view, 59x28x80, bypass hole at z=28):

      orientation="vertical"   cut line runs along the view's vertical axis, so offset_m
                               shifts along the *horizontal* axis. Negative = left.
                               A hole axis 15.55 mm left of centre needs offset_m=-0.01555.
      orientation="horizontal" cut line runs along the horizontal axis, so offset_m shifts
                               along the *vertical* axis — BUT the view sketch's y axis
                               points at -Z for a top view, so the sign is inverted inside
                               this function. Callers pass the model-space offset: to cut
                               at z=28 on a 0..80 part (centre 40) you pass -0.012, and
                               the tool negates it internally. Measured: -0.012 -> 59x80
                               (correct cross-section), +0.012 -> 14x50 (wrong plane).

    **Cut lines do not follow a moved parent view, and cannot be moved afterwards**
    (measured 2026-11-01: `IDrSection.GetLineInfo` is model-space and does not change when the
    view moves; `SetLineInfo` writes the section PLANE, not the drawn line; the view's sketch
    is unreachable). Place the parent view FIRST, then create the section, then never move it —
    or delete and rebuild. See check_capabilities `section_line_follows_view`.

    Both orientations take `offset_m` as model-space distance from the part centre, i.e.
    `offset_m = (target_model_coordinate - centre) * scale` in SHEET metres.

    The cut line is drawn in the view's own sketch — the view sketch coordinate system is
    `sketch = (paper - view.Position) / view.ScaleDecimal` (calibrated with Convert
    Entities). SolidWorks requires the line to live in that sketch and to be selected
    before CreateSectionViewAt5, otherwise it returns null.

    orientation: "horizontal" cuts left-right (shows the vertical section), "vertical"
    cuts top-bottom. offset_m shifts the line off the view centre (sheet metres).
    flip reverses the viewing direction (Options = 4). depth > 0 gives a partial section.

    P1-4 (0.8.21) — two additions, both aimed at "one wrong parameter must not require
    rebuilding the whole sheet":

    cut_plane_normal / cut_plane_origin
        MODEL coordinates (metres), given together. They override orientation + offset_m:
        the normal picks the cut direction (dominant part axis, must be >= 0.9 aligned;
        oblique planes have no single-line sketch representation and are refused) and the
        origin supplies the position. Mapping model -> view requires knowing which model
        axis each sheet axis shows; that is matched by size against the part box. Refusals
        are structured: `cut_plane_incomplete`, `cut_plane_oblique`,
        `cut_plane_not_in_view_plane`, `cut_plane_unmapped` (with `mapping_reason` such as
        `axes_ambiguous` for a cube, where guessing would silently cut the wrong plane).

    verify_cut (default True)
        After creation, compare the section's deflated outline against the source view's:
        a cut perpendicular to the sheet's horizontal axis cannot change the width, so
        `section_width / source_width` must be 1.0 ± 10% (height for vertical cuts).
        `edge_count == 0` means an empty view. Both failures return errors
        (`section_empty`, `section_suspicious`) WITH the view kept on the sheet, so you can
        inspect it instead of re-cutting blindly. Partial sections (depth>0) report
        `sanity.checked=false` because they shrink by design.
    """
    import win32com.client as win

    doc = _open_drawing(drawing_path)
    views = [v for v in _iter_views(doc) if getattr(v, "Type", None) != 1]
    if not views:
        return error("no_views", "The drawing has no model view to section")
    target = None
    for candidate in views:
        if view_name is None or candidate.Name == view_name:
            target = candidate
            break
    if target is None:
        return error("view_not_found", "No view named %r" % view_name,
                     "Use drawing_audit to list view names.")

    doc.ActivateView(target.Name)
    pos = list(target.Position)
    scale = float(target.ScaleDecimal)
    outline = list(target.GetOutline)
    width = outline[2] - outline[0]
    height = outline[3] - outline[1]

    # P1-4: a cut plane in model coordinates has to be translated into the view sketch.
    cut_plane = None
    if cut_plane_normal is not None or cut_plane_origin is not None:
        if cut_plane_normal is None or cut_plane_origin is None:
            return error("cut_plane_incomplete",
                         "cut_plane_normal and cut_plane_origin must be passed together")
        try:
            normal = [float(v) for v in cut_plane_normal]
            origin = [float(v) for v in cut_plane_origin]
        except Exception:
            return error("bad_input", "cut_plane_normal/origin must be 3 numbers each")
        if len(normal) != 3 or len(origin) != 3:
            return error("bad_input", "cut_plane_normal/origin must be 3 numbers each")
        length = sum(v * v for v in normal) ** 0.5
        if length <= 1e-9:
            return error("bad_input", "cut_plane_normal has zero length")
        unit = [v / length for v in normal]

        axis_map, reason = _view_axis_map(target)
        # _view_axis_map reads the part through open_document, which switches the
        # active document; bring the drawing back before touching its sketch.
        doc = _open_drawing(drawing_path)
        doc.ActivateView(target.Name)
        if axis_map is None:
            return error(
                "cut_plane_unmapped",
                "Cannot map model coordinates onto view %r (%s)" % (target.Name, reason),
                "Pass orientation + offset_m instead, or pick a view whose displayed "
                "axes match the part box unambiguously.",
                cut_plane_normal=normal, cut_plane_origin=origin,
                mapping_reason=reason)
        axes = axis_map["axes"]
        centre_m = axis_map["centre_m"]
        dominant = max(range(3), key=lambda i: abs(unit[i]))
        if abs(unit[dominant]) < 0.9:
            return error(
                "cut_plane_oblique",
                "cut_plane_normal %r is not axis-aligned (dominant component %.3f)"
                % (normal, abs(unit[dominant])),
                "Only planes perpendicular to a part axis are supported — an oblique "
                "plane has no single-line representation in the view sketch.",
                cut_plane_normal=normal, cut_plane_origin=origin,
                mapping=axis_map)
        if dominant == axes[0]:
            orientation = "vertical"
        elif dominant == axes[1]:
            orientation = "horizontal"
        else:
            return error(
                "cut_plane_not_in_view_plane",
                "cut_plane_normal is parallel to the viewing direction of %r"
                % target.Name,
                "A plane parallel to the sheet cannot be drawn as a cut line; pick a "
                "normal along one of the two displayed axes.",
                cut_plane_normal=normal, cut_plane_origin=origin,
                mapping=axis_map)
        offset_m = (origin[dominant] - centre_m[dominant]) * scale
        cut_plane = {
            "normal_model": [round(v, 6) for v in normal],
            "origin_model": [round(v, 6) for v in origin],
            "dominant_axis": int(dominant),
            "derived_orientation": str(orientation),
            "derived_offset_m": round(float(offset_m), 6),
            "mapping": axis_map,
        }
    doc.ActivateView(target.Name)

    # `width`/`height` come from GetOutline and are ALREADY sheet metres — dividing by
    # scale again produced a 1035 mm line on a 420 mm sheet (measured 2026-11-01:
    # 0.1669 * 0.62 / 0.1 = 1.0347, which is exactly the ±1.0347 that GetLineInfo
    # reported). Use the deflated outline so the 11.9 mm phantom margin does not
    # stretch the line either.
    precise_box = precise_outline(target) or outline
    half = max(precise_box[2] - precise_box[0],
               precise_box[3] - precise_box[1]) * 0.62
    offset = float(offset_m) / scale
    # A top view's sketch y axis runs at -Z, so a model-space offset applied to the
    # horizontal cut line lands mirrored. Measured on 灌装阀阀体: offset -0.012 produced
    # a 59x80 cross-section while +0.012 produced 14x50. Negate here so callers can
    # think in model coordinates for both orientations.
    if str(orientation).lower() != "vertical":
        offset = -offset

    sm = doc.SketchManager
    sm.InsertSketch(True)
    try:
        # The sketch lives on the SHEET, so the line needs the view's own position added.
        # Without `pos` it was drawn at x = 0 (the sheet edge) and only its middle
        # happened to cross the view — that is the "cut line is nowhere near the view"
        # the report saw.
        if str(orientation).lower() == "vertical":
            sm.CreateLine(pos[0] + offset, pos[1] - half, 0.0,
                          pos[0] + offset, pos[1] + half, 0.0)
        else:
            sm.CreateLine(pos[0] - half, pos[1] + offset, 0.0,
                          pos[0] + half, pos[1] + offset, 0.0)
    except Exception as exc:
        try:
            sm.InsertSketch(True)
        except Exception:
            pass
        return error("cut_line_failed", str(exc),
                     "Could not draw the cut line in the view sketch.")

    raw = sm.ActiveSketch.GetSketchSegments
    segments = list(raw) if isinstance(raw, (tuple, list)) else ([raw] if raw else [])
    segments = [s for s in segments if s is not None]
    selected = False
    for segment in segments:
        try:
            seg = win.Dispatch(segment)
            if seg.Select2(False, 0) and doc.SelectionManager.GetSelectedObjectCount2(-1):
                selected = True
                break
        except Exception:
            continue
    sm.InsertSketch(True)
    if not selected:
        return error("cut_line_not_selected",
                     "Drew the cut line but could not select it",
                     "SolidWorks needs the cut line selected before creating the section.")

    if position:
        place_x, place_y = float(position[0]), float(position[1])
    else:
        # Default to the free area at the sheet's right side. Placing the section
        # directly under the source view pushes it into the parts list once that
        # column starts at ~14% of the width.
        sheet_box = list(doc.GetFirstView.GetOutline)
        sheet_w = sheet_box[2] - sheet_box[0]
        sheet_h = sheet_box[3] - sheet_box[1]
        # A section stays x-aligned with its parent, so it must go above/below it.
        place_x = outline[0] + width / 2
        place_y = max(sheet_h * 0.14, outline[1] - max(height * 0.55, 0.06))

    options = 4 if flip else 0  # swCreateSectionView_ChangeDirection
    section = doc.CreateSectionViewAt5(
        place_x, place_y, 0.0, str(label), options, None, float(depth)
    )
    if section is None:
        return error(
            "section_failed",
            "CreateSectionViewAt5 returned null for %r" % target.Name,
            "The cut line must lie inside the view and cross its geometry; "
            "or an earlier failed attempt left the view in a bad state — reopen the drawing.",
            cut_plane=cut_plane,
        )
    try:
        section.ScaleDecimal = scale
    except Exception:
        pass
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass
    # Deliberately no repack here: repacking moves the section back into the parts
    # list whenever the sheet is tight. Call drawing_audit to check the result and
    # pass `position` to place it precisely.
    layout = None
    _save_drawing(doc)
    try:
        so = list(section.GetOutline)
        box = {"width_m": round(so[2] - so[0], 4), "height_m": round(so[3] - so[1], 4)}
    except Exception:
        box = None
    precise = precise_outline(section)
    edge_count = _view_edge_count(section)
    sanity = (_section_sanity(target, section, orientation, depth)
              if verify_cut else
              {"checked": False, "reason": "verify_cut=False"})
    diag = {
        "drawing": doc.GetTitle, "section": section.Name, "source_view": target.Name,
        "label": str(label), "orientation": str(orientation), "flip": bool(flip),
        "offset_m": round(float(offset_m), 6), "outline": box,
        "precise_outline_m": ([round(v, 5) for v in precise] if precise else None),
        "edge_count": edge_count, "sanity": sanity, "cut_plane": cut_plane,
        "layout": layout,
    }
    if edge_count == 0:
        return error(
            "section_empty",
            "The section view has 0 visible edges — the cut line missed the geometry.",
            "Move `offset_m` inside the view's outline, or pass cut_plane_origin at "
            "real geometry. The empty view is left on the sheet for inspection.",
            **diag)
    if sanity.get("checked") and sanity.get("verdict") == "suspicious":
        return error(
            "section_suspicious",
            "The section kept %s mm of %s mm (ratio %.3f) — the cut plane probably "
            "missed the intended cross-section."
            % (sanity.get("section_kept_mm"), sanity.get("source_kept_mm"),
               sanity.get("ratio") or 0.0),
            "Check offset_m sign and magnitude against the source view, or pass "
            "cut_plane_origin in model coordinates. The view is kept for inspection.",
            **diag)
    return drawing_ok(**diag)


@tool_guard("detail_failed")
def add_detail_view(drawing_path, view_name=None, center=None, radius_m=0.02,
                    position=None, scale=2.0, label="A", style=0, shape=0,
                    flip=False, show_label=True, circle_segments=32):
    """Add a DETAIL (enlarged) view around a circular area of a model view.

    CreateDetailViewAt4 takes TWELVE arguments — the last one is the circle's segment
    count. Passing 0 there raises "non-optional parameter" and no view is created, which
    is what made this look unsupported. 32 works.

    center/radius_m are SHEET metres — point the boundary at real detail (a hole, a
    corner, a bend), otherwise the enlarged view comes out empty. The circle is drawn in
    the source view's sketch
    (coordinates = (paper - view.Position) / view.ScaleDecimal) and selected before the
    call — otherwise SolidWorks has no boundary to enlarge.

    P2-9 (0.8.20): `scale` is the REQUESTED detail scale and is no longer overwritten by
    the parent's scale (that bug made a 2:1 request silently come out 1:2). After creation
    the new view's visible edges are counted: 0 means the boundary missed the geometry and
    the call returns `detail_empty`. If SolidWorks keeps the parent scale instead of the
    requested one, the call returns `detail_scale_locked` rather than ok:true with a quiet
    `scale_locked` flag.
    """
    import win32com.client as win

    # P2-9: `scale` is the REQUESTED detail scale. Never overwrite it with the
    # parent view's scale — that old bug made a 2:1 request silently come out 1:2.
    requested_scale = float(scale)
    doc = _open_drawing(drawing_path)
    views = [v for v in _iter_views(doc) if getattr(v, "Type", None) != 1]
    if not views:
        return error("no_views", "The drawing has no model view to enlarge")
    target = None
    for candidate in views:
        if view_name is None or candidate.Name == view_name:
            target = candidate
            break
    if target is None:
        return error("view_not_found", "No view named %r" % view_name,
                     "Use drawing_audit to list view names.")

    doc.ActivateView(target.Name)
    pos = list(target.Position)
    parent_scale = float(target.ScaleDecimal)
    outline = list(target.GetOutline)
    if center:
        sheet_x, sheet_y = float(center[0]), float(center[1])
    else:
        sheet_x = (outline[0] + outline[2]) / 2
        sheet_y = (outline[1] + outline[3]) / 2
    # The boundary circle is drawn in the PARENT view's sketch, so convert with the
    # parent's scale, not the detail's requested one.
    sketch_x = (sheet_x - pos[0]) / parent_scale
    sketch_y = (sheet_y - pos[1]) / parent_scale
    sketch_r = float(radius_m) / parent_scale

    # P1-2: decide where the detail view will live BEFORE drawing the boundary, so a
    # crowded sheet returns `no_space` with the drawing untouched instead of producing
    # a view that hangs off the edge or stacks on its neighbour.
    if position:
        place_x, place_y = float(position[0]), float(position[1])
    else:
        ratio = requested_scale / max(parent_scale, 1e-6)
        est_half_w = float(radius_m) * ratio * 1.3 + 0.008
        est_half_h = float(radius_m) * ratio * 1.05 + 0.014
        spot = _free_spot_for_view(
            doc, est_half_w, est_half_h,
            prefer_x=min(max(outline[2] + est_half_w + 0.02, est_half_w + 0.01),
                         0.42 - est_half_w - 0.01),
            prefer_y=max(est_half_h + 0.01, (outline[1] + outline[3]) / 2),
        )
        if spot is None:
            return error(
                "no_space",
                "图纸放不下约 %.2f×%.2f m 的局部视图（%g:1）"
                % (est_half_w * 2, est_half_h * 2, requested_scale),
                "已有视图：%s。可减小 scale / radius_m、显式传 position，或换更大的图幅。"
                % _view_outline_hint(doc),
                estimated_size_m=[round(est_half_w * 2, 4), round(est_half_h * 2, 4)],
                requested_scale=requested_scale,
            )
        place_x, place_y = spot

    sm = doc.SketchManager
    sm.InsertSketch(True)
    try:
        sm.CreateCircleByRadius(sketch_x, sketch_y, 0.0, sketch_r)
    except Exception as exc:
        try:
            sm.InsertSketch(True)
        except Exception:
            pass
        return error("boundary_failed", str(exc),
                     "Could not draw the detail boundary circle in the view sketch.")

    raw = sm.ActiveSketch.GetSketchSegments
    segments = list(raw) if isinstance(raw, (tuple, list)) else ([raw] if raw else [])
    segments = [s for s in segments if s is not None]
    selected = False
    for segment in segments:
        try:
            seg = win.Dispatch(segment)
            if seg.Select2(False, 0) and doc.SelectionManager.GetSelectedObjectCount2(-1):
                selected = True
                break
        except Exception:
            continue
    sm.InsertSketch(True)
    if not selected:
        return error("boundary_not_selected",
                     "Drew the detail circle but could not select it",
                     "SolidWorks needs the circle selected as the detail boundary.")

    detail = doc.CreateDetailViewAt4(
        place_x, place_y, 0.0, int(style), float(parent_scale), float(parent_scale),
        str(label), bool(show_label), bool(flip), False, int(shape),
        int(circle_segments),
    )
    if detail is None:
        return error(
            "detail_failed",
            "CreateDetailViewAt4 returned null for %r" % target.Name,
            "The boundary circle must lie inside the view geometry; reopen the "
            "drawing if earlier attempts left it in a bad state.",
        )
    # A detail view inherits the parent's scale (UseParentScale = True by default) and
    # silently ignores ScaleDecimal until that flag is cleared. Setting both right
    # after creation can be swallowed, so retry against the REQUESTED scale until the
    # read-back matches. Measured: _apply_detail_scale(2.0) does take effect and the
    # view is renamed "局部视图 A (2 : 1)".
    applied = _apply_detail_scale(doc, detail, requested_scale)
    if applied is None or abs(applied - requested_scale) > 0.01:
        # A freshly created detail view refuses the scale until it is committed once.
        try:
            doc.SaveSilent
        except Exception:
            pass
        applied = _apply_detail_scale(doc, detail, requested_scale)
        try:
            doc.SaveSilent
        except Exception:
            pass
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass
    _save_drawing(doc)
    try:
        d_o = list(detail.GetOutline)
        box = {"width_m": round(d_o[2] - d_o[0], 4), "height_m": round(d_o[3] - d_o[1], 4)}
    except Exception:
        box = None
    # P2-13: report the deflated box too — a detail view is smaller than the source, so
    # the fixed 11.9 mm inflation is a much larger relative error here (a 30 mm detail
    # reads 41.9 mm, i.e. 40% too wide).
    precise = precise_outline(detail)
    final_scale = applied if applied is not None else float(detail.ScaleDecimal)
    locked = abs(float(final_scale) - requested_scale) > 0.01

    # P1-2: the view exists, so its real box is known. Move it somewhere that neither
    # overlaps a neighbour nor leaves the sheet; if the sheet is full, remove the view
    # we just made and fail loudly — a view hanging off the paper is worse than a clear
    # `no_space`, because drawing_audit only catches it later.
    placement = {"requested": [round(place_x, 4), round(place_y, 4)], "x_locked": True}
    try:
        if precise and len(precise) >= 4:
            placement["measured"] = [round(v, 4) for v in precise]
            inside = (precise[0] >= -0.001 and precise[1] >= -0.001
                      and precise[2] <= 0.421 and precise[3] <= 0.298)
            placement["inside_sheet"] = bool(inside)
    except Exception:
        pass

    # P2-9: a boundary circle that missed the geometry produces an EMPTY detail view.
    # `GetVisibleComponents` is None in that case, so _view_edge_count is 0 (measured
    # on door.SLDPRT: boundary inside the parent outline -> 1 edge, outside -> 0).
    edge_count = _view_edge_count(detail)
    diag = {
        "detail": detail.Name,
        "source_view": target.Name,
        "boundary_center_m": [round(sheet_x, 4), round(sheet_y, 4)],
        "boundary_radius_m": float(radius_m),
        "source_outline_m": [round(v, 4) for v in outline],
        "edge_count": edge_count,
        "scale": final_scale,
        "requested_scale": requested_scale,
        "parent_scale": parent_scale,
        "outline": box,
        "precise_outline_m": ([round(v, 5) for v in precise] if precise else None),
        "placement": placement,
    }
    if edge_count == 0:
        return error(
            "detail_empty",
            "The detail view was created but has 0 visible edges — its boundary "
            "circle fell outside the source view's geometry.",
            "Point `center` at real geometry (a hole, corner or bend) inside the "
            "source view's outline; use drawing_audit / list_drawing_edges to find it.",
            **diag)
    if locked:
        return error(
            "detail_scale_locked",
            "The detail view kept scale %.4g, not the requested %.4g — SolidWorks "
            "locked it to the parent view's scale."
            % (float(final_scale), requested_scale),
            "Save and reopen the drawing, then set the scale in the UI; or pass a "
            "scale equal to the parent's to create without a scale conflict.",
            scale_locked=True,
            scale_note=("局部视图默认继承父视图比例（UseParentScale）；本次请求的 "
                        "%.4g 未生效，实际 %.4g。" % (requested_scale, float(final_scale))),
            **diag)
    # diag already carries every measured field (detail, source_view, boundary_*,
    # edge_count, scale, requested_scale, parent_scale, outline, precise_outline_m),
    # so success and failure report exactly the same numbers.
    return drawing_ok(drawing=doc.GetTitle, label=str(label),
                      scale_locked=False, scale_note=None, **diag)


@tool_guard("list_edges_failed")
def list_drawing_edges(drawing_path, view_name=None, include_points=True):
    """List a view's visible edges with stable indices, so dimensions can be added
    without guessing sheet coordinates (which need a model->sheet transform this binding
    exposes unreliably).

    Performance notes — this used to be the slowest call in the toolkit (measured 335 s on
    a 1094-edge assembly view), almost all of it in fetching vertex coordinates:

        direct `edge.GetStartVertex`       335 s   (property dispatch, 3 COM calls per edge)
        sw_session.call_method             179 s   (DISPATCH_METHOD, 2x faster)
        GetCurveParams2 + evaluate          28 s   (12x faster, but only gives parameters)

    Two mitigations, both measured:

    * `include_points=False` returns indices only. Enumerating edges costs ~66 s on that
      view; adding coordinates is what pushes it past 300 s. Callers that only need to pick
      an edge by index should not pay for coordinates.
    * Results are cached per (drawing, view, include_points). The second call is instant.
      The cache is dropped when the drawing's mtime changes or after EDGE_CACHE_LIMIT
      entries, so a re-saved drawing re-reads instead of serving stale geometry.
    """
    import win32com.client as win

    from ...sw_session import call_method

    path = os.path.abspath(drawing_path)
    try:
        stamp = (path, os.path.getmtime(path), view_name, bool(include_points))
    except OSError:
        stamp = None
    if stamp is not None and stamp in _EDGE_CACHE:
        cached = dict(_EDGE_CACHE[stamp])
        cached["cached"] = True
        return ok(**cached)

    doc = _open_drawing(path)
    target = None
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        if view_name is None or view.Name == view_name:
            target = view
            break
    if target is None:
        return error("view_not_found", "No model view found (view_name=%r)" % view_name)

    rows = []
    raw_edges = []
    for component in (target.GetVisibleComponents or ()):
        try:
            raw_edges.extend(list(target.GetVisibleEntities2(component, 1) or ()))
        except Exception:
            continue
    if not raw_edges:
        # Some views return nothing per-component; fall back to the view itself.
        try:
            raw_edges = list(target.GetVisibleEntities2(target, 1) or ())
        except Exception:
            raw_edges = []

    for index, raw in enumerate(raw_edges):
        item = {"index": index}
        if include_points:
            try:
                edge = win.Dispatch(raw)
                # call_method: attribute access is dispatched as PROPERTYGET and costs
                # roughly twice as much per edge.
                p1 = win.Dispatch(call_method(edge, "GetStartVertex")).GetPoint
                p2 = win.Dispatch(call_method(edge, "GetEndVertex")).GetPoint
                item["model_start"] = [round(v, 4) for v in p1]
                item["model_end"] = [round(v, 4) for v in p2]
            except Exception:
                pass
        rows.append(item)

    payload = {"drawing": doc.GetTitle, "view": target.Name,
               "edge_count": len(rows), "edges": rows}
    if stamp is not None:
        if len(_EDGE_CACHE) >= EDGE_CACHE_LIMIT:
            _EDGE_CACHE.clear()
        _EDGE_CACHE[stamp] = dict(payload)
    result = ok(**payload)
    result["cached"] = False
    return result


@tool_guard("circle_scan_failed")
def list_circular_edges(drawing_path, view_name=None, min_diameter_mm=0.0):
    """Circular edges of a view, with sheet position and diameter.

    Reads ONLY `ICurve.CircleParams` — never vertices. That matters: a view's visible
    entities are SILHOUETTE edges and `GetStartVertex` returns None for every one of them,
    so the previous vertex-based scan answered "0 circles" on a valve body that visibly had
    thirteen (measured 2026-10-31). CircleParams returns centre, axis and radius in one
    call, which is also cheaper than two vertex reads per edge.

    Sheet positions use the same size-matched axis mapping as `add_section_view`: the
    model's own bounding box, not a vertex bounding box. When the mapping is unavailable
    (an ambiguous box, a section/detail view, a missing referenced model) the row keeps its
    diameter and reports `x: null` plus `position_unavailable` — a made-up coordinate is
    worse than none, because callers feed it straight into `add_hole_callout`.
    """
    import win32com.client as win

    doc = _open_drawing(drawing_path)
    out = []
    for view in _iter_views(doc):
        if getattr(view, "Type", None) == 1:
            continue
        if view_name and view.Name != view_name:
            continue
        try:
            edges = [win.Dispatch(e) for e in (view.GetVisibleEntities2(view, 1) or ())]
        except Exception:
            continue

        circles = []
        for edge in edges:
            try:
                curve = win.Dispatch(edge.GetCurve)
                if curve is None or not curve.IsCircle:
                    continue
                params = list(curve.CircleParams)
                radius_m = float(params[6])
                diameter = round(radius_m * 2000, 2)
                if diameter < float(min_diameter_mm):
                    continue
                circles.append({
                    "diameter_mm": diameter,
                    "centre_model": [float(params[0]), float(params[1]),
                                     float(params[2])],
                    # Circle normal. A circle is a hole MOUTH only when its normal is
                    # (anti)parallel to the direction the view looks along; otherwise
                    # you are seeing the cylinder wall, and AddHoleCallout2 refuses.
                    "axis": [float(params[3]), float(params[4]), float(params[5])],
                })
            except Exception:
                continue
        if not circles:
            continue

        axis_map, reason = _view_axis_map(view)
        # Reading the part box switches SolidWorks' active document; come back.
        doc = _open_drawing(drawing_path)
        try:
            doc.ActivateView(view.Name)
        except Exception:
            pass
        pos = list(view.Position)
        scale = float(view.ScaleDecimal)

        for item in circles:
            row = {"view": view.Name, "diameter_mm": item["diameter_mm"],
                   "x": None, "y": None}
            depth_axis = None
            if axis_map:
                axes = axis_map["axes"]
                centre_m = axis_map["centre_m"]
                row["x"] = round(pos[0] + (item["centre_model"][axes[0]]
                                           - centre_m[axes[0]]) * scale, 5)
                row["y"] = round(pos[1] + (item["centre_model"][axes[1]]
                                           - centre_m[axes[1]]) * scale, 5)
                # The two axes the view DISPLAYS leave exactly one it looks along.
                depth_axis = 3 - axes[0] - axes[1]
            else:
                row["position_unavailable"] = reason

            # ---- annotatability, predicted instead of discovered ----------------
            # A report burned 13 guaranteed-failing calls on an outline view because
            # nothing in this list said which circles can take a callout. The rule is
            # geometric: the circle's normal must run along the view direction.
            row["view_kind"] = _view_kind(view)
            axis_vec = item.get("axis") or [0.0, 0.0, 0.0]
            if depth_axis is not None and any(abs(v) > 1e-9 for v in axis_vec):
                dominant = max(range(3), key=lambda i: abs(axis_vec[i]))
                faces_viewer = (dominant == depth_axis
                                and abs(axis_vec[depth_axis]) >= 0.9)
                row["facing"] = "front" if faces_viewer else "side"
                row["annotatable_hint"] = faces_viewer
            else:
                row["facing"] = "unknown"
                row["annotatable_hint"] = "unknown"
            out.append(row)
    return drawing_ok(drawing=doc.GetTitle, count=len(out), circles=out)


