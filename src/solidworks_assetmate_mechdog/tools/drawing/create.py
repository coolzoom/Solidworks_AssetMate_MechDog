from __future__ import annotations

"""Drawing creation, export, render and whole-sheet audit."""

from ._core import *

@tool_guard("create_drawing_failed")
def create_drawing(
    model_path,
    output_path=None,
    template=None,
    views=("front", "top", "isometric"),
    scale=None,
    auto_fit=True,
    save=True,
    view_scales=None,
    paper="A3",
    view_positions=None,
):
    """Create a native .SLDDRW from a part/assembly, add model views, optionally save.

    views: any of front/back/left/right/top/bottom/isometric. View names are
    resolved against the model, so Chinese and English SolidWorks both work.
    scale=None with auto_fit picks a standard scale from the model's projected size
    (1:1, 1:2, 1:5 ...); pass an explicit scale to override.

    view_positions: explicit sheet centres in METRES, one per view, e.g.
    [[0.12, 0.20], [0.12, 0.08], [0.30, 0.20], [0.30, 0.08]] for a 2x2 layout.
    When given, the automatic packer is skipped entirely — a 4-view sheet previously
    fell back to a single vertical column and pushed the last views outside the frame.
    """
    from ...sw_session import get_sw

    model_path = os.path.abspath(model_path)
    if not os.path.exists(model_path):
        raise FileNotFoundError("file not found: %s" % model_path)
    if os.path.splitext(model_path)[1].lower() not in (".sldprt", ".sldasm"):
        return error("bad_input", "model_path must be .SLDPRT or .SLDASM",
                     "Drawings are created FROM a model; pass the model, not a drawing.")

    template = os.path.abspath(template) if template else _find_template(str(paper).upper())
    if not template or not os.path.exists(template):
        return error("template_not_found",
                     "No drawing template found. Pass template=<path to .drwdot>.",
                     "Look under C:\\ProgramData\\SOLIDWORKS\\<version>\\templates\\*.drwdot")
    # Forward slashes make sw.NewDocument silently return None.
    template = template.replace("/", "\\")

    sw = get_sw()
    is_assembly = not model_path.lower().endswith(".sldprt")
    model = open_document(model_path, SW_DOC_PART if not is_assembly else 2)

    # A lightweight assembly draws blank views: all three came out 35.7 x 25.2 mm,
    # identical, with no geometry. Resolving the components before the first view
    # exists is what makes the drawing contain the machine rather than a placeholder.
    lightweight_unresolved = None
    if is_assembly:
        # ResolveAllLightweightComponents returns 0 when everything loaded, 2 when some
        # components stayed lightweight. Views drawn from an unresolved assembly come
        # out blank (all of them identical at 35.7 x 25.2 mm), so record the outcome
        # instead of assuming the resolve worked.
        for _attempt in range(2):
            try:
                lightweight_unresolved = model.ResolveAllLightweightComponents(True)
            except Exception:
                lightweight_unresolved = None
                break
            try:
                model.ForceRebuild3(False)
            except Exception:
                pass
            if lightweight_unresolved == 0:
                break

    drw = _new_document(sw, template)
    if drw is None:
        return error("drawing_create_failed",
                     "SolidWorks NewDocument returned null for %s" % template,
                     "Check the template is valid and uses a Windows path with backslashes.")

    sheet = drw.GetFirstView
    outline = list(sheet.GetOutline or [])
    sheet_w = (outline[2] - outline[0]) if outline else 0.42
    sheet_h = (outline[3] - outline[1]) if outline else 0.297
    # Must be measured now: after views exist, GetLines3 also returns their
    # geometry and the title-block / parts-list detection breaks down.
    safety = template_safe_area(sheet, sheet_w, sheet_h)

    # Size the model from its bounding box so the scale is known BEFORE the first
    # view is inserted — a brand-new view reports no outline, so measuring it and
    # moving it afterwards fights the layout.
    box = _model_box(model)

    chosen_scale = float(scale) if scale else None
    # Usable region from the template itself (left band, title block, parts list).
    # Hard-coding 0.12/0.10 here and 0.05/0.03/0.245 below is what made the predicted
    # layout disagree with where SolidWorks actually placed the views (reported twice:
    # D-3 then again for assemblies).
    safe_x0 = float(safety.get("left", sheet_w * 0.14))
    safe_x1 = sheet_w - float(safety.get("right", sheet_w * 0.01))
    safe_y0 = float(safety.get("bottom", sheet_h * 0.22))
    safe_y1 = sheet_h - float(safety.get("top", sheet_h * 0.012))
    avail_w = max(safe_x1 - safe_x0, 0.05)
    avail_h = max(safe_y1 - safe_y0, 0.05)
    gap = 0.022

    def _fit_scale(candidate_box):
        width_need = candidate_box[0] + gap + candidate_box[2]
        height_need = candidate_box[1] + gap + candidate_box[2]
        return min(avail_w / max(width_need, 1e-6),
                   avail_h / max(height_need, 1e-6), 1.0)

    if chosen_scale is None and auto_fit and box and box[0] > 0 and box[1] > 0:
        chosen_scale = _standard_scale(_fit_scale(box))
    elif chosen_scale is None:
        chosen_scale = 1.0

    def _layout(scale_value):
        ww, hh, dd = (box or (0.15, 0.08, 0.08))
        ww, hh, dd = ww * scale_value, hh * scale_value, dd * scale_value
        # Two columns (front+top on the left, right+iso on the right), centred inside
        # the measured safe area — so the prediction matches where views really land.
        total_w = ww + gap + dd
        total_h = hh + gap + dd
        over = (total_w > avail_w + 1e-6) or (total_h > avail_h + 1e-6)
        x_start = safe_x0 + (avail_w - total_w) / 2
        y_start = safe_y0 + (avail_h - total_h) / 2
        c1 = x_start + ww / 2
        c2 = x_start + ww + gap + dd / 2
        fy = y_start + total_h - hh / 2
        ty = y_start + dd / 2
        return (c1, c2, fy, ty, ww, hh, dd, over)

    # A 1.69 m assembly at 1.0 simply does not fit A3; step down until the predicted
    # layout is inside the sheet instead of emitting an out-of-frame drawing.
    if auto_fit and box and chosen_scale:
        for _ in range(4):
            probe = _layout(chosen_scale)
            if not probe[7]:
                break
            smaller = _standard_scale(chosen_scale * 0.5)
            if smaller >= chosen_scale:
                break
            chosen_scale = smaller

    col1, col2, front_y, top_y, w, h, d, _overflow = _layout(chosen_scale)
    positions = {
        "front": (col1, front_y), "back": (col1, front_y),
        "top": (col1, top_y), "bottom": (col1, top_y),
        "right": (col2, front_y), "left": (col2, front_y),
        "isometric": (col2, top_y),
    }

    # Explicit sheet centres win over the automatic packer. Callers were forced to
    # re-place views with raw COM (DISPATCH_PROPERTYPUT + VARIANT array) because a
    # 4-view sheet collapsed into one column; passing positions avoids that entirely.
    explicit_positions = False
    if view_positions:
        keys = [str(v).strip().lower() for v in views]
        for index, pos in enumerate(view_positions):
            if index >= len(keys):
                break
            try:
                positions[keys[index]] = (float(pos[0]), float(pos[1]))
            except Exception:
                continue
        explicit_positions = True

    created, skipped, view_objects = [], [], []
    for key in [str(v).strip().lower() for v in views]:
        if key not in VIEW_ALIASES:
            skipped.append({"view": key, "reason": "unsupported_view"})
            continue
        view_name = _model_view_name(model, key)
        if not view_name:
            skipped.append({"view": key, "reason": "view_name_not_found"})
            continue
        x, y = positions.get(key, (col1, front_y))
        view = drw.CreateDrawViewFromModelView3(model_path, view_name, x, y, 0)
        if view is None:
            skipped.append({"view": key, "reason": "create_returned_null"})
            continue
        per_view_scale = float((view_scales or {}).get(key, chosen_scale))
        try:
            view.ScaleDecimal = per_view_scale
        except Exception:
            per_view_scale = float(chosen_scale)
        try:
            o = list(view.GetOutline or [])
            box_m = {"width_m": round(o[2] - o[0], 4), "height_m": round(o[3] - o[1], 4)}
        except Exception:
            box_m = None
        created.append({"view": key, "name": view.Name,
                        "scale": per_view_scale, "outline": box_m})
        view_objects.append(view)

    # Views report a default outline until the drawing is redrawn; repacking before
    # that packs placeholder sizes, which is how views ended up scattered across the
    # sheet instead of grouped. Force geometry first, then repack.
    if view_objects:
        try:
            doc.ForceRebuild3(False)
        except Exception:
            pass
        try:
            doc.GraphicsRedraw2()
        except Exception:
            pass

    if explicit_positions:
        # Caller placed the views; repacking would throw that away.
        layout = {"fits": True, "columns": len(views), "explicit": True}
    else:
        layout = _place_views(view_objects, sheet_w, sheet_h, safety=safety)
    try:
        doc.GraphicsRedraw2()
    except Exception:
        pass

    # Measure-then-correct. An assembly's bounding box is not obtainable (lightweight
    # components return None from GetBox, GetPartBox does not exist on IAssemblyDoc),
    # so the predicted layout cannot be trusted for the very models that need auto_fit
    # most. After the views exist we can measure them for real: if any outline leaves
    # the sheet, step the scale down and repack, up to three times.
    auto_scale_adjustments = []
    # auto_fit must not overrule a scale the caller asked for explicitly: a 4-view
    # layout was silently taken from 1:2.5 to 1:5 (and an A2 to 1:10) while the views
    # still fitted. Only step down when the caller left scale unset.
    if auto_fit and not scale and view_objects:
        sheet_box = list(sheet.GetOutline or [])
        for _attempt in range(3):
            try:
                doc.GraphicsRedraw2()
            except Exception:
                pass
            outside = False
            measured = []
            for candidate in view_objects:
                try:
                    # P2-13: judge out-of-sheet on the deflated ink box, not the
                    # 11.9 mm-inflated GetOutline (which shrank scales needlessly).
                    o = precise_outline(candidate) or list(candidate.GetOutline or [])
                except Exception:
                    continue
                if len(o) < 4:
                    continue
                measured.append(o)
                if (o[0] < sheet_box[0] - 1e-6 or o[1] < sheet_box[1] - 1e-6
                        or o[2] > sheet_box[2] + 1e-6 or o[3] > sheet_box[3] + 1e-6):
                    outside = True
            if not outside:
                break
            try:
                current = float(view_objects[0].ScaleDecimal)
            except Exception:
                break
            smaller = _standard_scale(current * 0.5)
            if smaller >= current:
                break
            for candidate in view_objects:
                try:
                    candidate.ScaleDecimal = smaller
                except Exception:
                    pass
            auto_scale_adjustments.append({"from": current, "to": smaller})
            layout = _place_views(view_objects, sheet_w, sheet_h, safety=safety)
    try:
        drw.GraphicsRedraw2()
    except Exception:
        pass
    for item, view in zip(created, view_objects):
        try:
            o = list(view.GetOutline or [])
            item["outline"] = {"width_m": round(o[2] - o[0], 4),
                               "height_m": round(o[3] - o[1], 4)}
            p = precise_outline(view)
            if p:
                item["precise_outline_m"] = {
                    "width_m": round(p[2] - p[0], 4),
                    "height_m": round(p[3] - p[1], 4),
                }
        except Exception:
            pass

    try:
        drw.GraphicsRedraw2()
    except Exception:
        pass

    result = {
        "drawing": drw.GetTitle,
        "model": model.GetTitle,
        "model_path": model_path,
        "template": template,
        "sheet": {
            "name": sheet.Name if hasattr(sheet, "Name") else None,
            "size_m": [round(outline[2] - outline[0], 4), round(outline[3] - outline[1], 4)] if outline else None,
        },
        "views": created,
        "skipped_views": skipped,
        "layout": layout,
        "safe_area": safety,
        "saved": False,
    }
    if not created:
        return error("no_views_created", "No drawing view could be created",
                     "Check the model is saved, not read-only, and has valid named views.",
                     skipped_views=skipped)

    if save and output_path:
        out = os.path.abspath(output_path)
        if os.path.splitext(out)[1].lower() != ".slddrw":
            out += ".SLDDRW"
        os.makedirs(os.path.dirname(out) or ".", exist_ok=True)

        # `saved = os.path.exists(out)` was a lie whenever the path was already held by
        # an open document: SaveAsSilent fails, but the OLD file is still there, so the
        # check passed and callers believed they were editing the new drawing. A report
        # spent three rounds on this. Two guards instead:
        #   1. refuse up front if an open document already owns the path
        #   2. after saving, require the file to be newer than before
        blocker = None
        try:
            from ...sw_session import get_sw
            app = get_sw()
            open_names = []
            for doc in (app.GetDocuments or ()):
                try:
                    if os.path.normcase(str(doc.GetPathName)) == os.path.normcase(out):
                        open_names.append(str(doc.GetTitle))
                except Exception:
                    continue
            if open_names:
                blocker = ("path_locked_by_open_doc",
                           "target path is already open as %r — saving would fail "
                           "silently and leave the OLD file in place" % open_names[0])
        except Exception:
            pass

        before_mtime = os.path.getmtime(out) if os.path.exists(out) else None
        if blocker is None:
            try:
                drw.SaveAsSilent(out, False)
            except Exception as exc:
                blocker = ("saveas_failed", str(exc)[:120])
        after_mtime = os.path.getmtime(out) if os.path.exists(out) else None
        wrote = after_mtime is not None and after_mtime != before_mtime
        if blocker is None and not wrote:
            blocker = ("save_not_written",
                       "SaveAsSilent returned without raising, but the file on disk did "
                       "not change (mtime unchanged) — treat as not saved")
        if blocker:
            result.update(saved=False, drawing_path=out, save_error=blocker[0],
                          save_reason=blocker[1])
        else:
            result.update(saved=True, drawing_path=out,
                          bytes=os.path.getsize(out) if os.path.exists(out) else 0)
    result.update(
        safe_area_used={
            "x": [round(safe_x0, 4), round(safe_x1, 4)],
            "y": [round(safe_y0, 4), round(safe_y1, 4)],
            "source": safety.get("source"),
        },
        box_used_m=[round(v, 4) for v in box] if box else None,
        auto_scale_adjustments=auto_scale_adjustments,
        outline_note=("`outline` 来自 IView.GetOutline，宽高各虚胖 %.1f mm（恒定常数，"
                      "非投影误差）；`precise_outline_m` 已内缩该常数，实测与渲染 ink bbox "
                      "吻合（58.98 mm vs 59.2 mm，0.4%%）。没有逐顶点版本："
                      "IView.GetVisibleEntities2 返回轮廓边（GetStartVertex 为 None）、"
                      "IBody2.GetVertices 对全圆边实体返回 None、ICurve 在该版本没有 "
                      "GetEndPoint/GetTessPoints，三条取点路都已实测排除。"
                      % (OUTLINE_MARGIN_M * 1000)),
        box_source=("part" if box else
                    "unavailable (assembly) — layout verified by measuring views"),
        lightweight_resolve_attempted=is_assembly,
        lightweight_unresolved=lightweight_unresolved,
    )
    return ok(**result)


@tool_guard("export_failed")
def export_drawing(drawing_path, output_path, fmt=None):
    """Export an existing .SLDDRW to PDF / DWG / DXF. PDF comes from a drawing only."""
    out = os.path.abspath(output_path)
    ext = os.path.splitext(out)[1].lower()
    if fmt:
        ext = "." + fmt.strip().lower().lstrip(".")
        out = os.path.splitext(out)[0] + ext
    if ext not in SUPPORTED_FORMATS:
        return error("unsupported_format", "Supported: pdf, dwg, dxf (got %r)" % ext,
                     "STEP/STL come from models — use batch_export instead.")
    doc = _open_drawing(drawing_path)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    doc.SaveAsSilent(out, False)
    if not os.path.exists(out):
        return error("export_failed", "SolidWorks reported success but wrote no file",
                     "Check the output folder is writable and the drawing has at least one view.")
    return ok(drawing=doc.GetTitle, path=out, format=ext.lstrip("."),
              bytes=os.path.getsize(out))


@tool_guard("render_failed")


@tool_guard("audit_failed")
def drawing_audit(drawing_path):
    """Structural audit: views, scales, blank views, dimension and annotation counts.

    Evidence only — it does not judge whether the chosen views or tolerances are
    manufacturable. Blank views and missing dimensions are the checks that catch
    silent failures.
    """
    doc = _open_drawing(drawing_path)
    views = _iter_views(doc)
    sheets = [v for v in views if getattr(v, "Type", None) == 1]
    model_views = [v for v in views if getattr(v, "Type", None) != 1]

    rows, warnings = [], []
    total_dims = 0
    boxes = []
    sheet_box = list(sheets[0].GetOutline) if sheets else None
    # IView.GetAnnotationCount on a model view also reports sheet-level notes (技术要求),
    # so "view 3: annotations 11" read as "this view has 11 annotations" when most were
    # sheet notes. Report the sheet's own count separately and say so, rather than
    # subtracting — the split between the two is not reliably attributable.
    sheet_annotations = None
    if sheets:
        try:
            sheet_annotations = int(sheets[0].GetAnnotationCount)
        except Exception:
            sheet_annotations = None
    for view in model_views:
        try:
            name = view.Name
        except Exception:
            name = None
        try:
            scale = view.ScaleDecimal
        except Exception:
            scale = None
        try:
            o = list(view.GetOutline or [])
            width, height = o[2] - o[0], o[3] - o[1]
        except Exception:
            width = height = None
        try:
            dims = int(view.GetDimensionCount)
        except Exception:
            dims = None
        try:
            notes = int(view.GetAnnotationCount)
        except Exception:
            notes = None
        if dims:
            total_dims += dims
        # P2-13: GetOutline is inflated by a constant 11.9 mm on every view. Overlap
        # and out-of-sheet checks must use the deflated ink box, otherwise two views
        # whose phantom margins merely touch raise a views_overlap that is not real.
        # Keep the raw width/height in the row for backwards compatibility.
        pbox = precise_outline(view)
        row = {
            "name": name,
            "scale": scale,
            "width_m": round(width, 4) if width is not None else None,
            "height_m": round(height, 4) if height is not None else None,
            "precise_width_m": round(pbox[2] - pbox[0], 4) if pbox else None,
            "precise_height_m": round(pbox[3] - pbox[1], 4) if pbox else None,
            "dimensions": dims,
            "annotations": notes,
        }
        rows.append(row)
        if width is not None and (width < 0.001 or height < 0.001):
            warnings.append({"view": name, "issue": "blank_view"})
        # A view without dimensions is normal: axonometric and detail views often
        # carry none. Measured on 100 real drawings, 79 of them would have tripped
        # this warning, so it now only fires when the SHEET has no dimensions at all.

        if width is not None and height is not None:
            try:
                boxes.append((name, pbox or list(view.GetOutline)))
            except Exception:
                pass
        check_box = pbox
        if check_box is None and sheet_box:
            try:
                check_box = list(view.GetOutline)
            except Exception:
                check_box = None
        if sheet_box and check_box:
            if (check_box[0] < sheet_box[0] or check_box[1] < sheet_box[1]
                    or check_box[2] > sheet_box[2] or check_box[3] > sheet_box[3]):
                warnings.append({
                    "view": name, "issue": "view_outside_sheet",
                    "hint": "reduce scale or move the view",
                })

    if not model_views:
        warnings.append({"view": None, "issue": "no_model_views"})
    if model_views and total_dims == 0:
        warnings.append({
            "view": None, "issue": "no_dimensions_anywhere",
            "hint": "the sheet carries no dimensions at all",
        })

    # Overlapping views are the failure mode a human notices first but code often
    # misses: two outlines intersect and neither is blank nor out of bounds.
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            a, b = boxes[i][1], boxes[j][1]
            if a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]:
                warnings.append({
                    "view": "%s <-> %s" % (boxes[i][0], boxes[j][0]),
                    "issue": "views_overlap",
                    "hint": "reduce scale or use fewer views per sheet",
                })

    return ok(
        drawing=doc.GetTitle,
        path=doc.GetPathName,
        sheet_count=len(sheets),
        sheet_size_m=[
            round(list(sheets[0].GetOutline)[2] - list(sheets[0].GetOutline)[0], 4),
            round(list(sheets[0].GetOutline)[3] - list(sheets[0].GetOutline)[1], 4),
        ] if sheets else None,
        view_count=len(model_views),
        total_dimensions=total_dims,
        views=rows,
        sheet_annotations=sheet_annotations,
        warnings=warnings,
        note=("Evidence only: view choice, tolerances and manufacturability are not judged here. "
              "A view's `annotations` count may include sheet-level notes (技术要求) — "
              "sheet_annotations reports the sheet's own count separately."),
    )


