"""Feature ledger — structured inventory of every feature, dimension and hole.

Process §2 requires this before any annotation: it is the basis for deciding what to
dimension, where to call out holes, and what to verify afterwards. Read-only analysis;
the model is never modified.
"""

import os

from ..sw_session import SW_DOC_PART, error, ok, open_document, tool_guard

# Feature types that carry geometry worth cataloguing; folders and bookkeeping are skipped.
FOLDER_TYPES = {
    "CommentsFolder", "FavoriteFolder", "HistoryFolder", "SelectionSetFolder",
    "SensorFolder", "DocsFolder", "DetailCabinet", "SurfaceBodyFolder",
    "SolidBodyFolder", "RefPlane", "OriginProfileFeature",
    "CutListFolder", "EnvFolder", "InkMarkupFolder", "EqnFolder", "MaterialFolder",
    "BomFeature", "Comment",
}

CUT_TYPES = {"CutExtrusion", "ICE", "HoleWzd", "Cut", "CutThin", "ExtruThin", "BossThin"}


def _flatten(obj):
    import win32com.client as win

    out = []

    def walk(x):
        if isinstance(x, (tuple, list)):
            for item in x:
                walk(item)
        elif x is not None:
            out.append(win.Dispatch(x))

    walk(obj)
    return out


def _feature_dimensions(feature):
    """Driving/reference dimensions attached to a feature, in millimetres."""
    rows = []
    try:
        display = feature.GetFirstDisplayDimension
    except Exception:
        return rows
    guard = 0
    while display is not None and guard < 80:
        guard += 1
        try:
            import win32com.client as win

            dim = win.Dispatch(display).GetDimension2(0)
            if dim is not None:
                name = ""
                try:
                    name = str(dim.FullName)
                except Exception:
                    try:
                        name = str(dim.Name)
                    except Exception:
                        name = ""
                kind = None
                try:
                    # swDimensionType_e: 0 linear, 1 angular, 2 diameter, 3 radial,
                    # 4 ordinate. Angles and ordinates are not lengths, so reconciliation
                    # must not compare them against drawing dimension values.
                    kind = int(dim.GetType)
                except Exception:
                    kind = None
                rows.append({
                    "name": name,
                    "value_mm": round(float(dim.SystemValue) * 1000, 3),
                    "dimension_type": kind,
                })
        except Exception:
            pass
        try:
            display = feature.GetNextDisplayDimension(display)
        except Exception:
            break
    return rows


def _cylindrical_faces(feature):
    """Cylindrical faces of one feature → hole candidates (diameter, axis, centre).

    A cylindrical face is not proof of a hole (it can be a boss or a fillet), so these
    are catalogued as candidates and flagged for the caller's judgement.
    """
    holes = []
    try:
        faces = feature.GetFaces
    except Exception:
        return holes
    for face in _flatten(faces):
        try:
            surface = face.GetSurface
            if surface is None or not surface.IsCylinder:
                continue
            params = list(surface.CylinderParams)  # ox,oy,oz, ax,ay,az, radius
            holes.append({
                "diameter_mm": round(float(params[6]) * 2000, 3),
                "axis": [round(float(v), 3) for v in params[3:6]],
                "centre_mm": [round(float(v) * 1000, 2) for v in params[0:3]],
            })
        except Exception:
            continue
    return holes


def _unused_is_inner_cylinder(face, params):
    """True when the cylindrical face faces its own axis (a bore), False for a boss or a
    bend fillet. Without this filter every R2.54 bend reads as a Ø5.1 hole."""
    try:
        normal = list(face.Normal)
        box = list(face.GetBox)
        centre = [(box[0] + box[3]) / 2, (box[1] + box[4]) / 2, (box[2] + box[5]) / 2]
        origin = params[0:3]
        axis = params[3:6]
        v = [centre[i] - origin[i] for i in range(3)]
        along = sum(v[i] * axis[i] for i in range(3))
        radial = [v[i] - along * axis[i] for i in range(3)]
        return sum(radial[i] * normal[i] for i in range(3)) < 0
    except Exception:
        return True  # keep it when we cannot tell; better an extra candidate than a miss


def _body_cylinders(model):
    """Every cylindrical face on the solid bodies → holes, regardless of which feature
    made it (formed tools and imported geometry included, which feature-scoped scanning
    misses). Wrapped defensively: GetBodies2 returns shapes that vary by body type.
    """
    import win32com.client as win

    out = []
    for args in ((-1, True), (-1, False), (0, True)):
        try:
            bodies = model.GetBodies2(*args)
        except Exception:
            continue
        if bodies is None:
            continue
        items = bodies if isinstance(bodies, (tuple, list)) else (bodies,)
        for item in items:
            if item is None:
                continue
            try:
                body = win.Dispatch(item)
                # IBody2.GetFaces is a METHOD (IComponent2/IFeature use properties) —
                # calling it as a property silently yields the bound method and the scan
                # finds nothing.
                faces = body.GetFaces()
            except Exception:
                continue
            face_items = faces if isinstance(faces, (tuple, list)) else (faces,)
            for face_item in face_items:
                if face_item is None:
                    continue
                try:
                    face = win.Dispatch(face_item)
                    surface = face.GetSurface
                    if surface is None or not surface.IsCylinder:
                        continue
                    params = list(surface.CylinderParams)
                    out.append({
                        "diameter_mm": round(float(params[6]) * 2000, 3),
                        "axis": [round(float(v), 3) for v in params[3:6]],
                        "centre_mm": [round(float(v) * 1000, 2) for v in params[0:3]],
                        "feature": None,
                    })
                except Exception:
                    continue
        if out:
            break
    return out


def _hole_key(hole):
    return (round(hole["diameter_mm"], 1),
            tuple(hole["axis"]),
            tuple(round(c, 1) for c in hole["centre_mm"]))


def _classify(bbox_mm, has_flat_pattern, types):
    """Heuristic part category (process §2.2). Advisory only."""
    if has_flat_pattern:
        return "钣金", "零件含 FlatPattern 特征"
    dims = sorted(bbox_mm)
    shortest, middle, longest = dims[0], dims[1], dims[2]
    if shortest <= 0:
        return "其他", "包围盒退化"
    slender = longest / shortest
    width = middle / shortest
    if slender >= 8 and width <= 3:
        return "轴类", "长径比 %.1f、中/短比 %.1f" % (slender, width)
    if shortest <= 6 and middle >= 20:
        return "板类", "最薄 %.1f mm、最宽 %.1f mm" % (shortest, middle)
    if longest / middle <= 1.6 and longest / shortest >= 1.6 and len(types) >= 8:
        return "箱体", "三方向尺寸接近且特征多"
    return "一般零件", "包围盒 %s" % (bbox_mm,)


@tool_guard("ledger_failed")
def build_feature_ledger(model_path, include_faces=True, config=None):
    """Build a structured ledger of features, dimensions, holes and model doubts.

    Output is the basis for annotation planning and self-check (process §2, §6.1):
    every entry carries the model object it came from so a dimension can be traced
    back to geometry instead of being guessed.

    include_faces=False skips cylindrical-face scanning (much faster on large models).
    """
    path = os.path.abspath(model_path)
    if not os.path.exists(path):
        raise FileNotFoundError("file not found: %s" % path)
    if not path.lower().endswith((".sldprt", ".sldasm")):
        return error("bad_input", "Expected a .SLDPRT/.SLDASM, got %s" % os.path.basename(path))

    model = open_document(path)  # type inferred from extension: assemblies failed here
    config_name = config
    if not config_name:
        try:
            config_name = model.ConfigurationManager.ActiveConfiguration.Name
        except Exception:
            config_name = None

    bbox_mm = None
    try:
        b = list(model.GetPartBox(True))
        bbox_mm = [round((b[3] - b[0]) * 1000, 2),
                   round((b[4] - b[1]) * 1000, 2),
                   round((b[5] - b[2]) * 1000, 2)]
    except Exception:
        pass

    features, holes, doubts, types = [], [], [], {}
    has_flat_pattern = False
    suppressed = []
    index = 0
    feature = model.FirstFeature
    guard = 0
    while feature is not None and guard < 800:
        guard += 1
        try:
            ftype = feature.GetTypeName2
            fname = feature.Name
            is_suppressed = bool(feature.IsSuppressed)
        except Exception:
            feature = feature.GetNextFeature
            continue

        if ftype in FOLDER_TYPES:
            feature = feature.GetNextFeature
            continue

        types[ftype] = types.get(ftype, 0) + 1
        if ftype == "FlatPattern":
            has_flat_pattern = True
        if is_suppressed:
            suppressed.append(fname)

        dims = _feature_dimensions(feature)
        entry = {
            "index": index,
            "name": fname,
            "type": ftype,
            "suppressed": is_suppressed,
            "dimensions": dims,
        }

        if include_faces and ftype in CUT_TYPES:
            found = _cylindrical_faces(feature)
            if found:
                entry["hole_candidates"] = found
                for hole in found:
                    holes.append(dict(hole, feature=fname))

        features.append(entry)
        index += 1
        feature = feature.GetNextFeature

    try:
        needs_rebuild = bool(model.Extension.NeedsRebuild2)
        if needs_rebuild:
            # A DRAWING edit (add/move/delete annotation) sets this flag on the
            # referenced model without any model change, and the doubt then pulls the
            # delivery verdict down for a defect that does not exist (reported
            # 2026-11-01 against an untouched model). Force one rebuild and re-read:
            # a real model change survives it, a drawing-induced flag does not.
            try:
                from ..sw_session import call_method as _rebuild_once

                _rebuild_once(model, "EditRebuild3")
                needs_rebuild = bool(model.Extension.NeedsRebuild2)
            except Exception:
                pass
    except Exception:
        needs_rebuild = None
    if needs_rebuild:
        doubts.append({"kind": "needs_rebuild",
                       "detail": "模型存在待重建的改动，尺寸可能与实体不一致"})
    if suppressed:
        doubts.append({"kind": "suppressed_features", "detail": suppressed})

    # Feature-scoped holes miss formed/imported ones, so merge a body-wide scan.
    # Cylindrical faces include bores, bosses and bend fillets; only `inner` ones are
    # reported as holes, the rest are counted so nothing disappears silently.
    cylinders_total, inner_count = 0, 0
    if include_faces:
        seen = {_hole_key(h) for h in holes}
        scanned = _body_cylinders(model)
        cylinders_total = len(scanned)
        # Keep every cylindrical face. Deciding bore vs boss vs bend fillet from the
        # face normal proved unreliable here (all 80 faces classified the same way,
        # contradicting the known holes), so no pre-filtering is applied.
        for hole in scanned:
            key = _hole_key(hole)
            if key in seen:
                continue
            seen.add(key)
            holes.append(hole)

    category, reason = _classify(bbox_mm or [0, 0, 0], has_flat_pattern, types)

    driving = sum(len(f["dimensions"]) for f in features)
    return ok(
        part=os.path.basename(path),
        path=path,
        config=config_name,
        units="mm",
        category=category,
        category_reason=reason,
        bbox_mm=bbox_mm,
        material=_material_of(model),
        features=features,
        holes=holes,
        doubts=doubts,
        stats={
            "features": len(features),
            "driving_dimensions": driving,
            "hole_candidates": len(holes),
            "patterns": sum(v for k, v in types.items() if "Pattern" in k),
            "cylinder_faces": cylinders_total,
            "feature_types": types,
        },
        notes=[
            "holes 是**全模型圆柱面清单**（直径/轴线/中心），不是孔清单：孔、凸台、"
            "折弯圆角都在其中。判「是否孔」需结合入口面与轴线，工具不做推断。",
            "非圆柱孔（异形孔、成型工具做的百叶窗、导入体上的孔）不在本表中，需人工核对。",
            "驱动尺寸来自特征，草图内部尺寸与基准面偏移未计入。",
        ],
    )


@tool_guard("features_failed")
def list_features(model_path, include_folders=False):
    """Flat feature-tree listing: name, type, suppression state. Read-only.

    The light counterpart of build_feature_ledger — no dimensions, no face
    scans, just the design-tree order so a caller can answer "what is in this
    model" and target features for list_mates / set_component_state style
    follow-ups without paying for the full ledger walk.
    """
    model = open_document(model_path)
    rows = []
    feature = model.FirstFeature
    guard = 0
    while feature is not None and guard < 2000:
        guard += 1
        try:
            ftype = feature.GetTypeName2
        except Exception:
            ftype = None
        if ftype in FOLDER_TYPES and not include_folders:
            feature = feature.GetNextFeature
            continue
        row = {"index": len(rows), "name": None, "type": ftype,
               "suppressed": None}
        try:
            row["name"] = str(feature.Name)
        except Exception:
            pass
        try:
            row["suppressed"] = bool(feature.IsSuppressed)
        except Exception:
            pass
        rows.append(row)
        feature = feature.GetNextFeature
    return ok(document=str(model.GetTitle), path=str(model.GetPathName),
              count=len(rows), features=rows)


def _material_of(model):
    """Material name, or None when the part has none assigned (never invented)."""
    try:
        import pythoncom

        from ..sw_session import variant

        name = model.GetMaterialPropertyName2(
            "", variant("", pythoncom.VT_BYREF | pythoncom.VT_BSTR))
        return str(name) if name else None
    except Exception:
        return None
