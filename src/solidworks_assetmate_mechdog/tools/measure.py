"""Entity measurement through IMeasure — the geometric half of sheet_metal_qc.

SolidWorks' own Measure tool answers distance/angle/area/diameter between
selected entities; the COM binding is IExtension.CreateMeasure followed by
Calculate() once the entities are selected. Selection is the awkward part of
COM automation (SelectByID2 wants a name or a point that lands exactly on the
entity), so each entity accepts both forms:

    {"type": "FACE",  "name": "Plane1"}
    {"type": "EDGE",  "point": [0.01, 0.02, 0.005]}   # model metres
    {"type": "FACE",  "nearest": [x, y, z]}           # GetEntityByName fallback off

Two entities give distance/angle/deltas; one gives perimeter/area/diameter —
the pieces the v0.9 min-flange and hole-to-bend checks need.
"""
from __future__ import annotations

from ..sw_session import (
    byref_dispatch_null,
    error,
    ok,
    open_document,
)

_ENTITY_TYPES = {"FACE": "FACE", "EDGE": "EDGE", "VERTEX": "VERTEX",
                 "PLANE": "PLANE", "DATUMPLANE": "PLANE", "AXIS": "AXIS",
                 "SKETCHSEGMENT": "SKETCHSEGMENT", "COMPONENT": "COMPONENT",
                 "EXTSKETCHPOINT": "EXTSKETCHPOINT"}


def _select_entity(doc, spec):
    """Select one entity for the measure; returns a readable row or None."""
    sel_type = str(spec.get("type", "FACE")).upper()
    sw_type = _ENTITY_TYPES.get(sel_type)
    if sw_type is None:
        raise ValueError("unknown entity type %r (known: %s)"
                         % (sel_type, sorted(_ENTITY_TYPES)))
    name = spec.get("name") or ""
    point = spec.get("point") or spec.get("nearest")
    x = y = z = 0.0
    if point:
        if len(point) != 3:
            raise ValueError("point must be [x, y, z] in model metres")
        x, y, z = (float(v) for v in point)
    selected = doc.Extension.SelectByID2(
        name, sw_type, x, y, z, True, 0, byref_dispatch_null(), 0)
    return bool(selected), {"type": sel_type, "name": name or None,
                            "point": [x, y, z] if point else None}


def measure_entities(path, entities):
    """Measure between one or two entities of a part/assembly (read-only).

    entities: list of 1-2 specs like {"type":"EDGE","point":[x,y,z]} — distances
    in model metres. With two entities the result carries distance/angle/xyz
    deltas; with one it carries perimeter/area/diameter as applicable. Values
    that do not apply to the selection are omitted, never zero-filled.
    """
    if not entities or len(entities) > 2:
        return error("bad_input", "entities must be a list of 1 or 2 specs")
    try:
        doc = open_document(path)
        measure = doc.Extension.CreateMeasure
        if measure is None:
            return error("measure_unavailable",
                         "Extension.CreateMeasure returned null on this binding")
        try:
            doc.ClearSelection2(True)
        except Exception:
            pass
        described = []
        for spec in entities:
            try:
                selected, row = _select_entity(doc, spec)
            except ValueError as exc:
                return error("bad_input", str(exc))
            row["selected"] = selected
            described.append(row)
            if not selected:
                return error(
                    "select_failed",
                    "Could not select entity %r — the name may be wrong or the "
                    "point may not land on it" % spec,
                    "list_drawing_edges / build_feature_ledger report model "
                    "coordinates that can feed the point form.",
                    entities=described)
        try:
            calculated = measure.Calculate()
        except Exception as exc:
            return error("measure_failed", "Calculate() raised: %s" % exc,
                         entities=described)
        result = {"document": str(doc.GetTitle), "entities": described,
                  "calculated": bool(calculated)}

        def _num(name):
            try:
                return float(getattr(measure, name))
            except Exception:
                return None

        if len(entities) == 2:
            result["units"] = {"distance": "m", "angle": "deg", "delta": "m"}
            for key in ("Distance", "DeltaX", "DeltaY", "DeltaZ", "Angle",
                        "CenterX", "CenterY", "CenterZ"):
                value = _num(key)
                if value is not None:
                    result[key.lower()] = value
            for key in ("IsBetween", "IsParallel", "IsPerpendicular", "IsIntersect"):
                try:
                    result[key.lower()] = bool(getattr(measure, key))
                except Exception:
                    pass
        else:
            result["units"] = {"length": "m", "area": "m^2", "diameter": "m"}
            for key in ("Area", "Perimeter", "Diameter", "TotalLength",
                        "CenterX", "CenterY", "CenterZ", "Radius"):
                value = _num(key)
                if value is not None:
                    result[key.lower()] = value
        return ok(**result)
    except (FileNotFoundError, ValueError, PermissionError) as exc:
        return error("bad_input", str(exc))
    except Exception as exc:
        return error("measure_failed", str(exc))
