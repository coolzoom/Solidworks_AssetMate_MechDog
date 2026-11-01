"""Bill of materials for an assembly (read-only, never modifies the model).

Design notes:
- Components are aggregated by part name (instance suffix stripped), so 46 components
  can collapse to 19 BOM rows.
- Mass/material are read per unique part document, not per component, to keep the slow
  COM calls down to one per part type.
- GetMaterialPropertyName2 needs a byref BSTR for the database argument; a plain string
  raises a type mismatch on pywin32.
- Every value that cannot be read is reported as null plus a warning, never guessed.
"""
from __future__ import annotations

import csv
import os

from ..sw_session import (
    tool_guard,
    SW_DOC_ASSEMBLY,
    error,
    get_components,
    ok,
    open_document,
    part_name,
)


def _byref_bstr():
    import pythoncom

    from ..sw_session import variant

    return variant("", pythoncom.VT_BYREF | pythoncom.VT_BSTR)


def _material_of(doc):
    try:
        value = doc.GetMaterialPropertyName2("", _byref_bstr())
        if value and str(value).strip():
            return str(value).strip()
    except Exception:
        pass
    try:
        value = doc.MaterialPropertyName
        return str(value).strip() if value and str(value).strip() else None
    except Exception:
        return None
    return None


def _mass_of(doc):
    try:
        return float(doc.Extension.CreateMassProperty.Mass)
    except Exception:
        return None


def _unique_parts(components):
    """part name -> representative component (first instance)."""
    parts = {}
    for comp in components:
        parts.setdefault(part_name(comp.Name2), comp)
    return parts


@tool_guard("bom_failed")
def export_bom(
    assembly_path,
    include_mass=True,
    include_material=True,
    csv_path=None,
    refresh=False,
    resolve_lightweight=True,
):
    """Aggregate components into BOM rows. Read-only; optionally writes a CSV.

    resolve_lightweight loads suppressed/lightweight parts first. Without it every
    lightweight component returns GetModelDoc2=None and mass/material come back null.
    """
    doc = open_document(assembly_path, SW_DOC_ASSEMBLY)
    if resolve_lightweight:
        try:
            doc.ResolveAllLightweightComponents(True)
        except Exception:
            pass
        # ResolveAllLightweightComponents can return 2 while components stay
        # lightweight (GetModelDoc2 None, mass/material null). A forced rebuild
        # is what actually loads them — measured 1/46 -> 46/46.
        try:
            doc.ForceRebuild3(False)
        except Exception:
            pass
    components = get_components(doc, refresh=refresh)
    if not components:
        return error("empty_assembly", "No components found in %s" % doc.GetTitle)

    counts = {}
    for comp in components:
        name = part_name(comp.Name2)
        counts[name] = counts.get(name, 0) + 1

    rows, missing = [], []
    for name in sorted(counts):
        comp = _unique_parts(components)[name]
        path = None
        mass = None
        material = None
        try:
            path = comp.GetPathName or None
        except Exception:
            pass

        model = None
        try:
            model = comp.GetModelDoc2
        except Exception:
            model = None

        if include_mass and model is not None:
            mass = _mass_of(model)
        if include_material and model is not None:
            material = _material_of(model)

        if include_mass and model is None:
            missing.append({"part": name, "missing": "document_not_loaded"})
        elif include_mass and mass is None:
            missing.append({"part": name, "missing": "mass_unavailable"})
        if include_material and model is not None and material is None:
            missing.append({"part": name, "missing": "material_not_set"})

        qty = counts[name]
        rows.append({
            "part": name,
            "qty": qty,
            "material": material,
            "mass_each_kg": round(mass, 6) if mass is not None else None,
            "mass_total_kg": round(mass * qty, 6) if mass is not None else None,
            "path": path,
            # A sub-assembly row already contains its children, so summing every row
            # double-counts. Needed by total_mass_leaf_only_kg below.
            "is_subassembly": str(path or "").lower().endswith((".sldasm", ".asm")),
        })

    known_mass = [r["mass_total_kg"] for r in rows if r["mass_total_kg"] is not None]
    # SolidWorks falls back to 1000 kg/m3 when a part has no material; call that
    # out so the number is not mistaken for a real material density.
    default_density_parts = [
        r["part"] for r in rows if r["material"] is None and r["mass_each_kg"] is not None
    ]
    result = {
        "assembly": doc.GetTitle,
        "path": doc.GetPathName,
        "components": len(components),
        "part_types": len(rows),
        "items": rows,
        # Rows flagged as sub-assemblies already contain their children's mass, so
        # adding both double-counts (measured +11.9% on a 7-component fixture).
        "total_mass_kg": round(sum(known_mass), 6) if known_mass else None,
        "total_mass_leaf_only_kg": round(sum(
            row["mass_total_kg"] for row in rows
            if not row.get("is_subassembly") and isinstance(row.get("mass_total_kg"), (int, float))
        ), 6) if known_mass else None,
        "total_mass_note": ("total_mass_kg 含子装配行与其子件行；按零件树去重用 "
                            "total_mass_leaf_only_kg，它与 get_mass_properties 对得上。"),
        "mass_complete": len(known_mass) == len(rows) if include_mass else False,
        "mass_basis": (
            "SolidWorks default density 1000 kg/m3 for %d part(s) without material: %s"
            % (len(default_density_parts), ", ".join(default_density_parts))
            if default_density_parts else "material density"
        ),
        "warnings": missing,
    }

    if csv_path:
        out = os.path.abspath(csv_path)
        os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
        with open(out, "w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.writer(fh)
            writer.writerow(["Part", "Qty", "Material", "MassEach_kg", "MassTotal_kg", "Path"])
            for row in rows:
                writer.writerow([
                    row["part"], row["qty"], row["material"] or "",
                    row["mass_each_kg"] if row["mass_each_kg"] is not None else "",
                    row["mass_total_kg"] if row["mass_total_kg"] is not None else "",
                    row["path"] or "",
                ])
        result["csv"] = out
    return ok(**result)
