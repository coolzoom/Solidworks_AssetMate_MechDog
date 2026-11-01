"""Read-only assembly inspection and preview export."""
from __future__ import annotations

import os

from ..sw_session import (
    tool_guard,
    SW_DOC_ASSEMBLY,
    call_method,
    error,
    get_components,
    ok,
    open_document,
    part_name,
    rgb_of,
)


def _component_info(comp):
    try:
        path = comp.GetPathName
    except Exception:
        path = None
    return {
        "component": comp.Name2,
        "part": part_name(comp.Name2),
        "path": path,
        "is_subassembly": bool(path and path.lower().endswith(".sldasm")),
        "rgb": rgb_of(comp.MaterialPropertyValues),
    }


@tool_guard("tree_failed")
def assembly_tree(assembly_path, max_depth=1, refresh=False):
    """Component tree. max_depth=1 is flat (fast); deeper levels use GetChildren."""
    doc = open_document(assembly_path, SW_DOC_ASSEMBLY)
    components = get_components(doc, refresh=refresh)  # one slow call, then memory only
    nodes = []
    for comp in components:
        info = _component_info(comp)
        if max_depth > 1 and info["is_subassembly"]:
            info["children"] = _children(comp, max_depth - 1)
        nodes.append(info)
    return ok(
        assembly=doc.GetTitle,
        path=doc.GetPathName,
        components=len(components),
        max_depth=max_depth,
        tree=nodes,
    )


def _children(comp, depth):
    try:
        children = comp.GetChildren or []
    except Exception:
        return []
    result = []
    for child in children:
        info = _component_info(child)
        if depth > 1 and info["is_subassembly"]:
            info["children"] = _children(child, depth - 1)
        result.append(info)
    return result


@tool_guard("mass_failed")
def mass_properties(assembly_path):
    """Mass/volume/area in SI units (kg, m^3, m^2) — works for parts and assemblies.

    Includes `mass_basis`: average density derived from mass/volume. SolidWorks
    falls back to 1000 kg/m3 for parts without material, so a value of exactly
    1000 is reported as a suspected placeholder rather than real engineering data.
    """
    # Was hard-coded to SW_DOC_ASSEMBLY, which made every part fail with a
    # misleading "file may be locked" message (reported 2026-10-28).
    doc = open_document(assembly_path)
    try:
        mass = doc.Extension.CreateMassProperty
        center = mass.CenterOfMass
        mass_kg = float(mass.Mass)
        volume_m3 = float(mass.Volume)
        density = mass_kg / volume_m3 if volume_m3 else None
        if density is not None and abs(density - 1000.0) < 1.0:
            basis = (
                "suspected default density: mass/volume = %.1f kg/m3, which matches "
                "SolidWorks' fallback for parts without material — verify material "
                "assignments before using this mass" % density
            )
        elif density is not None:
            basis = "average density %.1f kg/m3 from mass/volume" % density
        else:
            basis = "unknown (volume is zero)"
        return ok(
            assembly=doc.GetTitle,
            units={"mass": "kg", "volume": "m^3", "surface_area": "m^2", "center_of_mass": "m"},
            mass=mass_kg,
            volume=volume_m3,
            surface_area=float(mass.SurfaceArea),
            center_of_mass=[float(v) for v in center] if center else None,
            mass_basis=basis,
            average_density_kg_m3=round(density, 3) if density is not None else None,
        )
    except Exception as exc:
        return error(
            "mass_unavailable", str(exc),
            "CreateMassProperty failed on this binding; see docs/com-quirks.md "
            "section 5. Read per-part properties instead of guessing.",
        )


@tool_guard("preview_failed")
def export_preview(assembly_path, output_path, size="1600x1200", view="*Isometric", keep_bmp=False):
    """Save a view image. SolidWorks SaveBMP always writes BMP; PNG needs Pillow."""
    doc = open_document(assembly_path)  # type inferred from extension
    try:
        width, height = (int(x) for x in size.lower().split("x"))
    except Exception:
        width, height = 1600, 1200
    try:
        doc.ShowNamedView2(view, 7)
    except Exception:
        pass  # keep the current view if the named view is missing
    doc.ViewZoomtofit2()

    output_path = os.path.abspath(output_path)
    wants_png = output_path.lower().endswith(".png")
    bmp_path = output_path[:-4] + ".bmp" if wants_png else output_path
    if not doc.SaveBMP(bmp_path, width, height):
        return error("preview_failed", "SaveBMP returned failure: %s" % bmp_path)

    result = ok(assembly=doc.GetTitle, image=bmp_path, width=width, height=height, format="bmp")
    if wants_png:
        try:
            from PIL import Image
            Image.open(bmp_path).save(output_path)
            if not keep_bmp:
                try:
                    os.remove(bmp_path)  # avoid leaving a multi-MB intermediate
                except OSError:
                    pass
            result.update(image=output_path, format="png")
        except ImportError:
            result["note"] = "Saved BMP instead. Install Pillow (pip install '.[png]') for PNG output."
    return result


@tool_guard("capture_failed")
def capture_model_views(model_path, output_dir, views=None, size="1200x900"):
    """Export one image per named view — the multi-view counterpart of export_preview.

    views accepts standard names (front/top/right/isometric...) or SolidWorks
    native names ("*前视" / "*Front"); the trailing entries in the standard set
    cover the four oblique angles. Returns the written files plus any view name
    SolidWorks did not recognise.
    """
    doc = open_document(model_path)
    try:
        width, height = (int(x) for x in size.lower().split("x"))
    except Exception:
        width, height = 1200, 900
    wanted = views or ["*Isometric", "*Front", "*Top", "*Right"]
    out_dir = os.path.abspath(output_dir)
    os.makedirs(out_dir, exist_ok=True)
    written, failed = [], []
    base = os.path.splitext(os.path.basename(str(doc.GetPathName) or "model"))[0]
    for index, view in enumerate(wanted):
        name = str(view)
        try:
            doc.ShowNamedView2(name, 7)
        except Exception:
            # Localised installs use 前视/上视-style names; the standard set
            # keeps English defaults, and unknown names land in `failed`
            # rather than aborting the whole capture.
            failed.append({"view": name, "reason": "named_view_rejected"})
            continue
        doc.ViewZoomtofit2()
        target = os.path.join(out_dir, "%s_%d_%s.bmp" % (
            base, index, name.lstrip("*").replace(" ", "_")))
        try:
            if doc.SaveBMP(target, width, height):
                written.append({"view": name, "file": target,
                                "bytes": os.path.getsize(target)})
            else:
                failed.append({"view": name, "reason": "SaveBMP_false"})
        except Exception as exc:
            failed.append({"view": name, "reason": str(exc)[:160]})
    result = ok(model=str(doc.GetTitle), output_dir=out_dir,
                written=len(written), images=written)
    if failed:
        result["failed"] = failed
    return result


_MATE_TYPE_NAMES = {
    0: "coincident", 1: "concentric", 2: "perpendicular", 3: "parallel",
    4: "tangent", 5: "distance", 6: "angle", 8: "symmetric",
    9: "cam_follower", 10: "gear", 11: "width", 13: "rack_pinion",
    15: "path", 16: "lock", 17: "screw", 18: "linear_coupler",
    19: "universal_joint", 21: "slot", 22: "hinge",
}


def _mate_entities(mate):
    entities = []
    try:
        count = mate.GetMateEntityCount()
    except Exception:
        return entities
    for i in range(count):
        try:
            entity = mate.MateEntity(i)
            if entity is None:
                continue
            comp = entity.ReferenceComponent
            row = {"component": str(comp.Name2) if comp is not None else None}
            try:
                row["reference_type"] = int(entity.ReferenceType2)
            except Exception:
                pass
            entities.append(row)
        except Exception:
            continue
    return entities


@tool_guard("mates_failed")
def list_mates(assembly_path, component=None):
    """List assembly mates: type, partner components, per-entity reference types.

    Mates live under the MateGroup feature as sub-features whose
    GetSpecificFeature2 is an IMate2. component="Plate-1" filters to mates that
    involve that instance — the assembly-level equivalent of asking "how is this
    part held in place".
    """
    doc = open_document(assembly_path, SW_DOC_ASSEMBLY)
    feature = doc.FirstFeature
    mates, scanned_groups = [], 0
    guard = 0
    while feature is not None and guard < 2000:
        guard += 1
        try:
            ftype = feature.GetTypeName2
        except Exception:
            ftype = None
        if ftype == "MateGroup":
            scanned_groups += 1
            sub = feature.GetFirstSubFeature
            sub_guard = 0
            while sub is not None and sub_guard < 2000:
                sub_guard += 1
                try:
                    mate = sub.GetSpecificFeature2
                except Exception:
                    mate = None
                if mate is not None:
                    row = {"name": str(sub.Name)}
                    try:
                        row["type"] = _MATE_TYPE_NAMES.get(
                            int(mate.Type), "unknown(%s)" % mate.Type)
                    except Exception:
                        row["type"] = "unreadable"
                    entities = _mate_entities(mate)
                    row["entities"] = entities
                    components_involved = sorted({
                        e["component"] for e in entities if e.get("component")})
                    row["components"] = components_involved
                    if component is None or component in components_involved:
                        mates.append(row)
                sub = sub.GetNextSubFeature
        feature = feature.GetNextFeature
    result = ok(assembly=str(doc.GetTitle), mates=len(mates),
                mate_groups_scanned=scanned_groups, items=mates)
    if component is not None:
        result["filtered_for"] = component
    return result


@tool_guard("interference_failed")
def check_interference(assembly_path):
    """Detect overlapping component pairs in an assembly (read-only).

    InterferenceDetectionManager runs the same check as the UI's 干涉检查.
    Returns the colliding component pairs; overlap volume is reported when the
    binding exposes it (lightweight components must be resolved first or the
    check silently sees nothing).
    """
    doc = open_document(assembly_path, SW_DOC_ASSEMBLY)
    try:
        doc.ResolveAllLightweightComponents(True)
    except Exception:
        pass
    mgr = doc.InterferenceDetectionManager
    if mgr is None:
        return error("interference_unavailable",
                     "InterferenceDetectionManager returned null on this binding")
    try:
        mgr.ToolsCheckInterference2(0, None, True)
        count = mgr.GetInterferenceCount()
        pairs = []
        try:
            interferences = mgr.GetInterferences() or []
        except Exception:
            interferences = []
        for item in interferences:
            row = {"components": None, "volume_m3": None}
            try:
                comps = list(item.Components or [])
                row["components"] = sorted(
                    str(c.Name2) for c in comps if c is not None)
            except Exception:
                pass
            try:
                row["volume_m3"] = float(item.Volume)
            except Exception:
                pass
            pairs.append(row)
        return ok(assembly=str(doc.GetTitle), interferences=count,
                  pairs=pairs)
    finally:
        try:
            mgr.Done()
        except Exception:
            pass


@tool_guard("config_failed")
def list_configurations(model_path):
    """List a document's configurations and which one is active."""
    doc = open_document(model_path)
    names = []
    try:
        names = [str(n) for n in (doc.GetConfigurationNames or [])]
    except Exception:
        pass
    active = None
    try:
        active = str(doc.ConfigurationManager.ActiveConfiguration.Name)
    except Exception:
        pass
    return ok(document=str(doc.GetTitle), count=len(names), active=active,
              configurations=names)


@tool_guard("config_failed")
def activate_configuration(model_path, config_name, apply=False):
    """Switch to a named configuration (ShowConfiguration2). Dry-run by default.

    A dry run reports the target and the current active configuration without
    switching. apply=True switches and re-reads the active name — a refused
    switch (bad name, suppressed config) returns config_not_activated.
    """
    if not config_name:
        return error("bad_input", "config_name is required")
    doc = open_document(model_path)
    names = []
    try:
        names = [str(n) for n in (doc.GetConfigurationNames or [])]
    except Exception:
        pass
    active = None
    try:
        active = str(doc.ConfigurationManager.ActiveConfiguration.Name)
    except Exception:
        pass
    if config_name not in names:
        return error("config_not_found",
                     "No configuration named %r" % config_name,
                     "Available: %s" % names,
                     configurations=names, active=active)
    if not apply:
        return ok(dry_run=True, target=config_name, active=active,
                  configurations=names,
                  message="Nothing switched. apply=True activates it.")
    switched = doc.ShowConfiguration2(config_name)
    new_active = None
    try:
        new_active = str(doc.ConfigurationManager.ActiveConfiguration.Name)
    except Exception:
        pass
    if not switched or new_active != config_name:
        return error("config_not_activated",
                     "ShowConfiguration2 did not activate %r" % config_name,
                     active=new_active)
    return ok(action="activated", active=new_active, previous=active)


# swComponentSuppressionState_e for IComponent2.SetSuppression2.
_SUPPRESSION_STATES = {"lightweight": 0, "suppressed": 1, "resolved": 3}
_SUPPRESSION_READBACK = {0: "lightweight", 1: "suppressed", 3: "resolved"}


@tool_guard("state_failed")
def set_component_state(assembly_path, component, state, apply=False):
    """Set a component's suppression / visibility state. Dry-run by default.

    state: resolved | lightweight | suppressed | hidden | shown.
    `component` is the instance name from get_assembly_tree (e.g. "Plate-1");
    every matching instance is changed. The read-back uses GetSuppression, so
    a state SolidWorks refused is reported as not_applied rather than assumed.
    """
    if state not in _SUPPRESSION_STATES and state not in ("hidden", "shown"):
        return error("bad_input",
                     "state must be one of: resolved, lightweight, suppressed, "
                     "hidden, shown")
    doc = open_document(assembly_path, SW_DOC_ASSEMBLY)
    matches = [c for c in get_components(doc)
               if str(c.Name2) == component]
    if not matches:
        names = sorted(str(c.Name2) for c in get_components(doc))
        return error("component_not_found",
                     "No component named %r" % component,
                     "Components: %s" % names)
    before = []
    for comp in matches:
        row = {"component": str(comp.Name2)}
        try:
            row["suppression"] = _SUPPRESSION_READBACK.get(
                int(comp.GetSuppression), "unknown")
        except Exception:
            row["suppression"] = None
        try:
            row["visible"] = bool(comp.Visible)
        except Exception:
            row["visible"] = None
        before.append(row)
    if not apply:
        return ok(dry_run=True, target_state=state, matched=len(matches),
                  before=before,
                  message="Nothing changed. apply=True sets the state.")
    changed, failed = [], []
    for comp in matches:
        name = str(comp.Name2)
        try:
            if state == "hidden":
                comp.Visible = False
            elif state == "shown":
                comp.Visible = True
            else:
                comp.SetSuppression2(_SUPPRESSION_STATES[state])
            after = {"component": name}
            try:
                after["suppression"] = _SUPPRESSION_READBACK.get(
                    int(comp.GetSuppression), "unknown")
            except Exception:
                after["suppression"] = None
            try:
                after["visible"] = bool(comp.Visible)
            except Exception:
                after["visible"] = None
            changed.append(after)
        except Exception as exc:
            failed.append({"component": name, "error": str(exc)[:160]})
    result = ok(action="state_set", target_state=state,
                changed=len(changed), after=changed)
    if failed:
        result["failed"] = failed
    return result
