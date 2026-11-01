"""Custom property read/write — the data behind title blocks and BOM columns.

GB title blocks, BOM weight columns and PDM workflows all key off custom
properties, and until now the only way to fill them was the template's static
text. Reads go through ICustomPropertyManager (doc-level when config="", or a
named configuration), which reports both the stored expression and the resolved
value — "SW-Mass" stores a link, resolves to a number.

Writes are dry-run by default, like every mutating tool in this kit.
"""
from __future__ import annotations

from ..sw_session import error, ok, open_document, tool_guard


def _prop_manager(doc, config):
    try:
        return doc.Extension.CustomPropertyManager(config or "")
    except Exception:
        # Older bindings expose it directly on the document.
        try:
            return doc.CustomPropertyManager(config or "")
        except Exception:
            return None


def _read_prop(pm, name):
    """Return (value, resolved) for one property; either may be None."""
    value = resolved = None
    try:
        out = pm.Get5(name, True)  # (value, resolved, was_resolved, linked)
        if isinstance(out, (tuple, list)) and len(out) >= 2:
            value, resolved = out[0], out[1]
    except Exception:
        pass
    if value is None:
        try:
            out = pm.Get2(name)  # (value, resolved)
            if isinstance(out, (tuple, list)) and len(out) >= 2:
                value, resolved = out[0], out[1]
        except Exception:
            pass
    if value is None:
        try:
            value = pm.Get(name)
        except Exception:
            pass
    return value, resolved


@tool_guard("props_failed")
def get_custom_properties(path, config=None):
    """Read every custom property of a document (optionally per configuration).

    Returns each property's stored expression and its resolved value — the two
    differ for linked values like "SW-Mass@..." which is exactly what a BOM or
    title-block audit needs to see. Read-only.
    """
    doc = open_document(path)
    pm = _prop_manager(doc, config)
    if pm is None:
        return error("props_unavailable",
                     "CustomPropertyManager is not available on this binding")
    try:
        names = list(pm.GetNames or [])
    except Exception:
        names = []
    rows = []
    for name in names:
        value, resolved = _read_prop(pm, name)
        rows.append({"name": str(name),
                     "value": None if value is None else str(value),
                     "resolved": None if resolved is None else str(resolved)})
    return ok(document=str(doc.GetTitle), path=str(doc.GetPathName),
              configuration=config or "(document)",
              count=len(rows), properties=rows)


@tool_guard("prop_failed")
def set_custom_property(path, name, value, config=None, field_type=30, apply=False):
    """Add or update one custom property. Dry-run by default.

    field_type=30 is text (swCustomInfoText); other swCustomInfoType_e values:
    3 date, 11 yes/no, 64 number. Add3 with overwrite=swCustomPropertyReplaceValue
    updates in place. The write is verified by reading the value back.
    """
    if not name:
        return error("bad_input", "name is required")
    doc = open_document(path)
    pm = _prop_manager(doc, config)
    if pm is None:
        return error("props_unavailable",
                     "CustomPropertyManager is not available on this binding")
    old_value, old_resolved = _read_prop(pm, name)
    result = {
        "document": str(doc.GetTitle),
        "configuration": config or "(document)",
        "name": name,
        "previous_value": old_value,
        "previous_resolved": old_resolved,
        "new_value": str(value),
    }
    if not apply:
        result["dry_run"] = True
        result["action"] = "update" if old_value is not None else "add"
        return ok(**result)
    # Add3's last arg is OverwriteExisting — a plain Boolean, not the
    # swCustomInfoAddResult_e enum the return value uses.
    status = pm.Add3(name, int(field_type), str(value), True)
    new_value, new_resolved = _read_prop(pm, name)
    result.update(action="set", add_status=status,
                  applied_value=new_value, applied_resolved=new_resolved,
                  verified=str(new_value) == str(value))
    if not result["verified"]:
        return error(
            "prop_not_applied",
            "Add3 returned status %s but read-back shows %r, not %r"
            % (status, new_value, value),
            "The property may be linked to model data and reject direct writes.",
            **result)
    return ok(**result)


@tool_guard("prop_failed")
def delete_custom_property(path, name, config=None, apply=False):
    """Delete one custom property by name. Dry-run reports whether it exists."""
    if not name:
        return error("bad_input", "name is required")
    doc = open_document(path)
    pm = _prop_manager(doc, config)
    if pm is None:
        return error("props_unavailable",
                     "CustomPropertyManager is not available on this binding")
    old_value, _ = _read_prop(pm, name)
    if old_value is None:
        return error("prop_not_found",
                     "No custom property named %r" % name,
                     configuration=config or "(document)")
    if not apply:
        return ok(dry_run=True, name=name, previous_value=old_value,
                  configuration=config or "(document)",
                  message="Nothing deleted. apply=True removes it.")
    removed = pm.Delete2(name)
    new_value, _ = _read_prop(pm, name)
    return ok(action="deleted", name=name, previous_value=old_value,
              delete_return=removed, verified=new_value is None)
