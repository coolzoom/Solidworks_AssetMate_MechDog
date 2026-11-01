"""Component / part renaming.

Two very different operations, with different levels of API support:

1. File rename (mode="file") — implemented with the filesystem. SolidWorks'
   Extension.RenameDocument returns 2 and does nothing outside a PDM workspace,
   so the OS rename is the only reliable path. Every rename is dry-run first,
   refuses to overwrite, and rolls back on failure.

2. Instance rename (mode="instance") — NOT supported by the API on this build.
   IComponent2.Name is read-only ("Property can not be set") and Name2 silently
   ignores assignment. SolidWorks only exposes this through the FeatureManager UI.
   Returning a fake success would be worse than saying so.

Renaming a file breaks assembly references that point at the old name; pass
assembly_path to get a report of affected components before applying.
"""
from __future__ import annotations

import os

from ..sw_session import (
    tool_guard,
    SW_DOC_ASSEMBLY,
    SW_DOC_PART,
    error,
    get_components,
    ok,
    open_document,
)

ALLOWED_EXT = {".sldprt", ".sldasm", ".slddrw"}
ILLEGAL = set('<>:"/\\|?*')


def _target_name(path, mapping, prefix, suffix):
    base = os.path.splitext(os.path.basename(path))[0]
    ext = os.path.splitext(path)[1]
    new_base = mapping.get(base, mapping.get(os.path.basename(path), base)) if mapping else base
    new_base = f"{prefix or ''}{new_base}{suffix or ''}"
    return os.path.join(os.path.dirname(path), new_base + ext)


def _reference_report(assembly_path, old_names):
    """Which components in the assembly reference the files being renamed."""
    doc = open_document(assembly_path, SW_DOC_ASSEMBLY)
    affected = []
    for comp in get_components(doc):
        try:
            current = comp.GetPathName
        except Exception:
            continue
        if current and os.path.basename(current).lower() in old_names:
            affected.append({"component": comp.Name2, "references": current})
    return affected


def _resolve(assembly_path):
    doc = open_document(assembly_path, SW_DOC_ASSEMBLY)
    try:
        doc.ResolveAllLightweightComponents(True)
    except Exception:
        pass
    return doc


def _auto_rename(assembly_path, part_paths, mapping, prefix, suffix, save):
    """Rename files AND re-point the assembly references in one operation.

    Verified path (SW 2024): IComponent2.ReplaceReference(new_path) returns 1 and
    retargets that component immediately; the change persists after save+reopen.
    Side effect: instance names follow the new file name (P01-1 -> P01_new-1).
    """
    from ..sw_session import get_sw, invalidate_components
    sw = get_sw()
    doc = _resolve(assembly_path)
    components = get_components(doc, refresh=True)

    by_file = {}
    for comp in components:
        try:
            current = comp.GetPathName
        except Exception:
            continue
        by_file.setdefault(os.path.normcase(current), []).append(comp)

    plans = []
    for raw in part_paths:
        old = os.path.abspath(raw)
        new = _target_name(old, mapping or {}, prefix, suffix)
        if os.path.normcase(old) == os.path.normcase(new):
            continue
        refs = by_file.get(os.path.normcase(old), [])
        plans.append({"old": old, "new": new, "components": refs})

    result = {
        "mode": "auto",
        "assembly": doc.GetTitle,
        "planned": [
            {
                "old": p["old"],
                "new": p["new"],
                "components": [c.Name2 for c in p["components"]],
            }
            for p in plans
        ],
        "applied": False,
    }
    if not plans:
        return error("nothing_planned", "No matching components/files to rename")

    created, replaced, failed = [], [], []
    for plan in plans:
        try:
            if os.path.exists(plan["new"]):
                raise RuntimeError("target already exists: %s" % plan["new"])
            # The assembly locks its component files, so renaming on disk is impossible
            # while it is open. Save a copy under the new name instead, then retarget.
            part = open_document(plan["old"], SW_DOC_PART)
            part.SaveAsSilent(plan["new"], False)
            created.append(plan["new"])
            for comp in plan["components"]:
                try:
                    if comp.ReplaceReference(plan["new"]):
                        replaced.append({"component": comp.Name2, "file": plan["new"]})
                    else:
                        failed.append({"component": comp.Name2, "reason": "replace_returned_false"})
                except Exception as exc:
                    failed.append({"component": comp.Name2, "reason": str(exc)})
            if part.GetSaveFlag:
                part.SaveSilent
            sw.CloseDoc(part.GetTitle)
        except Exception as exc:
            failed.append({"file": plan["old"], "error": str(exc)})

    try:
        doc.ForceRebuild3(False)
    except Exception:
        pass
    if save:
        doc.SaveSilent

    invalidate_components(assembly_path)
    verify = []
    for comp in get_components(doc, refresh=True):
        try:
            current = comp.GetPathName
        except Exception:
            continue
        if os.path.normcase(current) in {os.path.normcase(p) for p in created}:
            verify.append({"component": comp.Name2, "file": current})

    result.update(
        applied=True,
        new_files=created,
        replaced=replaced,
        failed=failed,
        saved=bool(save),
        verified_components=verify,
        originals_kept=True,
        note="New files were created (originals kept). Instance names follow the new "
             "file name (P01-1 -> P01_new-1); no API exists for arbitrary instance names.",
    )
    return ok(**result)


@tool_guard("rename_failed")
def rename_components(
    part_paths=None,
    mapping=None,
    prefix=None,
    suffix=None,
    apply=False,
    assembly_path=None,
    mode="file",
    save=True,
):
    """Rename part/assembly FILES by mapping or prefix/suffix. Dry-run by default.

    mode="file"  rename on disk only; assembly references are reported, not changed.
    mode="auto"  rename + retarget assembly references via ReplaceReference, then save.
    mode="instance" reports that the SolidWorks API cannot set instance names directly.

    Note: after mode="auto", instance names follow the new file name
    (P01-1 -> P01_new-1). There is no API to set an arbitrary instance name.
    """
    if mode == "auto":
        if not assembly_path:
            return error("bad_input", "mode='auto' requires assembly_path")
        if not part_paths:
            return error("bad_input", "part_paths is required, e.g. ['D:/parts/P01.SLDPRT']")
        if not mapping and not prefix and not suffix:
            return error("bad_input", "Provide mapping, prefix, or suffix — otherwise nothing changes")
        try:
            if apply:
                return _auto_rename(assembly_path, part_paths, mapping, prefix, suffix, save)
            doc = _resolve(assembly_path)
            components = get_components(doc, refresh=True)
            by_file = {}
            for comp in components:
                try:
                    by_file.setdefault(os.path.normcase(comp.GetPathName), []).append(comp.Name2)
                except Exception:
                    continue
            planned, conflicts = [], []
            for raw in part_paths:
                old = os.path.abspath(raw)
                if not os.path.exists(old):
                    conflicts.append({"path": old, "reason": "not_found"})
                    continue
                new = _target_name(old, mapping or {}, prefix, suffix)
                if os.path.normcase(old) == os.path.normcase(new):
                    conflicts.append({"path": old, "reason": "name_unchanged"})
                    continue
                refs = by_file.get(os.path.normcase(old), [])
                if not refs:
                    conflicts.append({"path": old, "reason": "not_referenced_by_assembly"})
                    continue
                planned.append({"old": old, "new": new, "components": refs})
            return ok(
                mode="auto",
                assembly=doc.GetTitle,
                planned=planned,
                conflicts=conflicts,
                applied=False,
                dry_run=True,
                message="Nothing changed. apply=true renames files and re-points the references.",
            )
        except (FileNotFoundError, ValueError) as exc:
            return error("bad_input", str(exc))
        except Exception as exc:
            return error("rename_failed", str(exc))

    if mode != "file":
        return error(
            "instance_rename_unsupported",
            "SolidWorks exposes no writable API for component instance names on this build: "
            "IComponent2.Name is read-only and Name2 ignores assignment.",
            "Rename instances in the FeatureManager tree (slow double-click), or use "
            "mode='auto' to rename the file and let the instance name follow it.",
        )

    if not part_paths:
        return error("bad_input", "part_paths is required, e.g. ['D:/parts/P01.SLDPRT']")
    if not mapping and not prefix and not suffix:
        return error("bad_input", "Provide mapping, prefix, or suffix — otherwise nothing changes")

    changes, conflicts = [], []
    for raw in part_paths:
        path = os.path.abspath(raw)
        if not os.path.exists(path):
            conflicts.append({"path": path, "reason": "not_found"})
            continue
        if os.path.splitext(path)[1].lower() not in ALLOWED_EXT:
            conflicts.append({"path": path, "reason": "unsupported_extension"})
            continue
        target = _target_name(path, mapping or {}, prefix, suffix)
        if os.path.normcase(target) == os.path.normcase(path):
            conflicts.append({"path": path, "reason": "name_unchanged"})
            continue
        if any(ch in ILLEGAL for ch in os.path.basename(target)):
            conflicts.append({"path": path, "reason": "illegal_characters", "target": target})
            continue
        if os.path.exists(target):
            conflicts.append({"path": path, "reason": "target_exists", "target": target})
            continue
        changes.append({"old": path, "new": target})

    result = {
        "mode": "file",
        "planned": changes,
        "conflicts": conflicts,
        "applied": False,
    }
    if assembly_path and changes:
        old_names = {os.path.basename(c["old"]).lower() for c in changes}
        try:
            result["affected_components"] = _reference_report(assembly_path, old_names)
        except Exception as exc:
            result["reference_check_error"] = str(exc)

    if not apply:
        result["dry_run"] = True
        result["message"] = "Nothing renamed. apply=true performs the file renames."
        return ok(**result)

    done = []
    try:
        for change in changes:
            os.rename(change["old"], change["new"])
            done.append(change)
    except Exception as exc:
        for change in reversed(done):
            try:
                os.rename(change["new"], change["old"])
            except Exception:
                pass
        return error(
            "rename_failed",
            "%s; rolled back %d completed rename(s)" % (exc, len(done)),
            "Close the file in SolidWorks first — an open document cannot be renamed.",
        )

    result.update(applied=True, renamed=done)
    if result.get("affected_components"):
        result["warning"] = (
            "Renamed files are referenced by %d component(s); reopen the assembly or "
            "re-point the references, otherwise it will show missing components."
            % len(result["affected_components"])
        )
    return ok(**result)
