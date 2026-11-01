"""Batch export of models to neutral formats.

Currently STEP only. Verified on SW 2024/2026:
`doc.SaveAsSilent(path.step, False)` writes a valid STEP and leaves the document
path unchanged (it behaves as an export, not a save-as), so no re-open is needed.

PDF comes from .SLDDRW and DXF usually from sheet-metal flat patterns / drawings,
so they need separate source types and are deliberately rejected instead of being
silently mangled.
"""
from __future__ import annotations

import os

from ..sw_session import SW_DOC_ASSEMBLY, SW_DOC_PART, error, ok, open_document, tool_guard

SUPPORTED = {"step": ".STEP"}
SOURCE_TYPES = {".sldprt": SW_DOC_PART, ".sldasm": SW_DOC_ASSEMBLY}


@tool_guard("export_failed")
def batch_export(targets, output_dir, fmt="step", apply=False):
    """Export parts/assemblies to STEP. Dry-run by default; never modifies the model."""
    fmt_key = (fmt or "").strip().lower()
    if fmt_key not in SUPPORTED:
        return error(
            "unsupported_format",
            "Only 'step' is supported right now (requested %r)." % fmt,
            "PDF must come from a .SLDDRW and DXF from a sheet-metal flat pattern or "
            "drawing — those need a separate tool with a different source document.",
        )

    if not targets:
        return error("bad_input", "targets is required, e.g. ['D:/parts/P01.SLDPRT']")
    if not output_dir:
        return error("bad_input", "output_dir is required")

    ext = SUPPORTED[fmt_key]
    out_dir = os.path.abspath(output_dir)
    planned, conflicts = [], []

    for raw in targets:
        src = os.path.abspath(raw)
        if not os.path.exists(src):
            conflicts.append({"path": src, "reason": "not_found"})
            continue
        src_ext = os.path.splitext(src)[1].lower()
        if src_ext not in SOURCE_TYPES:
            conflicts.append({"path": src, "reason": "unsupported_source"})
            continue
        base = os.path.splitext(os.path.basename(src))[0]
        dst = os.path.join(out_dir, base + ext)
        if os.path.exists(dst):
            conflicts.append({"path": src, "reason": "target_exists", "target": dst})
            continue
        planned.append({"source": src, "target": dst})

    result = {
        "format": fmt_key,
        "output_dir": out_dir,
        "planned": planned,
        "conflicts": conflicts,
        "exported": [],
        "failed": [],
        "applied": False,
    }
    if not planned:
        result["message"] = "Nothing to export."
        return ok(**result)
    if not apply:
        result["dry_run"] = True
        result["message"] = "Nothing written. apply=true exports the planned files."
        return ok(**result)

    os.makedirs(out_dir, exist_ok=True)
    for item in planned:
        try:
            doc = open_document(item["source"], SOURCE_TYPES[os.path.splitext(item["source"])[1].lower()])
            doc.SaveAsSilent(item["target"], False)
            if not os.path.exists(item["target"]):
                raise RuntimeError("SolidWorks reported success but no file was written")
            result["exported"].append({
                "source": item["source"],
                "target": item["target"],
                "bytes": os.path.getsize(item["target"]),
            })
        except Exception as exc:
            result["failed"].append({"source": item["source"], "error": str(exc)})
    result["applied"] = True
    if not result["exported"]:
        return error(
            "export_failed",
            "All %d export(s) failed" % len(result["failed"]),
            "Check that the source is not open read-only and the output folder is writable.",
            failed=result["failed"],
        )
    return ok(**result)
