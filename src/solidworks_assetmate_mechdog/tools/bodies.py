"""Body-level appearance for multibody parts.

Measured on SW 2024 SP4 (2026-10-26):
- IBody2.MaterialPropertyValues is writable but silently ineffective, and reads back None.
- Extension.SetMaterialPropertyValues + SelectByID2('name','SOLIDBODY') works for SOME
  bodies only (small bodies yes, the main 758-face shell no) — unusable as a guarantee.
- IFace2.MaterialPropertyValues works for every face, reads back correctly, and costs
  ~12 s for 758 faces. This module therefore paints per BODY by iterating its faces.

Trade-off: face-level appearance overrides any existing face appearance, so operate on a
copy. There is no reliable per-body single-call API; pretending otherwise is worse.
"""
from __future__ import annotations

from ..sw_session import SW_DOC_PART, double_array, error, ok, open_document, tool_guard
from .recolor import DEFAULT_LIGHTING, normalize_palette


def _body_info(body, index):
    info = {"index": index}
    for attr, key in (("Name", "name"), ("GetFaceCount", "faces"),
                      ("GetEdgeCount", "edges"), ("GetVertexCount", "vertices")):
        try:
            value = getattr(body, attr)
            info[key] = value() if callable(value) else value
        except Exception:
            info[key] = None
    return info


def _bodies(doc):
    return list(doc.GetBodies2(0, False) or [])


@tool_guard("bodies_failed")
def list_bodies(part_path):
    """List solid bodies of a part. Order is not stable across calls — match by name."""
    doc = open_document(part_path, SW_DOC_PART)
    bodies = _bodies(doc)
    return ok(
        part=doc.GetTitle,
        path=doc.GetPathName,
        body_count=len(bodies),
        bodies=[_body_info(b, i) for i, b in enumerate(bodies)],
        note="Body order can change between calls; use body names as keys.",
    )


def _paint_body(body, rgb):
    """Paint every face of one body. Returns (faces_painted, faces_previously_colored)."""
    faces = list(body.GetFaces() or [])
    color = double_array([float(c) for c in rgb] + DEFAULT_LIGHTING)
    previously = 0
    for face in faces:
        try:
            if face.MaterialPropertyValues:
                previously += 1
        except Exception:
            pass
        face.MaterialPropertyValues = color
    return len(faces), previously


@tool_guard("recolor_bodies_failed")
def recolor_bodies(part_path, palette, apply=False, save=True):
    """Dry-run by default. Gives each named solid body its own color (face-level)."""
    palette = normalize_palette(palette)
    doc = open_document(part_path, SW_DOC_PART)
    bodies = _bodies(doc)
    if not bodies:
        return error("no_bodies", "No solid bodies found in %s" % doc.GetTitle)

    items = [_body_info(b, i) for i, b in enumerate(bodies)]
    present = {item["name"] for item in items}
    result = {
        "part": doc.GetTitle,
        "path": doc.GetPathName,
        "body_count": len(bodies),
        "bodies": items,
        "matched_bodies": len([n for n in palette if n in present]),
        "palette_keys_not_found": sorted(set(palette) - present),
        "applied": False,
        "mechanism": "face-level appearance (IBody2/SetMaterialPropertyValues are unreliable)",
    }
    if not apply:
        result["dry_run"] = True
        result["message"] = (
            "Nothing changed. apply=true paints every face of the matched bodies; "
            "face-level appearance overrides existing face colors, so test on a copy."
        )
        return ok(**result)

    painted, failed = [], []
    total_faces = 0
    overwritten = 0
    for body in bodies:
        name = body.Name
        rgb = palette.get(name)
        if rgb is None:
            continue
        try:
            count, previously = _paint_body(body, rgb)
            total_faces += count
            overwritten += previously
            painted.append({"body": name, "faces": count})
        except Exception as exc:
            failed.append({"body": name, "error": str(exc)})

    try:
        doc.ForceRebuild3(False)
    except Exception:
        doc.EditRebuild3
    if save:
        doc.SaveSilent  # property access performs the save; see quirks section 7

    result.update(
        applied=True,
        painted_bodies=painted,
        failed_bodies=failed,
        faces_painted=total_faces,
        faces_with_previous_color=overwritten,
        saved=bool(save),
        note="Per-body readback is unavailable; verify visually or by exporting a preview. "
             "If this part is used in an assembly that carries a component-level appearance, "
             "the assembly keeps showing the component color — component level overrides "
             "face level in the viewport.",
    )
    return ok(**result)
