"""Document lifecycle: list / open / activate / close / save / import / macro.

Every tool in this kit used to open documents implicitly through
sw_session.open_document() — fine for single-shot calls, but there was no way to
see what is loaded, close it (SolidWorks refuses new opens once ~40 documents
are loaded, which is why survey/sampling code must close each drawing), or save
to a chosen path. These tools expose that session layer directly.

Neutral-file import (STEP/IGES/x_t/SAT/STL) goes through ISldWorks.LoadFile4 —
OpenDoc6 only accepts .sldprt/.sldasm/.slddrw types.
"""
from __future__ import annotations

import os

from ..sw_session import (
    SW_DOC_ASSEMBLY,
    SW_DOC_DRAWING,
    SW_DOC_PART,
    call_method,
    check_path_allowed,
    error,
    get_sw,
    invalidate_components,
    ok,
    open_document,
    tool_guard,
)

_DOC_TYPE_NAMES = {SW_DOC_PART: "part", SW_DOC_ASSEMBLY: "assembly",
                   SW_DOC_DRAWING: "drawing"}

_NEUTRAL_EXTENSIONS = {".step", ".stp", ".iges", ".igs", ".x_t", ".x_b",
                       ".sat", ".stl"}


def _doc_row(doc):
    row = {"title": None, "path": None, "type": None, "unsaved_changes": None,
           "read_only": None}
    try:
        row["title"] = str(doc.GetTitle)
    except Exception:
        pass
    try:
        row["path"] = str(doc.GetPathName) or None
    except Exception:
        pass
    try:
        row["type"] = _DOC_TYPE_NAMES.get(int(doc.GetType), "unknown(%s)" % doc.GetType)
    except Exception:
        pass
    try:
        row["unsaved_changes"] = bool(doc.GetSaveFlag())
    except Exception:
        pass
    try:
        row["read_only"] = bool(doc.IsReadOnly)
    except Exception:
        pass
    return row


def list_documents():
    """Every document currently loaded in the SolidWorks session."""
    try:
        sw = get_sw()
        docs = sw.GetDocuments
        rows = []
        for doc in (docs or []):
            rows.append(_doc_row(doc))
        active = None
        try:
            active_doc = sw.ActiveDoc
            active = str(active_doc.GetTitle) if active_doc is not None else None
        except Exception:
            pass
        return ok(count=len(rows), active=active, documents=rows)
    except Exception as exc:
        return error("list_failed", str(exc))


def open_file(path, activate=True):
    """Open a native SolidWorks document and report what was loaded.

    Reuses sw_session.open_document so the active-document check, the byref
    error/warning VARIANTs and the component-cache invalidation all apply.
    Neutral formats (STEP/IGES/...) are rejected with a pointer to import_file.
    """
    try:
        path = check_path_allowed(path)
        ext = os.path.splitext(path)[1].lower()
        if ext in _NEUTRAL_EXTENSIONS:
            return error(
                "neutral_format",
                "%s is a neutral CAD file — import_file converts it to a part."
                % os.path.basename(path),
            )
        doc = open_document(path)
        if activate:
            try:
                _activate(doc)
            except Exception:
                pass
        return ok(action="opened", **_doc_row(doc))
    except PermissionError as exc:
        return error("path_not_allowed", str(exc))
    except (FileNotFoundError, ValueError) as exc:
        return error("bad_input", str(exc))
    except Exception as exc:
        return error("open_failed", str(exc))


def _activate(doc):
    """ActivateDoc3 takes four args; the fourth is a byref error VARIANT."""
    from ..sw_session import _com, variant

    pythoncom, _, _ = _com()
    errors = variant(0, pythoncom.VT_BYREF | pythoncom.VT_I4)
    return get_sw().ActivateDoc3(str(doc.GetTitle), True, 0, errors)


def _find_open_doc(ref):
    """Match an open document by title, full path, or basename."""
    sw = get_sw()
    docs = list(sw.GetDocuments or [])
    if not docs:
        return None, []
    ref_norm = os.path.normcase(ref)
    for doc in docs:
        try:
            if os.path.normcase(doc.GetPathName) == ref_norm:
                return doc, docs
        except Exception:
            pass
    base = os.path.basename(ref)
    matches = []
    for doc in docs:
        try:
            title = str(doc.GetTitle)
            path = str(doc.GetPathName or "")
            if title == ref or title == base or os.path.basename(path) == base:
                matches.append(doc)
        except Exception:
            pass
    if len(matches) == 1:
        return matches[0], docs
    return None, docs


def activate_document(ref):
    """Bring an already-open document to the front (ActiveDoc).

    ref accepts a title ("Part1"), a full path, or a basename. Refuses when the
    name is ambiguous instead of activating an arbitrary match.
    """
    try:
        doc, docs = _find_open_doc(ref)
        if doc is None:
            titles = []
            for d in docs:
                try:
                    titles.append(str(d.GetTitle))
                except Exception:
                    pass
            return error(
                "doc_not_open",
                "No unique open document matches %r" % ref,
                "Open it first with open_document. Open documents: %s" % titles,
            )
        activated = _activate(doc)
        result = activated if activated is not None else doc
        return ok(action="activated", **_doc_row(result))
    except Exception as exc:
        return error("activate_failed", str(exc))


def close_document(ref=None, save=False, all_docs=False):
    """Close a document by title/path/basename — or everything with all_docs=True.

    Batch readers must close each document after use: SolidWorks stops opening
    files once roughly 40 are loaded. save=True persists first (the default is
    to discard; a document with unsaved changes is reported so the discard is
    never silent).
    """
    try:
        sw = get_sw()
        if all_docs:
            docs = list(sw.GetDocuments or [])
            titles = []
            for d in docs:
                try:
                    titles.append(str(d.GetTitle))
                except Exception:
                    pass
            sw.CloseAllDocuments(save)
            return ok(action="closed_all", count=len(titles), saved=save,
                      closed=titles)
        if not ref:
            return error("bad_input", "ref (title/path/basename) or all_docs=True required")
        doc, _ = _find_open_doc(ref)
        if doc is None:
            return error("doc_not_open", "No unique open document matches %r" % ref)
        row = _doc_row(doc)
        if row.get("unsaved_changes") and not save:
            row["discarded_unsaved_changes"] = True
        title = str(doc.GetTitle)
        try:
            path = str(doc.GetPathName or "")
            if path:
                invalidate_components(path)
        except Exception:
            pass
        sw.CloseDoc(title)
        return ok(action="closed", saved=save, **row)
    except Exception as exc:
        return error("close_failed", str(exc))


def save_document(path=None, new_path=None, overwrite=False):
    """Save a document in place, or save-as to a new path.

    new_path with a neutral extension (.STEP/.IGES/.X_T/.SAT/.STL) behaves as an
    EXPORT: SaveAsSilent writes the file and leaves the document's own path
    unchanged (same quirk batch_export relies on). Native extensions do a real
    save-as; overwrite=False refuses an existing target instead of clobbering it.
    """
    try:
        if path:
            doc = open_document(path)
        else:
            doc = get_sw().ActiveDoc
            if doc is None:
                return error("no_active_doc", "No active document and no path given")
        row = _doc_row(doc)
        if not new_path:
            doc.SaveSilent
            return ok(action="saved", **row)
        target = check_path_allowed(new_path)
        ext = os.path.splitext(target)[1].lower()
        if os.path.exists(target) and not overwrite:
            return error("target_exists",
                         "%s already exists — pass overwrite=True" % target)
        if ext in _NEUTRAL_EXTENSIONS:
            doc.SaveAsSilent(target, False)
            if not os.path.exists(target):
                return error("export_failed",
                             "SolidWorks reported success but wrote no file",
                             target=target)
            return ok(action="exported", target=target,
                      bytes=os.path.getsize(target),
                      note="neutral export — document path unchanged", **row)
        doc.SaveAsSilent(target, False)
        return ok(action="saved_as", target=target, **_doc_row(doc))
    except PermissionError as exc:
        return error("path_not_allowed", str(exc))
    except (FileNotFoundError, ValueError) as exc:
        return error("bad_input", str(exc))
    except Exception as exc:
        return error("save_failed", str(exc))


@tool_guard("rebuild_failed")
def rebuild_document(path=None, force=False):
    """Rebuild a model and report what changed.

    EditRebuild3 performs a normal rebuild; force=True calls ForceRebuild3(False)
    which rebuilds every feature. needs_rebuild is re-read afterwards so a
    drawing-induced dirty flag (which a rebuild clears) is told apart from a
    real model change.
    """
    if path:
        doc = open_document(path)
    else:
        doc = get_sw().ActiveDoc
        if doc is None:
            return error("no_active_doc", "No active document and no path given")
    row = _doc_row(doc)
    try:
        needed_before = bool(doc.Extension.NeedsRebuild2)
    except Exception:
        needed_before = None
    if force:
        doc.ForceRebuild3(False)
    else:
        call_method(doc, "EditRebuild3")
    try:
        needs_rebuild = bool(doc.Extension.NeedsRebuild2)
    except Exception:
        needs_rebuild = None
    return ok(action="rebuilt", forced=force,
              needed_rebuild_before=needed_before,
              needs_rebuild=needs_rebuild, **row)


def import_file(path, save_as=None):
    """Import a neutral CAD file (STEP/IGES/Parasolid/ACIS/STL) as a part.

    ISldWorks.LoadFile4 handles formats OpenDoc6 cannot. Feature recognition is
    deliberately out of scope (it belongs to parametric modeling toolkits); the
    imported solid lands as an imported body.
    """
    try:
        path = check_path_allowed(path)
        if not os.path.exists(path):
            return error("bad_input", "file not found: %s" % path)
        ext = os.path.splitext(path)[1].lower()
        if ext not in _NEUTRAL_EXTENSIONS:
            return error("bad_input",
                         "not a neutral format: %s (supported: %s)"
                         % (ext or "none", ", ".join(sorted(_NEUTRAL_EXTENSIONS))))
        sw = get_sw()
        from ..sw_session import _com, variant

        pythoncom, _, _ = _com()
        errors = variant(0, pythoncom.VT_BYREF | pythoncom.VT_I4)
        doc = sw.LoadFile4(path, "r", None, errors)
        try:
            err = int(errors.value)
        except Exception:
            err = 0
        if err:
            return error("import_failed", "LoadFile4 returned error code %s" % err,
                         path=path)
        doc = doc if doc is not None else sw.ActiveDoc
        if doc is None:
            return error("import_failed", "LoadFile4 succeeded but produced no document")
        row = _doc_row(doc)
        if save_as:
            target = check_path_allowed(save_as)
            doc.SaveAsSilent(target, False)
            row["saved_as"] = target
        return ok(action="imported", source=path, **row)
    except PermissionError as exc:
        return error("path_not_allowed", str(exc))
    except Exception as exc:
        return error("import_failed", str(exc))


@tool_guard("macro_failed")
def run_macro(macro_path, module="", procedure="", apply=False):
    """Run a SolidWorks VBA macro (.swp/.swb). Dry-run validates the file only.

    Escape hatch for operations the typed tools do not cover — e.g. the BOM
    workaround scripts in 项目规划. Runs whatever the macro does with no
    sandboxing, so apply=True is required.
    """
    path = check_path_allowed(macro_path)
    if not os.path.exists(path):
        return error("bad_input", "file not found: %s" % path)
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".swp", ".swb"):
        return error("bad_input", "expected a .swp/.swb macro, got %s" % ext)
    if not apply:
        return ok(dry_run=True, macro=path, module=module or None,
                  procedure=procedure or None,
                  message="Nothing ran. apply=True executes the macro.")
    sw = get_sw()
    result = sw.RunMacro(path, module or "main", procedure or "main")
    # RunMacro return codes: 0 ok; -1 error; else the macro's own return.
    return ok(action="macro_ran", macro=path, return_code=result,
              note="0 = success; negative codes are SolidWorks-side errors")
