"""Create a populated native TopLevelOnly BOM in a SolidWorks drawing.

Verified on SOLIDWORKS 2024 SP4 (32.4.0), Python + pywin32 late binding.

Why the indirect sequence is needed:
- InsertBomTable4 with BomType=TopLevelOnly(2) creates only the header in the tested
  environment, and BomFeature.SetConfigurations does not repair it.
- InsertBomTable4 with BomType=Indented(3) and a valid configuration populates rows.
- Switching BomFeature.TableType to TopLevelOnly(2) then produces top-level rows plus
  one duplicate Quantity column. Hiding that duplicate column yields the target table.

Do NOT call SetConfigurations in this workflow: it clears the table again.
Do NOT try to add a Quantity column with SetColumnType: SOLIDWORKS forbids it.
"""
from __future__ import annotations

import os
from typing import Any

import pythoncom
import win32com.client as win

SW_DOC_ASSEMBLY = 2
SW_OPEN_SILENT = 1
BOM_TYPE_PARTS_ONLY = 1
BOM_TYPE_TOP_LEVEL_ONLY = 2
BOM_TYPE_INDENTED = 3
ANCHOR_TOP_LEFT = 1
COLUMN_TYPE_QUANTITY = 203


# One definition, in sw_session (reports caught this duplicated for four rounds: two copies
# drift, and the retry/classification logic only ever landed in one of them).
from ..sw_session import call_method  # noqa: E402,F401


def set_column_hidden(table: Any, index: int, hidden: bool) -> None:
    """Set ITableAnnotation.ColumnHidden via PROPERTYPUT, not Python assignment."""
    ole = table._oleobj_
    dispid = ole.GetIDsOfNames("ColumnHidden")
    ole.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, True, int(index), bool(hidden))


def connect_sw() -> Any:
    pythoncom.CoInitialize()
    return win.dynamic.Dispatch(win.GetActiveObject("SldWorks.Application"))


def open_assembly(sw: Any, assembly_path: str) -> Any:
    """Open or reuse the assembly. Reuses an already-open document by normalized path."""
    assembly_path = os.path.abspath(assembly_path)
    if not os.path.exists(assembly_path):
        raise FileNotFoundError(assembly_path)
    for doc in list(sw.GetDocuments or ()):
        try:
            if os.path.normcase(doc.GetPathName) == os.path.normcase(assembly_path):
                return doc
        except Exception:
            continue
    errors = win.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    doc = sw.OpenDoc6(
        assembly_path, SW_DOC_ASSEMBLY, SW_OPEN_SILENT, "", errors, warnings
    )
    if doc is None:
        raise RuntimeError(
            f"OpenDoc6 failed for {assembly_path}: "
            f"errors={errors.value}, warnings={warnings.value}"
        )
    return doc


def default_template(sw: Any, language: str = "chinese-simplified") -> str:
    path = os.path.join(
        str(sw.GetExecutablePath), "lang", language, "bom-standard.sldbomtbt"
    )
    if not os.path.exists(path):
        raise FileNotFoundError(f"BOM template not found: {path}")
    return path


def create_view(drawing: Any, assembly_path: str) -> Any:
    """Create one front view, handling localized SolidWorks view names."""
    last_error = None
    for name in ("*Front", "*前视", "前视", "*前視"):
        try:
            view = drawing.CreateDrawViewFromModelView3(
                assembly_path, name, 0.15, 0.15, 0.0
            )
            if view is not None:
                return view
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"Could not create drawing view: {last_error}")


def table_headers(table: Any) -> list[str]:
    return [
        str(table.GetColumnTitle2(i, True)).replace("<FONT size=3.5>", "")
        for i in range(int(table.ColumnCount))
    ]


def populate_top_level_bom(
    drawing: Any,
    view: Any,
    template: str,
    configuration: str = "Default",
    x: float = 0.24,
    y: float = 0.19,
) -> dict[str, Any]:
    """Convert an assembly view into a populated native TopLevelOnly BOM."""
    drawing.ActivateView(view.Name)
    drawing.ClearSelection2(True)

    # Step 1: create a populated Indented BOM. A direct TopLevelOnly BOM is empty here.
    table = view.InsertBomTable4(
        False,
        float(x), float(y),
        ANCHOR_TOP_LEFT,
        BOM_TYPE_INDENTED,
        str(configuration),
        str(template),
        False,
        0,
        False,
    )
    if table is None:
        from ..sw_session import sw_version as _sw_version
        raise RuntimeError(
            "InsertBomTable4(Indented) returned None [sw_version=%s]. This step is verified "
            "on SW2024 SP4 only — a SW2026 report fails right here, so version is the first "
            "thing to check. If you are on SW2024: (1) confirm the configuration name exists "
            "in the assembly, (2) confirm the BOM template path is valid for this install, "
            "(3) confirm the drawing view really references an assembly. "
            "See check_capabilities -> bom_table for the evidence trail."
            % (_sw_version() or "unknown"))
    table = win.Dispatch(table)
    call_method(drawing, "ForceRebuild3", False)
    if int(table.RowCount) <= 1:
        raise RuntimeError("Indented BOM has no data rows; check model/configuration")

    # Step 2: convert to TopLevelOnly. This yields the desired module rows and, in this
    # environment, one duplicate Quantity column.
    feature = table.BomFeature
    if feature is None:
        raise RuntimeError("BomFeature is unavailable; table is not a native BOM")
    feature.TableType = BOM_TYPE_TOP_LEVEL_ONLY
    call_method(drawing, "ForceRebuild3", False)

    # Step 3: hide only the duplicate LAST Quantity column. Do not SetConfigurations.
    if int(table.ColumnCount) == 5:
        set_column_hidden(table, int(table.ColumnCount) - 1, True)
        call_method(drawing, "ForceRebuild3", False)

    rows = int(table.RowCount)
    columns = int(table.ColumnCount)
    headers = table_headers(table)
    quantity_indexes = [
        i for i in range(columns) if int(table.GetColumnType(i)) == COLUMN_TYPE_QUANTITY
    ]
    if rows <= 1:
        raise RuntimeError(f"TopLevelOnly BOM has no data rows (rows={rows})")
    if columns != 4 or len(quantity_indexes) != 1:
        raise RuntimeError(
            f"Unexpected final BOM shape: rows={rows}, columns={columns}, "
            f"quantity_indexes={quantity_indexes}"
        )

    data = []
    for r in range(1, rows):
        data.append([str(table.Text(r, c)) for c in range(columns)])
    return {
        "rows": rows,
        "columns": columns,
        "data_rows": rows - 1,
        "headers": headers,
        "quantity_column_index": quantity_indexes[0],
        "data": data,
        "referenced_model": str(feature.GetReferencedModelName),
        "table_type": int(feature.TableType),
        "native_bom": True,
    }


def create_drawing_with_bom(
    assembly_path: str,
    output_drawing: str,
    configuration: str = "Default",
    template: str | None = None,
) -> dict[str, Any]:
    """End-to-end: assembly -> new A3 drawing -> populated native 4-column BOM."""
    assembly_path = os.path.abspath(assembly_path)
    output_drawing = os.path.abspath(output_drawing)
    os.makedirs(os.path.dirname(output_drawing), exist_ok=True)

    sw = connect_sw()
    open_assembly(sw, assembly_path)
    template = template or default_template(sw)

    # NewDocument needs a valid template; fall back to SolidWorks' default if the
    # language-specific Tutorial path is absent.
    drawing_template = os.path.join(
        str(sw.GetExecutablePath), "data", "templates", "gb.drwdot"
    )
    drawing = sw.NewDocument(drawing_template, 9, 0, 0)  # 9 = A3
    if drawing is None:
        drawing = sw.NewDocument("", 0, 0, 0)
    if drawing is None:
        raise RuntimeError("NewDocument failed")

    view = create_view(drawing, assembly_path)
    # Save once before annotation work; this also avoids new-document annotation quirks.
    drawing.SaveAs3(output_drawing, 0, 0)
    result = populate_top_level_bom(drawing, view, template, configuration)
    drawing.SaveAs3(output_drawing, 0, 0)
    result.update({
        "ok": True,
        "assembly": assembly_path,
        "drawing": output_drawing,
        "configuration": configuration,
        "template": template,
    })
    return result
