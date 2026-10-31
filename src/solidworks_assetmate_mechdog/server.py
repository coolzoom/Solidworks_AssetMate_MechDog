"""MCP entry point. Registered name: Solidworks_AssetMate_MechDog.

Every tool returns a JSON string (never prose) so agents can keep reasoning.
"""
from __future__ import annotations

import json

from mcp.server.fastmcp import FastMCP

from . import SERVER_NAME, __version__
from . import sw_session
from .tools import (assembly, bodies, bom, capabilities, drawing, export as export_tools,
                    layout, ledger, material as material_tools, pipeline, planner, recolor,
                    rename, sheet_metal, survey)

mcp = FastMCP(SERVER_NAME)
# FastMCP (mcp 1.x) has no `version` argument: the low-level server falls back to
# the mcp package version, so serverInfo.version reported 1.30.0 instead of this
# toolkit's version. Set it explicitly when the attribute exists.
if getattr(mcp, "_mcp_server", None) is not None:
    mcp._mcp_server.version = __version__


def _json(payload):
    """Serialise a tool result, and log it when it reports failure.

    Hooking here instead of in 51 individual wrappers keeps the coverage complete without
    touching every tool: every result passes through this one function. The tool name is
    taken from the caller's frame, which is the `@mcp.tool()` function itself.
    """
    import sys

    tool = None
    try:
        tool = sys._getframe(1).f_code.co_name
    except Exception:
        pass
    try:
        if isinstance(payload, dict):
            sw_session.note_last_call(tool=tool, args=None)
            if payload.get("ok") is False:
                sw_session.log_failure(tool, payload)
    except Exception:
        # Diagnostics must never break a tool.
        pass
    return json.dumps(payload, ensure_ascii=False, indent=2)


def _install_excepthook():
    """Flush unhandled tracebacks to stderr AND the failure log (P0-1, suggestion 4).

    The report noted there was no error output at all before the outage, so the user could
    not tell whether the MCP process or SolidWorks had broken. A bare traceback on stderr
    is also invisible inside most MCP hosts, so it is mirrored into
    logs/mcp-YYYYMMDD.log where `recent_failures` can read it back after the fact.
    """
    import sys
    import traceback

    previous = sys.excepthook

    def hook(exc_type, exc, tb):
        try:
            sys.stderr.write("".join(traceback.format_exception(exc_type, exc, tb)))
            sys.stderr.flush()
        except Exception:
            pass
        try:
            sw_session.log_failure(
                "unhandled", {"stage": "excepthook"},
                exc=exc if isinstance(exc, BaseException) else None,
                extra={"exc_type": getattr(exc_type, "__name__", None)})
        except Exception:
            pass
        try:
            previous(exc_type, exc, tb)
        except Exception:
            pass

    sys.excepthook = hook


_install_excepthook()


@mcp.tool()
def health_check(deep: bool = False) -> str:
    """Report whether the MCP bridge and SolidWorks are both alive.

    Tells apart three states that used to look identical: the bridge is dead (a stale COM
    pointer — `reconnect()` retries), SolidWorks is not responding, or everything is fine.
    `deep=True` also reads ActiveDoc, which is where a stale pointer usually surfaces.
    """
    return json.dumps(sw_session.health(deep=deep), ensure_ascii=False, indent=2)


@mcp.tool()
def ping() -> str:
    """Cheapest possible "is it alive" probe (P0-1).

    Answers in ~18 ms and never touches ActiveDoc, so it cannot hang on a modal dialog.
    Returns the SolidWorks revision, the round-trip latency, the last tool that ran, and
    the self-healing counters (`bridge.stale_drops` / `bridge.reconnects`). Use it to tell
    "alive but busy" from "dead"; use health(deep=True) when you also need ActiveDoc.
    """
    return _json(sw_session.ping())


@mcp.tool()
def recent_failures(tail: int = 20) -> str:
    """Read back the last N failures from logs/mcp-YYYYMMDD.log.

    A mid-run disconnect used to leave no trace at all. Every failing tool call now appends
    one JSON line (tool, args, payload, SW version, health snapshot, traceback).
    """
    import os

    path = os.path.join(sw_session.log_dir(), "mcp-%s.log" % __import__("datetime").datetime.now().strftime("%Y%m%d"))
    if not os.path.exists(path):
        return json.dumps({"ok": True, "log": path, "entries": [], "note": "no failures today"},
                          ensure_ascii=False, indent=2)
    with open(path, encoding="utf-8") as f:
        lines = [ln for ln in f.read().splitlines() if ln.strip()]
    entries = []
    for line in lines[-max(1, int(tail)):]:
        try:
            entries.append(json.loads(line))
        except Exception:
            continue
    return json.dumps({"ok": True, "log": path, "count": len(entries), "entries": entries},
                      ensure_ascii=False, indent=2)


@mcp.tool()
def produce_drawing(
    model_path: str,
    output_path: str | None = None,
    paper: str = "A3",
    views: list | None = None,
    category: str | None = None,
    do_hole_callouts: bool | None = None,
    do_surface_finish: bool | None = None,
    prune_to_plan: bool = True,
) -> str:
    """Model in, annotated and judged drawing out — the whole chain in one call.

    ledger → plan → views → dimensions → center marks → surface finish (category default)
    → hole callouts → audit → reconciliation → delivery verdict.

    Defaults come from the part category (钣金 / 一般零件 / 轴类 / 装配体) and are echoed in
    `defaults_applied`. A failing step is recorded and the chain continues.

    prune_to_plan=True makes the plan binding: dimensions the plan did not keep are deleted
    right after insertion. Set False for the old "import everything" behaviour.
    """
    return _json(pipeline.produce_drawing(
        model_path, output_path=output_path, paper=paper,
        views=tuple(views) if views else None, category=category,
        do_hole_callouts=do_hole_callouts, do_surface_finish=do_surface_finish,
        prune_to_plan=prune_to_plan))


@mcp.tool()
def plan_drawing(
    model_path: str,
    paper: str = "A3",
    views: list | None = None,
    include_unmarked: bool = False,
) -> str:
    """Decide before drawing: which dimensions to keep, how many sheets, what views.

    Read-only plan (nothing is created). Real drawings average 14 dimensions; importing
    every driving dimension puts ~44 on the sheet, which is transcription, not drafting.
    `keep_dimensions` is the shortlist to feed auto_dimension.
    """
    return _json(planner.plan_drawing(
        model_path, paper=paper,
        views=tuple(views) if views else ("front", "right"),
        include_unmarked=include_unmarked))


@mcp.tool()
def check_capabilities(operation: str | None = None) -> str:
    """What this toolkit can and cannot do — answer before a long pipeline, not after.

    Each entry states the limit, the evidence behind it, and the workaround. Use it to
    avoid a 400-second run that ends in "0 hole callouts".
    """
    return _json(capabilities.check_capabilities(operation=operation))


@mcp.tool()
def recolor_assembly(
    assembly_path: str,
    palette: dict,
    apply: bool = False,
    save: bool = True,
) -> str:
    """Recolor a SolidWorks assembly by part type, e.g. {"Crank": [0.88,0.2,0.18]}.

    Palette keys are component names without the '-N' instance suffix; RGB is 0..1.
    SAFE DEFAULT: apply=false is a dry run listing parts, current colors and
    matched components — with zero matches it still returns the part list so you
    can fix the keys. apply=true with zero matches returns no_match plus
    available_parts. Changes component appearance only, not material/density.
    Returns previous colors for undo; readback covers at most 8 part types
    (1 instance per palette key) and flags when the getter is unreliable.
    """
    return _json(recolor.recolor_assembly(assembly_path, palette, apply=apply, save=save))


@mcp.tool()
def list_bodies(part_path: str) -> str:
    """List solid bodies in a part: name, faces, edges, vertices.

    Use body names as palette keys — the returned order can change between calls.
    """
    return _json(bodies.list_bodies(part_path))


@mcp.tool()
def recolor_bodies(
    part_path: str,
    palette: dict,
    apply: bool = False,
    save: bool = True,
) -> str:
    """Recolor each solid body of a multibody PART a different color.

    Palette keys are body names from list_bodies, e.g. {"Hull": [0.8,0.2,0.2]}.
    SAFE DEFAULT: apply=false lists bodies and matched names without changing
    anything. Each paint is applied only after verifying exactly one body is
    selected (guards against a document-wide recolor).
    """
    return _json(bodies.recolor_bodies(part_path, palette, apply=apply, save=save))


@mcp.tool()
def export_bom(
    assembly_path: str,
    include_mass: bool = True,
    include_material: bool = True,
    csv_path: str | None = None,
    refresh: bool = False,
    resolve_lightweight: bool = True,
) -> str:
    """Export a BOM: one row per part type with qty, material, unit and total mass.

    Read-only. Lightweight components are resolved first (resolve_lightweight=false
    skips that, at the cost of null mass/material). Missing values are reported in
    `warnings`, never guessed. Optionally writes a UTF-8 CSV for Excel.
    """
    return _json(bom.export_bom(
        assembly_path, include_mass=include_mass, include_material=include_material,
        csv_path=csv_path, refresh=refresh, resolve_lightweight=resolve_lightweight,
    ))


@mcp.tool()
def rename_components(
    part_paths: list[str],
    mapping: dict | None = None,
    prefix: str | None = None,
    suffix: str | None = None,
    apply: bool = False,
    assembly_path: str | None = None,
    mode: str = "file",
    save: bool = True,
) -> str:
    """Rename part/assembly FILES by mapping or prefix/suffix. Dry-run by default.

    mode="file"  rename on disk only.
    mode="auto"  rename + retarget assembly references (ReplaceReference) + save.
                 Requires assembly_path. Instance names follow the new file name.
    mode="instance" is unsupported: SolidWorks has no writable instance-name API.
    """
    return _json(rename.rename_components(
        part_paths=part_paths, mapping=mapping, prefix=prefix, suffix=suffix,
        apply=apply, assembly_path=assembly_path, mode=mode, save=save,
    ))


@mcp.tool()
def batch_export(
    targets: list[str],
    output_dir: str,
    fmt: str = "step",
    apply: bool = False,
) -> str:
    """Export parts/assemblies to STEP in one pass (dry-run by default).

    Read-only on the model: SaveAsSilent() behaves as an export and leaves the
    document path unchanged. Existing targets are reported as conflicts, never
    overwritten. PDF needs a .SLDDRW and DXF a flat pattern — those are rejected
    with a hint rather than guessed at.
    """
    return _json(export_tools.batch_export(targets, output_dir, fmt=fmt, apply=apply))


@mcp.tool()
def set_view(
    drawing_path: str,
    view_name: str,
    x_m: float | None = None,
    y_m: float | None = None,
    scale: float | None = None,
    save: bool = True,
    break_alignment: bool = False,
) -> str:
    """Move and/or rescale one view (sheet metres).

    Writes IView.Position through DISPATCH_PROPERTYPUT with a VARIANT array — a plain
    assignment silently lands the view elsewhere.

    `break_alignment=True` calls IView.RemoveAlignment first (the UI's 解除对齐), so a section
    view can leave the strip beside its parent. It clears THIS view's alignment only; views
    that inherit from it stay as they are.
    """
    return _json(layout.set_view(drawing_path, view_name, x_m=x_m, y_m=y_m,
                                 scale=scale, save=save,
                                 break_alignment=break_alignment))


@mcp.tool()
def place_views(drawing_path: str, placements: list, save: bool = True) -> str:
    """Place several views at once: [{view, x, y, scale}] in sheet metres."""
    return _json(layout.place_views(drawing_path, placements, save=save))


@mcp.tool()
def delete_dimensions(
    drawing_path: str,
    keep_values: list | None = None,
    view_name: str | None = None,
    keep_tolerance_mm: float = 0.05,
    save: bool = True,
) -> str:
    """Delete dimensions whose value is not in keep_values (the plan_drawing shortlist).

    SolidWorks has no bulk delete; each dimension is selected by position and removed with
    EditDelete. Measured 2026-10-29: 14 dimensions in ONE call took < 5 s.
    """
    return _json(layout.delete_dimensions(drawing_path, keep_values=keep_values,
                                          view_name=view_name,
                                          keep_tolerance_mm=keep_tolerance_mm, save=save))


@mcp.tool()
def add_note(
    drawing_path: str, text: str, x_m: float, y_m: float,
    height_mm: float | None = None, bold: bool = False, save: bool = True,
    avoid_collisions: bool = False,
) -> str:
    """Add a drawing note (技术要求). x/y are sheet metres.

    `avoid_collisions=True` walks the note DOWN until it clears every model view, so the
    first line is not overprinted by a section arrow. The receipt reports
    `position_requested_m` vs `position_m` plus `collisions_avoided`.
    """
    return _json(layout.add_note(drawing_path, text, x_m, y_m,
                                 height_mm=height_mm, bold=bold, save=save,
                                 avoid_collisions=avoid_collisions))


@mcp.tool()
def layout_dimensions(
    drawing_path: str,
    view_name: str | None = None,
    min_gap_m: float = 0.006,
    cluster_tol_m: float = 0.0015,
    dry_run: bool = False,
    save: bool = True,
) -> str:
    """Pull apart dimensions that auto_dimension stacked on the same axis.

    Clusters by x (tolerance `cluster_tol_m`), sorts each cluster by y, pushes any closer
    than `min_gap_m` upward. Already-clear dimensions are not touched.
    Use dry_run=True first to see how many collisions it would fix.
    """
    return _json(drawing.layout_dimensions(drawing_path, view_name=view_name,
                                           min_gap_m=min_gap_m,
                                           cluster_tol_m=cluster_tol_m,
                                           dry_run=dry_run, save=save))


@mcp.tool()
def delete_view(drawing_path: str, view_name: str, save: bool = True) -> str:
    """Delete a drawing view by name.

    Create-only APIs meant one wrong `add_section_view` forced rebuilding the whole sheet.
    `view_name` comes from `drawing_audit` or `view_positions`.
    """
    return _json(layout.delete_view(drawing_path, view_name, save=save))


@mcp.tool()
def delete_annotation(drawing_path: str, name: str, save: bool = True,
                      position_m: list[float] | None = None) -> str:
    """Delete one annotation (hole callout / GT / surface finish / note / dimension) by name.

    Names come from `list_drawing_annotations` (e.g. "细节项目272"). All four kinds are
    verified: note, geometric tolerance, surface finish and dimension. Success is measured
    across every view, because deleting a dimension does not change the sheet's own count.

    `position_m` disambiguates repeated names — SolidWorks reuses D1/D2/D3 per sketch, so
    one sheet can hold eight annotations called "D1"; pass the position the listing reported.
    """
    return _json(layout.delete_annotation(drawing_path, name, save=save,
                                          position_m=position_m))


@mcp.tool()
def move_annotation(drawing_path: str, name: str, x_m: float, y_m: float,
                    save: bool = True, position_m: list[float] | None = None) -> str:
    """Move an annotation or dimension to a sheet position in metres.

    Returns `requested` vs `position_after_m` separately — section and detail views keep
    projection alignment and may refuse the requested spot.

    `position_m` disambiguates repeated names (see delete_annotation).
    """
    return _json(layout.move_annotation(drawing_path, name, x_m, y_m, save=save,
                                        position_m=position_m))


@mcp.tool()
def set_dimension_position(drawing_path: str, name: str, x_m: float, y_m: float,
                           save: bool = True) -> str:
    """Move ONE dimension to a sheet position in metres, by its reported name.

    P0-2: single-dimension counterpart to layout_dimensions' bulk auto-placement. Verifies
    the target is a dimension (not a note/GTOL), and reports value plus before/after
    positions. Use list_drawing_annotations first to get the `name`.
    """
    return _json(layout.set_dimension_position(drawing_path, name, x_m, y_m, save=save))


@mcp.tool()
def add_bom_table(
    drawing_path: str,
    view_name: str | None = None,
    configuration: str = "Default",
    template: str | None = None,
    save: bool = True,
) -> str:
    """Insert a POPULATED native BOM (materials list) into an assembly drawing.

    Returns 10 rows x 4 columns with a quantity column on the reference assembly.

    Verified on SW2024 SP4 only. On other builds `InsertBomTable4` may return None;
    `check_capabilities` reports the version-qualified status and `bom_native` includes
    the SW version in the failure message. Wrapper/implementation signature drift is
    guarded by `static_check.check_wrapper_signatures` (a regression assertion, not a
    runtime concern — see CONTRIBUTING).
    """
    return _json(layout.add_bom_table(drawing_path, view_name=view_name,
                                      configuration=configuration,
                                      template=template, save=save))


@mcp.tool()
def add_sheet_line(
    drawing_path: str, x1_m: float, y1_m: float, x2_m: float, y2_m: float, save: bool = True,
) -> str:
    """Draw a line in true sheet space (metres) — for BOM tables and frames.

    Drawing sketches take sheet metres, not millimetres: passing 236/62 produced a scaled
    fan of diagonals instead of a horizontal line.
    """
    return _json(layout.add_sheet_line(drawing_path, x1_m, y1_m, x2_m, y2_m, save=save))


@mcp.tool()
def sheet_metal_qc(model_path: str, include_cylinders: bool = False) -> str:
    """Sheet-metal manufacturability checks: thickness, bend count, bend relief.

    Each check returns pass / fail / warn / not_checked. `not_checked` means it was NOT
    evaluated — geometry-distance checks (min flange, hole-to-bend) are not implemented yet
    and say so instead of returning a green tick.

    include_cylinders=True adds a cylinder radius inventory (every hole and bend fillet,
    grouped by radius). Costs about 90 extra seconds on a 556-face part, hence opt-in.
    """
    return _json(sheet_metal.sheet_metal_qc(model_path, include_cylinders=include_cylinders))


@mcp.tool()
def flat_pattern_report(model_path: str) -> str:
    """Where the flat-pattern size comes from — and what this toolkit did not do.

    The blank is computed by SolidWorks' SheetMetal engine. AssetMate only projects the
    flattened state into a view; it does not read the K-factor or bend table, does not
    verify the blank, and must not have process parameters invented on the drawing.
    """
    return _json(sheet_metal.flat_pattern_report(model_path))


@mcp.tool()
def add_geometric_tolerance(
    drawing_path: str,
    x: float,
    y: float,
    tolerance: str = "0.05",
    datum: str = "A",
    symbol: str = "none",
    view_name: str | None = None,
    edge_index: int | None = None,
    second_datum: str = "",
    third_datum: str = "",
    save: bool = True,
) -> str:
    """Insert a NATIVE geometric tolerance frame (形位公差).

    The API works on SW2026 — it only fails when pywin32 dispatches it as a property. This
    calls the same DISPID as DISPATCH_METHOD, which returns the Gtol object.

    symbol: none / position / flatness / perpendicularity / parallelism / concentricity /
    circularity / cylindricity / profile_line / profile_surface / angularity / symmetry /
    runout / total_runout. The symbol is written as readable TEXT into the first frame
    cell (SolidWorks does not parse <MOD-*> format codes here) and verified by read-back;
    a failed write returns gtol_symbol_not_applied instead of a frame with no symbol.
    Pass view_name + edge_index to attach the leader to geometry.
    """
    return _json(drawing.add_geometric_tolerance(
        drawing_path, x, y, tolerance=tolerance, datum=datum, symbol=symbol,
        view_name=view_name, edge_index=edge_index, second_datum=second_datum,
        third_datum=third_datum, save=save))


@mcp.tool()
def add_surface_finish(
    drawing_path: str,
    x: float,
    y: float,
    roughness: str = "3.2",
    sym_type: str = "machining",
    leader_type: str = "none",
    lay_symbol: int = 0,
    arrow_type: int = 0,
    prod_method: str = "",
    sample_len: str = "",
    min_roughness: str = "",
    roughness_spacing: str = "",
    save: bool = True,
    verify_render: bool = False,
    verify_dwg: bool = False,
) -> str:
    """Insert a NATIVE surface finish symbol (Ra value) at a sheet point.

    x/y are sheet metres; roughness is text like "3.2". Needs InsertSurfaceFinishSymbol3's
    full 14 arguments — a shorter call fails with "non-optional parameter".

    The API ignores x/y and drops the symbol at the sheet origin, so the tool moves it
    with SetPosition and reports the real `position_actual_m`. Set verify_render=True to
    diff a before/after PDF render at (x,y) and prove the symbol is visible on the sheet
    (`visibility.visible`); verify_dwg=True additionally checks a DWG export is non-empty
    (DWG is vector data, so visibility itself is proven through the PDF).
    """
    return _json(drawing.add_surface_finish(
        drawing_path, x, y, roughness=roughness, sym_type=sym_type,
        leader_type=leader_type, lay_symbol=lay_symbol, arrow_type=arrow_type,
        prod_method=prod_method, sample_len=sample_len,
        min_roughness=min_roughness, roughness_spacing=roughness_spacing, save=save,
        verify_render=verify_render, verify_dwg=verify_dwg,
    ))


@mcp.tool()
def verify_dimension_attachments(drawing_path: str, view_name: str | None = None) -> str:
    """Report what each dimension is attached to (edge / face / sketch segment / nothing).

    Dimensions with no attachment are untraceable and count as incomplete (process §6.1).
    """
    return _json(drawing.verify_dimension_attachments(drawing_path, view_name=view_name))


@mcp.tool()
def prune_orphan_dimensions(
    drawing_path: str, view_name: str | None = None,
    dry_run: bool = True, save: bool = True, confirm: bool = False,
) -> str:
    """Delete dimensions attached to no geometry, which `delete_dimensions` cannot target.

    `auto_dimension` imports every driving dimension and some land unattached. Clearing them
    used to be impossible: `delete_dimensions` works by VALUE, so it cannot tell three 23 mm
    dimensions apart. This deletes by name + position instead.

    Dry-run by default — returns view / name / value / position per candidate. With
    `dry_run=False` it deletes them and RE-RUNS the attachment check, so `orphan_count_after`
    is measured, not assumed.

    `confirm=True` is required for an actual delete: `orphan` only means
    `GetAttachedEntities2` returned nothing, and dimensions moved with SetPosition can read
    as orphan while still leading to real geometry (a report found the 50 mm overall length
    in that list).
    """
    return _json(drawing.prune_orphan_dimensions(drawing_path, view_name=view_name,
                                                 dry_run=dry_run, save=save,
                                                 confirm=confirm))


@mcp.tool()
def list_circular_edges(
    drawing_path: str,
    view_name: str | None = None,
    min_diameter_mm: float = 0.0,
) -> str:
    """List circular edges of a view: sheet position + diameter. Use before callouts."""
    return _json(drawing.list_circular_edges(
        drawing_path, view_name=view_name, min_diameter_mm=min_diameter_mm))


@mcp.tool()
def add_hole_callouts(
    drawing_path: str,
    view_name: str | None = None,
    min_diameter_mm: float = 0.0,
    max_callouts: int = 40,
    budget_s: float = 25.0,
    leader_x_m: float | None = None,
    leader_y_m: float | None = None,
    save: bool = True,
) -> str:
    """Native hole callouts on a view's visible circular edges, verified by IsHoleCallout.

    Only this view's visible edges are enumerated (a whole-model scan took >400 s; one view
    has ~9 circular edges). Callouts are stacked to the right of the view by default; pass
    leader_x_m / leader_y_m (sheet metres) to pin the text inside the frame (P1-7).
    """
    return _json(drawing.add_hole_callouts(
        drawing_path, view_name=view_name, min_diameter_mm=min_diameter_mm,
        max_callouts=max_callouts, budget_s=budget_s,
        leader_x_m=leader_x_m, leader_y_m=leader_y_m, save=save))


@mcp.tool()
def reconcile_ledger(
    model_path: str,
    drawing_path: str,
    tolerance_mm: float = 0.5,
) -> str:
    """Match the model's driving dimensions against the drawing's, and list what is missing.

    The §6.1 check that counting cannot do: returns coverage, unmatched ledger dimensions
    grouped by feature, and drawing-only values. Projection means the same edge dimensions
    differently per view, so treat the unmatched list as candidates to confirm.
    """
    return _json(drawing.reconcile_ledger(
        model_path, drawing_path, tolerance_mm=tolerance_mm))


@mcp.tool()
def survey_drawings(
    folder: str,
    sample_size: int = 100,
    seed: int = 28,
    html_path: str | None = None,
) -> str:
    """Audit a random sample of real drawings and report the statistical baseline.

    Returns distributions (views / dimensions / notes / center marks) and the share of
    drawings carrying surface finish, geometric tolerance and balloons — the numbers every
    verdict threshold should be based on. Pass html_path to also write a standalone page.

    Each drawing is closed right after it is read: SolidWorks refuses to open anything once
    ~40 documents are loaded, and one drawing takes 4-11 s, so sampling beats full scans.
    """
    return _json(survey.survey_drawings(
        folder, sample_size=sample_size, seed=seed, html_path=html_path,
    ))


@mcp.tool()
def delivery_check(
    model_path: str,
    drawing_path: str,
    ai_supplied: list[str] | None = None,
    require_gdt: bool = False,
    strict_attachments: bool = False,
) -> str:
    """Aggregate ledger + audit + annotations into a delivery verdict (process §7).

    Returns 可制造 / 待设计审核 / 仅供参考·未完成 plus the blockers and open items.
    ai_supplied lists values the caller supplied (material, tolerance, fit) — anything
    listed downgrades the verdict to 待设计审核.
    """
    return _json(drawing.delivery_check(
        model_path, drawing_path, ai_supplied=ai_supplied, require_gdt=require_gdt,
        strict_attachments=strict_attachments,
    ))


@mcp.tool()
def list_drawing_annotations(drawing_path: str, view_name: str | None = None) -> str:
    """Inventory every annotation: dimensions (with value + tolerance), notes,
    geometric tolerances (frame content), center marks, hole callouts.

    Use it to prove what a drawing does or does not carry — e.g. "no geometric
    tolerance anywhere" — which the CAD API cannot create but can read.
    """
    return _json(drawing.list_drawing_annotations(drawing_path, view_name=view_name))


@mcp.tool()
def add_dimension_tolerance(
    drawing_path: str,
    value_mm: float,
    tolerance_type: str = "bilateral",
    upper: float = 0.1,
    lower: float = 0.1,
    view_name: str | None = None,
    apply: bool = False,
) -> str:
    """Apply a NATIVE tolerance to an existing drawing dimension, located by its value in mm.

    tolerance_type: bilateral | symmetric | limit | basic | min | max | fit |
    fit_with_tolerance | none. upper/lower in millimetres. Dry-run by default.
    """
    return _json(drawing.add_dimension_tolerance(
        drawing_path, value_mm, tolerance_type=tolerance_type, upper=upper,
        lower=lower, view_name=view_name, apply=apply,
    ))


@mcp.tool()
def material_library(database: str = "SolidWorks Materials.sldmat") -> str:
    """List material names available in the installed SolidWorks material library."""
    return _json(material_tools.material_library(database=database))


@mcp.tool()
def set_part_material(
    paths: list[str],
    material_name: str,
    database: str = "SolidWorks Materials.sldmat",
    config: str = "",
    save: bool = True,
    apply: bool = False,
) -> str:
    """Assign a material to one or more parts; returns previous material and mass.

    Dry-run by default. Without a material SolidWorks uses 1000 kg/m³, so BOM masses are
    wrong until this is set.
    """
    return _json(material_tools.set_part_material(
        paths, material_name, database=database, config=config, save=save, apply=apply,
    ))


@mcp.tool()
def assign_materials(
    paths: list[str],
    rules: list | None = None,
    database: str = "SolidWorks Materials.sldmat",
    config: str = "",
    save: bool = True,
    apply: bool = False,
) -> str:
    """Pick a material per part from name keywords (标准件→碳钢, 结构件→铝合金), then assign.

    The chosen material is an AI-supplied value: it is reported per part with its reason
    and should be listed as pending confirmation in the delivery record.
    """
    return _json(material_tools.assign_materials(
        paths, rules=rules, database=database, config=config, save=save, apply=apply,
    ))


@mcp.tool()
def build_feature_ledger(
    model_path: str,
    include_faces: bool = True,
    config: str | None = None,
) -> str:
    """Build a feature ledger: every feature, its driving dimensions, hole candidates
    and model doubts (process §2). This is the basis for annotation planning and
    self-check — read-only, the model is never modified.

    include_faces=False skips cylindrical-face scanning (faster on large models).
    """
    return _json(ledger.build_feature_ledger(
        model_path, include_faces=include_faces, config=config,
    ))


@mcp.tool()
def create_drawing(
    model_path: str,
    output_path: str | None = None,
    template: str | None = None,
    views: list[str] | None = None,
    scale: float | None = None,
    auto_fit: bool = True,
    view_scales: dict | None = None,
    paper: str = "A3",
    view_positions: list | None = None,
) -> str:
    """Create a native .SLDDRW from a part/assembly, add model views, optionally save.

    views: front/back/left/right/top/bottom/isometric (default front+top+isometric).
    View names are resolved against the model, so Chinese (`*前视`) and English
    (`*Front`) SolidWorks both work. Units are meters; an A3 sheet is 0.42 x 0.297.
    view_positions: one sheet position per view, in metres, e.g.
    [[0.10, 0.20], [0.28, 0.20]] — same order as `views`. Omit to let auto_fit place them.
    """
    return _json(drawing.create_drawing(
        model_path, output_path=output_path, template=template,
        views=views or ["front", "top", "isometric"], scale=scale, auto_fit=auto_fit,
        view_scales=view_scales, paper=paper, view_positions=view_positions,
    ))


@mcp.tool()
def export_drawing(drawing_path: str, output_path: str, fmt: str | None = None) -> str:
    """Export a .SLDDRW to PDF / DWG / DXF.

    PDF can only come from a drawing — for STEP/STL of models use batch_export.
    """
    return _json(drawing.export_drawing(drawing_path, output_path, fmt=fmt))


@mcp.tool()
def render_drawing(pdf_path: str, output_dir: str | None = None, scale: float = 2.0) -> str:
    """Render each PDF page to PNG so the drawing can be visually checked.

    Requires pypdfium2 (pip install pypdfium2). Returns one PNG per page.
    """
    return _json(drawing.render_drawing(pdf_path, output_dir=output_dir, scale=scale))


@mcp.tool()
def add_section_view(
    drawing_path: str,
    view_name: str | None = None,
    orientation: str = "horizontal",
    offset_m: float = 0.0,
    position: list[float] | None = None,
    label: str = "A",
    flip: bool = False,
    depth: float = 0.0,
    cut_plane_normal: list[float] | None = None,
    cut_plane_origin: list[float] | None = None,
    verify_cut: bool = True,
) -> str:
    """Create a SECTION VIEW through a model view (exposes internal features).

    The cut line is drawn in the view's sketch automatically. orientation is
    "horizontal" or "vertical"; offset_m shifts it off the view centre; flip reverses
    the viewing direction; depth > 0 gives a partial section.

    cut_plane_normal / cut_plane_origin (MODEL metres, pass together) override
    orientation + offset_m with an explicit axis-aligned plane; unmappable or oblique
    planes are refused with a structured error instead of cutting the wrong plane.

    verify_cut (default True) checks the result: `edge_count == 0` -> `section_empty`,
    and a section that does not keep the source view's width (or height, for vertical
    cuts) within 10% -> `section_suspicious`. Both keep the view on the sheet.
    """
    return _json(drawing.add_section_view(
        drawing_path, view_name=view_name, orientation=orientation,
        offset_m=offset_m, position=position, label=label, flip=flip, depth=depth,
        cut_plane_normal=cut_plane_normal, cut_plane_origin=cut_plane_origin,
        verify_cut=verify_cut,
    ))


@mcp.tool()
def add_detail_view(
    drawing_path: str,
    view_name: str | None = None,
    center: list[float] | None = None,
    radius_m: float = 0.02,
    position: list[float] | None = None,
    scale: float = 2.0,
    label: str = "A",
    style: int = 0,
    shape: int = 0,
    flip: bool = False,
    show_label: bool = True,
    circle_segments: int = 32,
) -> str:
    """Add a DETAIL (enlarged) view around a circular area of a model view.

    center/radius_m are sheet metres (default: the view centre, 20 mm radius).
    scale is the REQUESTED detail scale (2.0 = 2:1); it is no longer overwritten by the
    parent view's scale. circle_segments must be > 0 — passing 0 is what makes
    SolidWorks reject the call. Returns error(detail_empty) if the new view has zero
    visible edges, and error(detail_scale_locked) if SolidWorks refuses the requested
    scale — neither is reported as a silent ok:true anymore.
    """
    return _json(drawing.add_detail_view(
        drawing_path, view_name=view_name, center=center, radius_m=radius_m,
        position=position, scale=scale, label=label, style=style, shape=shape,
        flip=flip, show_label=show_label, circle_segments=circle_segments,
    ))


@mcp.tool()
def add_flat_pattern_view(
    drawing_path: str,
    model_path: str,
    config_name: str | None = None,
    position: list[float] | None = None,
    scale: float | None = None,
    hide_bend_lines: bool = False,
    flip: bool = False,
) -> str:
    """Add a sheet-metal FLAT PATTERN view (unfolded blank with bend lines).

    The part must have a Flat-Pattern feature. Dimension sheet metal on the flat
    pattern, not the folded view.
    """
    return _json(drawing.add_flat_pattern_view(
        drawing_path, model_path, config_name=config_name, position=position,
        scale=scale, hide_bend_lines=hide_bend_lines, flip=flip,
    ))


@mcp.tool()
def auto_dimension(
    drawing_path: str,
    all_views: bool = True,
    include_unmarked: bool = False,
    eliminate_duplicates: bool = True,
    dimension_filter: list | None = None,
    place_outside: bool = False,
    outside_gap_m: float = 0.012,
    outside_step_m: float = 0.008,
) -> str:
    """Insert the model's driving dimensions into the drawing (Insert Model Items).

    Marked-for-drawing dimensions only by default; include_unmarked=True brings in
    every driving dimension. Returns inserted count and the post-insert audit.

    dimension_filter: `plan_drawing()["keep_dimensions"]`. Everything inserted that the
    plan did not keep is deleted afterwards, so "keep 8 of 41" really ends with ~8
    dimensions. Aborts without deleting if nothing matches the filter.

    place_outside=True moves everything this call inserted to just outside its view, stacked
    by side (gap + n × step). Use it instead of scaling the view up to hide overprinting:
    SolidWorks' own ~10 mm offset does not scale with the part.
    """
    return _json(drawing.auto_dimension(
        drawing_path, all_views=all_views, include_unmarked=include_unmarked,
        eliminate_duplicates=eliminate_duplicates, dimension_filter=dimension_filter,
        place_outside=place_outside, outside_gap_m=outside_gap_m,
        outside_step_m=outside_step_m,
    ))


@mcp.tool()
def list_drawing_edges(drawing_path: str, view_name: str | None = None) -> str:
    """List a view's visible edges with stable indices and model coordinates.

    Use the index with add_drawing_dimension(edge_index=...) when sheet coordinates
    are inconvenient (drawing edges report MODEL coordinates under this binding).
    """
    return _json(drawing.list_drawing_edges(drawing_path, view_name=view_name))


@mcp.tool()
def add_drawing_dimension(
    drawing_path: str,
    x: float,
    y: float,
    value: float | None = None,
    offset_x: float = 0.0,
    offset_y: float = -0.015,
    view_name: str | None = None,
    tolerance: float = 0.01,
    edge_index: int | None = None,
) -> str:
    """Add ONE dimension on a projected edge at sheet point (x, y) in meters.

    Manual dimensioning path: InsertModelItems can only carry dimensions the model
    already has, so IR-rebuilt parts need this. value is meters; omit to keep the
    computed value. Selection falls back EDGE -> SILHOUETTE -> SKETCHSEGMENT.
    """
    return _json(drawing.add_drawing_dimension(
        drawing_path, x, y, value=value, offset_x=offset_x, offset_y=offset_y,
        view_name=view_name, tolerance=tolerance, edge_index=edge_index,
    ))


@mcp.tool()
def add_center_marks(
    drawing_path: str,
    include_slots: bool = True,
    extended_lines: bool = True,
) -> str:
    """Auto-insert center marks and centerlines on every hole/slot in every view.

    Returns views_marked (reliable) rather than a mark count, which under-reports
    on this binding.
    """
    return _json(drawing.add_center_marks(
        drawing_path, include_slots=include_slots, extended_lines=extended_lines,
    ))


@mcp.tool()
def add_hole_callout(drawing_path: str, x: float, y: float,
                    view_name: str | None = None,
                    model_path: str | None = None,
                    save: bool = True) -> str:
    """Insert a native hole callout at a sheet point (meters) on a hole edge.

    Returns callout_failed when the point is not on a hole's projected circle —
    a missing callout is reported, never faked.
    """
    return _json(drawing.add_hole_callout(
        drawing_path, x, y, view_name=view_name, model_path=model_path, save=save,
    ))


@mcp.tool()
def set_title_block(
    drawing_path: str,
    name: str | None = None,
    code: str | None = None,
    material: str | None = None,
    apply: bool = False,
    include_reversed: bool = True,
) -> str:
    """Fill title-block placeholders (图样名称 / 图样代号 / 材质) in a drawing.

    `图样代号` also appears rotated 180 deg at the sheet's top-left corner (GB
    folding convention). include_reversed=False leaves that one untouched.

    GB templates store these as literal quoted text, so nothing is linked to the
    model. Dry-run by default; only the fields you pass are changed, and fields
    that cannot be found are reported instead of invented.
    """
    return _json(drawing.set_title_block(
        drawing_path, name=name, code=code, material=material, apply=apply,
        include_reversed=include_reversed,
    ))


@mcp.tool()
def drawing_audit(drawing_path: str) -> str:
    """Audit a drawing: views, scales, blank views, dimension and annotation counts.

    Evidence only — it never judges view choice, tolerances or manufacturability.
    Blank views and zero dimensions are reported as warnings, not silently passed.
    """
    return _json(drawing.drawing_audit(drawing_path))


@mcp.tool()
def get_assembly_tree(assembly_path: str, max_depth: int = 1, refresh: bool = False) -> str:
    """List assembly components with part names, paths, colors and instance info.

    max_depth=1 is flat and fast; higher values recurse into subassemblies.
    The component list is cached per document; set refresh=true after editing
    the assembly's structure outside this server.
    """
    return _json(assembly.assembly_tree(assembly_path, max_depth=max_depth, refresh=refresh))


@mcp.tool()
def get_mass_properties(assembly_path: str) -> str:
    """Read whole-assembly mass, volume, surface area and center of mass (SI units).

    Returns a structured error when CreateMassProperty is unavailable on the
    current binding (see docs/com-quirks.md section 5) instead of a guess.
    """
    return _json(assembly.mass_properties(assembly_path))


@mcp.tool()
def export_preview(
    assembly_path: str,
    output_path: str,
    size: str = "1600x1200",
    view: str = "*Isometric",
    keep_bmp: bool = False,
) -> str:
    """Export the current view to an image, e.g. 'D:/out/assembly.bmp' or '.png'.

    size is 'WxH'. SolidWorks writes BMP natively; PNG requires Pillow and the
    intermediate BMP is deleted unless keep_bmp=true.
    """
    return _json(assembly.export_preview(
        assembly_path, output_path, size=size, view=view, keep_bmp=keep_bmp,
    ))


def main():
    mcp.run()


if __name__ == "__main__":
    main()
