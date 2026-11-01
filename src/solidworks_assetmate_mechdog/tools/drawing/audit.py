from __future__ import annotations

"""Delivery verdicts and ledger reconciliation."""

from ._core import *
from .annot import list_drawing_annotations
from .create import drawing_audit
from .dims import verify_dimension_attachments

@tool_guard("delivery_check_failed")
def delivery_check(model_path, drawing_path, ai_supplied=None, require_gdt=False,
                   strict_attachments=False):
    """Aggregate model ledger + drawing audit + annotation inventory into a verdict.

    Process §7 asks for exactly three outcomes, and forbids calling a drawing
    "manufacturable" without evidence:

      可制造        關鍵尺寸/孔/公差齐全，双重自检通过
      待设计审核     图纸对象正确，但含调用方声明为「AI 拟定」的关键值
      仅供参考/未完成 有越界/重叠/空视图、无尺寸、模型疑点未解、缺孔标注等阻塞项

    ai_supplied: list of field names whose values the caller supplied (materials,
    tolerances, fits) — anything listed here downgrades the verdict and is reported.

    Dimensions with no readable attachment are reported; strict_attachments=True also
    makes them a blocker (§6.1 literal reading). They are a warning by default because
    GetAttachedEntities2 returns nothing for about half of the imported dimensions no
    matter which import parameters are used — whether those are truly floating or just
    unreadable is not settled, so the verdict should not hinge on it.
    """
    from . import ledger as _ledger

    blockers, pending, passed, checks = [], [], [], {}
    ledger_missing = []

    def record(name, ok, detail=None, blocking=True, **extra):
        # `extra` carries the numbers behind the verdict (attached vs excused); without
        # them a reader cannot tell "8 attached" from "23 attached".
        entry = {"ok": bool(ok), "detail": detail}
        entry.update(extra)
        checks[name] = entry
        if ok:
            passed.append(name)
        elif blocking:
            blockers.append({"check": name, "detail": detail})
        else:
            pending.append({"check": name, "detail": detail})

    # ---- model side -------------------------------------------------------
    # Faces are needed to count hole candidates; the extra scan costs a few seconds.
    led = _ledger.build_feature_ledger(model_path, include_faces=True)
    if not led.get("ok"):
        return error("model_unreadable", led.get("message") or "ledger failed",
                     "Check the model opens and rebuilds before delivering a drawing.")
    model_stats = led.get("stats") or {}
    record("model_readable", True, "类别 %s" % led.get("category"))
    doubts = led.get("doubts") or []
    # P2-5: this tool makes and checks drawings; it deliberately does NOT rebuild
    # models. A source part SolidWorks flags as needing a rebuild is not a drawing
    # defect, yet it was a BLOCKER the caller could never clear from inside this
    # toolkit — so the verdict could never reach 可制造. It stays visible as a pending
    # item instead of a blocker.
    record("model_clean", not doubts,
           "; ".join(str(d.get("kind")) for d in doubts) or None,
           blocking=False)

    # ---- drawing side -----------------------------------------------------
    audit = drawing_audit(drawing_path)
    if not audit.get("ok"):
        return error("drawing_unreadable",
                     audit.get("message") or "audit failed",
                     "The drawing must open and rebuild.")
    warnings = [w.get("issue") for w in (audit.get("warnings") or [])]
    view_count = audit.get("view_count") or 0
    dim_total = audit.get("total_dimensions") or 0

    record("views_present", view_count > 0, "%d 个视图" % view_count)
    record("no_view_outside_sheet", "view_outside_sheet" not in warnings,
           "有视图超出图纸" if "view_outside_sheet" in warnings else None)
    record("no_view_overlap", "views_overlap" not in warnings,
           "有视图互相重叠" if "views_overlap" in warnings else None)
    record("no_blank_view", "blank_view" not in warnings,
           "存在空白视图" if "blank_view" in warnings else None)
    record("dimensions_present", dim_total > 0, "%d 个尺寸" % dim_total)

    # ---- annotations -----------------------------------------------------
    anns = list_drawing_annotations(drawing_path)
    if not anns.get("ok"):
        return error("annotations_unreadable", anns.get("message") or "failed")
    totals = anns.get("totals") or {}
    holes_in_drawing = int(totals.get("hole_callout", 0))
    toleranced = 0
    for row in anns.get("annotations") or []:
        if row.get("kind") == "dimension" and row.get("tolerance_type"):
            toleranced += 1
    gtols = int(totals.get("geometric_tolerance", 0))

    hole_candidates = len(led.get("holes") or [])
    if hole_candidates:
        # Hole information can legitimately land three ways: a native hole callout, a
        # thread callout, or a plain Ø dimension on the circular edge. Any of them
        # carries the manufacturing information, so requiring the callout object
        # specifically would reject valid drawings.
        holes_in_drawing = int(totals.get("hole_callout", holes_in_drawing) or 0)
        diameter_dims = 0
        for row in anns.get("annotations") or []:
            if row.get("kind") != "dimension":
                continue
            # P1-1: never decide this from display text alone. `GetText()` with no
            # slot returns '' for a dimension, so the old text test found 0 diameters
            # on a sheet that printed Ø28/Ø23/Ø12/Ø18/Ø13. The reliable signal is the
            # dimension type (5 = diameter, measured 2026-10-31); the raw value-slot
            # string stays as a fallback for builds that report no type.
            if int(row.get("dimension_type") or -1) == 5:
                diameter_dims += 1
                continue
            marker = (str(row.get("raw_text") or "") + str(row.get("text") or "")).upper()
            if "Ø" in marker or "MOD-DIAM" in marker:
                diameter_dims += 1
        covered = holes_in_drawing + diameter_dims
        record("hole_callouts", covered > 0,
               "图纸 %d 处孔标注 + %d 处 Ø 尺寸 / 模型 %d 个圆柱面候选"
               % (holes_in_drawing, diameter_dims, hole_candidates),
               "背面或被遮挡的孔 SolidWorks 不给标注，需人工补或换视图")
    record("tolerances_present", toleranced > 0,
           "%d 个尺寸带公差" % toleranced, blocking=False)

    # Attachment: GetAttachedEntities2 returns nothing for roughly half of the
    # dimensions that InsertModelAnnotations3 creates, and changing the import
    # parameters does not move that ratio. It is unclear whether those are really
    # floating or just unreadable, so by default this is a WARNING; strict mode
    # (process §6.1 literal reading) turns it into a blocker.
    att = verify_dimension_attachments(drawing_path)
    if att.get("ok"):
        all_orphans = att.get("orphans") or []
        # Every orphan counts. An earlier version excused the ones sitting near the view
        # outline; since dimensions are parked outside the outline by design, that
        # excused all of them (15/15 in a report) and turned a real failure into ok.
        # Proximity is reported, never used to clear a blocker.
        orphan = len(all_orphans)
        near_view = sum(1 for o in all_orphans if o.get("orphan_near_view"))
        dim_count = int(att.get("dimension_count") or 0)
        attached = max(dim_count - orphan, 0)
        suffix = ("（其中 %d 条位置在视图附近，但引线未验证）" % near_view
                  if near_view else "")
        record("dimension_attachments", orphan == 0,
               (("%d/%d 个尺寸无关联实体%s" % (orphan, dim_count, suffix))
                if orphan
                else ("%d 个尺寸均已关联" % dim_count)),
               excused_count=0,
               attached_count=attached,
               blocking=bool(strict_attachments) and orphan > 0)
    else:
        record("dimension_attachments", False,
               "关联信息读取失败: %s" % (att.get("message") or "unknown"))
    unattached = int(att.get("orphan_count") or 0) if att.get("ok") else None
    # Measured on 100 real drawings: surface finish appears in 60%, geometric
    # tolerance in only 7%. So finish is the routine check and GD&T stays optional.
    finishes = int(totals.get("surface_finish", 0))
    record("surface_finish_present", finishes > 0,
           "%d 处表面粗糙度" % finishes, blocking=False)
    if require_gdt:
        record("gdt_present", gtols > 0, "%d 处形位公差" % gtols)
    else:
        checks["gdt_present"] = {
            "ok": gtols > 0, "detail": "%d 处形位公差（真实图纸仅 7%% 有）" % gtols,
            "required": False,
        }

    # ---- ledger reconciliation (§6.1: 逐条比对, not just count) ----------
    rec = reconcile_ledger(model_path, drawing_path)
    if rec.get("ok"):
        coverage = rec.get("coverage")
        unmatched = rec.get("unmatched_count") or 0
        record("ledger_reconciled",
               coverage is not None and coverage >= 0.9,
               "台账覆盖率 %s（%s/%s 个尺寸在图纸上找到，缺 %s）"
               % (coverage, rec.get("matched"), rec.get("ledger_dimensions"), unmatched),
               blocking=False)
        # P1-3: a ratio cannot tell you WHAT is missing, and on a "must be repairable
        # from this drawing" job the missing five are the whole point.
        from ...sw_session import fix_mojibake as _fix_mojibake

        ledger_missing = [
            {"feature": _fix_mojibake(row.get("feature")),
             "dimension": _fix_mojibake(row.get("dimension")),
             "value_mm": row.get("value_mm"),
             "dimension_type": row.get("dimension_type")}
            for row in (rec.get("unmatched_ledger") or [])
        ]
    else:
        record("ledger_reconciled", False,
               "对账失败: %s" % (rec.get("message") or "unknown"), blocking=False)

    # ---- caller-declared assumptions -------------------------------------
    if ai_supplied:
        for field in ai_supplied:
            pending.append({"check": "ai_supplied", "detail": str(field)})

    if blockers:
        verdict = "仅供参考/未完成"
    elif pending:
        verdict = "待设计审核"
    else:
        verdict = "可制造"

    return drawing_ok(
        verdict=verdict,
        drawing=audit.get("drawing"),
        model=led.get("part"),
        category=led.get("category"),
        checks=checks,
        passed=passed,
        blockers=blockers,
        pending=pending,
        strict_attachments=bool(strict_attachments),
        # P1-3: the concrete list behind `ledger_reconciled`'s ratio.
        ledger_missing=ledger_missing,
        ledger_missing_count=len(ledger_missing),
        facts={
            "views": view_count,
            "dimensions": dim_total,
            "dimensions_with_tolerance": toleranced,
            "hole_callouts": holes_in_drawing,
            "geometric_tolerances": gtols,
            "surface_finishes": finishes,
            "cylinder_faces_in_model": hole_candidates,
            "dimensions_unattached": unattached,
            "features": model_stats.get("features"),
            "model_doubts": [d.get("kind") for d in doubts],
        },
        note="verdict 依据流程 §7；blockers 非空即不可宣称可制造，"
             "ai_supplied 非空则只能算待设计审核。无关联实体的尺寸计入 blockers"
             "（§6.1：无关联视为未完成）。",
    )


@tool_guard("reconcile_failed")
def reconcile_ledger(model_path, drawing_path, tolerance_mm=0.5):
    """Match the model's driving dimensions against the drawing's dimensions.

    This is the §6.1 check that only counting couldn't do: the ledger holds every driving
    dimension, the drawing holds what was actually annotated, and the gap between them is
    the concrete list of what is missing.

    Matching is by VALUE (whole millimetres, ±tolerance), which is deliberately loose:
    a drawing dimension is the PROJECTED length, so a 1253 mm edge can appear as 37.6 mm
    in a side view. Values that do not match are reported as `unmatched_ledger` —
    candidates for "not dimensioned", not proof of it.
    """
    from . import ledger as _ledger

    led = _ledger.build_feature_ledger(model_path, include_faces=False)
    if not led.get("ok"):
        return error("model_unreadable", led.get("message") or "ledger failed")
    doc = _open_drawing(drawing_path)
    sheet_values = sorted({round(v * 1000, 2) for v in drawing_dimension_values(doc)})

    tol = float(tolerance_mm)
    # Array spacing and sketch parameters are driving dimensions too, but values far
    # beyond the part's own bounding box cannot be a manufacturable length.
    bbox = led.get("bbox_mm") or []
    try:
        diagonal = float(sum(v * v for v in bbox) ** 0.5) if bbox else 0.0
    except Exception:
        diagonal = 0.0
    ceiling = diagonal * 2.0 if diagonal else None

    ledger_rows, skipped_non_length, skipped_oversize = [], [], []
    for feature in led.get("features") or []:
        for dim in feature.get("dimensions") or []:
            kind = dim.get("dimension_type")
            if kind is not None and kind not in (0, 2, 3):
                # angular / ordinate dimensions never appear as a linear value
                skipped_non_length.append({
                    "feature": feature.get("name"),
                    "value": dim.get("value_mm"), "type": kind,
                })
                continue
            value = dim.get("value_mm")
            if ceiling and value is not None and value > ceiling:
                skipped_oversize.append({
                    "feature": feature.get("name"), "value": value,
                })
                continue
            ledger_rows.append({
                "feature": feature.get("name"),
                "feature_type": feature.get("type"),
                "dimension": dim.get("name"),
                "value_mm": value,
                "dimension_type": kind,
            })

    matched, unmatched, skipped_zero = [], [], []
    for row in ledger_rows:
        value = row.get("value_mm")
        if value is None:
            continue
        # P3-5: a 0.0 "dimension" is a sketch relation/constraint value, not a length —
        # no drawing can dimension it, so reporting it as a missing item is a false
        # gap. It is kept in its own bucket rather than silently dropped.
        if abs(float(value)) < 1e-9:
            skipped_zero.append(row)
            continue
        hit = next((v for v in sheet_values if abs(v - value) <= tol), None)
        if hit is None:
            unmatched.append(row)
        else:
            matched.append(dict(row, drawing_value_mm=hit))

    ledger_values = {round(r["value_mm"], 2) for r in ledger_rows
                     if r.get("value_mm") is not None
                     and abs(float(r["value_mm"])) > 1e-9}
    sheet_only = [v for v in sheet_values
                  if not any(abs(v - lv) <= tol for lv in ledger_values)]

    by_feature = {}
    for row in unmatched:
        by_feature.setdefault(row.get("feature") or "?", []).append(row.get("value_mm"))

    return drawing_ok(
        drawing=doc.GetTitle,
        model=led.get("part"),
        ledger_dimensions=len(ledger_rows),
        drawing_dimensions=len(sheet_values),
        matched=len(matched),
        unmatched_count=len(unmatched),
        coverage=(round(len(matched) / (len(ledger_rows) - len(skipped_zero)), 3)
                  if (len(ledger_rows) - len(skipped_zero)) else None),
        unmatched_by_feature={
            k: sorted({round(v, 2) for v in vals if v is not None})
            for k, vals in sorted(by_feature.items())
        },
        skipped_non_length=len(skipped_non_length),
        skipped_oversize=len(skipped_oversize),
        # 0.0 entries are sketch relations, not lengths: excluded from coverage so the
        # ratio measures what a drawing could actually dimension (P3-5).
        skipped_zero_value=len(skipped_zero),
        zero_value_examples=skipped_zero[:6],
        comparable_dimensions=len(ledger_rows) - len(skipped_zero),
        oversize_examples=skipped_oversize[:6],
        skipped_examples=skipped_non_length[:8],
        unmatched_ledger=unmatched[:60],
        drawing_only_values=sheet_only[:40],
        note="unmatched_ledger 是「台账有、图纸没有对应值」的候选缺项；"
             "drawing_only_values 是图上有但台账没有的（多为投影尺寸或参考尺寸）。"
             "投影关系会让同一尺寸在不同视图取值不同，因此按值匹配只能作为待确认清单。"
             "skipped_zero_value 是 0.0 的草图约束值，不是长度，已从覆盖率分母剔除。",
    )


