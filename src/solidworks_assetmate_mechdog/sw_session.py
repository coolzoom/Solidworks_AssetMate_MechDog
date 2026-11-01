"""Shared SolidWorks COM session layer.

Rules baked in from real debugging (see docs/com-quirks.md):
- OpenDoc6 Errors/Warnings must be byref VARIANTs.
- Reuse the already-active document; never reopen blindly (memory vs disk).
- GetComponents is slow: callers fetch once, then group in memory.
- A cached dispatcher can die silently: probe it, and drop it on the first failure.

pywin32 is imported lazily so the package stays importable off-Windows.
"""
from __future__ import annotations

import os
import time

SW_DOC_PART = 1
SW_DOC_ASSEMBLY = 2
SW_DOC_DRAWING = 3
SW_OPEN_SILENT = 1

_SW = None
_LAST_PROBE = 0.0
# Measured 18 ms for a RevisionNumber round trip: cheap for a human, too expensive to pay
# on every get_sw() call (open_document calls it several times per tool).
LIVENESS_PROBE_INTERVAL_S = 2.0
# P0-1: enough state for ping() to distinguish "alive but busy" from "link is dead".
_BRIDGE_STATE = {"stale_drops": 0, "reconnects": 0, "last_error": None,
                 "last_reconnect": None}
# Must exist from import time: log_failure() and ping() read it, and the FIRST failure is
# exactly when no tool has called note_last_call() yet. Without this default the failure
# logger itself raised NameError and wrote to _logger_error.log instead of mcp-*.log.
_LAST_CALL = {"tool": None, "args": None}
# Measured: GetComponents costs ~130 s on a 46-component assembly (2026-10-26).
# Cache per document path so tree -> dry-run -> apply -> tree stays interactive.
# Pass refresh=True after the user adds/removes components outside this server.
_COMPONENTS_CACHE = {}


def _com():
    try:
        import pythoncom
        import win32com.client as win
        from win32com.client import VARIANT
    except ImportError as exc:  # non-Windows or pywin32 missing
        raise RuntimeError(
            "pywin32 is required and SolidWorks must run on this Windows machine. "
            "Install with: pip install pywin32"
        ) from exc
    return pythoncom, win, VARIANT


def sw_version(sw=None):
    """Return the SolidWorks revision string, or None.

    Capability conclusions are version-dependent: `InsertBalloon` is UNKNOWNNAME on SW2026
    and absent on SW2024; `InsertBomTable4(Indented)` works on SW2024 SP4 and returns None
    on SW2026. Without the revision in the payload, a failure is indistinguishable from a
    wrong-configuration failure.
    """
    try:
        app = sw or get_sw()
        return str(app.RevisionNumber)
    except Exception:
        return None


def get_sw(probe=True):
    """Connect to the running SolidWorks instance, or start one. Cached.

    Always wrapped in a DYNAMIC (late-bound) dispatcher on purpose. If
    %TEMP%\\gen_py holds a makepy cache for the SldWorks type library, win32com
    hands back an early-bound wrapper whose OpenDoc6 declares Errors/Warnings as
    plain byref ints and calls int() on our byref VARIANT — every tool dies with
    "int() argument must be ... not 'VARIANT'". Late binding accepts the VARIANT
    and does not depend on whether some other program generated gen_py.

    P0-1 (0.8.22): the cached dispatcher used to be returned forever. When the COM link
    died, every later call kept failing against the same dead pointer for minutes — the
    reported "all calls fail while SolidWorks itself is fine" outage. The cache is now
    probed with `RevisionNumber` at most once per LIVENESS_PROBE_INTERVAL_S and dropped on
    the first failure, so the next call re-acquires through GetActiveObject. `probe=False`
    skips the check when the caller already knows the link works (e.g. reconnect()).
    """
    global _SW, _LAST_PROBE
    if _SW is not None:
        if not probe or (time.time() - _LAST_PROBE) < LIVENESS_PROBE_INTERVAL_S:
            return _SW
        _LAST_PROBE = time.time()
        try:
            _SW.RevisionNumber
            return _SW
        except Exception as exc:
            _BRIDGE_STATE["stale_drops"] = _BRIDGE_STATE.get("stale_drops", 0) + 1
            _BRIDGE_STATE["last_error"] = "stale dispatcher: %s" % str(exc)[:180]
            log_failure("get_sw", {"stage": "liveness_probe"}, exc=exc)
            _SW = None
    _, win, _ = _com()
    from win32com.client import dynamic as dyn

    try:
        obj = win.GetActiveObject("SldWorks.Application")
    except Exception:
        obj = win.Dispatch("SldWorks.Application")
    try:
        _SW = dyn.Dispatch(obj)
    except Exception:
        # Last resort: an early-bound wrapper here would reintroduce the byref
        # VARIANT crash, so make the fallback visible in logs.
        import logging

        logging.getLogger(__name__).warning(
            "win32com.client.dynamic.Dispatch failed; using the raw COM object. "
            "If every tool fails with 'int() argument ... not VARIANT', this is why."
        )
        _SW = obj
    return _SW


def reconnect(reason=None):
    """Drop the cached dispatcher and connect again.

    The cached `_SW` never noticed when the COM connection died: after a `Not connected`
    every subsequent call kept failing against a dead pointer for minutes, which is exactly
    what a report described ("all calls fail, SolidWorks itself is fine"). Clearing the
    cache lets the next `get_sw()` re-acquire through GetActiveObject.

    P0-1 (0.8.22): `reason` is recorded so ping() can explain why the link was rebuilt;
    `probe=False` avoids paying for a liveness check on a dispatcher that was just created.
    """
    global _SW, _LAST_PROBE
    _SW = None
    _LAST_PROBE = 0.0
    _BRIDGE_STATE["reconnects"] = _BRIDGE_STATE.get("reconnects", 0) + 1
    try:
        from datetime import datetime

        _BRIDGE_STATE["last_reconnect"] = datetime.now().isoformat(timespec="seconds")
    except Exception:
        pass
    if reason:
        _BRIDGE_STATE["last_error"] = str(reason)[:200]
    return get_sw(probe=False)


def health(deep=False):
    """Report whether the bridge is alive, and whether SolidWorks is alive.

    Distinguishes three states that used to look identical from the outside:
      * bridge alive + SW responding  -> ok
      * bridge alive + SW not responding -> the COM call is the problem
      * bridge dead (cached pointer stale) -> reconnect() may fix it

    `deep=True` also touches ActiveDoc, which is where a stale pointer usually shows up.
    """
    # Provenance first: with several editable installs on one machine, "which package is
    # this MCP actually running" is the only question that matters (asked for three rounds).
    import sys as _sys

    try:
        import solidworks_assetmate_mechdog as _pkg

        _package_path = getattr(_pkg, "__file__", None)
        _package_version = getattr(_pkg, "__version__", None)
    except Exception:
        _package_path, _package_version = None, None
    state = {"sw_version": None, "active_doc": None, "reconnected": False, "error": None,
             "package_path": _package_path,
             "package_version": _package_version,
             "python_executable": _sys.executable}
    try:
        app = get_sw()
        state["sw_version"] = str(app.RevisionNumber)
    except Exception as exc:
        state["error"] = "connect failed: %s" % str(exc)[:100]
        try:
            app = reconnect()
            state["reconnected"] = True
            state["sw_version"] = str(app.RevisionNumber)
            state["error"] = None
        except Exception as exc2:
            state["error"] = "%s; reconnect failed: %s" % (state["error"], str(exc2)[:80])
            return state
    if deep:
        try:
            doc = app.ActiveDoc
            state["active_doc"] = str(doc.GetTitle) if doc is not None else None
        except Exception as exc:
            state["error"] = "ActiveDoc unreadable: %s" % str(exc)[:100]
            try:
                app = reconnect()
                state["reconnected"] = True
                doc = app.ActiveDoc
                state["active_doc"] = str(doc.GetTitle) if doc is not None else None
                state["error"] = None
            except Exception as exc2:
                state["error"] = "ActiveDoc still unreadable after reconnect: %s" % str(exc2)[:80]
    return state


LOG_DIRNAME = "logs"


def log_dir():
    import os
    path = os.path.join(os.getcwd(), LOG_DIRNAME)
    os.makedirs(path, exist_ok=True)
    return path


def note_last_call(tool=None, args=None):
    """Remember what was running, so a crash report can name it."""
    global _LAST_CALL
    _LAST_CALL = {"tool": tool, "args": args}


def log_failure(tool, payload, exc=None, extra=None, args=None):
    """Append one JSON line per failure to logs/mcp-YYYYMMDD.log.

    Without this, a mid-run disconnect left no trace at all — the reporter could not tell
    whether the MCP process had died or SolidWorks had. The line records the tool, its
    arguments, the error, the SW version and the last call before the failure.
    """
    import json
    import os
    import traceback
    from datetime import datetime

    try:
        today = datetime.now().strftime("%Y%m%d")
        path = os.path.join(log_dir(), "mcp-%s.log" % today)
        record = {
            "ts": datetime.now().isoformat(timespec="seconds"),
            "pid": os.getpid(),
            "tool": tool,
            "args": args if args is not None else _LAST_CALL.get("args"),
            "payload": payload if isinstance(payload, dict) else str(payload)[:400],
            "last_call_before": _LAST_CALL,
            "sw_version": sw_version(),
            "health": health(),
        }
        if exc is not None:
            record["exception"] = "%s: %s" % (type(exc).__name__, str(exc)[:300])
            record["traceback"] = traceback.format_exc()[-1200:]
        if extra:
            record["extra"] = extra
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False, default=str) + chr(10))
        return path
    except Exception as exc2:
        # Logging must never be the reason a tool fails — but staying silent about the
        # logger itself is how the first version swallowed a NameError and returned None.
        try:
            with open(os.path.join(log_dir(), "_logger_error.log"), "a", encoding="utf-8") as f:
                f.write("%s: %s" % (type(exc2).__name__, exc2) + chr(10))
        except Exception:
            pass
        return None


def variant(value, vartype):
    pythoncom, _, VARIANT = _com()
    return VARIANT(vartype, value)


# ---------------------------------------------------------------------------
# Dedicated COM thread
#
# FastMCP runs every sync tool on an arbitrary anyio worker thread, so without
# marshalling each call lands on a different apartment. A dispatcher created on
# thread A and invoked from thread B bypasses marshalling and intermittently
# raises RPC_E_WRONG_THREAD / "Not connected" — the same outward symptom as the
# bridge-lost outage P0-1 fixed at the pointer level. Marshalling every tool
# body onto ONE STA thread (the pattern haunchen/solidworks-mcp and
# yezg9920-kely/solidworks-mcp both use) removes the apartment variable
# entirely: every COM object is created and consumed on the same thread.
#
# run_on_com_thread() runs the callable on that worker. Off-Windows or without
# pywin32 it falls back to inline execution so imports and static tests still
# work; ASSETMATE_COM_INLINE=1 forces the fallback (debug escape hatch).
# ---------------------------------------------------------------------------

import queue as _queue
import threading as _threading

_COM_THREAD = None
_COM_TASKS = None


def _com_worker(tasks):
    """Single-threaded COM apartment: initialise STA, then serve the queue."""
    try:
        pythoncom, _, _ = _com()
        pythoncom.CoInitialize()
    except Exception:
        pass
    while True:
        item = tasks.get()
        if item is None:
            return
        fn, answer = item
        try:
            answer.put((True, fn()))
        except BaseException as exc:
            answer.put((False, exc))


def _com_thread():
    global _COM_THREAD, _COM_TASKS
    if _COM_TASKS is None:
        _COM_TASKS = _queue.Queue()
        _COM_THREAD = _threading.Thread(
            target=_com_worker, args=(_COM_TASKS,), daemon=True, name="sw-com")
        _COM_THREAD.start()
    return _COM_THREAD


def run_on_com_thread(fn):
    """Run fn() on the single dedicated COM thread and return its result.

    Re-entrant: a tool body already executing on the COM thread (e.g. run_batch
    dispatching inner tools) runs inline instead of deadlocking on its own queue.
    """
    if os.environ.get("ASSETMATE_COM_INLINE"):
        return fn()
    if _COM_THREAD is not None and _threading.current_thread() is _COM_THREAD:
        return fn()
    thread = _com_thread()
    answer = _queue.Queue()
    _COM_TASKS.put((fn, answer))
    succeeded, value = answer.get()
    if succeeded:
        return value
    raise value


# ---------------------------------------------------------------------------
# Path sandbox (borrowed from jaylamping/solidworks-mcp's ALLOWED_ROOTS idea)
#
# ASSETMATE_ALLOWED_ROOTS is a ';' separated list of directories. When set,
# open_document() refuses paths outside them — a path typo or a crafted request
# cannot make the server read a file the operator did not intend to expose.
# Unset (default) preserves the old unrestricted behaviour; the sandbox is an
# opt-in hardening switch, echoed in ping() so the operator can prove it is on.
# ---------------------------------------------------------------------------

def allowed_roots():
    raw = os.environ.get("ASSETMATE_ALLOWED_ROOTS", "").strip()
    if not raw:
        return []
    roots = []
    for chunk in raw.split(";"):
        chunk = chunk.strip()
        if chunk:
            roots.append(os.path.normcase(os.path.abspath(chunk)))
    return roots


def check_path_allowed(path):
    """Raise PermissionError when path escapes ASSETMATE_ALLOWED_ROOTS."""
    roots = allowed_roots()
    if not roots:
        return os.path.abspath(path)
    resolved = os.path.normcase(os.path.abspath(path))
    for root in roots:
        if resolved == root or resolved.startswith(root + os.sep):
            return os.path.abspath(path)
    raise PermissionError(
        "path outside ASSETMATE_ALLOWED_ROOTS: %s (allowed: %s)"
        % (path, "; ".join(roots)))


def double_array(values):
    """VT_ARRAY | VT_R8 — a plain Python list silently fails for many SW APIs."""
    pythoncom, _, _ = _com()
    return variant(list(values), pythoncom.VT_ARRAY | pythoncom.VT_R8)


# P0-1: COM failures need two different responses, and conflating them is why the reported
# outage lasted minutes.
#   transient   — SolidWorks is busy and refused the call; waiting then retrying works.
#   bridge lost — the dispatcher itself is dead; retrying the same pointer can never work,
#                 so the cache must be dropped and the caller told to retry the tool.
TRANSIENT_HRESULTS = (
    -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED        "Call was rejected by callee"
    -2147417846,  # 0x8001010A RPC_E_SERVERCALL_RETRYLATER
    -2147417847,  # 0x80010109 RPC_E_SERVERCALL_RETRYLATER variant
    -2147417849,  # 0x80010107 RPC_E_SERVERCALL_RETRYLATER variant
)
BRIDGE_LOST_HRESULTS = (
    -2147220995,  # 0x800401FD CO_E_OBJNOTCONNECTED        "Not connected"
    -2147417848,  # 0x80010108 RPC_E_DISCONNECTED
    -2147023170,  # 0x800706BE RPC_S_CALL_FAILED
    -2147023174,  # 0x800706BA RPC_S_SERVER_UNAVAILABLE
)
TRANSIENT_HINTS = ("call was rejected", "server busy", "retry later", "被调用方拒绝", "忙")
BRIDGE_LOST_HINTS = ("not connected", "object is not connected",
                     "rpc server unavailable", "remote procedure call failed",
                     "rpc 服务器", "服务器运行失败")

COM_RETRIES = 2
COM_RETRY_BACKOFF_S = 0.25


class BridgeLost(RuntimeError):
    """The cached COM dispatcher is dead. The cache has already been dropped.

    Raised after a `Not connected`-class failure so the caller can say "retry this call"
    instead of surfacing a raw COM hresult. Retrying the same tool will re-acquire the link
    through GetActiveObject.
    """


def _hresult(exc):
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0]
    return None


def _error_text(exc):
    return " ".join(v for v in getattr(exc, "args", ()) if isinstance(v, str)).lower()


def is_transient_com_error(exc):
    """SolidWorks refused the call because it is busy — safe to retry."""
    return _hresult(exc) in TRANSIENT_HRESULTS or any(
        hint.lower() in _error_text(exc) for hint in TRANSIENT_HINTS)


def is_bridge_lost_error(exc):
    """The COM link itself is gone — retrying the same pointer cannot help."""
    return _hresult(exc) in BRIDGE_LOST_HRESULTS or any(
        hint.lower() in _error_text(exc) for hint in BRIDGE_LOST_HINTS)


def call_method(obj, name, *args):
    """Invoke a SolidWorks method through DISPATCH_METHOD, bypassing pywin32's guesswork.

    win32com decides method-vs-property from the type library, and for a number of
    SolidWorks members it decides WRONG: it dispatches them as PROPERTYGET and the call
    raises `DISP_E_MEMBERNOTFOUND` (0x80020003), which reads exactly like "this API does
    not exist". It does exist. Calling the same DISPID with DISPATCH_METHOD works.

    Confirmed cases (each cost hours before being found):
      * `ModelDoc2.InsertGtol`      — geometric tolerance "impossible" for two days
      * `INote.GetAnnotation`       — notes could not be positioned
      * `IFeature.FirstFeature` / `GetNextFeature` / `GetDefinition` — feature walk

    So: for any of those (and anything that reports MEMBERNOTFOUND on a method that the
    documentation says exists) use this instead of attribute access. `static_check.py`
    fails the build when a known trap is called directly.

    P0-1 (0.8.22): bounded retry, and the two failure classes are handled differently.
    A busy SolidWorks answers "Call was rejected by callee"; sleeping 0.25 s then 0.5 s and
    asking again succeeds (observed while a rebuild was running). A dead link is NOT
    retried — the dispatcher is worthless — so the cache is dropped and `BridgeLost` is
    raised with a message that says the next call will reconnect.

    Note on timeouts: this is not a thread-level watchdog. pywin32 dispatches on the STA
    thread that owns the object, so running the call on a worker thread and abandoning it
    after N seconds risks corrupting the COM apartment rather than cancelling anything. The
    bounded retry plus the liveness probe in get_sw() cover the reported failure mode (a
    bus/link failure that never recovers); a single genuinely hung call still blocks, which
    is a SolidWorks-side condition no client can safely interrupt.
    """
    import pythoncom

    ole = obj._oleobj_
    dispid = ole.GetIDsOfNames(name)
    flag = getattr(obj, "_FlagAsMethod", None)
    if flag is not None:
        try:
            flag(name)
        except Exception:
            pass
    last = None
    for attempt in range(COM_RETRIES + 1):
        try:
            if args:
                return ole.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, True, *args)
            return ole.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, True)
        except Exception as exc:
            last = exc
            if is_bridge_lost_error(exc):
                _BRIDGE_STATE["last_error"] = "%s.%s: %s" % (
                    type(obj).__name__, name, str(exc)[:160])
                try:
                    reconnect(reason="%s.%s lost the link" % (type(obj).__name__, name))
                except Exception:
                    pass
                raise BridgeLost(
                    "SolidWorks COM link was lost during %s(); the cached dispatcher has "
                    "been dropped, so calling this tool again will reconnect. Original: %s"
                    % (name, str(exc)[:180])) from exc
            if not is_transient_com_error(exc) or attempt >= COM_RETRIES:
                raise
            time.sleep(COM_RETRY_BACKOFF_S * (attempt + 1))
    raise last if last is not None else RuntimeError("call_method exhausted retries")


def fix_mojibake(text):
    """Recover GB2312 text that SolidWorks handed back decoded as Latin-1.

    Feature names arrive as `ÂÝÎÆ¿××êÍ·Éî¶È` (= 螺纹孔钻头深度) — the GB2312 bytes were
    decoded one byte at a time. Because those bytes are all Latin-1 printable, a lossless
    re-encode is possible. Real Chinese fails at `.encode("latin-1")` and is returned
    untouched, so this is safe to call on every string (reports asked for it three rounds
    running, and mojibake in `keep_dimensions` also weakens `dimension_filter` matching).
    """
    if not isinstance(text, str) or not text:
        return text
    try:
        repaired = text.encode("latin-1").decode("gb2312")
    except Exception:
        return text
    # Only accept it when the repair actually produced CJK.
    return repaired if any("\u4e00" <= ch <= "\u9fff" for ch in repaired) else text


def runtime_info():
    """Report WHICH copy of this package is actually running, not just its version.

    A report burned a whole round chasing fixes that were already in the repository: the MCP
    host pointed at an editable install of an older copy (0.8.14) while the working tree held
    a newer one, and three source directories existed side by side. A version number alone
    cannot settle that — the path can. Echoed by `check_capabilities` and `ping`.
    """
    import sys

    info = {"server_version": None, "server_path": None,
            "python_executable": None, "python_version": None, "pid": None}
    try:
        import solidworks_assetmate_mechdog as package

        info["server_version"] = getattr(package, "__version__", None)
        info["server_path"] = os.path.dirname(os.path.abspath(package.__file__))
    except Exception:
        pass
    try:
        info["python_executable"] = sys.executable
        info["python_version"] = sys.version.split()[0]
        info["pid"] = os.getpid()
    except Exception:
        pass
    return info


def ping():
    """Cheap liveness report: is the bridge alive, and is SolidWorks answering?

    P0-1 asked for a `ping()` so the host can tell "alive but busy" from "dead". Unlike
    `health(deep=True)` this never touches ActiveDoc, so it answers in ~18 ms and cannot
    itself hang on a modal dialog.

    It always performs a real round trip (a cached pointer proves nothing), and if that
    fails it rebuilds the link once and asks again — so a stale dispatcher makes ping
    report `recovered: true` instead of a false outage. `bridge` carries the counters
    behind the self-healing: stale dispatchers dropped and reconnects performed.
    """
    started = time.time()
    payload = {
        "sw_version": None,
        "latency_ms": None,
        "last_call": dict(_LAST_CALL) if isinstance(_LAST_CALL, dict) else None,
        "bridge": None,
        "probe_interval_s": LIVENESS_PROBE_INTERVAL_S,
        "com_thread": _COM_THREAD.name if _COM_THREAD is not None else "inline",
        "allowed_roots": allowed_roots() or None,
        "recovered": False,
        # Which copy is running — see runtime_info() for why the path matters.
        "runtime": runtime_info(),
    }
    for attempt in (0, 1):
        try:
            app = get_sw(probe=False)
            payload["sw_version"] = str(app.RevisionNumber)
            payload["latency_ms"] = round((time.time() - started) * 1000, 1)
            payload["recovered"] = attempt > 0
            payload["bridge"] = dict(_BRIDGE_STATE)
            return ok(**payload)
        except Exception as exc:
            payload["bridge"] = dict(_BRIDGE_STATE)
            if attempt == 0:
                payload["first_error"] = str(exc)[:160]
                try:
                    reconnect(reason="ping: %s" % str(exc)[:120])
                except Exception as reconnect_exc:
                    payload["reconnect_error"] = str(reconnect_exc)[:160]
                    break
                continue
            break
    log_failure("ping", {"stage": "ping"}, exc=None,
                extra={"payload": {k: v for k, v in payload.items() if k != "bridge"}})
    return error(
        "sw_unreachable",
        "SolidWorks is not answering: %s" % (payload.get("first_error") or "unknown"),
        "Start SolidWorks or close a modal dialog, then call ping again. "
        "health(deep=True) additionally reads ActiveDoc.",
        **payload)


# Members win32com dispatches as properties although they are methods. Kept here so the
# static check and the helper share one list.
PROPERTY_TRAP_METHODS = (
    "InsertGtol",
    "GetAnnotation",
    "GetAnnotation2",
    "FirstFeature",
    "GetNextFeature",
    "GetDefinition",
    "SetPosition2",
    "Select3",
    "SetScale",
)


def byref_dispatch_null():
    """Null Callout argument for SelectByID2 (byref IDispatch).

    SelectByID2 fails with '类型不匹配' unless the Callout argument is passed
    this way (docs/com-quirks.md, section 3).
    """
    pythoncom, _, _ = _com()
    return variant(None, pythoncom.VT_BYREF | pythoncom.VT_DISPATCH)


_DOC_TYPE_BY_EXT = {
    ".sldprt": SW_DOC_PART,
    ".sldasm": SW_DOC_ASSEMBLY,
    ".slddrw": SW_DOC_DRAWING,
}


def open_document(path, doc_type=None):
    """Open path silently, reusing the active document when it is the same file.

    doc_type defaults to being inferred from the extension. Hard-coding the assembly type
    made `get_mass_properties` and `export_preview` fail on every part with a misleading
    "file may be locked" message, when the real cause was `OpenDoc6` returning None for a
    type mismatch (reported 2026-10-28).
    """
    path = check_path_allowed(path)
    if not os.path.exists(path):
        # FileNotFoundError(path) stringifies to a bare path; keep the reason.
        raise FileNotFoundError(
            "file not found: %s (check drive, folder and extension "
            ".SLDASM/.SLDPRT)" % path
        )
    extension = os.path.splitext(path)[1].lower()
    if doc_type is None:
        doc_type = _DOC_TYPE_BY_EXT.get(extension, SW_DOC_ASSEMBLY)
    sw = get_sw()
    doc = sw.ActiveDoc
    if doc is not None and os.path.normcase(doc.GetPathName) == os.path.normcase(path):
        return doc
    pythoncom, _, _ = _com()
    errors = variant(0, pythoncom.VT_BYREF | pythoncom.VT_I4)
    warnings = variant(0, pythoncom.VT_BYREF | pythoncom.VT_I4)
    doc = sw.OpenDoc6(path, doc_type, SW_OPEN_SILENT, "", errors, warnings)
    if doc is None:
        # Keep the COM detail: errors=2097152 is the type-mismatch branch, and dropping it
        # sent one reporter down two wrong paths (read-only, same-name lock).
        error_code = None
        try:
            error_code = int(errors.value)
        except Exception:
            pass
        hint = ("OpenDoc6 returned null with errors=%s for doc_type=%s (%s). "
                "The most common cause is a document-type mismatch, not a file lock."
                % (error_code, doc_type, extension or "no extension"))
        raise RuntimeError(
            "SolidWorks could not open the document. It may be open read-only "
            "elsewhere, or another file with the same name is already open; "
            "doc_type/extension mismatch is another common cause. %s "
            "(see docs/com-quirks.md, section 2)." % hint
        )
    # A reopened document means previously cached component pointers are stale.
    invalidate_components(path)
    return doc


def get_components(doc, refresh=False):
    """Fetch once per document and cache. Never call this in a loop (quirks section 4).

    Cache validity is checked two ways because a stale IComponent2 does NOT always
    raise after the document is closed and reopened from outside: it may silently
    return empty strings (Name2 == ""). So we require (a) the same COM document
    identity and (b) a non-empty first component name.
    """
    key = os.path.normcase(doc.GetPathName)
    if not refresh:
        entry = _COMPONENTS_CACHE.get(key)
        if entry:
            entry_ole, cached = entry
            same_doc = False
            try:
                same_doc = entry_ole == doc._oleobj_
            except Exception:
                same_doc = False
            if same_doc and cached:
                try:
                    if str(cached[0].Name2 or "").strip():
                        return cached
                except Exception:
                    pass
            _COMPONENTS_CACHE.pop(key, None)
    components = list(doc.GetComponents(False) or [])
    if components:
        try:
            _COMPONENTS_CACHE[key] = (doc._oleobj_, components)
        except Exception:
            pass
    return components


def invalidate_components(path=None):
    if path is None:
        _COMPONENTS_CACHE.clear()
    else:
        _COMPONENTS_CACHE.pop(os.path.normcase(os.path.abspath(path)), None)


def part_name(component_name):
    """'C Linkage-6' -> 'C Linkage' (instance suffix stripped)."""
    head, sep, tail = component_name.rpartition("-")
    return head if sep and tail.isdigit() else component_name


def rgb_of(values):
    if not values:
        return None
    return [round(float(v), 4) for v in values[:3]]


def ok(**payload):
    payload["ok"] = True
    return payload


def error(code, message, hint=None, **extra):
    payload = {"ok": False, "error": code, "message": message}
    if hint:
        payload["hint"] = hint
    if extra:
        payload.update(extra)
    return payload
