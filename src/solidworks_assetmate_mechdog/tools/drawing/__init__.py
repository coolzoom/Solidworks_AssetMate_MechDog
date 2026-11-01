"""Drawing creation, export, render and audit.

Verified on SolidWorks 2024 SP4 Chinese UI (2026-10-27). Traps that cost hours:

- sw.NewDocument(template, ...) returns None when the template path uses forward
  slashes. Windows backslashes only.
- Named views follow the UI language: Chinese builds answer `*前视`, not `*Front`.
  Always resolve the view name from the model's own GetModelViewNames.
- IDrawingDoc.GetViews / GetFirstView / GetOutline are PROPERTIES under late
  binding, not methods. GetViews returns a nested tuple of raw PyIDispatch.
- Coordinates are meters (A3 = 0.420 x 0.297).

Scope: geometry + views + export + audit. Dimensioning helpers live elsewhere and
are deliberately separate so a failed annotation never blocks a usable drawing."""

from ._core import *
from .create import (
    create_drawing,
    export_drawing,
    drawing_audit,
)
from .dims import (
    auto_dimension,
    add_drawing_dimension,
    add_dimension_tolerance,
    layout_dimensions,
    verify_dimension_attachments,
    prune_orphan_dimensions,
)
from .annot import (
    add_center_marks,
    add_hole_callout,
    add_hole_callouts,
    set_title_block,
    list_drawing_annotations,
    add_surface_finish,
    add_geometric_tolerance,
)
from .views import (
    add_flat_pattern_view,
    add_section_view,
    add_detail_view,
    list_drawing_edges,
    list_circular_edges,
)
from .audit import (
    delivery_check,
    reconcile_ledger,
)
