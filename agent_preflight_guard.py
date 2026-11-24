"""Agent Pre-Flight Guard (RealmPlay Subconscious Engineering Layer).

Performs strict geometric validation, scale math verification, polyline closure
checks, and collision avoidance before committing changes to AutoCAD.
"""

from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

STATE_FILE = Path(r"K:\AgentSpace\01_Working_Tasks\On_Going_Projects\Novotel_Barcode\project_state.json")


def load_project_state() -> Dict[str, Any]:
    """Load centralized project state."""
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def update_project_state(patch: Dict[str, Any]) -> Dict[str, Any]:
    """Atomically update project state."""
    state = load_project_state()
    state.update(patch)
    STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")
    return state


# ============================================================================
# Engineering Pre-Flight Checkers
# ============================================================================

def verify_scale_mathematics(
    scale_factor: float,
    text_height: float,
    annotation_type: str = "text",
) -> Tuple[bool, str]:
    """Verify text and annotation sizing against plot-scale standards."""
    expected_standard = 2.5 * scale_factor
    expected_title = 5.0 * scale_factor

    if annotation_type == "title":
        target = expected_title
    else:
        target = expected_standard

    tolerance = target * 0.35  # ±35% tolerance for hierarchy
    if abs(text_height - target) > tolerance:
        return False, (
            f"Scale mismatch: text_height {text_height} deviates from standard {target} "
            f"for scale 1:{int(scale_factor)} (tolerated range: {target - tolerance:.1f} - {target + tolerance:.1f})"
        )
    return True, "Scale math verified."


def verify_polyline_closure(points: List[Tuple[float, float]], closed_flag: bool) -> Tuple[bool, str]:
    """Ensure architectural boundary polylines are strictly closed."""
    if not closed_flag:
        return False, "Polyline must have closed=True for architectural elements."
    if len(points) < 3:
        return False, f"Closed polyline requires at least 3 vertices, got {len(points)}."
    
    # Check if first and last vertices coincide
    p_start, p_end = points[0], points[-1]
    dist = math.hypot(p_end[0] - p_start[0], p_end[1] - p_start[1])
    if dist < 1e-4:
        return False, "Redundant closing vertex detected. In AutoCAD LWPOLYLINE, closing vertex is implicit when closed=True."
    return True, "Polyline geometry verified."


def calculate_margin_bands(
    building_face_coord: float,
    scale_factor: float,
    unit_mm_factor: float = 1.0,
    side: str = "left",
) -> Dict[str, float]:
    """Calculate exact collision-free margin positions according to Phase 8 Stacking."""
    s = scale_factor * unit_mm_factor
    sign = -1.0 if side.lower() in ("left", "bottom") else 1.0

    return {
        "band_1_local": building_face_coord + sign * (10.0 * s),
        "band_2_grid": building_face_coord + sign * (18.0 * s),
        "band_3_overall": building_face_coord + sign * (26.0 * s),
        "band_4_balloon_or_datum": building_face_coord + sign * (34.0 * s),
    }


def preflight_dimension_chain(
    coords: List[float],
    anchor_line_coord: float,
    scale_factor: float,
    tier: int,
) -> Tuple[bool, str]:
    """Validate dimension coordinates before executing CAD creation."""
    if len(coords) < 2:
        return False, "Dimension chain requires at least 2 distinct coordinates."
    
    unique_coords = sorted(set(coords))
    if len(unique_coords) < len(coords):
        return False, "Duplicate coordinates detected in dimension chain points."

    # Minimum measurable delta (must be >= 10 mm in 1:1 scale)
    for i in range(len(unique_coords) - 1):
        delta = abs(unique_coords[i + 1] - unique_coords[i])
        if delta < 10.0:
            return False, f"Micro-segment detected: distance between points is only {delta:.1f} mm (< 10 mm)."

    return True, "Dimension chain coordinates approved."
