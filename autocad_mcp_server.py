# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "mcp[cli]<2",
#     "pywin32>=306",
#     "Pillow>=10.0.0",
# ]
# ///

import contextlib
import io
import logging
import math
import os
import sys
import tempfile
import time
import unicodedata
from typing import List, Dict, Any, Optional

from PIL import Image as PILImage, ImageFilter
import pythoncom
import pywintypes
import win32com.client
import win32con
import win32gui
from mcp.server.fastmcp import FastMCP, Image

# CRITICAL FOR ANTIGRAVITY: Log strictly to stderr, never stdout!
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("autocad-mcp")

mcp = FastMCP("AutoCAD-Antigravity-MCP")

# Error code when AutoCAD is temporarily busy (RPC_E_CALL_REJECTED)
RPC_E_CALL_REJECTED = -2147418111

# Commands forbidden from raw execution to prevent hanging AutoCAD
BLOCKED_COMMANDS = {"QUIT", "_QUIT", "EXIT", "_EXIT", "CLOSE", "CLOSEALL", "QSAVE", "SAVEAS"}

# ============================================================================
# Win32 OS-Level Window Handle & Force-ESC Recovery
# ============================================================================

def _find_autocad_hwnd() -> Optional[int]:
    """Find AutoCAD main window handle across all desktops even when COM is completely locked."""
    matched: List[int] = []

    # 1. Search Default desktop first
    try:
        import win32service
        desk = win32service.OpenDesktop("Default", 0, False, win32con.GENERIC_ALL)
        hwnds = desk.EnumDesktopWindows()
        for h in hwnds:
            title = win32gui.GetWindowText(h)
            cls_name = win32gui.GetClassName(h)
            if "AutoCAD" in title and (cls_name.startswith("Afx") or "AutoCAD" in cls_name):
                matched.append(int(h))
    except Exception:
        pass

    if not matched:
        def _enum_cb(hwnd, _):
            title = win32gui.GetWindowText(hwnd)
            cls_name = win32gui.GetClassName(hwnd)
            if "AutoCAD" in title and (cls_name.startswith("Afx") or "AutoCAD" in cls_name):
                matched.append(hwnd)
        with contextlib.suppress(Exception):
            win32gui.EnumWindows(_enum_cb, None)

    return matched[0] if matched else None


def force_escape_autocad(hwnd: Optional[int] = None, presses: int = 2) -> bool:
    """
    Send asynchronous ESC keystrokes directly to AutoCAD's window queue via PostMessage.
    Works in the background without stealing window focus from Antigravity.
    """
    target_hwnd = hwnd or _find_autocad_hwnd()
    if not target_hwnd:
        return False

    for _ in range(presses):
        win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
        time.sleep(0.04)
        win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)
        time.sleep(0.08)
    return True


def com_retry(func, retries: int = 6, delay: float = 0.35):
    """Retry COM operations; if AutoCAD is stuck in a command, auto-send ESC to recover."""
    for attempt in range(retries):
        try:
            return func()
        except pywintypes.com_error as exc:
            if exc.hresult == RPC_E_CALL_REJECTED and attempt < retries - 1:
                logger.warning("AutoCAD busy/locked (%d/%d) — sending background ESC...", attempt + 1, retries)
                if attempt >= 1:
                    force_escape_autocad(presses=2)
                time.sleep(delay * (attempt + 1))
                continue
            raise


def ensure_command_idle(acad, doc) -> int:
    """Check CMDACTIVE bitmask and cancel any hung prompt before running new tools."""
    cmd_state = int(com_retry(lambda: doc.GetVariable("CMDACTIVE")))
    if cmd_state > 0:
        logger.warning("Detected active AutoCAD command (CMDACTIVE=%d). Auto-canceling...", cmd_state)
        force_escape_autocad(hwnd=getattr(acad, "HWND", None), presses=3)
        com_retry(lambda: doc.SendCommand("\x03\x03"))
        time.sleep(0.15)
        cmd_state = int(com_retry(lambda: doc.GetVariable("CMDACTIVE")))
    return cmd_state


@contextlib.contextmanager
def suppress_autocad_dialogs(doc):
    """
    Temporarily disable all AutoCAD modal popups and attribute prompts
    (FILEDIA, CMDDIA, ATTDIA, ATTREQ) and guarantee restoration afterward.
    """
    sysvars = {"FILEDIA": 0, "CMDDIA": 0, "ATTDIA": 0, "ATTREQ": 0}
    saved = {}
    try:
        for var, target_val in sysvars.items():
            with contextlib.suppress(Exception):
                saved[var] = com_retry(lambda v=var: doc.GetVariable(v))
                com_retry(lambda v=var, tv=target_val: doc.SetVariable(v, tv))
        yield
    finally:
        for var, orig_val in saved.items():
            with contextlib.suppress(Exception):
                com_retry(lambda v=var, ov=orig_val: doc.SetVariable(v, ov))


def ensure_not_in_block_editor(doc) -> bool:
    """If AutoCAD was left inside BEDIT mode (BLOCKEDITOR=1), close it safely so ModelSpace works."""
    in_bedit = int(com_retry(lambda: doc.GetVariable("BLOCKEDITOR")))
    if in_bedit == 1:
        logger.warning("AutoCAD is trapped inside Block Editor (BEDIT). Exiting safely...")
        com_retry(lambda: doc.SendCommand("\x03\x03._BCLOSE\n_Save\n"))
        time.sleep(0.25)
        return True
    return False


def _coerce_dyn_variant(prop, new_val: Any):
    """
    Match exact COM VARIANT type required by AutoCAD Dynamic Block properties.
    Prevents '-2145386476 Invalid input' on Flip states (VT_I2) and Distances (VT_R8).
    """
    cur_val = prop.Value
    allowed = list(getattr(prop, "AllowedValues", []) or [])

    if allowed and set(allowed).issubset({0, 1}) and len(allowed) == 2:
        return win32com.client.VARIANT(pythoncom.VT_I2, int(new_val))

    if isinstance(cur_val, str) or (allowed and isinstance(allowed[0], str)):
        str_val = str(new_val)
        if allowed and str_val not in allowed:
            raise ValueError(f"Value '{str_val}' not in AllowedValues: {allowed}")
        return win32com.client.VARIANT(pythoncom.VT_BSTR, str_val)

    if isinstance(cur_val, (int, float)):
        return win32com.client.VARIANT(pythoncom.VT_R8, float(new_val))

    return new_val


@contextlib.contextmanager
def autocad_session(auto_cancel_active: bool = True):
    """
    Thread-safe context manager for AutoCAD COM connection.
    When running under Antigravity or standard MCP, initializes COM cleanly on the active thread.
    If executed from a sandboxed desktop, safely switches thread desktop before COM initialization.
    """
    try:
        import win32service
        with contextlib.suppress(Exception):
            desk = win32service.OpenDesktop("Default", 0, False, win32con.MAXIMUM_ALLOWED)
            desk.SetThreadDesktop()
    except Exception:
        pass

    pythoncom.CoInitialize()
    try:
        def _connect():
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
            doc = acad.ActiveDocument
            return acad, doc

        acad, doc = com_retry(_connect)
        if auto_cancel_active:
            ensure_command_idle(acad, doc)
        yield acad, doc
    except pywintypes.com_error as exc:
        force_escape_autocad(presses=3)
        logger.error("COM connection error: %s", exc)
        raise RuntimeError(
            "AutoCAD was locked or unreachable. Background ESC has been triggered—please retry."
        ) from exc
    finally:
        pythoncom.CoUninitialize()


@mcp.tool()
def cancel_active_command() -> Dict[str, Any]:
    """Force-cancel any stuck command or prompt in AutoCAD without user intervention."""
    sent_os_esc = force_escape_autocad(presses=3)
    with autocad_session(auto_cancel_active=True) as (_, doc):
        rem = int(com_retry(lambda: doc.GetVariable("CMDACTIVE")))
        return {
            "os_escape_sent": sent_os_esc,
            "cmdactive_after": rem,
            "status": "IDLE" if rem == 0 else "STILL_ACTIVE",
        }

# ============================================================================
# Typography & Bounding-Box Helpers
# ============================================================================

def pt3d(x: float, y: float, z: float = 0.0) -> win32com.client.VARIANT:
    """Convert 3D coordinates into an AutoCAD-compatible COM Double Array."""
    return win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [float(x), float(y), float(z)],
    )

def pt2d_flat(points: List[List[float]]) -> win32com.client.VARIANT:
    """Flatten [[x1, y1], [x2, y2], ...] into a 2D LightWeightPolyline COM array."""
    flat = []
    for pt in points:
        flat.extend([float(pt[0]), float(pt[1])])
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)

def _visual_text_len(text: str) -> float:
    """Compute visual character width (ignores zero-width Thai combining vowels/tone marks)."""
    max_w = 0.0
    for line in text.replace("\\P", "\n").splitlines() or [""]:
        w = 0.0
        for ch in line:
            if unicodedata.combining(ch) or unicodedata.category(ch) == "Mn":
                continue
            if unicodedata.east_asian_width(ch) in ("W", "F"):
                w += 1.0
            else:
                w += 0.72
        max_w = max(max_w, w)
    return max(max_w, 1.0)


def _read_bbox(ent) -> Dict[str, float]:
    """Read exact 2D BoundingBox from an AutoCAD entity."""
    bmin, bmax = ent.GetBoundingBox()
    return {
        "x1": round(float(bmin[0]), 2),
        "y1": round(float(bmin[1]), 2),
        "x2": round(float(bmax[0]), 2),
        "y2": round(float(bmax[1]), 2),
        "width": round(float(bmax[0] - bmin[0]), 2),
        "height": round(float(bmax[1] - bmin[1]), 2),
    }


def _clamp_to_zone(
    ent, zone_bounds: Optional[List[float]], margin: float = 0.0
) -> Dict[str, Any]:
    """Keep an entity's BoundingBox strictly inside [zx1, zy1, zx2, zy2] with margin."""
    bbox = _read_bbox(ent)
    if not zone_bounds or len(zone_bounds) != 4:
        return {"bbox": bbox, "nudged": [0.0, 0.0], "fits_in_zone": True}

    zx1, zy1 = min(zone_bounds[0], zone_bounds[2]) + margin, min(zone_bounds[1], zone_bounds[3]) + margin
    zx2, zy2 = max(zone_bounds[0], zone_bounds[2]) - margin, max(zone_bounds[1], zone_bounds[3]) - margin

    dx, dy = 0.0, 0.0
    if bbox["x2"] > zx2:
        dx = zx2 - bbox["x2"]
    if bbox["x1"] + dx < zx1:
        dx = zx1 - bbox["x1"]

    if bbox["y1"] < zy1:
        dy = zy1 - bbox["y1"]
    if bbox["y2"] + dy > zy2:
        dy = zy2 - bbox["y2"]

    if abs(dx) > 1e-3 or abs(dy) > 1e-3:
        ent.Move(pt3d(0.0, 0.0, 0.0), pt3d(dx, dy, 0.0))
        ent.Update()
        bbox = _read_bbox(ent)

    fits = (
        bbox["x1"] >= zx1 - 1e-2
        and bbox["x2"] <= zx2 + 1e-2
        and bbox["y1"] >= zy1 - 1e-2
        and bbox["y2"] <= zy2 + 1e-2
    )
    return {
        "bbox": bbox,
        "nudged": [round(dx, 2), round(dy, 2)],
        "fits_in_zone": fits,
    }

# ============================================================================
# Read & Inspection Tools
# ============================================================================

@mcp.tool()
def get_drawing_context() -> Dict[str, Any]:
    """
    Read active drawing metadata, units, UCS origin, and INSBASE.
    Always call this BEFORE drawing to prevent coordinate shifts on templates.
    """
    with autocad_session() as (_, doc):
        def _read():
            insbase = list(doc.GetVariable("INSBASE"))
            ucsorg = list(doc.GetVariable("UCSORG"))
            extmin = list(doc.GetVariable("EXTMIN"))
            extmax = list(doc.GetVariable("EXTMAX"))
            layers = [doc.Layers.Item(i).Name for i in range(doc.Layers.Count)]
            return {
                "drawing_name": doc.Name,
                "full_path": doc.FullName,
                "active_layer": doc.ActiveLayer.Name,
                "insbase_origin": insbase,
                "ucs_origin": ucsorg,
                "extents_min": extmin,
                "extents_max": extmax,
                "insunits": doc.GetVariable("INSUNITS"),
                "layer_count": len(layers),
                "layers": layers,
            }
        return com_retry(_read)

@mcp.tool()
def audit_entities_and_hatches(layer_filter: str = "") -> Dict[str, Any]:
    """
    Count ModelSpace entities by type/layer and detect duplicate or overlapping Hatches.
    Optionally filter by a specific layer name.
    """
    with autocad_session() as (_, doc):
        def _audit():
            ms = doc.ModelSpace
            counts_by_type: Dict[str, int] = {}
            seen_hatches: Dict[str, List[str]] = {}
            duplicate_hatches: List[Dict[str, Any]] = []

            for i in range(ms.Count):
                ent = ms.Item(i)
                layer = ent.Layer
                if layer_filter and layer.lower() != layer_filter.lower():
                    continue

                obj_name = ent.ObjectName
                counts_by_type[obj_name] = counts_by_type.get(obj_name, 0) + 1

                if obj_name == "AcDbHatch":
                    try:
                        area = round(float(ent.Area), 2)
                        min_pt, max_pt = ent.GetBoundingBox()
                        sig = (
                            f"{layer}|area:{area}|"
                            f"min:{round(min_pt[0],1)},{round(min_pt[1],1)}|"
                            f"max:{round(max_pt[0],1)},{round(max_pt[1],1)}"
                        )
                        handle = ent.Handle
                        if sig in seen_hatches:
                            duplicate_hatches.append({
                                "layer": layer,
                                "pattern": ent.PatternName,
                                "area": area,
                                "original_handle": seen_hatches[sig][0],
                                "duplicate_handle": handle,
                            })
                        seen_hatches[sig] = [handle]
                    except Exception:
                        pass

            return {
                "total_scanned": sum(counts_by_type.values()),
                "counts_by_type": counts_by_type,
                "duplicate_hatch_count": len(duplicate_hatches),
                "suspected_duplicate_hatches": duplicate_hatches,
            }
        return com_retry(_audit)

# ============================================================================
# Visual QA/QC Inspection Tool (.PNG Viewport Capture)
# ============================================================================

@mcp.tool()
def capture_drawing_view(
    mode: str = "extents",
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    enhance_thin_lines: bool = True,
    max_width: int = 1920,
) -> Image:
    """
    Capture the current AutoCAD drawing directly from the viewport buffer as a PNG image
    for visual QA/QC inspection (works even if other windows cover AutoCAD).
    - mode="extents": Zoom to full drawing/title-block extents to inspect overall framing.
    - mode="window": Zoom into rectangular box (x1, y1) to (x2, y2) to inspect text/leaders/tables.
    """
    preview_dir = os.path.join(tempfile.gettempdir(), "autocad_mcp_previews")
    os.makedirs(preview_dir, exist_ok=True)
    raw_png = os.path.join(preview_dir, f"cad_capture_{int(time.time() * 1000)}.png")
    clean_path = raw_png.replace("\\", "/")

    with autocad_session() as (acad, doc):
        old_filedia = com_retry(lambda: doc.GetVariable("FILEDIA"))
        try:
            def _prepare_and_export():
                if acad.WindowState == 2:
                    acad.WindowState = 1

                doc.SetVariable("FILEDIA", 0)

                if mode.lower() == "window":
                    if x1 == x2 or y1 == y2:
                        raise ValueError("For mode='window', (x1, y1) and (x2, y2) must form a valid box.")
                    min_x, max_x = min(x1, x2), max(x1, x2)
                    min_y, max_y = min(y1, y2), max(y1, y2)
                    acad.ZoomWindow(pt3d(min_x, min_y, 0.0), pt3d(max_x, max_y, 0.0))
                else:
                    acad.ZoomExtents()

                doc.Regen(1)
                doc.SendCommand(f"._-PNGOUT\n{clean_path}\n\n")

            com_retry(_prepare_and_export)

            for _ in range(25):
                if os.path.exists(raw_png) and os.path.getsize(raw_png) > 0:
                    time.sleep(0.15)
                    break
                time.sleep(0.2)
            else:
                raise RuntimeError("Timed out waiting for AutoCAD PNGOUT file generation.")

        finally:
            with contextlib.suppress(Exception):
                com_retry(lambda: doc.SetVariable("FILEDIA", old_filedia))

    with PILImage.open(raw_png) as img:
        rgb = img.convert("RGB")

        if enhance_thin_lines:
            corner_r, corner_g, corner_b = rgb.getpixel((2, 2))
            luminance = 0.299 * corner_r + 0.587 * corner_g + 0.114 * corner_b
            if luminance < 128:
                rgb = rgb.filter(ImageFilter.MaxFilter(3))
            else:
                rgb = rgb.filter(ImageFilter.MinFilter(3))

        if rgb.width > max_width:
            ratio = max_width / float(rgb.width)
            new_height = int(rgb.height * ratio)
            rgb = rgb.resize((max_width, new_height), PILImage.Resampling.LANCZOS)

        buf = io.BytesIO()
        rgb.save(buf, format="PNG", optimize=True)
        png_bytes = buf.getvalue()

    logger.info("Captured CAD view (%s) -> %s (%d bytes)", mode, raw_png, len(png_bytes))
    return Image(data=png_bytes, format="png")

# ============================================================================
# Basic Drawing & Entity Tools
# ============================================================================

@mcp.tool()
def ensure_layer(layer_name: str, color_index: int = 7) -> str:
    """Create a layer if it does not exist and set its ACI color (1=Red, 2=Yellow, 3=Green, 7=White)."""
    with autocad_session() as (_, doc):
        def _create():
            layer = doc.Layers.Add(layer_name)
            layer.color = int(color_index)
            return f"Layer '{layer_name}' ready (ACI color {color_index})."
        return com_retry(_create)

@mcp.tool()
def draw_line(
    x1: float, y1: float, x2: float, y2: float, layer: str = ""
) -> Dict[str, str]:
    """Draw a 2D line in ModelSpace between (x1, y1) and (x2, y2)."""
    with autocad_session() as (_, doc):
        def _draw():
            line = doc.ModelSpace.AddLine(pt3d(x1, y1), pt3d(x2, y2))
            if layer:
                doc.Layers.Add(layer)
                line.Layer = layer
            line.Update()
            return {"status": "created", "handle": line.Handle, "layer": line.Layer}
        return com_retry(_draw)

@mcp.tool()
def draw_polyline(
    points: List[List[float]], closed: bool = False, layer: str = ""
) -> Dict[str, str]:
    """
    Draw a 2D LightWeightPolyline from a list of [x, y] coordinate pairs.
    Example points: [[0, 0], [100, 0], [100, 50], [0, 50]]
    """
    if len(points) < 2:
        raise ValueError("At least two [x, y] points are required.")

    with autocad_session() as (_, doc):
        def _draw():
            pline = doc.ModelSpace.AddLightWeightPolyline(pt2d_flat(points))
            pline.Closed = bool(closed)
            if layer:
                doc.Layers.Add(layer)
                pline.Layer = layer
            pline.Update()
            return {"status": "created", "handle": pline.Handle, "closed": str(closed)}
        return com_retry(_draw)

@mcp.tool()
def edit_text_by_handle(handle: str, new_text: str) -> str:
    """Modify the text string of a Text or MText entity using its Handle."""
    with autocad_session() as (_, doc):
        def _edit():
            ent = doc.HandleToObject(handle)
            obj_name = ent.ObjectName
            if obj_name in ("AcDbText", "AcDbMText"):
                ent.TextString = new_text
                ent.Update()
                return f"Updated {obj_name} '{handle}' text to: {new_text}"
            raise ValueError(f"Entity '{handle}' is {obj_name}, not Text/MText.")
        return com_retry(_edit)

@mcp.tool()
def delete_entity_by_handle(handle: str) -> str:
    """Delete a specific entity (such as a duplicate Hatch) using its hexadecimal Handle."""
    with autocad_session() as (_, doc):
        def _delete():
            ent = doc.HandleToObject(handle)
            obj_type = ent.ObjectName
            ent.Delete()
            return f"Deleted {obj_type} with handle '{handle}'."
        return com_retry(_delete)

# ============================================================================
# Transform Tools: Move, Base-Point Copy, Scale
# ============================================================================

@mcp.tool()
def move_entities(
    handles: List[str],
    from_x: float,
    from_y: float,
    to_x: float,
    to_y: float,
) -> Dict[str, Any]:
    """
    Move one or multiple entities together from a base point (from_x, from_y)
    to a destination point (to_x, to_y), preserving their relative spacing.
    """
    p1 = pt3d(from_x, from_y, 0.0)
    p2 = pt3d(to_x, to_y, 0.0)

    with autocad_session() as (_, doc):
        def _move():
            moved = []
            for h in handles:
                ent = doc.HandleToObject(h)
                ent.Move(p1, p2)
                ent.Update()
                moved.append(h)
            return {
                "status": "moved",
                "moved_count": len(moved),
                "handles": moved,
                "delta": [round(to_x - from_x, 3), round(to_y - from_y, 3)],
            }
        return com_retry(_move)

@mcp.tool()
def copy_entities_basepoint(
    handles: List[str],
    base_x: float,
    base_y: float,
    target_x: float,
    target_y: float,
    target_layer: str = "",
) -> Dict[str, Any]:
    """
    Copy one or multiple entities using Base-Point Copy methodology from (base_x, base_y)
    to (target_x, target_y). Optionally place the copies onto target_layer.
    Returns the new handles of all duplicated objects.
    """
    p1 = pt3d(base_x, base_y, 0.0)
    p2 = pt3d(target_x, target_y, 0.0)

    with autocad_session() as (_, doc):
        def _copy():
            if target_layer:
                doc.Layers.Add(target_layer)

            new_handles = []
            for h in handles:
                src = doc.HandleToObject(h)
                cloned = src.Copy()
                cloned.Move(p1, p2)
                if target_layer:
                    cloned.Layer = target_layer
                cloned.Update()
                new_handles.append(cloned.Handle)

            return {
                "status": "copied",
                "copied_count": len(new_handles),
                "source_handles": handles,
                "new_handles": new_handles,
            }
        return com_retry(_copy)

@mcp.tool()
def scale_entities(
    handles: List[str],
    base_x: float,
    base_y: float,
    scale_factor: float,
) -> Dict[str, Any]:
    """
    Scale one or multiple entities around anchor point (base_x, base_y) by scale_factor.
    Example: 0.5 shrinks to half size, 2.0 doubles size.
    """
    if scale_factor <= 0:
        raise ValueError("scale_factor must be greater than 0.")

    base_pt = pt3d(base_x, base_y, 0.0)

    with autocad_session() as (_, doc):
        def _scale():
            scaled = []
            for h in handles:
                ent = doc.HandleToObject(h)
                ent.ScaleEntity(base_pt, float(scale_factor))
                ent.Update()
                scaled.append(h)
            return {
                "status": "scaled",
                "scaled_count": len(scaled),
                "handles": scaled,
                "scale_factor": scale_factor,
            }
        return com_retry(_scale)

# ============================================================================
# Spatial & Anti-Clutter Tool: Empty Space Finder
# ============================================================================

@mcp.tool()
def find_free_zones(
    min_x: float = 0.0,
    min_y: float = 0.0,
    max_x: float = 0.0,
    max_y: float = 0.0,
    clearance_margin: float = 0.0,
    grid_cols: int = 24,
    grid_rows: int = 16,
    top_k: int = 5,
) -> Dict[str, Any]:
    """
    Analyze spatial density inside a region (defaults to drawing extents if 0,0,0,0)
    and return the top_k largest non-overlapping empty rectangular zones for clean,
    uncluttered placement of tables, notes, legends, or callouts.
    """
    with autocad_session() as (_, doc):
        def _scan():
            if min_x == max_x or min_y == max_y:
                ext_min = doc.GetVariable("EXTMIN")
                ext_max = doc.GetVariable("EXTMAX")
                rx1, ry1 = float(ext_min[0]), float(ext_min[1])
                rx2, ry2 = float(ext_max[0]), float(ext_max[1])
            else:
                rx1, rx2 = min(min_x, max_x), max(min_x, max_x)
                ry1, ry2 = min(min_y, max_y), max(min_y, max_y)

            region_w = rx2 - rx1
            region_h = ry2 - ry1
            if region_w <= 0 or region_h <= 0:
                raise RuntimeError("Drawing region is empty or invalid.")

            cell_w = region_w / grid_cols
            cell_h = region_h / grid_rows
            pad = clearance_margin if clearance_margin > 0 else min(cell_w, cell_h) * 0.25

            grid = [[0 for _ in range(grid_cols)] for _ in range(grid_rows)]
            ms = doc.ModelSpace

            def _mark_box(bx1: float, by1: float, bx2: float, by2: float):
                c1 = max(0, int((bx1 - rx1) / cell_w))
                c2 = min(grid_cols - 1, int((bx2 - rx1) / cell_w))
                r1 = max(0, int((by1 - ry1) / cell_h))
                r2 = min(grid_rows - 1, int((by2 - ry1) / cell_h))
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        grid[r][c] = 1

            for i in range(ms.Count):
                ent = ms.Item(i)
                try:
                    if not ent.Visible:
                        continue
                    bmin, bmax = ent.GetBoundingBox()
                except Exception:
                    continue

                ex1, ey1, ex2, ey2 = bmin[0], bmin[1], bmax[0], bmax[1]
                if ex2 < rx1 or ex1 > rx2 or ey2 < ry1 or ey1 > ry2:
                    continue

                bw, bh = ex2 - ex1, ey2 - ey1
                if bw > 0.85 * region_w and bh > 0.85 * region_h and ent.ObjectName != "AcDbHatch":
                    _mark_box(ex1 - pad, ey1 - pad, ex2 + pad, ey1 + pad)
                    _mark_box(ex1 - pad, ey2 - pad, ex2 + pad, ey2 + pad)
                    _mark_box(ex1 - pad, ey1 - pad, ex1 + pad, ey2 + pad)
                    _mark_box(ex2 - pad, ey1 - pad, ex2 + pad, ey2 + pad)
                else:
                    _mark_box(ex1 - pad, ey1 - pad, ex2 + pad, ey2 + pad)

            occupied_cells = sum(sum(row) for row in grid)
            total_cells = grid_rows * grid_cols
            occupancy_pct = round((occupied_cells / total_cells) * 100.0, 1)

            ascii_map = [
                "".join("#" if grid[r][c] else "." for c in range(grid_cols))
                for r in range(grid_rows - 1, -1, -1)
            ]

            work_grid = [row[:] for row in grid]
            free_zones = []

            for rank in range(1, top_k + 1):
                best_area = 0
                best_rect = None
                heights = [0] * grid_cols

                for r in range(grid_rows):
                    for c in range(grid_cols):
                        heights[c] = heights[c] + 1 if work_grid[r][c] == 0 else 0

                    stack = []
                    for c in range(grid_cols + 1):
                        cur_h = heights[c] if c < grid_cols else 0
                        start_c = c
                        while stack and stack[-1][1] > cur_h:
                            idx, h_val = stack.pop()
                            area = h_val * (c - idx)
                            if area > best_area:
                                best_area = area
                                best_rect = (r - h_val + 1, idx, r, c - 1)
                            start_c = idx
                        stack.append((start_c, cur_h))

                if not best_rect or best_area == 0:
                    break

                r1, c1, r2, c2 = best_rect
                for rr in range(r1, r2 + 1):
                    for cc in range(c1, c2 + 1):
                        work_grid[rr][cc] = 2

                zx1 = round(rx1 + c1 * cell_w, 2)
                zy1 = round(ry1 + r1 * cell_h, 2)
                zx2 = round(rx1 + (c2 + 1) * cell_w, 2)
                zy2 = round(ry1 + (r2 + 1) * cell_h, 2)

                free_zones.append({
                    "zone_rank": rank,
                    "x1": zx1,
                    "y1": zy1,
                    "x2": zx2,
                    "y2": zy2,
                    "center_x": round((zx1 + zx2) / 2.0, 2),
                    "center_y": round((zy1 + zy2) / 2.0, 2),
                    "width": round(zx2 - zx1, 2),
                    "height": round(zy2 - zy1, 2),
                    "sheet_share_pct": round((best_area / total_cells) * 100.0, 1),
                })

            return {
                "scanned_region": {"x1": round(rx1, 2), "y1": round(ry1, 2), "x2": round(rx2, 2), "y2": round(ry2, 2)},
                "occupancy_pct": occupancy_pct,
                "whitespace_pct": round(100.0 - occupancy_pct, 1),
                "layout_status": "CROWDED" if occupancy_pct > 65.0 else "CLEAN",
                "ascii_density_map_top_down": ascii_map,
                "recommended_free_zones": free_zones,
            }
        return com_retry(_scan)

# ============================================================================
# Smart Annotation Tools: MText, MLeader, Table
# ============================================================================

@mcp.tool()
def draw_mtext(
    text: str,
    x: float,
    y: float,
    text_height: float = 250.0,
    box_width: float = 0.0,
    layer: str = "G-ANNO-TEXT",
    background_mask: bool = True,
    zone_bounds: Optional[List[float]] = None,
    zone_margin: float = 0.0,
) -> Dict[str, Any]:
    """
    Create multi-line text (Top-Left anchored at x, y) with Thai line-spacing clearance,
    optional background mask, and automatic clamping inside zone_bounds [zx1, zy1, zx2, zy2].
    """
    formatted_text = text.replace("\r\n", "\\P").replace("\n", "\\P")

    with autocad_session() as (_, doc):
        def _create():
            if layer:
                doc.Layers.Add(layer)

            est_w = _visual_text_len(formatted_text) * text_height * 1.15
            max_zone_w = None
            if zone_bounds and len(zone_bounds) == 4:
                max_zone_w = abs(zone_bounds[2] - zone_bounds[0]) - (2.0 * zone_margin)

            wrap_w = box_width if box_width > 0 else est_w
            if max_zone_w and max_zone_w > text_height * 3:
                wrap_w = min(wrap_w, max_zone_w)

            mt = doc.ModelSpace.AddMText(pt3d(x, y, 0.0), float(wrap_w), formatted_text)
            mt.AttachmentPoint = 1
            mt.InsertionPoint = pt3d(x, y, 0.0)
            mt.Height = float(text_height)
            mt.LineSpacingFactor = 1.22
            if background_mask:
                with contextlib.suppress(Exception):
                    mt.BackgroundFill = True
            if layer:
                mt.Layer = layer
            mt.Update()

            placement = _clamp_to_zone(mt, zone_bounds, zone_margin)
            return {
                "status": "created",
                "handle": mt.Handle,
                "layer": mt.Layer,
                **placement,
            }
        return com_retry(_create)


@mcp.tool()
def draw_mleader(
    target_x: float,
    target_y: float,
    text_x: float,
    text_y: float,
    text: str,
    text_height: float = 250.0,
    arrow_size: float = 0.0,
    target_clearance: float = 0.0,
    landing_length: float = 0.0,
    layer: str = "G-ANNO-TEXT",
) -> Dict[str, Any]:
    """
    Draw a clean callout leader with automatic horizontal shoulder (dogleg) and optional
    target_clearance pullback so the arrowhead never covers column/wall centerlines.
    - (target_x, target_y): Point on the structural element being called out.
    - (text_x, text_y): Note location placed in clear space / free_zone.
    """
    formatted_text = text.replace("\r\n", "\\P").replace("\n", "\\P")
    arr_sz = arrow_size if arrow_size > 0 else text_height * 1.0
    landing = landing_length if landing_length > 0 else text_height * 2.2

    dir_sign = 1.0 if text_x >= target_x else -1.0
    elbow_x = text_x - (dir_sign * landing)
    elbow_y = text_y

    vx, vy = elbow_x - target_x, elbow_y - target_y
    seg_len = math.hypot(vx, vy)
    if target_clearance > 0 and seg_len > target_clearance * 1.5:
        ratio = target_clearance / seg_len
        tip_x = target_x + vx * ratio
        tip_y = target_y + vy * ratio
    else:
        tip_x, tip_y = target_x, target_y

    pts_3d = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [float(tip_x), float(tip_y), 0.0, float(elbow_x), float(elbow_y), 0.0],
    )

    with autocad_session() as (_, doc):
        def _create():
            if layer:
                doc.Layers.Add(layer)

            try:
                idx_ref = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                ml = doc.ModelSpace.AddMLeader(pts_3d, idx_ref)
                ml.ContentType = 2
                ml.TextString = formatted_text
                ml.TextHeight = float(text_height)
                ml.ArrowheadSize = float(arr_sz)
                ml.DoglegLength = float(landing)
                ml.LandingGap = float(text_height * 0.4)
                with contextlib.suppress(Exception):
                    ml.TextLineSpacingFactor = 1.22
                    ml.TextBackgroundFill = True
                if layer:
                    ml.Layer = layer
                ml.Update()
                return {
                    "status": "created_mleader",
                    "handles": [ml.Handle],
                    "arrow_tip": [round(tip_x, 2), round(tip_y, 2)],
                    "bbox": _read_bbox(ml),
                }
            except Exception:
                wrap_w = _visual_text_len(formatted_text) * text_height * 1.15
                mt = doc.ModelSpace.AddMText(pt3d(text_x, text_y, 0.0), float(wrap_w), formatted_text)
                mt.AttachmentPoint = 4 if dir_sign > 0 else 6
                mt.InsertionPoint = pt3d(text_x, text_y, 0.0)
                mt.Height = float(text_height)
                mt.LineSpacingFactor = 1.22
                with contextlib.suppress(Exception):
                    mt.BackgroundFill = True

                ldr_pts = win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_R8,
                    [
                        float(tip_x), float(tip_y), 0.0,
                        float(elbow_x), float(elbow_y), 0.0,
                        float(text_x), float(text_y), 0.0,
                    ],
                )
                ldr = doc.ModelSpace.AddLeader(ldr_pts, mt, 2)
                ldr.ArrowheadSize = float(arr_sz)
                if layer:
                    mt.Layer = layer
                    ldr.Layer = layer
                mt.Update()
                ldr.Update()

                b1, b2 = _read_bbox(mt), _read_bbox(ldr)
                merged_bbox = {
                    "x1": min(b1["x1"], b2["x1"]),
                    "y1": min(b1["y1"], b2["y1"]),
                    "x2": max(b1["x2"], b2["x2"]),
                    "y2": max(b1["y2"], b2["y2"]),
                }
                merged_bbox["width"] = round(merged_bbox["x2"] - merged_bbox["x1"], 2)
                merged_bbox["height"] = round(merged_bbox["y2"] - merged_bbox["y1"], 2)
                return {
                    "status": "created_leader_pair",
                    "handles": [ldr.Handle, mt.Handle],
                    "arrow_tip": [round(tip_x, 2), round(tip_y, 2)],
                    "bbox": merged_bbox,
                }
        return com_retry(_create)


@mcp.tool()
def draw_table(
    x: float,
    y: float,
    headers: List[str],
    rows: List[List[str]],
    title: str = "",
    text_height: float = 250.0,
    layer: str = "G-ANNO-TEXT",
    zone_bounds: Optional[List[float]] = None,
    zone_margin: float = 0.0,
) -> Dict[str, Any]:
    """
    Create an engineering schedule/BOQ Table at Top-Left (x, y) with auto-calculated
    column widths, dynamic row wrapping (never overflows cells), Thai vertical clearance,
    and automatic clamping inside zone_bounds [zx1, zy1, zx2, zy2].
    """
    if not headers:
        raise ValueError("headers must contain at least 1 column.")

    num_cols = len(headers)
    has_title = bool(title.strip())
    header_row_idx = 1 if has_title else 0
    data_start_idx = header_row_idx + 1
    total_rows = data_start_idx + len(rows)

    pad_x = text_height * 2.4
    pad_y = text_height * 1.15

    col_widths = []
    for c in range(num_cols):
        max_chars = _visual_text_len(str(headers[c])) * 1.15
        for r_data in rows:
            cell_val = str(r_data[c]) if c < len(r_data) else ""
            max_chars = max(max_chars, _visual_text_len(cell_val))
        col_widths.append(max(text_height * 5.0, max_chars * text_height + pad_x))

    if zone_bounds and len(zone_bounds) == 4:
        avail_w = abs(zone_bounds[2] - zone_bounds[0]) - (2.0 * zone_margin)
        total_w = sum(col_widths)
        min_viable_w = num_cols * text_height * 4.0
        if total_w > avail_w and avail_w >= min_viable_w:
            scale_ratio = avail_w / total_w
            col_widths = [max(text_height * 4.0, cw * scale_ratio) for cw in col_widths]

    def _calc_row_height(cells: List[str], row_txt_h: float) -> float:
        max_lines = 1
        for c, val in enumerate(cells[:num_cols]):
            usable_w = max(row_txt_h * 2.0, col_widths[c] - pad_x)
            explicit_lines = str(val).replace("\\P", "\n").splitlines() or [""]
            wrapped_count = 0
            for ln in explicit_lines:
                line_w = _visual_text_len(ln) * row_txt_h
                wrapped_count += max(1, math.ceil(line_w / usable_w))
            max_lines = max(max_lines, wrapped_count)
        return (max_lines * row_txt_h * 1.55) + pad_y

    with autocad_session() as (_, doc):
        def _create():
            if layer:
                doc.Layers.Add(layer)

            default_rh = text_height * 2.6
            default_cw = sum(col_widths) / num_cols
            tbl = doc.ModelSpace.AddTable(
                pt3d(x, y, 0.0), total_rows, num_cols, float(default_rh), float(default_cw)
            )
            tbl.RegenerateTableSuppressed = True
            try:
                if not has_title and num_cols > 1:
                    with contextlib.suppress(Exception):
                        tbl.UnmergeCells(0, 0, 0, num_cols - 1)

                for c, cw in enumerate(col_widths):
                    tbl.SetColumnWidth(c, float(cw))

                if has_title:
                    title_h = text_height * 1.35
                    tbl.SetRowHeight(0, float(_calc_row_height([title], title_h)))
                    tbl.SetCellTextHeight(0, 0, float(title_h))
                    tbl.SetCellAlignment(0, 0, 5)
                    tbl.SetText(0, 0, title)

                hdr_h = text_height * 1.1
                tbl.SetRowHeight(header_row_idx, float(_calc_row_height(headers, hdr_h)))
                for c, h_text in enumerate(headers):
                    tbl.SetCellTextHeight(header_row_idx, c, float(hdr_h))
                    tbl.SetCellAlignment(header_row_idx, c, 5)
                    tbl.SetText(header_row_idx, c, str(h_text))

                for idx, r_data in enumerate(rows):
                    r_idx = data_start_idx + idx
                    tbl.SetRowHeight(r_idx, float(_calc_row_height(r_data, text_height)))
                    for c in range(num_cols):
                        val_str = str(r_data[c]) if c < len(r_data) else ""
                        tbl.SetCellTextHeight(r_idx, c, float(text_height))
                        clean_num = val_str.replace(",", "").replace(".", "").replace("-", "").strip()
                        align = 6 if (clean_num.isdigit() and clean_num != "") else 4
                        tbl.SetCellAlignment(r_idx, c, align)
                        tbl.SetText(r_idx, c, val_str)

                if layer:
                    tbl.Layer = layer
            finally:
                tbl.RegenerateTableSuppressed = False
                with contextlib.suppress(Exception):
                    tbl.RecomputeTableBlock(True)
                tbl.Update()

            placement = _clamp_to_zone(tbl, zone_bounds, zone_margin)
            return {
                "status": "created",
                "handle": tbl.Handle,
                "rows": total_rows,
                "columns": num_cols,
                "column_widths": [round(w, 2) for w in col_widths],
                **placement,
            }
        return com_retry(_create)

# ============================================================================
# Orthographic Projection & Hatching Tools (Plan -> Elevation / Section)
# ============================================================================

@mcp.tool()
def extract_plan_projection(
    line_x1: float,
    line_y1: float,
    line_x2: float,
    line_y2: float,
    look_direction: str = "N",
    max_depth: float = 10000.0,
    cut_tolerance: float = 25.0,
) -> Dict[str, Any]:
    """
    Project 2D Floor Plan geometry onto a Section Cut Line or Elevation Viewing Plane
    defined by (line_x1, line_y1) -> (line_x2, line_y2).
    - look_direction: "N" (+Y), "S" (-Y), "E" (+X), "W" (-X), or "LEFT"/"RIGHT" of vector p1->p2.
    Returns horizontal 'u' spans [0 .. total_width] strictly ordered from Viewer's LEFT to RIGHT,
    and 'depth' [0 .. max_depth] sorted front-to-back, flagging cut elements vs background elements.
    """
    dx_line, dy_line = line_x2 - line_x1, line_y2 - line_y1
    line_len = math.hypot(dx_line, dy_line)
    if line_len < 1e-3:
        raise ValueError("Cut/view line length must be greater than 0.")

    tx, ty = dx_line / line_len, dy_line / line_len
    ld = look_direction.strip().upper()

    if ld == "N":
        dx, dy = 0.0, 1.0
    elif ld == "S":
        dx, dy = 0.0, -1.0
    elif ld == "E":
        dx, dy = 1.0, 0.0
    elif ld == "W":
        dx, dy = -1.0, 0.0
    elif ld == "RIGHT":
        dx, dy = ty, -tx
    else:
        dx, dy = -ty, tx

    ux, uy = dy, -dx

    u1 = line_x1 * ux + line_y1 * uy
    u2 = line_x2 * ux + line_y2 * uy
    if u1 <= u2:
        ox, oy = line_x1, line_y1
        total_width = u2 - u1
    else:
        ox, oy = line_x2, line_y2
        total_width = u1 - u2

    def _to_uv(px: float, py: float):
        rx, ry = px - ox, py - oy
        return (rx * ux + ry * uy), (rx * dx + ry * dy)

    def _classify_entity(layer_name: str, obj_name: str, name_hint: str = "") -> str:
        key = f"{layer_name} {name_hint}".upper()
        if any(k in key for k in ("GRID", "AXIS", "CENTER")):
            return "grid"
        if any(k in key for k in ("COL", "PIER", "POST", "ST-C")):
            return "column"
        if any(k in key for k in ("WALL", "PART", "MUR", "A-WALL")):
            return "wall"
        if any(k in key for k in ("DOOR", "DR", "D-", "ประตู")):
            return "door"
        if any(k in key for k in ("WIN", "GLAZ", "W-", "หน้าต่าง")):
            return "window"
        if any(k in key for k in ("FURN", "SAN", "FIXT", "CAB", "WC", "SINK")):
            return "fixture_or_furniture"
        if "TEXT" in obj_name.upper():
            return "annotation"
        return "geometry"

    with autocad_session() as (_, doc):
        def _project():
            ms = doc.ModelSpace
            cut_items: List[Dict[str, Any]] = []
            beyond_items: List[Dict[str, Any]] = []
            grids_and_notes: List[Dict[str, Any]] = []

            for i in range(ms.Count):
                ent = ms.Item(i)
                try:
                    if not ent.Visible:
                        continue
                    bmin, bmax = ent.GetBoundingBox()
                except Exception:
                    continue

                obj_name = ent.ObjectName
                layer = ent.Layer
                pts = []
                name_hint = ""
                text_val = ""

                if obj_name == "AcDbLine":
                    sp, ep = ent.StartPoint, ent.EndPoint
                    pts = [(sp[0], sp[1]), (ep[0], ep[1])]
                elif obj_name in ("AcDbPolyline", "AcDb2dPolyline"):
                    coords = list(ent.Coordinates)
                    step = 2 if obj_name == "AcDbPolyline" else 3
                    pts = [(coords[k], coords[k + 1]) for k in range(0, len(coords) - 1, step)]
                elif obj_name == "AcDbBlockReference":
                    name_hint = getattr(ent, "EffectiveName", ent.Name)
                    pts = [
                        (bmin[0], bmin[1]), (bmax[0], bmin[1]),
                        (bmax[0], bmax[1]), (bmin[0], bmax[1]),
                    ]
                elif obj_name in ("AcDbText", "AcDbMText"):
                    text_val = ent.TextString
                    ip = ent.InsertionPoint
                    pts = [(ip[0], ip[1])]
                else:
                    pts = [
                        (bmin[0], bmin[1]), (bmax[0], bmin[1]),
                        (bmax[0], bmax[1]), (bmin[0], bmax[1]),
                    ]

                if not pts:
                    continue

                uv_pts = [_to_uv(px, py) for px, py in pts]
                u_vals = [p[0] for p in uv_pts]
                v_vals = [p[1] for p in uv_pts]

                u_min, u_max = min(u_vals), max(u_vals)
                v_min, v_max = min(v_vals), max(v_vals)

                if u_max < -cut_tolerance or u_min > total_width + cut_tolerance:
                    continue
                if v_max < -cut_tolerance or v_min > max_depth:
                    continue

                clamped_u1 = round(max(0.0, u_min), 2)
                clamped_u2 = round(min(total_width, u_max), 2)
                category = _classify_entity(layer, obj_name, name_hint)
                is_cut = (v_min <= cut_tolerance and v_max >= -cut_tolerance)

                record = {
                    "handle": ent.Handle,
                    "category": category,
                    "layer": layer,
                    "u_start": clamped_u1,
                    "u_end": clamped_u2,
                    "projected_width": round(clamped_u2 - clamped_u1, 2),
                    "depth_min": round(max(0.0, v_min), 2),
                    "depth_max": round(v_max, 2),
                }
                if name_hint:
                    record["block_name"] = name_hint
                if text_val:
                    record["text"] = text_val

                if category in ("grid", "annotation"):
                    grids_and_notes.append(record)
                elif is_cut:
                    cut_items.append(record)
                else:
                    beyond_items.append(record)

            cut_items.sort(key=lambda r: r["u_start"])
            beyond_items.sort(key=lambda r: (r["depth_min"], r["u_start"]))
            grids_and_notes.sort(key=lambda r: r["u_start"])

            return {
                "projection_width": round(total_width, 2),
                "look_direction": ld,
                "viewer_left_plan_xy": [round(ox, 2), round(oy, 2)],
                "horizontal_unit_vec": [round(ux, 4), round(uy, 4)],
                "depth_unit_vec": [round(dx, 4), round(dy, 4)],
                "cut_elements_count": len(cut_items),
                "cut_elements": cut_items,
                "beyond_elements_count": len(beyond_items),
                "beyond_elements": beyond_items[:150],
                "grids_and_notes": grids_and_notes[:80],
            }
        return com_retry(_project)


@mcp.tool()
def draw_hatch_closed(
    outer_handle: str,
    pattern_name: str = "ANSI31",
    pattern_scale: float = 20.0,
    pattern_angle_deg: float = 0.0,
    inner_handles: Optional[List[str]] = None,
    layer: str = "A-SECT-HTCH",
) -> Dict[str, Any]:
    """
    Create an associative Hatch inside a closed polyline (outer_handle) for section cuts
    (e.g., 'ANSI31', 'AR-CONC', 'SOLID') or elevation materials.
    Optionally pass inner_handles (e.g., window frames inside a wall) as island cutouts.
    """
    angle_rad = math.radians(pattern_angle_deg)

    with autocad_session() as (_, doc):
        def _hatch():
            if layer:
                doc.Layers.Add(layer)

            outer_ent = doc.HandleToObject(outer_handle)
            hatch = doc.ModelSpace.AddHatch(0, pattern_name, True)
            outer_loop = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [outer_ent]
            )
            hatch.AppendOuterLoop(outer_loop)

            for ih in (inner_handles or []):
                inner_ent = doc.HandleToObject(ih)
                inner_loop = win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [inner_ent]
                )
                hatch.AppendInnerLoop(inner_loop)

            if pattern_name.upper() != "SOLID":
                hatch.PatternScale = float(pattern_scale)
                hatch.PatternAngle = float(angle_rad)

            hatch.Evaluate()
            if layer:
                hatch.Layer = layer
            hatch.Update()

            return {
                "status": "created",
                "handle": hatch.Handle,
                "pattern": pattern_name,
                "area": round(float(hatch.Area), 2),
                "layer": hatch.Layer,
            }
        return com_retry(_hatch)

# ============================================================================
# Popup-Free Block, Attribute, Dynamic Block & Complex Hatch Tools
# ============================================================================

@mcp.tool()
def create_block_with_attributes(
    block_name: str,
    polylines: Optional[List[Dict[str, Any]]] = None,
    circles: Optional[List[Dict[str, Any]]] = None,
    attributes: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """
    Create or overwrite a Block Definition directly in AutoCAD's Block Table (NO popup dialog).
    All coordinates are local relative to block origin (0, 0).
    - polylines: [{"points": [[x1,y1], [x2,y2], ...], "closed": True, "layer": "0"}]
    - circles: [{"cx": 0, "cy": 0, "radius": 150, "layer": "0"}]
    - attributes: [{"tag": "DOOR_NO", "prompt": "Door Number", "default": "D1", "x": 0, "y": 0, "height": 150}]
    """
    with autocad_session() as (_, doc):
        ensure_not_in_block_editor(doc)

        def _build():
            blk = doc.Blocks.Add(pt3d(0.0, 0.0, 0.0), block_name)

            for i in range(blk.Count - 1, -1, -1):
                with contextlib.suppress(Exception):
                    blk.Item(i).Delete()

            created_count = 0
            for pl in (polylines or []):
                pts = pl.get("points", [])
                if len(pts) >= 2:
                    ent = blk.AddLightWeightPolyline(pt2d_flat(pts))
                    ent.Closed = bool(pl.get("closed", False))
                    lyr = pl.get("layer", "0")
                    if lyr != "0":
                        doc.Layers.Add(lyr)
                    ent.Layer = lyr
                    created_count += 1

            for cir in (circles or []):
                ent = blk.AddCircle(
                    pt3d(cir.get("cx", 0.0), cir.get("cy", 0.0), 0.0),
                    float(cir.get("radius", 100.0)),
                )
                lyr = cir.get("layer", "0")
                if lyr != "0":
                    doc.Layers.Add(lyr)
                ent.Layer = lyr
                created_count += 1

            att_tags = []
            for att in (attributes or []):
                clean_tag = str(att.get("tag", "TAG")).strip().replace(" ", "_").upper()
                prompt_str = str(att.get("prompt", clean_tag))
                default_val = str(att.get("default", ""))
                txt_h = float(att.get("height", 150.0))
                ax, ay = float(att.get("x", 0.0)), float(att.get("y", 0.0))

                att_def = blk.AddAttribute(
                    txt_h, 0, prompt_str, pt3d(ax, ay, 0.0), clean_tag, default_val
                )
                att_def.Alignment = 10
                att_def.TextAlignmentPoint = pt3d(ax, ay, 0.0)
                att_tags.append(clean_tag)

            doc.Regen(1)
            return {
                "status": "block_defined",
                "block_name": block_name,
                "geometry_count": created_count,
                "attribute_tags": att_tags,
            }
        return com_retry(_build)


@mcp.tool()
def inspect_block_reference(handle: str) -> Dict[str, Any]:
    """
    Inspect a placed Block or Dynamic Block by handle to read all editable Attributes
    and Dynamic Block Parameters (including AllowedValues for Visibility/Flip states).
    Always call this before modifying an unfamiliar Dynamic Block!
    """
    with autocad_session() as (_, doc):
        def _inspect():
            blk_ref = doc.HandleToObject(handle)
            if blk_ref.ObjectName != "AcDbBlockReference":
                raise ValueError(f"Handle {handle} is {blk_ref.ObjectName}, not a BlockReference.")

            attrs = {}
            if blk_ref.HasAttributes:
                for att in blk_ref.GetAttributes():
                    attrs[att.TagString] = att.TextString

            dyn_props = []
            is_dyn = bool(getattr(blk_ref, "IsDynamicBlock", False))
            if is_dyn:
                for prop in blk_ref.GetDynamicBlockProperties():
                    if prop.PropertyName == "Origin":
                        continue
                    dyn_props.append({
                        "name": prop.PropertyName,
                        "current_value": prop.Value,
                        "read_only": bool(prop.ReadOnly),
                        "allowed_values": list(getattr(prop, "AllowedValues", []) or []),
                        "description": getattr(prop, "Description", ""),
                    })

            return {
                "handle": handle,
                "effective_name": getattr(blk_ref, "EffectiveName", blk_ref.Name),
                "is_dynamic_block": is_dyn,
                "attributes": attrs,
                "dynamic_properties": dyn_props,
                "bbox": _read_bbox(blk_ref),
            }
        return com_retry(_inspect)


@mcp.tool()
def insert_or_configure_block(
    block_name: str = "",
    x: float = 0.0,
    y: float = 0.0,
    scale: float = 1.0,
    rotation_deg: float = 0.0,
    layer: str = "",
    existing_handle: str = "",
    attribute_values: Optional[Dict[str, str]] = None,
    dynamic_properties: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Insert a new Block/Dynamic Block (or update an existing one via existing_handle)
    and set Attribute text values and Dynamic Block parameters (Stretch width, Flip 0/1,
    Visibility state) in a single popup-free operation.
    """
    rot_rad = math.radians(rotation_deg)

    with autocad_session() as (_, doc):
        ensure_not_in_block_editor(doc)
        with suppress_autocad_dialogs(doc):
            def _apply():
                if existing_handle:
                    blk_ref = doc.HandleToObject(existing_handle)
                else:
                    if not block_name:
                        raise ValueError("Provide either 'block_name' to insert or 'existing_handle' to update.")
                    if layer:
                        doc.Layers.Add(layer)
                    blk_ref = doc.ModelSpace.InsertBlock(
                        pt3d(x, y, 0.0),
                        block_name,
                        float(scale),
                        float(scale),
                        float(scale),
                        float(rot_rad),
                    )
                    if layer:
                        blk_ref.Layer = layer

                updated_dyn = {}
                if dynamic_properties and getattr(blk_ref, "IsDynamicBlock", False):
                    target_map = {k.strip().lower(): v for k, v in dynamic_properties.items()}
                    for prop in blk_ref.GetDynamicBlockProperties():
                        p_key = prop.PropertyName.strip().lower()
                        if p_key in target_map and not prop.ReadOnly:
                            coerced = _coerce_dyn_variant(prop, target_map[p_key])
                            prop.Value = coerced
                            updated_dyn[prop.PropertyName] = target_map[p_key]

                updated_atts = {}
                if attribute_values and blk_ref.HasAttributes:
                    att_map = {k.strip().upper(): str(v) for k, v in attribute_values.items()}
                    for att in blk_ref.GetAttributes():
                        tag_upper = att.TagString.strip().upper()
                        if tag_upper in att_map:
                            att.TextString = att_map[tag_upper]
                            att.Update()
                            updated_atts[att.TagString] = att_map[tag_upper]

                blk_ref.Update()
                return {
                    "status": "updated" if existing_handle else "inserted",
                    "handle": blk_ref.Handle,
                    "effective_name": getattr(blk_ref, "EffectiveName", blk_ref.Name),
                    "updated_attributes": updated_atts,
                    "updated_dynamic_properties": updated_dyn,
                    "bbox": _read_bbox(blk_ref),
                }
            return com_retry(_apply)


@mcp.tool()
def edit_block_in_place(
    block_name: str,
    add_polylines: Optional[List[Dict[str, Any]]] = None,
    scale_all_factor: float = 1.0,
    recolor_bylayer: bool = False,
    sync_attributes: bool = False,
) -> Dict[str, Any]:
    """
    Modify an existing Block Definition directly inside doc.Blocks WITHOUT opening BEDIT.
    Updates every placed instance across the entire drawing automatically!
    """
    with autocad_session() as (_, doc):
        ensure_not_in_block_editor(doc)
        with suppress_autocad_dialogs(doc):
            def _edit():
                blk = doc.Blocks.Item(block_name)
                modified_items = 0

                if abs(scale_all_factor - 1.0) > 1e-4 or recolor_bylayer:
                    origin = pt3d(0.0, 0.0, 0.0)
                    for i in range(blk.Count):
                        ent = blk.Item(i)
                        if abs(scale_all_factor - 1.0) > 1e-4:
                            ent.ScaleEntity(origin, float(scale_all_factor))
                        if recolor_bylayer:
                            ent.color = 256
                        ent.Update()
                        modified_items += 1

                added_items = 0
                for pl in (add_polylines or []):
                    pts = pl.get("points", [])
                    if len(pts) >= 2:
                        ent = blk.AddLightWeightPolyline(pt2d_flat(pts))
                        ent.Closed = bool(pl.get("closed", False))
                        ent.Layer = pl.get("layer", "0")
                        added_items += 1

                if sync_attributes:
                    doc.SendCommand(f"._ATTSYNC\n_Name\n{block_name}\n")

                doc.Regen(1)
                return {
                    "status": "block_edited_silently",
                    "block_name": block_name,
                    "modified_existing_entities": modified_items,
                    "added_entities": added_items,
                }
            return com_retry(_edit)


@mcp.tool()
def hatch_by_internal_point(
    pick_x: float,
    pick_y: float,
    pattern_name: str = "ANSI31",
    pattern_scale: float = 20.0,
    pattern_angle_deg: float = 0.0,
    layer: str = "C-REPR-HTCH",
) -> Dict[str, Any]:
    """
    Create a Hatch by picking an interior point (pick_x, pick_y) surrounded by intersecting
    lines/walls (no single closed polyline required), suppressing the Hatch dialog completely.
    """
    with autocad_session() as (_, doc):
        ensure_not_in_block_editor(doc)
        with suppress_autocad_dialogs(doc):
            def _pick_hatch():
                if layer:
                    doc.Layers.Add(layer)
                old_clayer = doc.GetVariable("CLAYER")
                old_hpname = doc.GetVariable("HPNAME")
                try:
                    if layer:
                        doc.SetVariable("CLAYER", layer)
                    before_cnt = doc.ModelSpace.Count

                    if pattern_name.upper() == "SOLID":
                        macro = f"._-BHATCH\n_Properties\nSOLID\n{pick_x},{pick_y}\n\n"
                    else:
                        macro = (
                            f"._-BHATCH\n_Properties\n{pattern_name}\n"
                            f"{pattern_scale}\n{pattern_angle_deg}\n"
                            f"{pick_x},{pick_y}\n\n"
                        )
                    doc.SendCommand(macro)
                    time.sleep(0.2)

                    after_cnt = doc.ModelSpace.Count
                    if after_cnt > before_cnt:
                        new_ent = doc.ModelSpace.Item(after_cnt - 1)
                        if new_ent.ObjectName == "AcDbHatch":
                            return {
                                "status": "created",
                                "handle": new_ent.Handle,
                                "area": round(float(new_ent.Area), 2),
                                "layer": new_ent.Layer,
                            }
                    raise RuntimeError(
                        f"Boundary around ({pick_x}, {pick_y}) is not closed or out of view. "
                        "Zoom onto the target zone or use 'draw_hatch_closed' with a closed polyline."
                    )
                finally:
                    with contextlib.suppress(Exception):
                        doc.SetVariable("CLAYER", old_clayer)
                        doc.SetVariable("HPNAME", old_hpname)
            return com_retry(_pick_hatch)

# ============================================================================
# Drafting Scale & Linetype Helpers
# ============================================================================

def _unit_mm_factor(doc) -> float:
    """Return drawing units per 1 mm on paper (1.0 for mm drawings, 0.001 for meter drawings)."""
    insunits = int(com_retry(lambda: doc.GetVariable("INSUNITS")))
    if insunits == 6:   # Meters
        return 0.001
    if insunits == 5:   # Centimeters
        return 0.1
    return 1.0          # Default Millimeters (insunits == 4 or unitless metric)


def _ensure_linetype(doc, linetype_name: str) -> str:
    """Ensure standard linetype (CENTER, DASHED, PHANTOM) is loaded without popup errors."""
    target = linetype_name.upper()
    if target in ("CONTINUOUS", "BYLAYER", "BYBLOCK"):
        return linetype_name

    for i in range(doc.Linetypes.Count):
        if doc.Linetypes.Item(i).Name.upper() == target:
            return doc.Linetypes.Item(i).Name

    for lin_file in ("acadiso.lin", "acad.lin"):
        with contextlib.suppress(Exception):
            doc.Linetypes.Load(target, lin_file)
            return target
    return "Continuous"


def _style_arch_dimension(dim, txt_h: float, precision: int = 2, scale_factor_override: float = 0.0):
    """Apply clean Architectural Tick (45-deg oblique) & plot-scale sizing to a Dimension."""
    dim.TextHeight = float(txt_h)
    dim.ArrowheadSize = float(txt_h * 0.75)
    with contextlib.suppress(Exception):
        # 4 = acArrowArchTick (45-degree architectural slash)
        dim.Arrowhead1Type = 4
        dim.Arrowhead2Type = 4
        dim.ExtensionLineOffset = float(txt_h * 0.6)
        dim.ExtensionLineExtend = float(txt_h * 0.5)
        dim.TextGap = float(txt_h * 0.35)
        dim.UnitsFormat = 2  # Decimal
        dim.PrimaryUnitsPrecision = int(precision)
        if scale_factor_override > 0:
            dim.LinearScaleFactor = float(scale_factor_override)


# ============================================================================
# 1. Elevation / Section Level Datum Marker (▽ FFL +0.15 m.)
# ============================================================================

@mcp.tool()
def draw_level_datum(
    x: float,
    y: float,
    level_value_m: float,
    prefix_label: str = "FFL",
    side: str = "right",
    target_x: Optional[float] = None,
    scale_factor: float = 100.0,
    layer: str = "G-ANNO-LEVL",
) -> Dict[str, Any]:
    """
    Draw a standard architectural Level Datum symbol (inverted triangle ▽ + horizontal shelf
    + optional dashed witness line from target_x + formatted elevation text like 'FFL +0.15 m.').
    - (x, y): Exact location of the ▽ bottom tip (y must match the true elevation Y coordinate).
    - side: 'right' extends text shelf to +X; 'left' extends text shelf to -X.
    """
    with autocad_session() as (_, doc):
        def _draw():
            u_mm = _unit_mm_factor(doc)
            s_unit = scale_factor * u_mm
            txt_h = 2.5 * s_unit
            tri_h = 3.2 * s_unit
            tri_half_w = 1.85 * s_unit
            dir_sign = 1.0 if side.strip().lower() != "left" else -1.0

            doc.Layers.Add(layer)
            handles = []

            # Format elevation string (+0.15 m., ±0.00 m., -0.45 m.)
            if abs(level_value_m) < 1e-4:
                num_str = "±0.00 m."
            elif level_value_m > 0:
                num_str = f"+{level_value_m:.2f} m."
            else:
                num_str = f"{level_value_m:.2f} m."

            full_label = f"{prefix_label.strip()} {num_str}".strip()
            shelf_len = max(18.0 * s_unit, (_visual_text_len(full_label) + 1.5) * txt_h)

            # 1. Optional dashed projection line from building element to datum tip
            if target_x is not None and abs(target_x - x) > tri_half_w:
                lt_dash = _ensure_linetype(doc, "DASHED")
                ext_line = doc.ModelSpace.AddLine(pt3d(target_x, y, 0.0), pt3d(x, y, 0.0))
                ext_line.Layer = layer
                ext_line.Linetype = lt_dash
                ext_line.color = 8  # Subtle gray witness line
                handles.append(ext_line.Handle)

            # 2. Inverted triangle ▽ touching (x, y)
            tri_pts = [
                [x, y],
                [x - tri_half_w, y + tri_h],
                [x + tri_half_w, y + tri_h],
            ]
            tri = doc.ModelSpace.AddLightWeightPolyline(pt2d_flat(tri_pts))
            tri.Closed = True
            tri.Layer = layer
            handles.append(tri.Handle)

            # Classic architectural half-shaded triangle for crisp visual contrast
            half_pts = [
                [x, y],
                [x, y + tri_h],
                [x - (tri_half_w * dir_sign), y + tri_h],
            ]
            half_tri = doc.ModelSpace.AddLightWeightPolyline(pt2d_flat(half_pts))
            half_tri.Closed = True
            half_tri.Layer = layer
            handles.append(half_tri.Handle)
            with contextlib.suppress(Exception):
                h_solid = doc.ModelSpace.AddHatch(0, "SOLID", True)
                h_solid.AppendOuterLoop(
                    win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [half_tri])
                )
                h_solid.Evaluate()
                h_solid.Layer = layer
                handles.append(h_solid.Handle)

            # 3. Horizontal baseline & riser shelf
            base_tick = doc.ModelSpace.AddLine(
                pt3d(x - 2.5 * s_unit, y, 0.0),
                pt3d(x + 2.5 * s_unit, y, 0.0),
            )
            base_tick.Layer = layer
            handles.append(base_tick.Handle)

            shelf_end_x = x + dir_sign * (tri_half_w + shelf_len)
            shelf = doc.ModelSpace.AddLightWeightPolyline(
                pt2d_flat([[x, y + tri_h], [shelf_end_x, y + tri_h]])
            )
            shelf.Layer = layer
            handles.append(shelf.Handle)

            # 4. Datum Label above shelf
            txt_x = x + dir_sign * (tri_half_w + 1.0 * s_unit)
            txt_y = y + tri_h + (0.8 * s_unit)
            mt = doc.ModelSpace.AddMText(pt3d(txt_x, txt_y, 0.0), float(shelf_len), full_label)
            # 7 = BottomLeft, 9 = BottomRight
            mt.AttachmentPoint = 7 if dir_sign > 0 else 9
            mt.InsertionPoint = pt3d(txt_x, txt_y, 0.0)
            mt.Height = float(txt_h)
            mt.Layer = layer
            mt.Update()
            handles.append(mt.Handle)

            return {
                "status": "created",
                "label": full_label,
                "handles": handles,
                "tip_xy": [round(x, 2), round(y, 2)],
            }
        return com_retry(_draw)


# ============================================================================
# 2. Architectural Aligned & Chained Dimensions (10*S / +8*S Tier Math)
# ============================================================================

@mcp.tool()
def draw_aligned_dim(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    offset_distance: float = 0.0,
    orientation: str = "aligned",
    scale_factor: float = 100.0,
    precision: int = 0,
    layer: str = "G-ANNO-DIMS",
) -> Dict[str, Any]:
    """
    Draw a single Architectural Dimension with 45-degree ArchTicks.
    - orientation: 'vertical' (90 deg for Elevation/Section heights), 'horizontal' (0 deg), or 'aligned'.
    - offset_distance: Signed offset from measured points (defaults to 10 * S if 0).
    """
    with autocad_session() as (_, doc):
        def _dim():
            u_mm = _unit_mm_factor(doc)
            s_unit = scale_factor * u_mm
            txt_h = 2.5 * s_unit
            dist = offset_distance if abs(offset_distance) > 1e-4 else (10.0 * s_unit)

            doc.Layers.Add(layer)
            p1 = pt3d(x1, y1, 0.0)
            p2 = pt3d(x2, y2, 0.0)
            mode = orientation.strip().lower()

            if mode == "vertical":
                ref_x = max(x1, x2) if dist >= 0 else min(x1, x2)
                dim_loc = pt3d(ref_x + dist, (y1 + y2) / 2.0, 0.0)
                dim = doc.ModelSpace.AddDimRotated(p1, p2, dim_loc, math.pi / 2.0)
            elif mode == "horizontal":
                ref_y = max(y1, y2) if dist >= 0 else min(y1, y2)
                dim_loc = pt3d((x1 + x2) / 2.0, ref_y + dist, 0.0)
                dim = doc.ModelSpace.AddDimRotated(p1, p2, dim_loc, 0.0)
            else:
                dx, dy = x2 - x1, y2 - y1
                seg_len = math.hypot(dx, dy)
                if seg_len < 1e-4:
                    raise ValueError("Dimension points cannot be identical.")
                nx, ny = -dy / seg_len, dx / seg_len
                dim_loc = pt3d((x1 + x2) / 2.0 + nx * dist, (y1 + y2) / 2.0 + ny * dist, 0.0)
                dim = doc.ModelSpace.AddDimAligned(p1, p2, dim_loc)

            _style_arch_dimension(dim, txt_h, precision=precision)
            dim.Layer = layer
            dim.Update()
            return {
                "status": "created",
                "handle": dim.Handle,
                "measurement": round(float(dim.Measurement), 3),
                "bbox": _read_bbox(dim),
            }
        return com_retry(_dim)


@mcp.tool()
def draw_dim_chain(
    coords: List[float],
    anchor_line_coord: float,
    side: str = "left",
    tier: int = 1,
    include_overall: bool = True,
    scale_factor: float = 100.0,
    precision: int = 0,
    layer: str = "G-ANNO-DIMS",
) -> Dict[str, Any]:
    """
    Draw a continuous Architectural Dimension String (e.g. GL -> FFL -> Sill -> Head -> CL
    or Grid-to-Grid bays) strictly following Drafting Math:
    Tier 1 = 10*S from anchor_line_coord, each outer tier adds +8*S!
    - side: 'left' or 'right' (creates vertical string); 'top' or 'bottom' (creates horizontal string).
    """
    pts_sorted = sorted(set(float(v) for v in coords))
    if len(pts_sorted) < 2:
        raise ValueError("Provide at least 2 distinct coordinates in 'coords'.")

    with autocad_session() as (_, doc):
        def _chain():
            u_mm = _unit_mm_factor(doc)
            s_unit = scale_factor * u_mm
            txt_h = 2.5 * s_unit
            tier_offset = (10.0 + max(0, tier - 1) * 8.0) * s_unit
            overall_offset = tier_offset + (8.0 * s_unit)

            doc.Layers.Add(layer)
            side_l = side.strip().lower()
            is_vert = side_l in ("left", "right")
            sign = -1.0 if side_l in ("left", "bottom") else 1.0

            created_handles = []
            dim_line_pos = anchor_line_coord + (sign * tier_offset)

            for i in range(len(pts_sorted) - 1):
                c1, c2 = pts_sorted[i], pts_sorted[i + 1]
                if is_vert:
                    p1 = pt3d(anchor_line_coord, c1, 0.0)
                    p2 = pt3d(anchor_line_coord, c2, 0.0)
                    dloc = pt3d(dim_line_pos, (c1 + c2) / 2.0, 0.0)
                    dim = doc.ModelSpace.AddDimRotated(p1, p2, dloc, math.pi / 2.0)
                else:
                    p1 = pt3d(c1, anchor_line_coord, 0.0)
                    p2 = pt3d(c2, anchor_line_coord, 0.0)
                    dloc = pt3d((c1 + c2) / 2.0, dim_line_pos, 0.0)
                    dim = doc.ModelSpace.AddDimRotated(p1, p2, dloc, 0.0)

                _style_arch_dimension(dim, txt_h, precision=precision)
                dim.Layer = layer
                dim.Update()
                created_handles.append(dim.Handle)

            # Optional Overall Dimension on next outer tier (+8*S)
            if include_overall and len(pts_sorted) > 2:
                ov_pos = anchor_line_coord + (sign * overall_offset)
                c_first, c_last = pts_sorted[0], pts_sorted[-1]
                if is_vert:
                    p1 = pt3d(anchor_line_coord, c_first, 0.0)
                    p2 = pt3d(anchor_line_coord, c_last, 0.0)
                    dloc = pt3d(ov_pos, (c_first + c_last) / 2.0, 0.0)
                    ov_dim = doc.ModelSpace.AddDimRotated(p1, p2, dloc, math.pi / 2.0)
                else:
                    p1 = pt3d(c_first, anchor_line_coord, 0.0)
                    p2 = pt3d(c_last, anchor_line_coord, 0.0)
                    dloc = pt3d((c_first + c_last) / 2.0, ov_pos, 0.0)
                    ov_dim = doc.ModelSpace.AddDimRotated(p1, p2, dloc, 0.0)

                _style_arch_dimension(ov_dim, txt_h, precision=precision)
                ov_dim.Layer = layer
                ov_dim.Update()
                created_handles.append(ov_dim.Handle)

            return {
                "status": "created_chain",
                "segment_count": len(pts_sorted) - 1,
                "included_overall": bool(include_overall and len(pts_sorted) > 2),
                "outermost_line_coord": round(
                    anchor_line_coord + sign * (overall_offset if (include_overall and len(pts_sorted) > 2) else tier_offset),
                    2,
                ),
                "handles": created_handles,
            }
        return com_retry(_chain)


# ============================================================================
# 3. Structural Grid Line & Grid Number Balloon
# ============================================================================

@mcp.tool()
def draw_grid_balloon(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    grid_label: str,
    bubble_at: str = "end",
    scale_factor: float = 100.0,
    grid_layer: str = "S-GRID",
    bubble_layer: str = "G-ANNO-SYMB",
) -> Dict[str, Any]:
    """
    Draw a CENTER-linetype Grid Axis from (x1, y1) to (x2, y2) with circular Grid Balloon(s)
    (radius = 5.0 * S, text = 4.0 * S) placed cleanly at 'end' (x2,y2), 'start' (x1,y1), or 'both'.
    """
    dx, dy = x2 - x1, y2 - y1
    seg_len = math.hypot(dx, dy)
    if seg_len < 1e-4:
        raise ValueError("Grid line start and end points must differ.")
    ux, uy = dx / seg_len, dy / seg_len

    with autocad_session() as (_, doc):
        def _grid():
            u_mm = _unit_mm_factor(doc)
            s_unit = scale_factor * u_mm
            radius = 5.0 * s_unit
            txt_h = 3.8 * s_unit

            doc.Layers.Add(grid_layer)
            doc.Layers.Add(bubble_layer)
            lt_center = _ensure_linetype(doc, "CENTER")

            handles = []
            axis_line = doc.ModelSpace.AddLine(pt3d(x1, y1, 0.0), pt3d(x2, y2, 0.0))
            axis_line.Layer = grid_layer
            axis_line.Linetype = lt_center
            axis_line.LinetypeScale = float(scale_factor * 0.5)
            handles.append(axis_line.Handle)

            ends = []
            loc = bubble_at.strip().lower()
            if loc in ("start", "both"):
                ends.append((x1 - ux * radius, y1 - uy * radius))
            if loc in ("end", "both"):
                ends.append((x2 + ux * radius, y2 + uy * radius))

            for cx, cy in ends:
                cir = doc.ModelSpace.AddCircle(pt3d(cx, cy, 0.0), float(radius))
                cir.Layer = bubble_layer
                handles.append(cir.Handle)

                mt = doc.ModelSpace.AddMText(pt3d(cx, cy, 0.0), float(radius * 1.8), grid_label.strip())
                mt.AttachmentPoint = 5  # MiddleCenter
                mt.InsertionPoint = pt3d(cx, cy, 0.0)
                mt.Height = float(txt_h)
                mt.Layer = bubble_layer
                mt.Update()
                handles.append(mt.Handle)

            return {"status": "created", "grid_label": grid_label, "handles": handles}
        return com_retry(_grid)


# ============================================================================
# 4. Architectural Section Cut Mark, Interior 4-Way Elevation & Enlarged Detail
# ============================================================================

@mcp.tool()
def draw_section_mark(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    section_id: str = "A",
    sheet_no: str = "A-301",
    look_direction: str = "N",
    scale_factor: float = 100.0,
    layer: str = "G-ANNO-SYMB",
) -> Dict[str, Any]:
    """
    Draw an Architectural Section Cut Marker at (x1, y1) -> (x2, y2) with split callout bubble
    (section_id / sheet_no) at start, thick tail fin at end, and filled directional pointer chevron.
    """
    dx_l, dy_l = x2 - x1, y2 - y1
    llen = math.hypot(dx_l, dy_l)
    if llen < 1e-4:
        raise ValueError("Section line length must be > 0.")
    tx, ty = dx_l / llen, dy_l / llen

    ld = look_direction.strip().upper()
    if ld == "N":
        vx, vy = 0.0, 1.0
    elif ld == "S":
        vx, vy = 0.0, -1.0
    elif ld == "E":
        vx, vy = 1.0, 0.0
    elif ld == "W":
        vx, vy = -1.0, 0.0
    elif ld == "RIGHT":
        vx, vy = ty, -tx
    else:
        vx, vy = -ty, tx

    with autocad_session() as (_, doc):
        def _sec():
            u_mm = _unit_mm_factor(doc)
            s_unit = scale_factor * u_mm
            r = 6.0 * s_unit
            txt_h = 3.0 * s_unit

            doc.Layers.Add(layer)
            handles = []

            # Bubble center just outside (x1, y1)
            bx, by = x1 - tx * r, y1 - ty * r
            cir = doc.ModelSpace.AddCircle(pt3d(bx, by, 0.0), float(r))
            cir.Layer = layer
            handles.append(cir.Handle)

            div = doc.ModelSpace.AddLine(pt3d(bx - r, by, 0.0), pt3d(bx + r, by, 0.0))
            div.Layer = layer
            handles.append(div.Handle)

            # Directional pointer hood wrapping bubble in look_direction
            px, py = -vy, vx
            apex_x, apex_y = bx + vx * (r * 1.65), by + vy * (r * 1.65)
            hood = doc.ModelSpace.AddLightWeightPolyline(
                pt2d_flat([
                    [bx - px * r, by - py * r],
                    [apex_x, apex_y],
                    [bx + px * r, by + py * r],
                ])
            )
            hood.Closed = False
            hood.Layer = layer
            handles.append(hood.Handle)

            # Top ID & Bottom Sheet text
            for val, cy_off in ((section_id, r * 0.45), (sheet_no, -r * 0.45)):
                mt = doc.ModelSpace.AddMText(pt3d(bx, by + cy_off, 0.0), float(r * 1.8), str(val))
                mt.AttachmentPoint = 5
                mt.InsertionPoint = pt3d(bx, by + cy_off, 0.0)
                mt.Height = float(txt_h * (1.1 if cy_off > 0 else 0.85))
                mt.Layer = layer
                mt.Update()
                handles.append(mt.Handle)

            # Gapped cut ticks (avoids slicing across interior room drawings) + tail flag at (x2, y2)
            seg_len = min(12.0 * s_unit, llen * 0.22)
            t1 = doc.ModelSpace.AddLine(pt3d(x1, y1, 0.0), pt3d(x1 + tx * seg_len, y1 + ty * seg_len, 0.0))
            t2 = doc.ModelSpace.AddLightWeightPolyline(
                pt2d_flat([
                    [x2 - tx * seg_len, y2 - ty * seg_len],
                    [x2, y2],
                    [x2 + vx * (r * 1.1), y2 + vy * (r * 1.1)],
                ])
            )
            t1.Layer = layer
            t2.Layer = layer
            handles.extend([t1.Handle, t2.Handle])

            return {"status": "created", "section_id": section_id, "handles": handles}
        return com_retry(_sec)


@mcp.tool()
def draw_interior_elevation_mark(
    cx: float,
    cy: float,
    sheet_no: str = "ID-201",
    active_views: Optional[Dict[str, str]] = None,
    scale_factor: float = 50.0,
    layer: str = "I-ANNO-SYMB",
) -> Dict[str, Any]:
    """
    Draw a 4-way Interior Elevation Symbol at room center (cx, cy).
    - active_views: Mapping of active compass directions to view tags,
      e.g. {"N": "A", "E": "B", "S": "C", "W": "D"} or {"N": "1", "W": "2"}.
    """
    views = active_views or {"N": "A", "E": "B", "S": "C", "W": "D"}
    dir_vecs = {"N": (0.0, 1.0), "E": (1.0, 0.0), "S": (0.0, -1.0), "W": (-1.0, 0.0)}

    with autocad_session() as (_, doc):
        def _int_mark():
            u_mm = _unit_mm_factor(doc)
            s_unit = scale_factor * u_mm
            r = 5.0 * s_unit
            txt_h = 2.3 * s_unit

            doc.Layers.Add(layer)
            handles = []

            cir = doc.ModelSpace.AddCircle(pt3d(cx, cy, 0.0), float(r))
            cir.Layer = layer
            handles.append(cir.Handle)

            center_mt = doc.ModelSpace.AddMText(pt3d(cx, cy, 0.0), float(r * 1.8), sheet_no)
            center_mt.AttachmentPoint = 5
            center_mt.InsertionPoint = pt3d(cx, cy, 0.0)
            center_mt.Height = float(txt_h * 0.85)
            center_mt.Layer = layer
            center_mt.Update()
            handles.append(center_mt.Handle)

            for d_key, tag in views.items():
                dk = d_key.strip().upper()
                if dk not in dir_vecs:
                    continue
                vx, vy = dir_vecs[dk]
                px, py = -vy, vx

                # Pointer triangle outside main circle
                p_tip = [cx + vx * (r * 1.85), cy + vy * (r * 1.85)]
                p_left = [cx + vx * (r * 1.02) - px * (r * 0.72), cy + vy * (r * 1.02) - py * (r * 0.72)]
                p_right = [cx + vx * (r * 1.02) + px * (r * 0.72), cy + vy * (r * 1.02) + py * (r * 0.72)]

                ptr = doc.ModelSpace.AddLightWeightPolyline(pt2d_flat([p_left, p_tip, p_right]))
                ptr.Closed = True
                ptr.Layer = layer
                handles.append(ptr.Handle)

                with contextlib.suppress(Exception):
                    h_fill = doc.ModelSpace.AddHatch(0, "SOLID", True)
                    h_fill.AppendOuterLoop(
                        win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [ptr])
                    )
                    h_fill.Evaluate()
                    h_fill.Layer = layer
                    handles.append(h_fill.Handle)

                # Place view letter/number cleanly just beyond pointer tip
                tx_pos = cx + vx * (r * 2.55)
                ty_pos = cy + vy * (r * 2.55)
                v_mt = doc.ModelSpace.AddMText(pt3d(tx_pos, ty_pos, 0.0), float(r), str(tag))
                v_mt.AttachmentPoint = 5
                v_mt.InsertionPoint = pt3d(tx_pos, ty_pos, 0.0)
                v_mt.Height = float(txt_h)
                v_mt.Layer = layer
                v_mt.Update()
                handles.append(v_mt.Handle)

            return {"status": "created", "active_views": views, "handles": handles}
        return com_retry(_int_mark)


@mcp.tool()
def draw_enlarged_detail_callout(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    bubble_x: float,
    bubble_y: float,
    detail_no: str = "DT-01",
    sheet_no: str = "A-501",
    shape: str = "rounded_box",
    scale_factor: float = 100.0,
    layer: str = "G-ANNO-SYMB",
) -> Dict[str, Any]:
    """
    Draw a dashed Enlarged Detail Callout boundary ('circle' or 'rounded_box' over [x1,y1..x2,y2])
    with a shouldered leader connecting to a split Detail Reference Bubble (detail_no / sheet_no).
    """
    bx1, bx2 = min(x1, x2), max(x1, x2)
    by1, by2 = min(y1, y2), max(y1, y2)
    cx, cy = (bx1 + bx2) / 2.0, (by1 + by2) / 2.0

    with autocad_session() as (_, doc):
        def _callout():
            u_mm = _unit_mm_factor(doc)
            s_unit = scale_factor * u_mm
            r_bub = 6.0 * s_unit
            txt_h = 2.8 * s_unit

            doc.Layers.Add(layer)
            lt_dash = _ensure_linetype(doc, "DASHED")
            handles = []

            # 1. Boundary Geometry
            if shape.strip().lower() == "circle":
                rad = max(math.hypot(bx2 - bx1, by2 - by1) / 2.0, 5.0 * s_unit)
                bnd = doc.ModelSpace.AddCircle(pt3d(cx, cy, 0.0), float(rad))
                vx, vy = bubble_x - cx, bubble_y - cy
                vlen = max(math.hypot(vx, vy), 1e-4)
                anchor_x, anchor_y = cx + (vx / vlen) * rad, cy + (vy / vlen) * rad
            else:
                # Chamfered/rounded detail box
                ch = min((bx2 - bx1) * 0.12, (by2 - by1) * 0.12, 4.0 * s_unit)
                box_pts = [
                    [bx1 + ch, by1], [bx2 - ch, by1], [bx2, by1 + ch], [bx2, by2 - ch],
                    [bx2 - ch, by2], [bx1 + ch, by2], [bx1, by2 - ch], [bx1, by1 + ch],
                ]
                bnd = doc.ModelSpace.AddLightWeightPolyline(pt2d_flat([box_pts[0], box_pts[1], box_pts[2], box_pts[3], box_pts[4], box_pts[5], box_pts[6], box_pts[7]]))
                bnd.Closed = True
                anchor_x = bx2 if bubble_x >= cx else bx1
                anchor_y = by2 if bubble_y >= cy else by1

            bnd.Layer = layer
            bnd.Linetype = lt_dash
            bnd.LinetypeScale = float(scale_factor * 0.4)
            handles.append(bnd.Handle)

            # 2. Leader with horizontal landing shoulder to Bubble edge
            dir_sign = 1.0 if bubble_x >= anchor_x else -1.0
            neck_x = bubble_x - dir_sign * (r_bub + 4.0 * s_unit)
            edge_x = bubble_x - dir_sign * r_bub
            ldr = doc.ModelSpace.AddLightWeightPolyline(
                pt2d_flat([[anchor_x, anchor_y], [neck_x, bubble_y], [edge_x, bubble_y]])
            )
            ldr.Layer = layer
            handles.append(ldr.Handle)

            # 3. Detail Reference Bubble
            cir = doc.ModelSpace.AddCircle(pt3d(bubble_x, bubble_y, 0.0), float(r_bub))
            div = doc.ModelSpace.AddLine(
                pt3d(bubble_x - r_bub, bubble_y, 0.0),
                pt3d(bubble_x + r_bub, bubble_y, 0.0),
            )
            cir.Layer = layer
            div.Layer = layer
            handles.extend([cir.Handle, div.Handle])

            for val, off_y in ((detail_no, r_bub * 0.45), (sheet_no, -r_bub * 0.45)):
                mt = doc.ModelSpace.AddMText(pt3d(bubble_x, bubble_y + off_y, 0.0), float(r_bub * 1.8), str(val))
                mt.AttachmentPoint = 5
                mt.InsertionPoint = pt3d(bubble_x, bubble_y + off_y, 0.0)
                mt.Height = float(txt_h * (1.05 if off_y > 0 else 0.85))
                mt.Layer = layer
                mt.Update()
                handles.append(mt.Handle)

            return {"status": "created", "detail_no": detail_no, "handles": handles}
        return com_retry(_callout)


# ============================================================================
# 5. PaperSpace Layout & Enlarged-Plan Viewports (1000 units = 1 m, paper in mm)
# ============================================================================

_STD_SCALES = [10, 20, 25, 30, 40, 50, 75, 100, 125, 150, 200, 250, 500]


def _detect_safe_zone_mm(doc, layout) -> Dict[str, Any]:
    """
    Detect ISO sheet size and the safe drawable box on a PaperSpace layout.
    Calibrated from the IX ONE ONE A3 template (right-hand title strip), then refined by
    scanning PaperSpace entities that sit inside the right-hand strip.
    """
    paper_w, paper_h = layout.GetPaperSize()
    if int(layout.PlotRotation) in (1, 3):
        paper_w, paper_h = paper_h, paper_w

    longer = max(paper_w, paper_h)
    iso = "A4" if longer < 330 else "A3" if longer < 500 else "A2" if longer < 700 else "A1"

    k = paper_w / 420.0

    # Sheet origin: this template does NOT start at paper (0,0). The title-block xref is
    # inserted at the sheet's lower-left corner (e.g. 1077.625, 8863.564), so every paper
    # coordinate must be offset by it or the viewport lands off the sheet.
    ox, oy, origin_src = 0.0, 0.0, "paper_zero"
    ps = doc.PaperSpace
    for i in range(ps.Count):
        ent = ps.Item(i)
        if ent.ObjectName == "AcDbBlockReference":
            nm = str(getattr(ent, "EffectiveName", ent.Name)).upper()
            if "TITLE" in nm:
                ip = ent.InsertionPoint
                ox, oy, origin_src = float(ip[0]), float(ip[1]), "titleblock_insertion"
                break
    if origin_src == "paper_zero":
        for i in range(ps.Count):
            ent = ps.Item(i)
            if ent.ObjectName == "AcDbViewport" and abs(ent.Width - paper_w) < 1.0 and abs(ent.Height - paper_h) < 1.0:
                cx, cy, _z = ent.Center
                ox, oy, origin_src = cx - paper_w / 2.0, cy - paper_h / 2.0, "fullsheet_viewport"
                break

    safe_min_x = 8.0 * k
    safe_min_y = 8.0 * k
    safe_max_y = paper_h - 11.0 * k
    safe_max_x = paper_w * 0.871
    source = "template_ratio"

    strip_left: List[float] = []
    for i in range(ps.Count):
        ent = ps.Item(i)
        if ent.ObjectName == "AcDbViewport":
            continue
        try:
            bb = _read_bbox(ent)
        except Exception:
            continue
        rx1, rx2 = bb["x1"] - ox, bb["x2"] - ox
        if bb["width"] < paper_w * 0.35 and rx1 > paper_w * 0.68 and rx2 <= paper_w + 5.0:
            strip_left.append(rx1)
    if strip_left:
        safe_max_x = min(strip_left) - 4.0 * k
        source = "scanned_titleblock"

    return {
        "iso": iso, "paper_w": paper_w, "paper_h": paper_h,
        "origin": [round(ox, 3), round(oy, 3)], "origin_source": origin_src,
        "min_x": round(ox + safe_min_x, 1), "min_y": round(oy + safe_min_y, 1),
        "max_x": round(ox + safe_max_x, 1), "max_y": round(oy + safe_max_y, 1),
        "source": source,
    }


def _activate_paperspace(doc, layout_name: str):
    layout = doc.Layouts.Item(layout_name)
    doc.ActiveLayout = layout
    if int(doc.ActiveSpace) != 0:
        doc.ActiveSpace = 0
    return layout


@mcp.tool()
def inspect_layout_template(layout_name: str) -> Dict[str, Any]:
    """
    Read-only survey of a PaperSpace tab BEFORE placing viewports: ISO size, safe drawable box
    (avoids the right-hand title block), existing viewports, and a suggested centre point.
    """
    if layout_name.strip().lower() == "model":
        return {"error": "Cannot inspect 'Model' as a PaperSpace sheet."}
    with autocad_session() as (_, doc):
        def _inspect():
            layout = _activate_paperspace(doc, layout_name)
            z = _detect_safe_zone_mm(doc, layout)
            vps = []
            ps = doc.PaperSpace
            for i in range(ps.Count):
                ent = ps.Item(i)
                if ent.ObjectName == "AcDbViewport" and i > 0:
                    cx, cy, _z = ent.Center
                    sc = float(ent.CustomScale)
                    vps.append({
                        "handle": ent.Handle,
                        "center_mm": [round(cx, 2), round(cy, 2)],
                        "size_mm": [round(ent.Width, 2), round(ent.Height, 2)],
                        "custom_scale": round(sc, 6),
                        "approx_scale": f"1:{round(1.0 / sc)}" if sc > 0 else "unscaled",
                        "locked": bool(ent.DisplayLocked),
                    })
            return {
                "layout_name": layout.Name,
                "iso_sheet_detected": z["iso"],
                "paper_size_mm": [round(z["paper_w"], 1), round(z["paper_h"], 1)],
                "sheet_origin_mm": z["origin"],
                "sheet_origin_source": z["origin_source"],
                "safe_drawable_box_mm": {
                    "min_x": z["min_x"], "min_y": z["min_y"], "max_x": z["max_x"], "max_y": z["max_y"],
                    "width": round(z["max_x"] - z["min_x"], 1),
                    "height": round(z["max_y"] - z["min_y"], 1),
                    "source": z["source"],
                    "note": "absolute PaperSpace coordinates (sheet origin already added)",
                },
                "safe_center_mm": [round((z["min_x"] + z["max_x"]) / 2, 1),
                                   round((z["min_y"] + z["max_y"]) / 2, 1)],
                "existing_viewports": vps,
            }
        return com_retry(_inspect)


@mcp.tool()
def place_enlarged_viewport_auto(
    layout_name: str,
    model_min_x: float,
    model_min_y: float,
    model_max_x: float,
    model_max_y: float,
    scale_denominator: int = 25,
    slot: str = "full",
    paper_center_x: Optional[float] = None,
    paper_center_y: Optional[float] = None,
    model_margin_mm: float = 600.0,
    vp_freeze_layers: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Place a locked, exact-scale PaperSpace viewport aimed at a ModelSpace box (1000 units = 1 m).
    Frame size = (model span + 2*margin) / scale. Rejected (nothing created) if it would overflow
    the safe zone, returning a recommended standard scale.
    slot: full | left_half | right_half | top_left | top_right | bottom_left | bottom_right.
    The frame goes on DEFPOINTS (non-plotting); PSLTSCALE is forced to 1.
    """
    if layout_name.strip().lower() == "model":
        return {"error": "Target layout must be a PaperSpace tab, not 'Model'."}
    with autocad_session() as (app, doc):
        def _place():
            layout = _activate_paperspace(doc, layout_name)
            z = _detect_safe_zone_mm(doc, layout)

            mid_x = (z["min_x"] + z["max_x"]) / 2.0
            mid_y = (z["min_y"] + z["max_y"]) / 2.0
            zones = {
                "full": (z["min_x"], z["min_y"], z["max_x"], z["max_y"]),
                "left_half": (z["min_x"], z["min_y"], mid_x - 3, z["max_y"]),
                "right_half": (mid_x + 3, z["min_y"], z["max_x"], z["max_y"]),
                "top_left": (z["min_x"], mid_y + 3, mid_x - 3, z["max_y"]),
                "top_right": (mid_x + 3, mid_y + 3, z["max_x"], z["max_y"]),
                "bottom_left": (z["min_x"], z["min_y"], mid_x - 3, mid_y - 3),
                "bottom_right": (mid_x + 3, z["min_y"], z["max_x"], mid_y - 3),
            }
            if slot not in zones:
                return {"error": f"Unknown slot '{slot}'. Use one of {sorted(zones)}."}
            z0x, z0y, z1x, z1y = zones[slot]
            avail_w, avail_h = z1x - z0x, z1y - z0y

            x0, x1 = sorted([model_min_x, model_max_x])
            y0, y1 = sorted([model_min_y, model_max_y])
            mcx, mcy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
            span_w = (x1 - x0) + 2.0 * model_margin_mm
            span_h = (y1 - y0) + 2.0 * model_margin_mm

            custom_scale = 1.0 / float(scale_denominator)
            vp_w, vp_h = span_w * custom_scale, span_h * custom_scale

            if vp_w > avail_w or vp_h > avail_h:
                need = max(span_w / max(avail_w, 1.0), span_h / max(avail_h, 1.0))
                rec = next((s for s in _STD_SCALES if s >= need), round(need))
                return {
                    "status": "rejected_overflow", "iso_sheet": z["iso"], "slot": slot,
                    "requested_scale": f"1:{scale_denominator}",
                    "required_frame_mm": [round(vp_w, 1), round(vp_h, 1)],
                    "available_zone_mm": [round(avail_w, 1), round(avail_h, 1)],
                    "recommended_standard_scale": f"1:{rec}",
                }

            cx = paper_center_x if paper_center_x is not None else (z0x + z1x) / 2.0
            cy = paper_center_y if paper_center_y is not None else (z0y + z1y) / 2.0
            cx = max(z0x + vp_w / 2.0, min(cx, z1x - vp_w / 2.0))
            cy = max(z0y + vp_h / 2.0, min(cy, z1y - vp_h / 2.0))

            try:
                doc.Layers.Item("DEFPOINTS")
            except Exception:
                doc.Layers.Add("DEFPOINTS")

            pvp = doc.PaperSpace.AddPViewport(pt3d(cx, cy, 0.0), vp_w, vp_h)
            pvp.Layer = "DEFPOINTS"
            pvp.Display(True)  # must be ON before aiming/scaling

            # Aim the camera at the model box centre, then apply the exact scale
            pvp.Target = pt3d(mcx, mcy, 0.0)
            pvp.CustomScale = custom_scale

            freeze_note = "none"
            if vp_freeze_layers:
                csv = ",".join(n.strip() for n in vp_freeze_layers if n.strip())
                if csv:
                    try:
                        doc.MSpace = True
                        doc.ActivePViewport = pvp
                        doc.SendCommand(f"._VPLAYER _Freeze {csv} _Current \n")
                        freeze_note = csv
                    except Exception as exc:
                        freeze_note = f"skipped ({exc.__class__.__name__}) - freeze manually via VPLAYER"
                    finally:
                        with contextlib.suppress(Exception):
                            doc.MSpace = False

            pvp.CustomScale = custom_scale
            pvp.DisplayLocked = True
            doc.SetVariable("PSLTSCALE", 1)
            doc.Regen(1)

            return {
                "status": "placed_and_locked",
                "handle": pvp.Handle, "layout": layout.Name,
                "iso_sheet": z["iso"], "slot": slot,
                "scale": f"1:{scale_denominator}",
                "paper_center_mm": [round(cx, 2), round(cy, 2)],
                "paper_frame_mm": [round(vp_w, 2), round(vp_h, 2)],
                "model_target": [round(mcx, 2), round(mcy, 2)],
                "vp_frozen_layers": freeze_note,
                "recommended_model_text_height_mm": round(2.5 * scale_denominator, 1),
                "next": "Verify with capture_drawing_view - target/scale were set via COM, not visually confirmed.",
            }
        return com_retry(_place)


# ============================================================================
# Safe Execution & Saving
# ============================================================================

@mcp.tool()
def send_command_safe(command_str: str) -> str:
    """
    Send a command-line string or AutoLISP expression to AutoCAD.
    Blocks QUIT/EXIT/SAVE commands that hang automated sessions.
    """
    tokens = {tok.strip().upper() for tok in command_str.replace("\n", " ").split()}
    blocked_hit = tokens.intersection(BLOCKED_COMMANDS)
    if blocked_hit:
        raise ValueError(
            f"Blocked unsafe command(s) {blocked_hit}. "
            "Use 'save_drawing' at the very end of the workflow instead."
        )

    with autocad_session() as (_, doc):
        def _exec():
            payload = command_str if command_str.endswith("\n") else f"{command_str}\n"
            doc.SendCommand(payload)
            return f"Executed command: {command_str.strip()}"
        return com_retry(_exec)

@mcp.tool()
def save_drawing() -> str:
    """Save the active drawing. Call this ONCE at the end of a workflow, not mid-task."""
    with autocad_session() as (_, doc):
        def _save():
            if not doc.FullName:
                raise RuntimeError("Drawing has never been saved. Save it manually once first.")
            doc.Save()
            return f"Saved '{doc.FullName}' successfully."
        return com_retry(_save)


# ==============================================================================
# SECTION: ADVANCED PAPERSPACE LAYOUT & TAB MANAGEMENT
# ==============================================================================

_INVALID_LAYOUT_CHARS = re.compile(r'[<>/\\":;?*|,=`]')


def _validate_layout_name(name: str) -> Optional[str]:
    clean = name.strip()
    if not clean:
        return "Layout name cannot be empty."
    if clean.lower() == "model":
        return "'Model' is a reserved system tab name."
    if len(clean) > 255:
        return f"Layout name '{clean[:30]}...' exceeds 255 characters."
    if _INVALID_LAYOUT_CHARS.search(clean):
        return f"Layout name '{clean}' contains forbidden AutoCAD characters (< > / \\ \" : ; ? * | , = `)."
    return None


def _get_ordered_ps_layouts(doc) -> List[Any]:
    """Return all PaperSpace layout COM objects sorted by visual TabOrder (1..N)."""
    ps_layouts = []
    for i in range(doc.Layouts.Count):
        lyt = doc.Layouts.Item(i)
        if lyt.Name.lower() != "model":
            ps_layouts.append(lyt)
    ps_layouts.sort(key=lambda x: x.TabOrder)
    return ps_layouts


@mcp.tool()
def list_layout_tabs() -> Dict[str, Any]:
    """
    List all PaperSpace Layout tabs in their exact left-to-right visual order,
    including paper dimensions, plot style (.ctb), and viewport count.
    """
    with autocad_session() as (_, doc):
        def _list():
            tabs = []
            for lyt in _get_ordered_ps_layouts(doc):
                pw, ph = lyt.GetPaperSize()
                if lyt.PlotRotation in (1, 3):
                    pw, ph = ph, pw

                longer = max(pw, ph)
                iso = "A4" if longer < 330 else "A3" if longer < 500 else "A2" if longer < 700 else "A1"

                blk = lyt.Block
                vp_count = 0
                for k in range(1, blk.Count):
                    if blk.Item(k).ObjectName == "AcDbViewport":
                        vp_count += 1

                tabs.append({
                    "tab_order": lyt.TabOrder,
                    "name": lyt.Name,
                    "iso_size": iso,
                    "paper_mm": [round(pw, 1), round(ph, 1)],
                    "plotter": lyt.ConfigName,
                    "ctb_style": lyt.StyleSheet,
                    "viewport_count": vp_count,
                })

            return {
                "active_layout": doc.ActiveLayout.Name,
                "total_paperspace_tabs": len(tabs),
                "tabs": tabs,
            }
        return com_retry(_list)


@mcp.tool()
def batch_clone_layouts(
    source_layout: str,
    new_layout_names: List[str],
    insert_after_layout: Optional[str] = None,
    clear_existing_viewports: bool = True,
    skip_if_exists: bool = True,
) -> Dict[str, Any]:
    """
    Clone a master/template Layout tab into multiple new PaperSpace tabs in one step.
    Copies exact Page Setup (A3/A4/A2/A1, plotter, CTB) and Title Block geometry
    WITHOUT modifying any Title Block text or sheet IDs.
    """
    with autocad_session() as (_, doc):
        def _clone():
            if not new_layout_names:
                return {"error": "new_layout_names list cannot be empty."}

            existing_map = {
                doc.Layouts.Item(i).Name.lower(): doc.Layouts.Item(i)
                for i in range(doc.Layouts.Count)
            }
            src_lyt = existing_map.get(source_layout.strip().lower())
            if src_lyt is None or src_lyt.Name.lower() == "model":
                return {"error": f"Source layout '{source_layout}' not found or is 'Model'."}

            cleaned_targets = []
            seen_lower = set()
            for raw_name in new_layout_names:
                err = _validate_layout_name(raw_name)
                if err:
                    return {"error": err}
                cname = raw_name.strip()
                if cname.lower() in seen_lower:
                    return {"error": f"Duplicate target name in request: '{cname}'."}
                seen_lower.add(cname.lower())
                cleaned_targets.append(cname)

            src_blk = src_lyt.Block
            entities_to_copy = []
            for idx in range(src_blk.Count):
                if idx == 0:
                    continue
                ent = src_blk.Item(idx)
                if clear_existing_viewports and ent.ObjectName == "AcDbViewport":
                    continue
                entities_to_copy.append(ent)

            var_objs = (
                win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, entities_to_copy
                )
                if entities_to_copy
                else None
            )

            created = []
            skipped = []
            newly_created_objs = []

            for target_name in cleaned_targets:
                if target_name.lower() in existing_map:
                    if skip_if_exists:
                        skipped.append(target_name)
                        continue
                    return {"error": f"Layout '{target_name}' already exists."}

                dst_lyt = doc.Layouts.Add(target_name)
                dst_lyt.CopyFrom(src_lyt)

                dst_blk = dst_lyt.Block
                if clear_existing_viewports and dst_blk.Count > 1:
                    for k in range(dst_blk.Count - 1, 0, -1):
                        if dst_blk.Item(k).ObjectName == "AcDbViewport":
                            dst_blk.Item(k).Delete()

                if var_objs is not None:
                    doc.CopyObjects(var_objs, dst_blk)

                existing_map[target_name.lower()] = dst_lyt
                newly_created_objs.append(dst_lyt)
                created.append(target_name)

            ordered_all = _get_ordered_ps_layouts(doc)
            new_set = {lyt.Name.lower() for lyt in newly_created_objs}
            base_list = [lyt for lyt in ordered_all if lyt.Name.lower() not in new_set]

            insert_idx = len(base_list)
            if insert_after_layout:
                anchor_key = insert_after_layout.strip().lower()
                for idx, lyt in enumerate(base_list):
                    if lyt.Name.lower() == anchor_key:
                        insert_idx = idx + 1
                        break

            final_sequence = (
                base_list[:insert_idx] + newly_created_objs + base_list[insert_idx:]
            )
            for pos, lyt in enumerate(final_sequence, start=1):
                if lyt.TabOrder != pos:
                    lyt.TabOrder = pos

            return {
                "status": "batch_clone_complete",
                "source_layout": src_lyt.Name,
                "created_layouts": created,
                "skipped_existing": skipped,
                "clear_existing_viewports": clear_existing_viewports,
                "final_tab_order": [lyt.Name for lyt in _get_ordered_ps_layouts(doc)],
            }
        return com_retry(_clone)


@mcp.tool()
def manage_layout_tabs(
    renames: Optional[Dict[str, str]] = None,
    tab_order: Optional[List[str]] = None,
    start_order_at: int = 1,
) -> Dict[str, Any]:
    """
    Batch rename and/or reorder PaperSpace Layout tabs in a single call.
    - renames: mapping of {'Old Tab Name': 'New Tab Name'} (safe against name-swap collisions).
    - tab_order: ordered list of tab names (use post-rename names if combining both).
      Unlisted tabs keep their relative order after the positioned block.
    """
    with autocad_session() as (_, doc):
        def _manage():
            renamed_log = []

            if renames:
                layout_lookup = {
                    lyt.Name.lower(): lyt for lyt in _get_ordered_ps_layouts(doc)
                }
                staged = []
                target_lower_set = set()

                for old_name, new_name in renames.items():
                    old_key = old_name.strip().lower()
                    if old_key not in layout_lookup:
                        return {"error": f"Layout to rename not found: '{old_name}'."}
                    err = _validate_layout_name(new_name)
                    if err:
                        return {"error": err}
                    clean_new = new_name.strip()
                    if clean_new.lower() in target_lower_set:
                        return {"error": f"Duplicate target name in renames: '{clean_new}'."}
                    target_lower_set.add(clean_new.lower())
                    staged.append((layout_lookup[old_key], old_name.strip(), clean_new))

                renaming_old_keys = {item[0].Name.lower() for item in staged}
                for _, _, clean_new in staged:
                    if (
                        clean_new.lower() in layout_lookup
                        and clean_new.lower() not in renaming_old_keys
                    ):
                        return {"error": f"Target name '{clean_new}' already exists on another tab."}

                for idx, (lyt_obj, _, _) in enumerate(staged):
                    lyt_obj.Name = f"__TMP_RENAME_{idx}_{lyt_obj.Handle}__"

                for lyt_obj, old_label, clean_new in staged:
                    lyt_obj.Name = clean_new
                    renamed_log.append({"from": old_label, "to": clean_new})

            if tab_order:
                current_tabs = _get_ordered_ps_layouts(doc)
                by_lower = {lyt.Name.lower(): lyt for lyt in current_tabs}

                requested_objs = []
                requested_keys = set()
                for name in tab_order:
                    key = name.strip().lower()
                    if key == "model":
                        continue
                    if key not in by_lower:
                        return {"error": f"Layout '{name}' in tab_order not found."}
                    if key not in requested_keys:
                        requested_objs.append(by_lower[key])
                        requested_keys.add(key)

                remaining = [
                    lyt for lyt in current_tabs if lyt.Name.lower() not in requested_keys
                ]
                insert_pos = max(0, min(start_order_at - 1, len(remaining)))
                final_tabs = (
                    remaining[:insert_pos] + requested_objs + remaining[insert_pos:]
                )

                for pos, lyt in enumerate(final_tabs, start=1):
                    if lyt.TabOrder != pos:
                        lyt.TabOrder = pos

            return {
                "status": "updated",
                "renamed": renamed_log,
                "current_tab_order": [
                    {"tab_order": lyt.TabOrder, "name": lyt.Name}
                    for lyt in _get_ordered_ps_layouts(doc)
                ],
            }
        return com_retry(_manage)

if __name__ == "__main__":
    logger.info("Starting AutoCAD MCP Server (stdio) for Google Antigravity...")
    mcp.run(transport="stdio")
