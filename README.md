# cadCOMPANION

High-performance, popup-free **AutoCAD Model Context Protocol (MCP) Server** and AI Drafting Companion designed for Google Antigravity, Claude Desktop, and Gemini Enterprise.

---

## 🚀 Key Features

### 1. Robust Win32 & COM Architecture
- **Zero Ghost Processes (`CAD_LAUNCH=never`)**: Connects safely to the active AutoCAD session via `GetActiveObject` without spawning zombie background processes.
- **Auto-ESC & Lock Recovery**: Traps `CMDACTIVE > 0` and unblocks frozen command prompts via asynchronous Win32 `VK_ESCAPE` message queue posting (`win32gui.PostMessage`).
- **Interactive Popup Suppression**: Automatically enforces `FILEDIA=0`, `CMDDIA=0`, `ATTDIA=0`, `ATTREQ=0`, and `BLOCKEDITOR=0` across all operations.
- **Cross-Desktop Compatibility**: Built-in support for switching to the interactive `Default` desktop (`win32service.OpenDesktop`) from sandboxed terminals.

### 2. Architectural Drafting & Construction Symbols
- **`draw_level_datum`**: Standard architectural level markers (▽ with half-shaded contrast, shelf, and dashed witness line, e.g., `FFL ±0.00 m.`).
- **`draw_aligned_dim` & `draw_dim_chain`**: Multi-tiered architectural dimension strings (10×S, +8×S margin stacking) with 45° ArchTicks.
- **`draw_grid_balloon`**: CENTER-linetype grid axes with circular numbered balloons outside dimension lines.
- **`draw_section_mark` & `draw_interior_elevation_mark`**: 4-way interior room markers and section cut arrows with sheet callout bubbles.
- **`draw_enlarged_detail_callout`**: Chamfered boundary callouts with shouldered leader landing lines.

### 3. Spatial Intelligence & Anti-Clutter
- **`find_free_zones`**: 24×16 grid density scanner using 2D histogram empty-rectangle algorithms to place tables, notes, and schedules into clean whitespace ($\ge 35\%$).
- **`capture_drawing_view`**: Native viewport image rendering via `._-PNGOUT` with Pillow edge enhancement for Visual QA/QC.
- **`agent_preflight_guard`**: Mathematical pre-flight check enforcing closed polylines, scale math ($2.5S$, $10S$, $+8S$), and duplicate hatch prevention.
- **`context_pruner`**: Dynamic domain skill routing (KTV, Bar, Wet Area, RCP) cutting prompt bloat by >55%.

---

## 📦 Quick Start

### Installation with `uv` (Recommended)

```bash
uv run --python 3.12 autocad_mcp_server.py
```

### Antigravity / Claude Desktop Configuration

Add the following to your `mcp_config.json`:

```json
{
  "mcpServers": {
    "cadcompanion": {
      "command": "uv",
      "args": [
        "run",
        "--python",
        "3.12",
        "<PATH_TO_REPO>/autocad_mcp_server.py"
      ],
      "env": {
        "PYTHONUTF8": "1"
      }
    }
  }
}
```

---

## 📄 License
MIT License.
