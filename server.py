"""
SOAR Desktop MCP - MCP Server
Lightweight desktop UI automation via Windows Accessibility Tree.

6 Tools:
  desktop_snapshot     - Capture Accessibility Tree of a window
  desktop_click        - Click a UI element
  desktop_type         - Type text into a UI element
  desktop_list_windows - List all visible windows
  desktop_focus_window - Bring a window to foreground
  desktop_press_key    - Simulate keyboard input

Open Source - MIT License
(c) 2026 Soar Design (Jimmy Kwok)
https://github.com/soardesignjimmy/soar-desktop-mcp
"""

from __future__ import annotations

import json
import logging
import sys

from mcp.server.fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stderr)],
)
log = logging.getLogger("soar-desktop-mcp")

# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------
mcp = FastMCP(
    "soar-desktop",
    version="1.0.0",
    description="SOAR Desktop MCP - Control any Windows application via Accessibility Tree",
)

# ---------------------------------------------------------------------------
# Lazy import desktop_lib (Windows-only)
# ---------------------------------------------------------------------------
_lib = None

def _get_lib():
    global _lib
    if _lib is None:
        import desktop_lib
        _lib = desktop_lib
    return _lib


# ---------------------------------------------------------------------------
# Tool: desktop_list_windows
# ---------------------------------------------------------------------------
@mcp.tool()
def desktop_list_windows() -> str:
    """List all visible windows on the desktop.
    Returns window titles, class names, PIDs, and positions.
    Use this to find the window you want to control.
    """
    lib = _get_lib()
    windows = lib.list_windows()
    return json.dumps(windows, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Tool: desktop_focus_window
# ---------------------------------------------------------------------------
@mcp.tool()
def desktop_focus_window(title: str = "", pid: int = 0) -> str:
    """Bring a window to the foreground.

    Args:
        title: Window title substring to match (case-insensitive)
        pid: Process ID of the window
    """
    lib = _get_lib()
    if not title and not pid:
        return json.dumps({"success": False, "error": "Provide title or pid"})
    result = lib.focus_window(title=title, pid=pid)
    return json.dumps(result, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Tool: desktop_snapshot
# ---------------------------------------------------------------------------
@mcp.tool()
def desktop_snapshot(
    title: str = "",
    pid: int = 0,
    max_depth: int = 6,
) -> str:
    """Capture the Accessibility Tree of a window.

    Returns a YAML-like tree with ref IDs for each element.
    Use ref IDs with desktop_click and desktop_type.
    If no title/pid given, captures the foreground window.

    Args:
        title: Window title substring to match
        pid: Process ID of the window
        max_depth: Maximum tree depth (default 6, reduce for large windows)
    """
    lib = _get_lib()
    use_fg = not title and not pid
    tree = lib.snapshot(title=title, pid=pid, max_depth=max_depth, use_foreground=use_fg)
    return tree


# ---------------------------------------------------------------------------
# Tool: desktop_click
# ---------------------------------------------------------------------------
@mcp.tool()
def desktop_click(
    ref: str = "",
    name: str = "",
    automation_id: str = "",
    button: str = "left",
    double_click: bool = False,
) -> str:
    """Click a UI element on the desktop.

    Target by ref (from snapshot), element name, or automationId.
    Ref is preferred - run desktop_snapshot first to get refs.

    Args:
        ref: Element reference from snapshot (e.g. "e5")
        name: Element name to search for
        automation_id: AutomationId of the element
        button: Mouse button - "left", "right", or "middle"
        double_click: Whether to double-click
    """
    lib = _get_lib()
    if not ref and not name and not automation_id:
        return json.dumps({"success": False, "error": "Provide ref, name, or automation_id"})
    result = lib.click_element(
        ref=ref, name=name, automation_id=automation_id,
        button=button, double_click=double_click,
    )
    return json.dumps(result, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Tool: desktop_type
# ---------------------------------------------------------------------------
@mcp.tool()
def desktop_type(
    text: str,
    ref: str = "",
    name: str = "",
    automation_id: str = "",
    clear_first: bool = True,
    submit: bool = False,
) -> str:
    """Type text into a UI element (text box, combo box, etc.).

    Target by ref (from snapshot), element name, or automationId.

    Args:
        text: Text to type
        ref: Element reference from snapshot (e.g. "e12")
        name: Element name to search for
        automation_id: AutomationId of the element
        clear_first: Clear existing text before typing (default True)
        submit: Press Enter after typing (default False)
    """
    lib = _get_lib()
    if not ref and not name and not automation_id:
        return json.dumps({"success": False, "error": "Provide ref, name, or automation_id"})
    result = lib.type_text(
        ref=ref, name=name, automation_id=automation_id,
        text=text, clear_first=clear_first, submit=submit,
    )
    return json.dumps(result, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Tool: desktop_press_key
# ---------------------------------------------------------------------------
@mcp.tool()
def desktop_press_key(key: str) -> str:
    """Press a key or key combination.

    Examples: "Enter", "Tab", "Escape", "Ctrl+S", "Alt+F4",
              "Shift+Ctrl+N", "F5", "Ctrl+C", "Ctrl+V"

    Args:
        key: Key name or combination (e.g. "Ctrl+S")
    """
    lib = _get_lib()
    result = lib.press_key(key)
    return json.dumps(result, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    log.info("SOAR Desktop MCP v1.0.0 starting (stdio) ...")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
