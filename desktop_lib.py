"""
SOAR Desktop MCP - Windows UI Automation Library
Lightweight wrapper around Windows UI Automation API for desktop control.

Open Source - MIT License
(c) 2026 Soar Design (Jimmy Kwok)
"""

from __future__ import annotations

import json
import sys
import time
from dataclasses import dataclass, field
from typing import Any, Optional

# ---------------------------------------------------------------------------
# Platform guard
# ---------------------------------------------------------------------------
if sys.platform != "win32":
    raise RuntimeError("SOAR Desktop MCP requires Windows (win32)")

try:
    import uiautomation as auto
except ImportError:
    raise ImportError(
        "uiautomation not installed. Run:  pip install uiautomation"
    )


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------
@dataclass
class UIElement:
    """Simplified representation of a UI Automation element."""
    ref: str                     # e.g. "e1", "e2" — short reference ID
    role: str                    # ControlType: Button, Edit, Text, etc.
    name: str                    # Element name / text
    automation_id: str = ""      # AutomationId (stable identifier)
    value: str = ""              # Current value (for editable elements)
    states: list[str] = field(default_factory=list)  # e.g. ["enabled", "focused"]
    rect: dict = field(default_factory=dict)          # {x, y, w, h}
    children: list[UIElement] = field(default_factory=list)
    depth: int = 0


# ---------------------------------------------------------------------------
# Reference tracking (session-scoped)
# ---------------------------------------------------------------------------
_ref_counter: int = 0
_ref_map: dict[str, Any] = {}   # ref -> uiautomation Control object


def _reset_refs():
    global _ref_counter, _ref_map
    _ref_counter = 0
    _ref_map = {}


def _next_ref(control) -> str:
    global _ref_counter
    _ref_counter += 1
    ref = f"e{_ref_counter}"
    _ref_map[ref] = control
    return ref


def resolve_ref(ref: str):
    """Look up a UI Automation control by its ref ID."""
    ctrl = _ref_map.get(ref)
    if ctrl is None:
        raise ValueError(f"Unknown ref '{ref}'. Run desktop_snapshot first.")
    return ctrl


# ---------------------------------------------------------------------------
# Core: list windows
# ---------------------------------------------------------------------------
def list_windows() -> list[dict]:
    """List all visible top-level windows with title and process info."""
    results = []
    root = auto.GetRootControl()
    for win in root.GetChildren():
        try:
            name = win.Name or ""
            if not name.strip():
                continue
            cls = win.ClassName or ""
            pid = win.ProcessId
            rect = win.BoundingRectangle
            results.append({
                "title": name,
                "class": cls,
                "pid": pid,
                "rect": {
                    "x": rect.left,
                    "y": rect.top,
                    "w": rect.width(),
                    "h": rect.height(),
                },
            })
        except Exception:
            continue
    return results


# ---------------------------------------------------------------------------
# Core: focus window
# ---------------------------------------------------------------------------
def focus_window(title: str = "", pid: int = 0) -> dict:
    """Bring a window to the foreground by title substring or PID."""
    root = auto.GetRootControl()
    for win in root.GetChildren():
        try:
            match = False
            if title and title.lower() in (win.Name or "").lower():
                match = True
            if pid and win.ProcessId == pid:
                match = True
            if match:
                win.SetActive()
                win.SetFocus()
                time.sleep(0.3)
                return {
                    "success": True,
                    "title": win.Name,
                    "pid": win.ProcessId,
                }
        except Exception:
            continue
    return {"success": False, "error": f"Window not found: title='{title}' pid={pid}"}


# ---------------------------------------------------------------------------
# Core: snapshot (Accessibility Tree)
# ---------------------------------------------------------------------------
def snapshot(
    title: str = "",
    pid: int = 0,
    max_depth: int = 6,
    use_foreground: bool = False,
) -> str:
    """
    Capture the Accessibility Tree of a window as YAML-like text.
    Returns a compact tree with ref IDs for each interactive element.
    """
    _reset_refs()

    # Find target window
    target = None
    if use_foreground:
        target = auto.GetForegroundControl()
    else:
        root = auto.GetRootControl()
        for win in root.GetChildren():
            try:
                match = False
                if title and title.lower() in (win.Name or "").lower():
                    match = True
                if pid and win.ProcessId == pid:
                    match = True
                if match:
                    target = win
                    break
            except Exception:
                continue

    if target is None:
        if not title and not pid:
            target = auto.GetForegroundControl()
        else:
            return f"# Window not found: title='{title}' pid={pid}"

    # Build tree
    lines = []
    lines.append(f"# Window: {target.Name}")
    lines.append(f"# Class: {target.ClassName}")
    lines.append(f"# PID: {target.ProcessId}")
    lines.append("")
    _walk(target, lines, depth=0, max_depth=max_depth)
    return "\n".join(lines)


def _walk(control, lines: list[str], depth: int, max_depth: int):
    """Recursively walk the UI tree and build YAML-like output."""
    if depth > max_depth:
        return

    indent = "  " * depth
    ctrl_type = _control_type_name(control)
    name = (control.Name or "").strip()
    auto_id = ""
    try:
        auto_id = control.AutomationId or ""
    except Exception:
        pass

    # Skip invisible / empty containers at deeper levels
    if depth > 1 and not name and ctrl_type in ("Pane", "Custom", "Group"):
        # Still walk children in case they have content
        try:
            for child in control.GetChildren():
                _walk(child, lines, depth, max_depth)
        except Exception:
            pass
        return

    # Assign ref for interactive / identifiable elements
    ref = _next_ref(control)

    # Build line
    parts = [f"{indent}- {ctrl_type}"]
    if name:
        parts.append(f'"{name}"')
    parts.append(f"[ref={ref}]")
    if auto_id:
        parts.append(f"[id={auto_id}]")

    # Add state info
    states = _get_states(control)
    if states:
        parts.append(f"[{', '.join(states)}]")

    # Add value for editable elements
    value = _get_value(control)
    if value:
        parts.append(f': "{value}"')

    lines.append(" ".join(parts))

    # Recurse children
    try:
        children = control.GetChildren()
        for child in children:
            _walk(child, lines, depth + 1, max_depth)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Core: click element
# ---------------------------------------------------------------------------
def click_element(
    ref: str = "",
    name: str = "",
    automation_id: str = "",
    button: str = "left",
    double_click: bool = False,
) -> dict:
    """Click a UI element by ref, name, or automationId."""
    ctrl = _find_element(ref, name, automation_id)
    if ctrl is None:
        return {"success": False, "error": _not_found_msg(ref, name, automation_id)}

    try:
        rect = ctrl.BoundingRectangle
        cx = rect.left + rect.width() // 2
        cy = rect.top + rect.height() // 2

        if double_click:
            auto.Click(cx, cy)
            time.sleep(0.05)
            auto.Click(cx, cy)
        elif button == "right":
            auto.RightClick(cx, cy)
        elif button == "middle":
            auto.MiddleClick(cx, cy)
        else:
            auto.Click(cx, cy)

        time.sleep(0.2)
        return {
            "success": True,
            "element": ctrl.Name or "",
            "type": _control_type_name(ctrl),
            "position": {"x": cx, "y": cy},
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


# ---------------------------------------------------------------------------
# Core: type text
# ---------------------------------------------------------------------------
def type_text(
    ref: str = "",
    name: str = "",
    automation_id: str = "",
    text: str = "",
    clear_first: bool = True,
    submit: bool = False,
) -> dict:
    """Type text into a UI element (Edit, ComboBox, etc.)."""
    ctrl = _find_element(ref, name, automation_id)
    if ctrl is None:
        return {"success": False, "error": _not_found_msg(ref, name, automation_id)}

    try:
        # Try ValuePattern first (most reliable)
        vp = ctrl.GetValuePattern()
        if vp:
            if clear_first:
                vp.SetValue("")
            vp.SetValue(text)
        else:
            # Fallback: focus + keyboard
            ctrl.SetFocus()
            time.sleep(0.1)
            if clear_first:
                auto.SendKeys("{Ctrl}a", waitTime=0.05)
            auto.SendKeys(text, interval=0.02, waitTime=0.05)

        if submit:
            auto.SendKeys("{Enter}", waitTime=0.1)

        return {
            "success": True,
            "element": ctrl.Name or "",
            "typed": text,
            "submitted": submit,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


# ---------------------------------------------------------------------------
# Core: press key
# ---------------------------------------------------------------------------
def press_key(key: str) -> dict:
    """
    Press a key or key combination.
    Examples: "Enter", "Tab", "Ctrl+S", "Alt+F4", "Shift+Ctrl+N"
    Uses SendKeys syntax: {Enter}, {Tab}, {Ctrl}, {Alt}, {Shift}, etc.
    """
    try:
        # Convert human-friendly format to SendKeys syntax
        sendkeys_str = _to_sendkeys(key)
        auto.SendKeys(sendkeys_str, waitTime=0.1)
        return {"success": True, "key": key, "sendkeys": sendkeys_str}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _find_element(ref: str, name: str, automation_id: str):
    """Find element by ref (preferred), name, or automationId."""
    if ref:
        return _ref_map.get(ref)
    if automation_id:
        try:
            root = auto.GetForegroundControl()
            found = root.Control(AutomationId=automation_id)
            if found and found.Exists(maxSearchSeconds=2):
                return found
        except Exception:
            pass
    if name:
        try:
            root = auto.GetForegroundControl()
            found = root.Control(Name=name)
            if found and found.Exists(maxSearchSeconds=2):
                return found
        except Exception:
            pass
    return None


def _not_found_msg(ref, name, automation_id) -> str:
    parts = []
    if ref:
        parts.append(f"ref='{ref}'")
    if name:
        parts.append(f"name='{name}'")
    if automation_id:
        parts.append(f"automationId='{automation_id}'")
    return f"Element not found: {', '.join(parts)}. Run desktop_snapshot first."


def _control_type_name(ctrl) -> str:
    """Get the short ControlType name."""
    try:
        ct = ctrl.ControlTypeName or "Unknown"
        return ct.replace("Control", "")
    except Exception:
        return "Unknown"


def _get_states(ctrl) -> list[str]:
    """Get relevant state flags for an element."""
    states = []
    try:
        if not ctrl.IsEnabled:
            states.append("disabled")
    except Exception:
        pass
    try:
        if ctrl.IsKeyboardFocusable:
            pass  # don't clutter, only show if actually focused
        if ctrl.HasKeyboardFocus:
            states.append("focused")
    except Exception:
        pass
    try:
        tp = ctrl.GetTogglePattern()
        if tp:
            state = tp.ToggleState
            states.append("checked" if state == 1 else "unchecked")
    except Exception:
        pass
    try:
        sp = ctrl.GetSelectionItemPattern()
        if sp and sp.IsSelected:
            states.append("selected")
    except Exception:
        pass
    return states


def _get_value(ctrl) -> str:
    """Get the current value for value-bearing elements."""
    try:
        vp = ctrl.GetValuePattern()
        if vp:
            val = vp.Value or ""
            if val and len(val) < 200:
                return val
    except Exception:
        pass
    return ""


def _to_sendkeys(key: str) -> str:
    """
    Convert human-friendly key names to SendKeys format.
    'Ctrl+S' -> '{Ctrl}s'
    'Alt+F4' -> '{Alt}{F4}'
    'Enter' -> '{Enter}'
    'Shift+Ctrl+N' -> '{Shift}{Ctrl}n'
    """
    parts = key.split("+")
    result = []
    for p in parts:
        p = p.strip()
        low = p.lower()
        if low in ("ctrl", "control"):
            result.append("{Ctrl}")
        elif low in ("alt",):
            result.append("{Alt}")
        elif low in ("shift",):
            result.append("{Shift}")
        elif low in ("win", "windows", "meta"):
            result.append("{Win}")
        elif low in ("enter", "return"):
            result.append("{Enter}")
        elif low in ("tab",):
            result.append("{Tab}")
        elif low in ("escape", "esc"):
            result.append("{Esc}")
        elif low in ("backspace", "back"):
            result.append("{Back}")
        elif low in ("delete", "del"):
            result.append("{Del}")
        elif low in ("space",):
            result.append("{Space}")
        elif low in ("up",):
            result.append("{Up}")
        elif low in ("down",):
            result.append("{Down}")
        elif low in ("left",):
            result.append("{Left}")
        elif low in ("right",):
            result.append("{Right}")
        elif low in ("home",):
            result.append("{Home}")
        elif low in ("end",):
            result.append("{End}")
        elif low in ("pageup", "pgup"):
            result.append("{PageUp}")
        elif low in ("pagedown", "pgdn"):
            result.append("{PageDown}")
        elif low.startswith("f") and low[1:].isdigit():
            result.append("{" + p + "}")
        else:
            result.append(p)
    return "".join(result)
