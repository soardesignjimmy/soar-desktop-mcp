# SOAR Desktop MCP

Control any Windows application through the Accessibility Tree via Model Context Protocol (MCP).

Lightweight, open-source desktop automation for Claude Cowork, Claude Code, or any MCP-compatible AI client.

## Features

| Tool | Description |
|------|-------------|
| `desktop_snapshot` | Capture Accessibility Tree of any window |
| `desktop_click` | Click UI elements by ref, name, or automation ID |
| `desktop_type` | Type text into input fields |
| `desktop_list_windows` | List all visible windows |
| `desktop_focus_window` | Bring a window to the foreground |
| `desktop_press_key` | Simulate keyboard shortcuts (Ctrl+S, Alt+F4, etc.) |

## Requirements

- Windows 10/11
- Python 3.10+

## Quick Start

```bash
# 1. Clone or download
git clone https://github.com/soardesignjimmy/soar-desktop-mcp.git
cd soar-desktop-mcp

# 2. Install
install.bat

# 3. Add to Claude Cowork / Claude Code config
```

Add to your MCP config:

```json
{
  "mcpServers": {
    "soar-desktop": {
      "command": "C:\\path\\to\\soar-desktop-mcp\\run_mcp.bat"
    }
  }
}
```

## Usage Example

```
User: List all open windows
AI:   [calls desktop_list_windows]

User: Click the Save button in Notepad
AI:   [calls desktop_focus_window title="Notepad"]
      [calls desktop_snapshot]
      [calls desktop_click ref="e15"]  # Save button
```

## How It Works

SOAR Desktop MCP uses Windows UI Automation API to read the Accessibility Tree of any application. The tree contains all UI elements (buttons, text fields, menus, etc.) with their properties and positions.

The AI reads the tree via `desktop_snapshot`, identifies elements by their ref IDs, then interacts using `desktop_click`, `desktop_type`, or `desktop_press_key`.

This is the same API that screen readers use - no application-specific integration needed.

## Architecture

```
server.py        <- MCP server (FastMCP, stdio transport)
desktop_lib.py   <- Windows UI Automation wrapper
install.bat      <- One-click installer
run_mcp.bat      <- Launcher
```

## Token Efficiency

- Tool definitions: ~800 tokens (one-time handshake)
- Typical snapshot: 2,000-15,000 tokens (depends on window complexity)
- Click/type result: ~200 tokens
- Average interaction round: ~5,000 tokens

## License

MIT License - (c) 2026 Soar Design (Jimmy Kwok)

## Links

- Website: [soarmcpsoftware.com](https://soarmcpsoftware.com)
- SOAR MCP Products: AutoCAD, Revit, Inventor, Word, Excel, PowerPoint, 3ds Max, Blender, Photoshop
