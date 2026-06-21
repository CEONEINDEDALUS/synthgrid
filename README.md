# ╔═══════════════════════════════════════════════════════════════════╗
# ║  SYNTHGRID — AI-Powered Excel Automation Node                    ║
# ║  v1.0.0 | Linux Native | Fedora/KDE Optimized                   ║
# ╚═══════════════════════════════════════════════════════════════════╝

## Overview

SynthGrid is a standalone desktop application that uses AI (Ollama or Groq)
to create, read, manipulate, and analyze Excel files (.xlsx) through natural
language commands. Built with a cybersigilism dark industrial aesthetic.

## Features

### AI Backends
- **Ollama (Local)**: Fully offline inference via Llama 3.1, Qwen 2.5, Mistral, etc.
- **Groq (Cloud)**: High-speed cloud inference via Groq API

### Excel Engine (MCP Parity)
- Workbook CRUD (create, open, save)
- Sheet management (create, delete, rename, copy)
- Cell read/write (single cells and ranges)
- Formula injection (any Excel formula)
- Formatting (fonts, colors, borders, alignment)
- Charts (bar, line, scatter)
- Pivot summaries
- Conditional formatting (color scales, cell rules)
- Data validation
- AutoFilter
- Freeze panes
- Row/column insert/delete
- Merge/unmerge cells

### UI / UX
- **Command Matrix**: Natural language input
- **Neural Uplink**: Real-time agent visibility console
- **Grid Render**: Live Excel data preview
- **Node Selector**: AI model switcher with connection testing

## Installation

```bash
# Install dependencies
pip install PyQt6 openpyxl pandas groq ollama

# Launch
python3 synthgrid.py
```

Or use the launch script:
```bash
chmod +x launch.sh
./launch.sh [GROQ_API_KEY]
```

## Configuration

### Ollama (Local)
1. Install Ollama: https://ollama.com
2. Pull a model: `ollama pull llama3.1`
3. In SynthGrid: Select OLLAMA backend, choose model, click TEST CONNECTION

### Groq (Cloud)
1. Get API key: https://console.groq.com
2. Set env var: `export GROQ_API_KEY=gsk_...`
   Or paste directly in the NODE SELECTOR panel
3. Select GROQ backend, choose model

## Usage Examples

```
Create a Q3 sales report with 5 products, monthly data, SUM totals, and a bar chart
Make row 1 headers bold, white text on dark background, and freeze the top row
Add conditional formatting to highlight values above 1000 in green
Create a budget sheet with income, expenses categories, and balance calculations
Add a pivot summary of the sales data grouped by product
```

## Keyboard Shortcuts

- `Ctrl+Enter` — Execute command
- `Ctrl+S` — Save workbook
- `Ctrl+O` — Open workbook
- `Ctrl+N` — New workbook

## Architecture

```
synthgrid.py
├── ExcelEngine        — openpyxl wrapper, 25+ tools
├── TOOL_SCHEMAS       — JSON schemas for LLM function calling
├── AgentWorker        — QThread: prompt → tools → result loop
├── MainWindow         — PyQt6 UI
│   ├── NodeSelector   — Backend/model config
│   ├── LogConsole     — Colored agent log
│   ├── GridView       — Live data table
│   └── Toolbar        — File operations
└── GLOBAL_STYLE       — Cybersigilism CSS
```

## Requirements

- Python 3.10+
- PyQt6
- openpyxl
- pandas
- groq (for Groq backend)
- ollama (Python client, for Ollama backend)
- Ollama daemon running (for local inference)
