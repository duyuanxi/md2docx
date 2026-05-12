# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Markdown to Word (.docx) converter with tkinter GUI. Supports template-based formatting (公务员/学生紧凑 + user-defined), system font detection, side-by-side Markdown editor + print-style Word preview with pagination, paper sizes (A4/A3/8K/16K), and AI-export blank line cleaning.

## Commands

```bash
# Run the GUI application
python app.py

# Double-click launcher (Windows, uses pythonw to suppress console)
start.bat

# Install dependencies
pip install -r requirements.txt

# Build standalone .exe with PyInstaller
pyinstaller md2docx.spec
# Output: dist/md2docx.exe (windowed, no console)
```

## Architecture

Three modules, ~1400 lines total:

- **`app.py`** (925 lines) — Full tkinter GUI. `App` class owns all widgets, theme, font detection (`_get_system_fonts()` using `tkinter.font.families()`), side-by-side PanedWindow (left: editable Markdown, right: canvas-based page-style preview with shadow frames), collapsible settings panel, template operations, real-time preview via `<KeyRelease>` + 40ms debounce. Calls `convert_text()` in background thread for export.

- **`converter.py`** (264 lines) — Markdown → .docx engine. `convert_text(md_text, docx_path, template, merge_blank)` is the main entry. Uses `markdown` lib → BeautifulSoup HTML walker → `python-docx` element builder. Handles: headings (with level-based spacing), paragraphs, inline styles, code blocks (gray bg per-line), blockquotes, tables, lists (tab-stop alignment). `_clean_ai_blank_lines()` strips whitespace-only lines and merges 3+ blank lines → 1. Paper size applied via `section.page_width/height`.

- **`templates.py`** (209 lines) — Built-in templates (`公务员`/`学生紧凑`), `TemplateManager` class with JSON persistence for user templates. Constants: `FONT_LIST`, `SIZE_LIST`, `COLOR_PRESETS`. User templates saved to `templates.json`.

## Key Design Decisions

- **Preview rendering**: Page-style — each page is a `tk.Frame` (shadow) containing a white `tk.Frame` with a `tk.Text` widget. Multi-page content splits into separate page widgets. Tags configured per-widget for actual template fonts/sizes/colors.
- **List formatting in docx**: Uses tab-stop alignment (`-` + `\t` + content) instead of hanging indent or `List Bullet` style, because CJK fonts lack standard bullet glyphs and `List Bullet` has font-baseline alignment issues.
- **Font detection**: `tkinter.font.families()` provides system fonts (232 detected on Windows). Priority CJK fonts listed first in dropdowns.
- **PyInstaller**: `console=False` in .spec for windowed mode. `DATA_DIR` in app.py uses `sys.frozen` detection for bundled data path.

## Dependencies

- `python-docx` — Word document generation
- `markdown` — Markdown → HTML parsing
- `beautifulsoup4` — HTML tree traversal
- `tkinter` — GUI (built-in)
- `pyinstaller` — Optional, for building standalone .exe
