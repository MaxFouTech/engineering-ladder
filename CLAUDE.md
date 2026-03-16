# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Engineering Ladder Assessment tool based on "The CTO's New Engineering Ladder" by Etienne de Bruin. It defines 5 rungs: Apprentice, Builder, Architect, Multiplier, Strategist — each with observable behaviors scored 1-4.

Two independent artifacts serve the same purpose:
- **`index.html`** — Self-contained single-page web app (HTML/CSS/JS, no build step). Hosted on GitHub Pages. Includes interactive scoring, JSON/Excel export/import, and a fixed left-side ladder navigation.
- **`create_ladder.py`** — Python script that generates `engineering_ladder_assessment.xlsx` using openpyxl. Produces a Summary sheet + one sheet per rung with formulas for auto-scoring.

## Commands

### Generate the Excel file
```
uv run create_ladder.py
```
Requires `openpyxl`. Output: `engineering_ladder_assessment.xlsx`.

### Run the web app locally
Open `index.html` directly in a browser — no server needed (all inline, no dependencies).

## Architecture Notes

- **`index.html`** is a large (~2100 lines) single file containing all markup, styles, and JavaScript inline. The ladder data (rung definitions, behaviors, salary ranges) is embedded as JS objects. CSS uses custom properties (`--dark-blue`, `--rung1`–`--rung5`, etc.) for theming. Features JSON export/import and Excel export (via SheetJS CDN).
- **`create_ladder.py`** mirrors the same rung data as a Python list of dicts (`rungs`). The `build_rung_sheet()` function generates one Excel tab; the Summary sheet references rung sheets via cross-sheet formulas.
- **`images/`** contains assets referenced by `index.html`. Naming convention: `rung{N}-{name}-{type}.{ext}` (e.g., `rung1-apprentice-illustration.jpg`, `rung3-architect-square.jpg`). Types include `illustration`, `infographic`, `square`, `frame`, and `animation-loop`.
- **`.nojekyll`** is present for GitHub Pages compatibility.

## Key Constraint

The rung data (behaviors, salary ranges, promotion signals) must stay in sync between `index.html` and `create_ladder.py` if both are updated.
