# AI Prompt Template — Paraglider Trim Tuning

Use this when asking Claude or ChatGPT to make changes to the app.
Paste the relevant file(s) only — not the whole codebase.

---

## Project context (always include this)

```
Paraglider line trim analysis tool — Trim Tuning v1.2
Built with: React 18, Vite 5, SheetJS (xlsx)
Deployed to: GitHub Pages at /Paraglider-trim-tuning/
```

## File structure

```
src/
  App.jsx                          ← Main app. Contains all state + TrimWorkflow component (the full app UI)
  utils/
    constants.js                   ← theme, PALETTE, DEFAULT_LOOP_SIZES, LOOP_TYPES, DIAGRAM_* constants
    math.js                        ← clamp, safeNum, median, deepClone
    trim.js                        ← bandForDelta, severity, chipColorFromLineId, groupColor
    parse.js                       ← rowsFromCSVText, rowsFromSheetAOA, parseWideTableFromRows
    groups.js                      ← makeDefaultRanges, buildInitialLineToGroup, getGroupOptions
    fileHelpers.js                 ← downloadJSON, readFileText
    audio.js                       ← playBeep
    index.js                       ← barrel re-export
  components/
    ui/index.jsx                   ← Panel, WarningBanner, NumInput, Select, Toggle, FactorySelect,
                                      ImportStatusRadio, StatPill, ControlPill, TogglePill, SegTabs
    DiagramPreview.jsx             ← SVG drag-and-drop line grouping diagram
    BlockTable.jsx                 ← Per-line measurement results table
    charts/
      RearViewChart.jsx            ← Wing rear-view chart with loop adjustments
      WingPitchViz.jsx             ← Wing pitch visualisation
      PitchTrimChart.jsx           ← Pitch trim delta chart
      DeltaLineChart.jsx           ← Per-line delta chart (before/after)
      WingProfileChart.jsx         ← Wing profile delta per maillon group
```

## Key concepts

- **Lines** are paraglider lines identified as A1-A16, B1-B16, C1-C16, D1-D16 (and BR for brakes)
- **Nominal (Soll)** = factory specified length in mm
- **Measured (Ist L / Ist R)** = actual measured length left/right side
- **Delta (Δ)** = measured − nominal (after correction offset)
- **Tolerance** = acceptable deviation in mm (default 10mm)
- **Band colours**: green ≤ 4mm, yellow > 4mm but within tolerance, red ≥ tolerance
- **Loop types**: SL, DL, TL, AS, AS+, AS++ — each has a fixed mm offset applied to trim cuts
- **Correction** = global offset applied to all measurements (accounts for measuring jig offset)
- **TrimWorkflow** = the main app component (Steps 1–4). Formerly misnamed "RiggingDiagramPanel"

## The 4 steps

1. **Import** — load CSV/XLSX or use manual grid or factory trim database
2. **Grouping** — assign lines to maillon groups via drag-and-drop diagram
3. **Loop sizing** — set loop type per group (affects trim cut calculation)
4. **Trim** — view colour-coded delta table + charts, export results

---

## Prompt templates

### To change something in the UI or logic:

```
Here is the relevant file from my paraglider trim tool:

[PASTE FILE HERE]

Project context: React 18 + Vite 5, deployed to GitHub Pages.
The main app component is called TrimWorkflow (in App.jsx).
Theme and constants are imported from ./utils/constants.js.
UI primitives (Panel, StatPill, etc.) are in ./components/ui/index.jsx.

Please [DESCRIBE CHANGE].

Return only the modified file, no explanation needed.
```

### To change a chart:

```
Here is the chart component from my paraglider trim tool:

[PASTE e.g. src/components/charts/RearViewChart.jsx]

This chart receives these props: [list them]
It uses: theme from utils/constants.js, severity/bandForDelta from utils/trim.js

Please [DESCRIBE CHANGE].

Return only the modified file.
```

### To change parsing logic:

```
Here is the CSV/XLSX parser from my paraglider trim tool:

[PASTE src/utils/parse.js]

Input format: CSV with row 0 = meta keys, row 1 = meta values,
row 2 = column headers (A, Soll, Ist L, Ist R, B, ...), rows 3+ = data.

Please [DESCRIBE CHANGE].

Return only the modified file.
```

### To add a new feature:

```
I'm building a feature for my paraglider trim tool.

Project: React 18 + Vite 5, dark theme (see theme object in utils/constants.js)
Main component: TrimWorkflow in src/App.jsx
UI primitives available: Panel, StatPill, ControlPill, TogglePill, SegTabs, Toggle, Select

Feature request: [DESCRIBE FEATURE]

Relevant existing files:
[PASTE only the files that touch this feature]

Please implement this. Keep the same code style (inline styles using theme object,
functional components, hooks). Return each modified file separately.
```

---

## Rules for AI when editing this codebase

- Use the `theme` object for all colours — never hardcode colour values
- Inline styles only (no CSS files, no Tailwind)
- Functional components + hooks only (no class components)
- Keep `safeNum()` for all numeric parsing — never use raw `Number()` or `parseFloat()` on user input
- `clamp()` for any bounded numeric input
- Do NOT rewrite files that aren't relevant to the change
- Do NOT rename existing exports — other files import them by name
- Build target is es2018+ — modern JS is fine, no IE polyfills needed
