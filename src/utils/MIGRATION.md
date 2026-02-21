# Splitting App12.jsx — Migration Guide

## What was extracted (safe to do now)

These are pure logic files — no JSX, no React, no state.  
They are drop-in replacements. Nothing about how the app *works* changes.

```
src/
  utils/
    constants.js     ← theme, PALETTE, DEFAULT_LOOP_SIZES, LOOP_TYPES,
                        DIAGRAM_* constants, ATTACHED_TEST_CSV, SITE_VERSION
    math.js          ← clamp, safeNum, median, deepClone
    trim.js          ← bandForDelta, severity, chipColorFromLineId, groupColor
    parse.js         ← rowsFromCSVText, rowsFromSheetAOA, parseWideTableFromRows
    groups.js        ← makeDefaultRanges, buildInitialLineToGroup, getGroupOptions
    fileHelpers.js   ← downloadJSON, readFileText
    audio.js         ← playBeep
    index.js         ← barrel re-export (import anything from "../utils")

  components/
    ui/
      index.jsx      ← Panel, WarningBanner, NumInput, Select, Toggle,
                        FactorySelect, ImportStatusRadio, StatPill,
                        ControlPill, TogglePill, SegTabs
```

## How to apply the split

### Step 1 — Copy the new files into your repo

Place all the files above into `src/utils/` and `src/components/ui/`
alongside your existing `App12.jsx`.

### Step 2 — Update App12.jsx imports

Replace the top of `App12.jsx`:

```js
// REMOVE these (they are now in utils/)
import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
// ... all the const/function declarations up to line ~1047
```

```js
// ADD these instead
import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

import {
  SITE_VERSION, DEFAULT_LOOP_SIZES, LOOP_TYPES, theme, PALETTE,
  DIAGRAM_SCALE, DIAGRAM_W, DIAGRAM_H, ATTACHED_TEST_CSV,
  clamp, safeNum, median, deepClone,
  bandForDelta, severity, chipColorFromLineId, groupColor,
  rowsFromCSVText, rowsFromSheetAOA, parseWideTableFromRows,
  makeDefaultRanges, buildInitialLineToGroup, getGroupOptions,
  downloadJSON, readFileText,
  playBeep,
} from "./utils/index.js";

import {
  Panel, WarningBanner, NumInput, Select, Toggle, FactorySelect,
  ImportStatusRadio, StatPill, ControlPill, TogglePill, SegTabs,
} from "./components/ui/index.jsx";
```

### Step 3 — Delete the duplicate declarations

Once the imports are in place, delete lines 11–1047 from `App12.jsx`
(everything from `const playBeep` up to but NOT including `export default function App()`).

### Step 4 — Test locally

```bash
npm run dev
```

Check the app loads and the example CSV still parses correctly.

### Step 5 — Rename App12.jsx → App.jsx (optional but tidy)

Update `src/main.jsx` to `import App from "./App"`.

---

## Next extractions (future steps)

Once Step 1-4 above is stable, the next things to pull out are:

| Component | Lines (approx) | Extract to |
|---|---|---|
| `DiagramPreview` | 406–638 | `components/DiagramPreview.jsx` |
| `BlockTable` | 639–752 | `components/BlockTable.jsx` |
| `RearViewChart` | 8723–9370 | `components/charts/RearViewChart.jsx` |
| `PitchTrimChart` | 9436–9547 | `components/charts/PitchTrimChart.jsx` |
| `DeltaLineChart` | 9548–9670 | `components/charts/DeltaLineChart.jsx` |
| `WingProfileChart` | 9671–9786 | `components/charts/WingProfileChart.jsx` |
| `WingPitchViz` | 9372–9435 | `components/charts/WingPitchViz.jsx` |
| `RiggingDiagramPanel` | 1500–3878 | `components/RiggingDiagramPanel.jsx` |

Each extraction follows the same pattern:
1. Cut the function out of App12.jsx
2. Paste into its own file with the right imports
3. Add `export` keyword
4. Import it back in App.jsx
