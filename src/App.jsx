import React, { useMemo, useState } from "react";

/**
 * Trim Tuning App — App.jsx (single file)
 *
 * Includes:
 * - Wide CSV import (German layout: Eingabe / Toleranz / Korrektur; A/B/C/D blocks)
 * - Compact A/B/C/D tables (2x2) with red/yellow highlight:
 *    RED:    |Δ| >= tolerance
 *    YELLOW: tolerance-3 <= |Δ| < tolerance
 * - Wing profile mapping line IDs -> groups (AR1/BR2/...)
 * - What-if group adjustments (mm) + factory target plan w/ A/B/C/D filter checkboxes
 * - BEFORE vs AFTER overlay line chart
 * - 3D-ish wing profile column chart (isometric SVG)
 * - Friendly wing profile editor modal + pop-out
 *
 * NEW:
 * - Maillon loop system:
 *   - Loop types with editable mm deltas
 *   - Per line/side loop selection (only one loop type per maillon per line side)
 *   - Bulk apply to All / Letter / Group; Mirror L->R, R->L; Sync
 * - Chart "Before trimming" = loops only, "After" = loops + adjustments, optional Original reference
 */

const BUILTIN_PROFILES = {
  "Speedster 3 (starter mapping)": {
    name: "Speedster 3 (starter mapping)",
    mmPerLoop: 10,
    mapping: {
      A: [
        [1, 4, "AR1"],
        [5, 8, "AR2"],
        [9, 16, "AR3"],
      ],
      B: [
        [1, 4, "BR1"],
        [5, 8, "BR2"],
        [9, 16, "BR3"],
      ],
      C: [
        [1, 4, "CR1"],
        [5, 8, "CR2"],
        [9, 12, "CR3"],
        [13, 16, "CR4"],
      ],
      D: [
        [1, 4, "DR1"],
        [5, 8, "DR2"],
        [9, 14, "DR3"],
      ],
    },
  },
};

/* ------------------------- CSV parsing ------------------------- */

function parseDelimited(text) {
  const lines = text
    .replace(/\uFEFF/g, "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  if (!lines.length) return { delim: ",", grid: [] };

  const first = lines[0];
  const counts = {
    ",": (first.match(/,/g) || []).length,
    ";": (first.match(/;/g) || []).length,
    "\t": (first.match(/\t/g) || []).length,
  };
  const delim =
    Object.entries(counts).sort((a, b) => b[1] - a[1])[0][1] > 0
      ? Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0]
      : ",";

  const grid = lines.map((l) => l.split(delim).map((c) => c.trim()));
  return { delim, grid };
}

function isWideFormat(grid) {
  if (!grid || grid.length < 4) return false;
  const r0 = (grid[0] || []).join(" ").toLowerCase();
  const r2 = (grid[2] || []).join(" ").toLowerCase();
  return (
    r0.includes("eingabe") &&
    r0.includes("toleranz") &&
    r0.includes("korrektur") &&
    r2.includes("soll")
  );
}

function n(x) {
  const v = parseFloat(String(x ?? "").replace(",", "."));
  return Number.isFinite(v) ? v : null;
}

function toCSV(delim, rows) {
  const esc = (v) => {
    const s = `${v ?? ""}`;
    if (/[",\n\r]/.test(s)) return `"${s.replaceAll('"', '""')}"`;
    return s;
  };
  return rows.map((r) => r.map(esc).join(delim)).join("\n") + "\n";
}

function downloadText(filename, text, mime = "text/csv") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function parseWide(grid) {
  const metaValues = grid[1] || [];
  const meta = {
    input1: metaValues[0] || "",
    input2: metaValues[1] || "",
    tolerance: n(metaValues[2]) ?? 0,
    correction: n(metaValues[3]) ?? 0,
  };

  const rows = [];
  for (let r = 3; r < grid.length; r++) {
    const row = grid[r] || [];
    const blocks = [
      { k: "A", i: 0 },
      { k: "B", i: 4 },
      { k: "C", i: 8 },
      { k: "D", i: 12 },
    ];

    const entry = { A: null, B: null, C: null, D: null };

    for (const b of blocks) {
      const line = (row[b.i] || "").trim();
      if (!line) continue;

      const nominal = n(row[b.i + 1]);
      const measL = n(row[b.i + 2]);
      const measR = n(row[b.i + 3]);

      entry[b.k] = { line, nominal, measL, measR };
    }

    if (entry.A || entry.B || entry.C || entry.D) rows.push(entry);
  }

  return { meta, rows };
}

/* ------------------------- Group mapping ------------------------- */

function parseLineId(lineId) {
  const m = String(lineId || "").trim().match(/^([A-Za-z])\s*0*([0-9]+)$/);
  if (!m) return null;
  return { prefix: m[1].toUpperCase(), num: parseInt(m[2], 10) };
}

function groupForLine(profile, lineId) {
  const p = parseLineId(lineId);
  if (!p) return null;
  const rules = profile?.mapping?.[p.prefix];
  if (!rules) return null;
  for (const [min, max, groupName] of rules) {
    if (p.num >= min && p.num <= max) return groupName;
  }
  return null;
}

function groupLetter(groupName) {
  const m = String(groupName || "").match(/^([A-D])R/i);
  return m ? m[1].toUpperCase() : null;
}

function groupSortKey(g) {
  const m = String(g).match(/^([A-D])R(\d+)$/i);
  if (m) return `${m[1].toUpperCase()}-${m[2].padStart(2, "0")}`;
  return g;
}

function extractGroupNames(wideRows, profile) {
  const set = new Set();
  for (const r of wideRows || []) {
    for (const k of ["A", "B", "C", "D"]) {
      const b = r?.[k];
      if (!b?.line) continue;
      const g = groupForLine(profile, b.line);
      if (g) set.add(g);
    }
  }
  if (!set.size && profile?.mapping) {
    for (const prefix of Object.keys(profile.mapping)) {
      for (const [, , g] of profile.mapping[prefix]) set.add(g);
    }
  }
  return Array.from(set).sort((a, b) => groupSortKey(a).localeCompare(groupSortKey(b)));
}

function getAllLinesFromWide(wideRows) {
  const seen = new Set();
  const out = [];
  for (const r of wideRows || []) {
    for (const letter of ["A", "B", "C", "D"]) {
      const b = r?.[letter];
      const lineId = b?.line;
      if (!lineId) continue;
      if (seen.has(lineId)) continue;
      seen.add(lineId);
      out.push({ lineId, letter });
    }
  }
  out.sort((a, b) => {
    const pa = parseLineId(a.lineId);
    const pb = parseLineId(b.lineId);
    const la = pa?.prefix || a.letter;
    const lb = pb?.prefix || b.letter;
    if (la !== lb) return la.localeCompare(lb);
    return (pa?.num ?? 0) - (pb?.num ?? 0);
  });
  return out;
}

function fmt(v, d = 0) {
  return Number.isFinite(v) ? v.toFixed(d) : "–";
}
function fmtSigned(v, d = 0) {
  if (!Number.isFinite(v)) return "–";
  return `${v > 0 ? "+" : ""}${v.toFixed(d)}`;
}
function avg(nums) {
  const v = nums.filter((x) => Number.isFinite(x));
  if (!v.length) return null;
  return v.reduce((a, b) => a + b, 0) / v.length;
}

/* ------------------------- Core math ------------------------- */

function getAdjustment(adjustments, groupName, side) {
  const key = `${groupName}|${side}`;
  return Number.isFinite(adjustments[key]) ? adjustments[key] : 0;
}

function deltaMm({ nominal, measured, correction, adjustment }) {
  if (nominal == null || measured == null) return null;
  return measured + (correction || 0) + (adjustment || 0) - nominal;
}

function severity(delta, tolerance) {
  if (!Number.isFinite(delta)) return "none";
  const a = Math.abs(delta);
  const tol = tolerance || 0;
  if (tol <= 0) return "ok";
  const warnBand = Math.max(0, tol - 3);
  if (a >= tol) return "red";
  if (a >= warnBand) return "yellow";
  return "ok";
}

/* ------------------------- Overlay Line Chart ------------------------- */

function makeLinePath(points) {
  if (!points.length) return "";
  return points
    .map((p, i) => `${i === 0 ? "M" : "L"} ${p.x.toFixed(2)} ${p.y.toFixed(2)}`)
    .join(" ");
}

function ChartOverlay({ series1, series2, series3, width = 1050, height = 360 }) {
  // series1 = optional reference (e.g. Original) dashed
  // series2 = "Before trimming" dotted
  // series3 = "After trimming" solid
  const padding = { l: 45, r: 16, t: 16, b: 28 };

  const maxLen = Math.max(
    0,
    ...(series1 ? series1.map((s) => s.values.length) : []),
    ...series2.map((s) => s.values.length),
    ...series3.map((s) => s.values.length)
  );

  const allVals = [
    ...(series1 ? series1.flatMap((s) => s.values) : []),
    ...series2.flatMap((s) => s.values),
    ...series3.flatMap((s) => s.values),
  ].filter(Number.isFinite);

  const minV = allVals.length ? Math.min(...allVals) : -10;
  const maxV = allVals.length ? Math.max(...allVals) : 10;

  const rangePad = (maxV - minV) * 0.1 || 5;
  const yMin = minV - rangePad;
  const yMax = maxV + rangePad;

  const plotW = width - padding.l - padding.r;
  const plotH = height - padding.t - padding.b;

  const xFor = (i) => (maxLen <= 1 ? padding.l : padding.l + (i / (maxLen - 1)) * plotW);
  const yFor = (v) => {
    const t = (v - yMin) / (yMax - yMin || 1);
    return padding.t + (1 - t) * plotH;
  };

  const gridCount = 6;
  const grid = Array.from({ length: gridCount + 1 }, (_, i) => {
    const t = i / gridCount;
    const v = yMin + t * (yMax - yMin);
    return { v, y: yFor(v) };
  });

  const colors = { A: "#b084ff", B: "#74d77f", C: "#ff6b6b", D: "#6ea8fe" };

  const drawSeries = (series, style) =>
    series.map((s) => {
      if (!s.visible) return null;
      const pts = s.values.map((v, i) => ({ x: xFor(i), y: yFor(Number.isFinite(v) ? v : 0) }));
      const path = makeLinePath(pts);
      return (
        <path
          key={`${style.keyPrefix}-${s.name}`}
          d={path}
          fill="none"
          stroke={colors[s.name] || "#eef1ff"}
          strokeWidth={style.strokeWidth}
          opacity={style.opacity}
          strokeDasharray={style.dash}
        />
      );
    });

  return (
    <div style={{ overflowX: "auto" }}>
      <svg
        width={width}
        height={height}
        style={{ background: "#0e1018", borderRadius: 14, border: "1px solid #2a2f3f" }}
      >
        {grid.map((g, i) => (
          <g key={i}>
            <line x1={padding.l} x2={width - padding.r} y1={g.y} y2={g.y} stroke="rgba(170,177,195,0.18)" />
            <text x={8} y={g.y + 4} fontSize="11" fill="rgba(170,177,195,0.75)">
              {g.v.toFixed(0)}
            </text>
          </g>
        ))}

        <line x1={padding.l} x2={width - padding.r} y1={yFor(0)} y2={yFor(0)} stroke="rgba(170,177,195,0.35)" />

        {series1 ? drawSeries(series1, { keyPrefix: "ref", strokeWidth: 2.0, opacity: 0.45, dash: "8 6" }) : null}
        {drawSeries(series2, { keyPrefix: "before", strokeWidth: 2.2, opacity: 0.7, dash: "2 6" })}
        {drawSeries(series3, { keyPrefix: "after", strokeWidth: 2.9, opacity: 0.95, dash: "" })}

        {Array.from({ length: Math.min(maxLen, 32) }, (_, i) => i).map((i) => {
          const show = maxLen <= 16 ? true : i % 2 === 0;
          if (!show) return null;
          return (
            <text
              key={i}
              x={xFor(i)}
              y={height - 10}
              fontSize="11"
              fill="rgba(170,177,195,0.75)"
              textAnchor="middle"
            >
              {i + 1}
            </text>
          );
        })}
      </svg>
    </div>
  );
}

/* ------------------------- Wing Profile 3D-ish Chart ------------------------- */

function Wing3D({ groupStats, width = 1050, height = 420 }) {
  const lanes = { A: [], B: [], C: [], D: [] };

  for (const s of groupStats || []) {
    const letter = groupLetter(s.groupName);
    const numMatch = String(s.groupName).match(/R(\d+)$/i);
    const groupNum = numMatch ? parseInt(numMatch[1], 10) : 99;
    if (!letter || !lanes[letter]) continue;
    lanes[letter].push({ ...s, groupNum });
  }
  for (const k of ["A", "B", "C", "D"]) lanes[k].sort((a, b) => a.groupNum - b.groupNum);

  const all = Object.values(lanes).flat();
  const vals = all.map((x) => x.meanDelta).filter(Number.isFinite);
  const minV = vals.length ? Math.min(...vals) : -10;
  const maxV = vals.length ? Math.max(...vals) : 10;
  const rangePad = (maxV - minV) * 0.1 || 5;
  const yMin = minV - rangePad;
  const yMax = maxV + rangePad;

  const pad = 16;
  const centerX = width / 2;
  const baseY = height - 70;

  const spanStep = 70;
  const laneStep = 55;

  const barW = 18;
  const depthSkewX = 10;
  const depthSkewY = 8;

  const yForVal = (v) => {
    const plotH = 220;
    const t = (v - yMin) / (yMax - yMin || 1);
    const yTop = baseY - plotH;
    const yBottom = baseY + 40;
    return yBottom - t * (yBottom - yTop);
  };

  const laneOrder = { A: 0, B: 1, C: 2, D: 3 };

  const fillColor = {
    A: "rgba(176,132,255,0.95)",
    B: "rgba(116,215,127,0.95)",
    C: "rgba(255,107,107,0.95)",
    D: "rgba(110,168,254,0.95)",
  };

  function drawBar({ x, y0, y1, fill }) {
    const topY = Math.min(y0, y1);
    const bottomY = Math.max(y0, y1);
    const h = Math.max(1, bottomY - topY);

    const fx = x - barW / 2;
    const fy = topY;
    const fw = barW;
    const fh = h;

    const top = [
      { x: fx, y: fy },
      { x: fx + fw, y: fy },
      { x: fx + fw + depthSkewX, y: fy - depthSkewY },
      { x: fx + depthSkewX, y: fy - depthSkewY },
    ];

    const side = [
      { x: fx + fw, y: fy },
      { x: fx + fw, y: fy + fh },
      { x: fx + fw + depthSkewX, y: fy + fh - depthSkewY },
      { x: fx + fw + depthSkewX, y: fy - depthSkewY },
    ];

    return (
      <g>
        <polygon points={side.map((p) => `${p.x},${p.y}`).join(" ")} fill="rgba(255,255,255,0.08)" />
        <rect x={fx} y={fy} width={fw} height={fh} fill={fill} opacity="0.75" />
        <polygon points={top.map((p) => `${p.x},${p.y}`).join(" ")} fill="rgba(255,255,255,0.12)" />
      </g>
    );
  }

  const yZero = yForVal(0);

  const bars = [];
  for (const letter of ["D", "C", "B", "A"]) {
    const lane = lanes[letter];
    const depth = laneOrder[letter];
    const laneY = baseY - depth * laneStep;

    for (const s of lane) {
      const v = Number.isFinite(s.meanDelta) ? s.meanDelta : 0;
      const yVal = yForVal(v);

      const i = Math.max(1, s.groupNum) - 1;
      const dir = s.side === "L" ? -1 : 1;
      const x = centerX + dir * (spanStep * (i + 0.3));

      const xSkewed = x + depth * depthSkewX;
      const laneYSkewed = laneY - depth * depthSkewY;

      bars.push({
        letter,
        groupName: s.groupName,
        x: xSkewed,
        y0: laneYSkewed + (yZero - baseY),
        y1: laneYSkewed + (yVal - baseY),
        fill: fillColor[letter],
      });
    }
  }

  return (
    <div style={{ overflowX: "auto" }}>
      <svg width={width} height={height} style={{ background: "#0e1018", borderRadius: 14, border: "1px solid #2a2f3f" }}>
        {["A", "B", "C", "D"].map((letter) => {
          const depth = laneOrder[letter];
          const y = baseY - depth * laneStep - depth * depthSkewY;
          return (
            <g key={letter}>
              <line x1={pad} x2={width - pad} y1={y} y2={y} stroke="rgba(170,177,195,0.14)" />
              <text x={width - pad} y={y - 6} fontSize="12" fill="rgba(170,177,195,0.8)" textAnchor="end">
                {letter} (front→back)
              </text>
            </g>
          );
        })}

        <line x1={width / 2} x2={width / 2} y1={pad} y2={height - pad} stroke="rgba(170,177,195,0.10)" />

        {bars.map((b, idx) => (
          <g key={idx}>
            {drawBar({ x: b.x, y0: b.y0, y1: b.y1, fill: b.fill })}
            <text x={b.x} y={b.y0 + 16} fontSize="10" fill="rgba(170,177,195,0.75)" textAnchor="middle">
              {b.groupName}
            </text>
          </g>
        ))}

        <text x={pad} y={24} fontSize="14" fill="rgba(238,241,255,0.92)" fontWeight="700">
          Wing profile (3D view) — AFTER (loops + adjustments)
        </text>
        <text x={pad} y={44} fontSize="12" fill="rgba(170,177,195,0.85)">
          Height = avg Δ (mm). Left bars on left, right bars on right. A is front row, D is back row.
        </text>
      </svg>
    </div>
  );
}

/* ------------------------- Main App ------------------------- */

export default function App() {
  const [delim, setDelim] = useState(",");
  const [meta, setMeta] = useState({ input1: "", input2: "", tolerance: 0, correction: 0 });
  const [wideRows, setWideRows] = useState([]);

  // Profiles JSON (persisted)
  const [profileKey, setProfileKey] = useState(Object.keys(BUILTIN_PROFILES)[0]);
  const [profileJson, setProfileJson] = useState(() => {
    const saved = localStorage.getItem("wingProfilesJson");
    return saved || JSON.stringify({ ...BUILTIN_PROFILES }, null, 2);
  });

  const profiles = useMemo(() => {
    try {
      const obj = JSON.parse(profileJson);
      if (obj && typeof obj === "object") return obj;
    } catch {}
    return { ...BUILTIN_PROFILES };
  }, [profileJson]);

  const activeProfile =
    profiles[profileKey] || Object.values(profiles)[0] || Object.values(BUILTIN_PROFILES)[0];

  // Adjustments (persisted)
  const [adjustments, setAdjustments] = useState(() => {
    try {
      const s = localStorage.getItem("groupAdjustments");
      return s ? JSON.parse(s) : {};
    } catch {
      return {};
    }
  });

  // Loop types (persisted) — mm delta applied to measured length
  const [loopTypes, setLoopTypes] = useState(() => {
    try {
      const s = localStorage.getItem("loopTypes");
      return s
        ? JSON.parse(s)
        : {
            SL: 0,
            DL: -7,
            AS: -10,
            "AS+": -16,
            PH: -18,
            "LF++": -23,
          };
    } catch {
      return {
        SL: 0,
        DL: -7,
        AS: -10,
        "AS+": -16,
        PH: -18,
        "LF++": -23,
      };
    }
  });

  // Loop setup per line+side (persisted) — stores the selected loop TYPE
  const [loopSetup, setLoopSetup] = useState(() => {
    try {
      const s = localStorage.getItem("loopSetup");
      return s ? JSON.parse(s) : {};
    } catch {
      return {};
    }
  });

  // UI: manual adjustments
  const [adjGroup, setAdjGroup] = useState("AR1");
  const [adjSide, setAdjSide] = useState("Both"); // Both | L | R
  const [adjMm, setAdjMm] = useState("0");

  // Factory target filters
  const [targetUseA, setTargetUseA] = useState(true);
  const [targetUseB, setTargetUseB] = useState(true);
  const [targetUseC, setTargetUseC] = useState(true);
  const [targetUseD, setTargetUseD] = useState(true);

  // Chart controls
  const [chartSideMode, setChartSideMode] = useState("Avg"); // Avg | L | R
  const [showA, setShowA] = useState(true);
  const [showB, setShowB] = useState(true);
  const [showC, setShowC] = useState(true);
  const [showD, setShowD] = useState(true);
  const [showOriginalRef, setShowOriginalRef] = useState(false); // optional
  const [beforeMode, setBeforeMode] = useState("LoopsOnly"); // LoopsOnly | Original

  // Bulk loop tool UI
  const [bulkScope, setBulkScope] = useState("All"); // All | A | B | C | D | Group
  const [bulkGroup, setBulkGroup] = useState("AR1");
  const [bulkSide, setBulkSide] = useState("Both"); // Both | L | R
  const [bulkLoopType, setBulkLoopType] = useState("SL");

  function persistAdjustments(next) {
    setAdjustments(next);
    localStorage.setItem("groupAdjustments", JSON.stringify(next));
  }
  function persistLoopTypes(next) {
    setLoopTypes(next);
    localStorage.setItem("loopTypes", JSON.stringify(next));
  }
  function persistLoopSetup(next) {
    setLoopSetup(next);
    localStorage.setItem("loopSetup", JSON.stringify(next));
  }

  function loopDeltaFor(lineId, side) {
    const key = `${lineId}|${side}`;
    const t = loopSetup[key] || "SL";
    const v = loopTypes?.[t];
    return Number.isFinite(v) ? v : 0;
  }

  function setCell(rowIndex, blockKey, field, value) {
    setWideRows((prev) => {
      const next = prev.slice();
      const row = { ...next[rowIndex] };
      const block = row[blockKey] ? { ...row[blockKey] } : null;
      if (!block) return prev;
      block[field] = value === "" ? null : n(value);
      row[blockKey] = block;
      next[rowIndex] = row;
      return next;
    });
  }

  function onImportFile(file) {
    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result || "");
      const parsed = parseDelimited(text);
      setDelim(parsed.delim);

      if (!isWideFormat(parsed.grid)) {
        alert("CSV not recognized. Expected wide format with Eingabe/Toleranz/Korrektur and A/B/C/D blocks.");
        return;
      }

      const w = parseWide(parsed.grid);
      setMeta(w.meta);
      setWideRows(w.rows);

      const groups = extractGroupNames(w.rows, activeProfile);
      if (groups.length) {
        setAdjGroup(groups[0]);
        setBulkGroup(groups[0]);
      }
    };
    reader.readAsText(file);
  }

  function exportWideCSV() {
    const rows = [];
    rows.push(["Input 1", "Input 2", "Tolerance", "Correction"]);
    rows.push([meta.input1, meta.input2, meta.tolerance, meta.correction]);
    rows.push([
      "A", "Nominal", "Measured L", "Measured R",
      "B", "Nominal", "Measured L", "Measured R",
      "C", "Nominal", "Measured L", "Measured R",
      "D", "Nominal", "Measured L", "Measured R",
    ]);

    for (const r of wideRows) {
      const out = [];
      for (const k of ["A", "B", "C", "D"]) {
        const b = r[k];
        if (!b) out.push("", "", "", "");
        else out.push(b.line ?? "", b.nominal ?? "", b.measL ?? "", b.measR ?? "");
      }
      rows.push(out);
    }

    downloadText(`measurement_export_${new Date().toISOString().slice(0, 10)}.csv`, toCSV(delim, rows), "text/csv");
  }

  const allLines = useMemo(() => getAllLinesFromWide(wideRows), [wideRows]);

  const compactBlocks = useMemo(() => {
    const blocks = { A: [], B: [], C: [], D: [] };
    for (let i = 0; i < wideRows.length; i++) {
      for (const k of ["A", "B", "C", "D"]) {
        const b = wideRows[i][k];
        if (!b) continue;
        blocks[k].push({ rowIndex: i, ...b });
      }
    }
    return blocks;
  }, [wideRows]);

  const allGroupNames = useMemo(() => extractGroupNames(wideRows, activeProfile), [wideRows, activeProfile]);

  function applyGroupAdjustment() {
    const mm = n(adjMm);
    if (!Number.isFinite(mm)) return;

    const next = { ...adjustments };
    if (adjSide === "Both") {
      next[`${adjGroup}|L`] = (Number.isFinite(next[`${adjGroup}|L`]) ? next[`${adjGroup}|L`] : 0) + mm;
      next[`${adjGroup}|R`] = (Number.isFinite(next[`${adjGroup}|R`]) ? next[`${adjGroup}|R`] : 0) + mm;
    } else {
      next[`${adjGroup}|${adjSide}`] = (Number.isFinite(next[`${adjGroup}|${adjSide}`]) ? next[`${adjGroup}|${adjSide}`] : 0) + mm;
    }
    persistAdjustments(next);
    setAdjMm("0");
  }

  function resetAdjustments() {
    persistAdjustments({});
  }

  // Bulk loop tools
  function applyBulkLoop() {
    if (!allLines.length) return;

    const selected = allLines.filter(({ lineId, letter }) => {
      if (bulkScope === "All") return true;
      if (["A", "B", "C", "D"].includes(bulkScope)) return letter === bulkScope;
      if (bulkScope === "Group") {
        const g = groupForLine(activeProfile, lineId);
        return g === bulkGroup;
      }
      return false;
    });

    const next = { ...loopSetup };
    for (const { lineId } of selected) {
      if (bulkSide === "Both" || bulkSide === "L") next[`${lineId}|L`] = bulkLoopType;
      if (bulkSide === "Both" || bulkSide === "R") next[`${lineId}|R`] = bulkLoopType;
    }
    persistLoopSetup(next);
  }

  function mirrorLoops(fromSide, toSide) {
    if (!allLines.length) return;
    const next = { ...loopSetup };
    for (const { lineId } of allLines) {
      const fromKey = `${lineId}|${fromSide}`;
      const toKey = `${lineId}|${toSide}`;
      next[toKey] = next[fromKey] || "SL";
    }
    persistLoopSetup(next);
  }

  function syncRightToLeft() {
    mirrorLoops("L", "R");
  }

  function resetLoopsToSL() {
    const ok = confirm("Reset ALL loop selections back to SL?");
    if (!ok) return;
    persistLoopSetup({});
  }

  /* ------------------------- Computations ------------------------- */
  // We compute three states:
  // - Original: no loops, no adjustments
  // - LoopsOnly: loops applied, no adjustments (Before trimming)
  // - After: loops + adjustments applied (After trimming)

  const computed = useMemo(() => {
    const corr = meta.correction || 0;
    const tol = meta.tolerance || 0;
    const mmPerLoop = Number.isFinite(activeProfile?.mmPerLoop) ? activeProfile.mmPerLoop : 10;
    const letters = ["A", "B", "C", "D"];

    const mapsOriginal = { A: new Map(), B: new Map(), C: new Map(), D: new Map() };
    const mapsLoopsOnly = { A: new Map(), B: new Map(), C: new Map(), D: new Map() };
    const mapsAfter = { A: new Map(), B: new Map(), C: new Map(), D: new Map() };

    const bucketAfter = new Map();

    for (const r of wideRows) {
      for (const letter of letters) {
        const b = r[letter];
        if (!b || b.nominal == null || !b.line) continue;

        const p = parseLineId(b.line);
        if (!p) continue;

        const groupName = groupForLine(activeProfile, b.line) || `${letter}?`;

        const loopL = loopDeltaFor(b.line, "L");
        const loopR = loopDeltaFor(b.line, "R");

        // Original
        const oL = deltaMm({ nominal: b.nominal, measured: b.measL, correction: corr, adjustment: 0 });
        const oR = deltaMm({ nominal: b.nominal, measured: b.measR, correction: corr, adjustment: 0 });
        mapsOriginal[letter].set(p.num, { dL: oL, dR: oR });

        // Loops only
        const lL = b.measL == null ? null : b.measL + loopL;
        const lR = b.measR == null ? null : b.measR + loopR;
        const loL = deltaMm({ nominal: b.nominal, measured: lL, correction: corr, adjustment: 0 });
        const loR = deltaMm({ nominal: b.nominal, measured: lR, correction: corr, adjustment: 0 });
        mapsLoopsOnly[letter].set(p.num, { dL: loL, dR: loR });

        // After (loops + adjustments)
        const adjL = getAdjustment(adjustments, groupName, "L");
        const adjR = getAdjustment(adjustments, groupName, "R");
        const aL = deltaMm({ nominal: b.nominal, measured: lL, correction: corr, adjustment: adjL });
        const aR = deltaMm({ nominal: b.nominal, measured: lR, correction: corr, adjustment: adjR });
        mapsAfter[letter].set(p.num, { dL: aL, dR: aR });

        if (Number.isFinite(aL)) {
          const key = `${groupName}|L`;
          if (!bucketAfter.has(key)) bucketAfter.set(key, []);
          bucketAfter.get(key).push(aL);
        }
        if (Number.isFinite(aR)) {
          const key = `${groupName}|R`;
          if (!bucketAfter.has(key)) bucketAfter.set(key, []);
          bucketAfter.get(key).push(aR);
        }
      }
    }

    const seriesFromMap = (m) => {
      const entries = Array.from(m.entries()).sort((a, b) => a[0] - b[0]);
      return entries.map(([, v]) => {
        if (chartSideMode === "L") return Number.isFinite(v.dL) ? v.dL : null;
        if (chartSideMode === "R") return Number.isFinite(v.dR) ? v.dR : null;
        const a = Number.isFinite(v.dL) ? v.dL : null;
        const b = Number.isFinite(v.dR) ? v.dR : null;
        if (a == null && b == null) return null;
        if (a == null) return b;
        if (b == null) return a;
        return (a + b) / 2;
      });
    };

    const originalSeries = {
      A: seriesFromMap(mapsOriginal.A),
      B: seriesFromMap(mapsOriginal.B),
      C: seriesFromMap(mapsOriginal.C),
      D: seriesFromMap(mapsOriginal.D),
    };
    const loopsOnlySeries = {
      A: seriesFromMap(mapsLoopsOnly.A),
      B: seriesFromMap(mapsLoopsOnly.B),
      C: seriesFromMap(mapsLoopsOnly.C),
      D: seriesFromMap(mapsLoopsOnly.D),
    };
    const afterSeries = {
      A: seriesFromMap(mapsAfter.A),
      B: seriesFromMap(mapsAfter.B),
      C: seriesFromMap(mapsAfter.C),
      D: seriesFromMap(mapsAfter.D),
    };

    const groupStatsAfter = [];
    for (const [key, arr] of bucketAfter.entries()) {
      const [groupName, side] = key.split("|");
      const mean = avg(arr);
      if (!Number.isFinite(mean)) continue;
      groupStatsAfter.push({ groupName, side, meanDelta: mean });
    }
    groupStatsAfter.sort((a, b) =>
      (groupSortKey(a.groupName) + a.side).localeCompare(groupSortKey(b.groupName) + b.side)
    );

    const suggestionsAfter = groupStatsAfter.map((s) => {
      const loopsSigned = Math.round(s.meanDelta / mmPerLoop);
      const action = loopsSigned > 0 ? "Shorten" : loopsSigned < 0 ? "Lengthen" : "No change";
      return {
        ...s,
        mmPerLoop,
        loopsSigned,
        loops: Math.abs(loopsSigned),
        outOfTol: tol > 0 ? Math.abs(s.meanDelta) >= tol : false,
        action,
      };
    });

    return { originalSeries, loopsOnlySeries, afterSeries, groupStatsAfter, suggestionsAfter };
  }, [
    wideRows,
    meta.correction,
    meta.tolerance,
    activeProfile,
    adjustments,
    chartSideMode,
    loopSetup,
    loopTypes,
  ]);

  const targetPlan = useMemo(() => {
    const mmPerLoop = Number.isFinite(activeProfile?.mmPerLoop) ? activeProfile.mmPerLoop : 10;

    const allow = (letter) => {
      if (letter === "A") return targetUseA;
      if (letter === "B") return targetUseB;
      if (letter === "C") return targetUseC;
      if (letter === "D") return targetUseD;
      return true;
    };

    const proposals = computed.suggestionsAfter
      .filter((s) => {
        const l = groupLetter(s.groupName);
        return l ? allow(l) : true;
      })
      .map((s) => {
        const loopsToZeroSigned = -Math.round(s.meanDelta / mmPerLoop);
        const extraMm = loopsToZeroSigned * mmPerLoop;
        return {
          groupName: s.groupName,
          side: s.side,
          currentMean: s.meanDelta,
          mmPerLoop,
          loopsToApplySigned: loopsToZeroSigned,
          extraMm,
          predictedMean: s.meanDelta + extraMm,
        };
      })
      .filter((p) => p.loopsToApplySigned !== 0);

    proposals.sort((a, b) => Math.abs(b.currentMean) - Math.abs(a.currentMean));
    return proposals;
  }, [computed.suggestionsAfter, activeProfile, targetUseA, targetUseB, targetUseC, targetUseD]);

  function applyTargetToAdjustments() {
    const next = { ...adjustments };
    for (const p of targetPlan) {
      const key = `${p.groupName}|${p.side}`;
      next[key] = (Number.isFinite(next[key]) ? next[key] : 0) + p.extraMm;
    }
    persistAdjustments(next);
  }

  /* ------------------------- Styles ------------------------- */

  const page = { minHeight: "100vh", background: "#0b0c10", color: "#eef1ff", fontFamily: "system-ui, sans-serif" };
  const wrap = { maxWidth: 1200, margin: "0 auto", padding: 16, display: "flex", flexDirection: "column", gap: 12 };
  const card = { border: "1px solid #2a2f3f", borderRadius: 14, background: "#11131a", padding: 12 };
  const muted = { color: "#aab1c3" };
  const btn = { padding: "10px 12px", borderRadius: 10, border: "1px solid #2a2f3f", background: "#0d0f16", color: "#eef1ff", cursor: "pointer", fontWeight: 650, fontSize: 13 };
  const btnDanger = { ...btn, border: "1px solid rgba(255,107,107,0.55)", background: "rgba(255,107,107,0.12)" };
  const btnWarn = { ...btn, border: "1px solid rgba(255,214,102,0.65)", background: "rgba(255,214,102,0.12)" };
  const input = { width: "100%", borderRadius: 10, border: "1px solid #2a2f3f", background: "#0d0f16", color: "#eef1ff", padding: "10px 10px", outline: "none" };
  const redCell = { border: "1px solid rgba(255,107,107,0.85)", background: "rgba(255,107,107,0.14)" };
  const yellowCell = { border: "1px solid rgba(255,214,102,0.95)", background: "rgba(255,214,102,0.14)" };

  const beforeSeries = beforeMode === "Original" ? computed.originalSeries : computed.loopsOnlySeries;

  return (
    <div style={page}>
      <div style={wrap}>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
          <div>
            <h1 style={{ margin: 0, fontSize: 22 }}>Trim Tuning</h1>
            <div style={{ ...muted, fontSize: 12, marginTop: 6 }}>
              Red: |Δ| ≥ tolerance. Yellow: within 3mm of tolerance. “Before trimming” can be loops-only.
            </div>
          </div>

          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
            <input
              type="file"
              accept=".csv,text/csv"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) onImportFile(f);
                e.target.value = "";
              }}
              style={{ color: "#aab1c3" }}
            />
            <button onClick={exportWideCSV} disabled={!wideRows.length} style={{ ...btn, opacity: wideRows.length ? 1 : 0.5 }}>
              Export CSV
            </button>
          </div>
        </div>

        <div style={{ ...card, borderColor: "rgba(255,204,102,0.5)", background: "rgba(255,204,102,0.08)" }}>
          <b>Safety:</b> This is analysis + “what-if” simulation. Real trimming can be dangerous—follow manufacturer/check-center procedures and verify after changes.
        </div>

        {/* Header + Profile */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div style={card}>
            <div style={{ fontWeight: 850, marginBottom: 10 }}>Session Header</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 10 }}>
              <Field label="Input 1" value={meta.input1} onChange={(v) => setMeta((m) => ({ ...m, input1: v }))} input={input} />
              <Field label="Input 2" value={meta.input2} onChange={(v) => setMeta((m) => ({ ...m, input2: v }))} input={input} />
              <Field label="Tolerance (mm)" value={meta.tolerance} onChange={(v) => setMeta((m) => ({ ...m, tolerance: n(v) ?? 0 }))} input={input} inputMode="numeric" />
              <Field label="Correction (mm)" value={meta.correction} onChange={(v) => setMeta((m) => ({ ...m, correction: n(v) ?? 0 }))} input={input} inputMode="numeric" />
            </div>
            <div style={{ ...muted, fontSize: 12, marginTop: 10 }}>
              Yellow rule: <b>Tolerance − 3 ≤ |Δ| &lt; Tolerance</b> (and Red at/over tolerance).
            </div>
          </div>

          <div style={card}>
            <div style={{ fontWeight: 850, marginBottom: 10 }}>Wing Profile</div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, alignItems: "end" }}>
              <div>
                <label style={{ ...muted, fontSize: 12 }}>Selected profile</label>
                <select value={profileKey} onChange={(e) => setProfileKey(e.target.value)} style={{ ...input, padding: "10px 10px" }}>
                  {Object.keys(profiles).map((k) => (
                    <option key={k} value={k}>{k}</option>
                  ))}
                </select>
              </div>

              <div>
                <label style={{ ...muted, fontSize: 12 }}>mm per loop</label>
                <input
                  value={activeProfile?.mmPerLoop ?? 10}
                  onChange={(e) => {
                    const v = n(e.target.value);
                    const next = { ...profiles };
                    const p = { ...(next[profileKey] || activeProfile) };
                    p.mmPerLoop = Number.isFinite(v) ? v : 10;
                    next[profileKey] = p;
                    const json = JSON.stringify(next, null, 2);
                    setProfileJson(json);
                    localStorage.setItem("wingProfilesJson", json);
                  }}
                  style={input}
                  inputMode="numeric"
                />
              </div>
            </div>

            <ProfileEditor
              profiles={profiles}
              profileKey={profileKey}
              setProfileKey={setProfileKey}
              profileJson={profileJson}
              setProfileJson={(next) => {
                setProfileJson(next);
                localStorage.setItem("wingProfilesJson", next);
              }}
            />
          </div>
        </div>

        {/* Loops card */}
        <div style={card}>
          <div style={{ fontWeight: 900, marginBottom: 8 }}>Maillon loop setup (before trimming)</div>
          <div style={{ color: "#aab1c3", fontSize: 12, lineHeight: 1.5 }}>
            Choose which <b>single loop type</b> is installed on each line’s maillon (Left/Right).
            Effective measured length is: <b>measured + loopDelta</b>.
          </div>

          <div style={{ height: 10 }} />

          <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
            <div style={{ fontWeight: 850, marginBottom: 8 }}>Loop types (editable)</div>
            <div style={{ color: "#aab1c3", fontSize: 12, marginBottom: 10 }}>
              Defaults are loaded. Edit the mm values to match your hardware/loops. Negative means the loop reduces length.
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 10 }}>
              {Object.entries(loopTypes).map(([name, mm]) => (
                <div key={name} style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 10, background: "#0d0f16" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                    <div style={{ fontWeight: 850 }}>{name}</div>
                    <div style={{ color: "#aab1c3", fontSize: 12 }}>mm</div>
                  </div>
                  <input
                    value={mm}
                    onChange={(e) => {
                      const v = n(e.target.value);
                      const next = { ...loopTypes, [name]: Number.isFinite(v) ? v : 0 };
                      persistLoopTypes(next);
                    }}
                    style={{
                      width: "100%",
                      borderRadius: 10,
                      border: "1px solid #2a2f3f",
                      background: "#0b0c10",
                      color: "#eef1ff",
                      padding: "10px 10px",
                      outline: "none",
                      marginTop: 8,
                      fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                      textAlign: "right",
                    }}
                    inputMode="numeric"
                  />
                </div>
              ))}
            </div>

            <div style={{ marginTop: 12, padding: 12, borderRadius: 14, border: "1px solid #2a2f3f", background: "#0b0c10" }}>
              <div style={{ fontWeight: 850, marginBottom: 8 }}>Bulk tools</div>
              <div style={{ color: "#aab1c3", fontSize: 12, marginBottom: 10 }}>
                Apply a loop type to many lines at once. This does <b>not</b> add multiple loops; it sets the single loop type on the maillon.
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 0.9fr 0.9fr 0.9fr auto", gap: 10, alignItems: "end" }}>
                <div>
                  <label style={{ color: "#aab1c3", fontSize: 12 }}>Scope</label>
                  <select value={bulkScope} onChange={(e) => setBulkScope(e.target.value)} style={{ ...input, padding: "10px 10px" }}>
                    <option value="All">All lines</option>
                    <option value="A">A lines</option>
                    <option value="B">B lines</option>
                    <option value="C">C lines</option>
                    <option value="D">D lines</option>
                    <option value="Group">Specific group…</option>
                  </select>
                </div>

                <div>
                  <label style={{ color: "#aab1c3", fontSize: 12 }}>Group</label>
                  <select
                    value={bulkGroup}
                    onChange={(e) => setBulkGroup(e.target.value)}
                    disabled={bulkScope !== "Group"}
                    style={{ ...input, padding: "10px 10px", opacity: bulkScope === "Group" ? 1 : 0.5 }}
                  >
                    {allGroupNames.map((g) => (
                      <option key={g} value={g}>{g}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label style={{ color: "#aab1c3", fontSize: 12 }}>Side</label>
                  <select value={bulkSide} onChange={(e) => setBulkSide(e.target.value)} style={{ ...input, padding: "10px 10px" }}>
                    <option value="Both">Both</option>
                    <option value="L">Left</option>
                    <option value="R">Right</option>
                  </select>
                </div>

                <div>
                  <label style={{ color: "#aab1c3", fontSize: 12 }}>Loop type</label>
                  <select value={bulkLoopType} onChange={(e) => setBulkLoopType(e.target.value)} style={{ ...input, padding: "10px 10px" }}>
                    {Object.keys(loopTypes).map((t) => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <button onClick={applyBulkLoop} style={btn} disabled={!allLines.length} title="Apply loop type to selected scope">
                    Apply
                  </button>
                </div>
              </div>

              <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button onClick={() => mirrorLoops("L", "R")} style={btn} disabled={!allLines.length}>Mirror L → R</button>
                <button onClick={() => mirrorLoops("R", "L")} style={btn} disabled={!allLines.length}>Mirror R → L</button>
                <button onClick={syncRightToLeft} style={btn} disabled={!allLines.length}>Set R = L</button>
                <button onClick={resetLoopsToSL} style={btnDanger} disabled={!allLines.length}>Reset loops to SL</button>
              </div>
            </div>
          </div>

          <div style={{ height: 12 }} />

          <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
            <div style={{ fontWeight: 850, marginBottom: 8 }}>Loops installed per line</div>
            <div style={{ color: "#aab1c3", fontSize: 12, marginBottom: 10 }}>
              Choose the loop type per line side. (Only one type per maillon.)
            </div>

            {!wideRows.length ? (
              <div style={{ color: "#aab1c3", fontSize: 12 }}>Upload a CSV to list the lines here.</div>
            ) : (
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 860 }}>
                  <thead>
                    <tr style={{ color: "#aab1c3", fontSize: 12 }}>
                      <th style={{ textAlign: "left", padding: "8px 8px" }}>Line</th>
                      <th style={{ textAlign: "left", padding: "8px 8px" }}>Group</th>
                      <th style={{ textAlign: "left", padding: "8px 8px" }}>Left loop</th>
                      <th style={{ textAlign: "right", padding: "8px 8px" }}>Left Δ(mm)</th>
                      <th style={{ textAlign: "left", padding: "8px 8px" }}>Right loop</th>
                      <th style={{ textAlign: "right", padding: "8px 8px" }}>Right Δ(mm)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {allLines.map(({ lineId, letter }) => {
                      const groupName = groupForLine(activeProfile, lineId) || `${letter}?`;

                      const kL = `${lineId}|L`;
                      const kR = `${lineId}|R`;
                      const tL = loopSetup[kL] || "SL";
                      const tR = loopSetup[kR] || "SL";

                      const dL = Number.isFinite(loopTypes?.[tL]) ? loopTypes[tL] : 0;
                      const dR = Number.isFinite(loopTypes?.[tR]) ? loopTypes[tR] : 0;

                      return (
                        <tr key={lineId} style={{ borderTop: "1px solid rgba(42,47,63,0.9)" }}>
                          <td style={{ padding: "8px 8px", fontWeight: 850 }}>{lineId}</td>
                          <td style={{ padding: "8px 8px", color: "#aab1c3", fontSize: 12 }}>{groupName}</td>

                          <td style={{ padding: "8px 8px" }}>
                            <select
                              value={tL}
                              onChange={(e) => {
                                const next = { ...loopSetup, [kL]: e.target.value };
                                persistLoopSetup(next);
                              }}
                              style={{ width: 130, borderRadius: 10, border: "1px solid #2a2f3f", background: "#0d0f16", color: "#eef1ff", padding: "8px 10px", outline: "none" }}
                            >
                              {Object.keys(loopTypes).map((name) => (
                                <option key={name} value={name}>{name}</option>
                              ))}
                            </select>
                          </td>

                          <td style={{ padding: "8px 8px", textAlign: "right", fontFamily: "ui-monospace, Menlo, Consolas, monospace", color: "#aab1c3" }}>
                            {dL > 0 ? `+${dL}` : `${dL}`}
                          </td>

                          <td style={{ padding: "8px 8px" }}>
                            <select
                              value={tR}
                              onChange={(e) => {
                                const next = { ...loopSetup, [kR]: e.target.value };
                                persistLoopSetup(next);
                              }}
                              style={{ width: 130, borderRadius: 10, border: "1px solid #2a2f3f", background: "#0d0f16", color: "#eef1ff", padding: "8px 10px", outline: "none" }}
                            >
                              {Object.keys(loopTypes).map((name) => (
                                <option key={name} value={name}>{name}</option>
                              ))}
                            </select>
                          </td>

                          <td style={{ padding: "8px 8px", textAlign: "right", fontFamily: "ui-monospace, Menlo, Consolas, monospace", color: "#aab1c3" }}>
                            {dR > 0 ? `+${dR}` : `${dR}`}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>

        {/* Adjustments + Target */}
        <div style={card}>
          <div style={{ fontWeight: 900, marginBottom: 8 }}>What-if group adjustments + Factory target</div>

          <div style={{ display: "grid", gridTemplateColumns: "1.2fr 0.8fr", gap: 12 }}>
            {/* Manual adjustments */}
            <div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 0.7fr 0.7fr 0.6fr", gap: 10, alignItems: "end" }}>
                <div>
                  <label style={{ ...muted, fontSize: 12 }}>Group</label>
                  <select value={adjGroup} onChange={(e) => setAdjGroup(e.target.value)} style={{ ...input, padding: "10px 10px" }}>
                    {allGroupNames.map((g) => (
                      <option key={g} value={g}>{g}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label style={{ ...muted, fontSize: 12 }}>Side</label>
                  <select value={adjSide} onChange={(e) => setAdjSide(e.target.value)} style={{ ...input, padding: "10px 10px" }}>
                    <option value="Both">Both</option>
                    <option value="L">Left</option>
                    <option value="R">Right</option>
                  </select>
                </div>
                <div>
                  <label style={{ ...muted, fontSize: 12 }}>Add mm</label>
                  <input value={adjMm} onChange={(e) => setAdjMm(e.target.value)} style={input} inputMode="numeric" />
                </div>
                <div>
                  <button onClick={applyGroupAdjustment} style={btn}>Apply</button>
                </div>
              </div>

              <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                <button onClick={resetAdjustments} style={btnDanger}>Reset all adjustments</button>
                <span style={{ ...muted, fontSize: 12 }}>(Positive mm = longer, Negative mm = shorter)</span>
              </div>

              <div style={{ marginTop: 10, ...muted, fontSize: 12 }}>
                Current adjustments:
                <div style={{ marginTop: 8, display: "grid", gap: 6 }}>
                  {Object.keys(adjustments).length ? (
                    Object.entries(adjustments)
                      .sort((a, b) => a[0].localeCompare(b[0]))
                      .map(([k, v]) => (
                        <div key={k} style={{ display: "flex", justifyContent: "space-between", gap: 10, padding: "8px 10px", border: "1px solid #2a2f3f", borderRadius: 12, background: "#0d0f16" }}>
                          <div style={{ fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                            {k} = {v > 0 ? `+${v}` : v} mm
                          </div>
                          <button
                            onClick={() => {
                              const next = { ...adjustments };
                              delete next[k];
                              persistAdjustments(next);
                            }}
                            style={{ ...btn, padding: "6px 8px", fontSize: 12 }}
                          >
                            remove
                          </button>
                        </div>
                      ))
                  ) : (
                    <div style={{ opacity: 0.8 }}>None</div>
                  )}
                </div>
              </div>
            </div>

            {/* Factory target */}
            <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
              <div style={{ fontWeight: 900, marginBottom: 8 }}>Factory target</div>
              <div style={{ ...muted, fontSize: 12, marginBottom: 10 }}>
                Generates loop-sized adjustments to bring each group average Δ toward 0 mm.
              </div>

              <div style={{ display: "grid", gap: 8, marginBottom: 10 }}>
                <div style={{ ...muted, fontSize: 12, fontWeight: 700 }}>Include in target:</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
                  <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                    <input type="checkbox" checked={targetUseA} onChange={(e) => setTargetUseA(e.target.checked)} /> A
                  </label>
                  <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                    <input type="checkbox" checked={targetUseB} onChange={(e) => setTargetUseB(e.target.checked)} /> B
                  </label>
                  <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                    <input type="checkbox" checked={targetUseC} onChange={(e) => setTargetUseC(e.target.checked)} /> C
                  </label>
                  <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                    <input type="checkbox" checked={targetUseD} onChange={(e) => setTargetUseD(e.target.checked)} /> D
                  </label>
                </div>
              </div>

              <button onClick={applyTargetToAdjustments} style={{ ...btnWarn, width: "100%" }} disabled={!targetPlan.length}>
                Apply target plan (adds to adjustments)
              </button>

              <div style={{ ...muted, fontSize: 12, marginTop: 10 }}>
                Target plan preview:
              </div>

              <div style={{ marginTop: 8, maxHeight: 240, overflow: "auto" }}>
                {!targetPlan.length ? (
                  <div style={{ ...muted, fontSize: 12 }}>No target changes (or no data loaded / all filtered out).</div>
                ) : (
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ color: "#aab1c3", fontSize: 12 }}>
                        <th style={{ textAlign: "left", padding: "6px 6px" }}>Group</th>
                        <th style={{ textAlign: "left", padding: "6px 6px" }}>Side</th>
                        <th style={{ textAlign: "right", padding: "6px 6px" }}>Now</th>
                        <th style={{ textAlign: "right", padding: "6px 6px" }}>Loops</th>
                      </tr>
                    </thead>
                    <tbody>
                      {targetPlan.map((p, idx) => (
                        <tr key={idx} style={{ borderTop: "1px solid rgba(42,47,63,0.9)" }}>
                          <td style={{ padding: "6px 6px" }}><b>{p.groupName}</b></td>
                          <td style={{ padding: "6px 6px" }}>{p.side}</td>
                          <td style={{ padding: "6px 6px", textAlign: "right", fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                            {fmtSigned(p.currentMean, 1)}
                          </td>
                          <td style={{ padding: "6px 6px", textAlign: "right", fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                            {p.loopsToApplySigned > 0 ? `+${p.loopsToApplySigned}` : `${p.loopsToApplySigned}`}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Overlay chart */}
        <div style={card}>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
            <div>
              <div style={{ fontWeight: 900 }}>Trim chart (Before vs After)</div>
              <div style={{ ...muted, fontSize: 12, marginTop: 6 }}>
                Dotted = Before trimming (loops only). Solid = After (loops + adjustments). Optional dashed = Original reference.
              </div>
            </div>

            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
              <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                Side mode
                <select value={chartSideMode} onChange={(e) => setChartSideMode(e.target.value)} style={{ ...input, padding: "6px 8px", width: 160 }}>
                  <option value="Avg">Avg (L/R)</option>
                  <option value="L">Left only</option>
                  <option value="R">Right only</option>
                </select>
              </label>

              <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                Before =
                <select value={beforeMode} onChange={(e) => setBeforeMode(e.target.value)} style={{ ...input, padding: "6px 8px", width: 160 }}>
                  <option value="LoopsOnly">Loops only</option>
                  <option value="Original">Original</option>
                </select>
              </label>

              <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                <input type="checkbox" checked={showOriginalRef} onChange={(e) => setShowOriginalRef(e.target.checked)} />
                Show original reference
              </label>

              <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                <input type="checkbox" checked={showA} onChange={(e) => setShowA(e.target.checked)} /> A
              </label>
              <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                <input type="checkbox" checked={showB} onChange={(e) => setShowB(e.target.checked)} /> B
              </label>
              <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                <input type="checkbox" checked={showC} onChange={(e) => setShowC(e.target.checked)} /> C
              </label>
              <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#aab1c3" }}>
                <input type="checkbox" checked={showD} onChange={(e) => setShowD(e.target.checked)} /> D
              </label>
            </div>
          </div>

          <div style={{ height: 10 }} />

          <ChartOverlay
            width={1050}
            height={360}
            series1={
              showOriginalRef
                ? [
                    { name: "A", values: computed.originalSeries.A.filter((x) => x !== null), visible: showA },
                    { name: "B", values: computed.originalSeries.B.filter((x) => x !== null), visible: showB },
                    { name: "C", values: computed.originalSeries.C.filter((x) => x !== null), visible: showC },
                    { name: "D", values: computed.originalSeries.D.filter((x) => x !== null), visible: showD },
                  ]
                : null
            }
            series2={[
              { name: "A", values: (beforeSeries.A || []).filter((x) => x !== null), visible: showA },
              { name: "B", values: (beforeSeries.B || []).filter((x) => x !== null), visible: showB },
              { name: "C", values: (beforeSeries.C || []).filter((x) => x !== null), visible: showC },
              { name: "D", values: (beforeSeries.D || []).filter((x) => x !== null), visible: showD },
            ]}
            series3={[
              { name: "A", values: computed.afterSeries.A.filter((x) => x !== null), visible: showA },
              { name: "B", values: computed.afterSeries.B.filter((x) => x !== null), visible: showB },
              { name: "C", values: computed.afterSeries.C.filter((x) => x !== null), visible: showC },
              { name: "D", values: computed.afterSeries.D.filter((x) => x !== null), visible: showD },
            ]}
          />
        </div>

        {/* 3D wing */}
        <div style={card}>
          <div style={{ fontWeight: 900, marginBottom: 8 }}>Wing profile (3D column chart)</div>
          <div style={{ ...muted, fontSize: 12, marginBottom: 10 }}>
            Shows the “shape” using group-average Δ after loops + adjustments. AR1 near center, higher Rn toward tips (mirrored left/right). A is front row, D is back.
          </div>
          <Wing3D groupStats={computed.groupStatsAfter} width={1050} height={420} />
        </div>

        {/* Compact tables */}
        <div style={card}>
          <div style={{ fontWeight: 900, marginBottom: 10 }}>Measurement tables (compact)</div>
          <div style={{ ...muted, fontSize: 12, marginBottom: 10 }}>
            Δ values include loops + adjustments (the current effective setup).
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 12 }}>
            <BlockTable title="A" rows={compactBlocks.A} meta={meta} activeProfile={activeProfile} adjustments={adjustments} loopDeltaFor={loopDeltaFor} input={input} redCell={redCell} yellowCell={yellowCell} setCell={setCell} blockKey="A" />
            <BlockTable title="B" rows={compactBlocks.B} meta={meta} activeProfile={activeProfile} adjustments={adjustments} loopDeltaFor={loopDeltaFor} input={input} redCell={redCell} yellowCell={yellowCell} setCell={setCell} blockKey="B" />
            <BlockTable title="C" rows={compactBlocks.C} meta={meta} activeProfile={activeProfile} adjustments={adjustments} loopDeltaFor={loopDeltaFor} input={input} redCell={redCell} yellowCell={yellowCell} setCell={setCell} blockKey="C" />
            <BlockTable title="D" rows={compactBlocks.D} meta={meta} activeProfile={activeProfile} adjustments={adjustments} loopDeltaFor={loopDeltaFor} input={input} redCell={redCell} yellowCell={yellowCell} setCell={setCell} blockKey="D" />
          </div>
        </div>
      </div>
    </div>
  );

  function Field({ label, value, onChange, inputMode, input }) {
    return (
      <div>
        <label style={{ fontSize: 12, color: "#aab1c3" }}>{label}</label>
        <input value={value} onChange={(e) => onChange(e.target.value)} style={input} inputMode={inputMode} />
      </div>
    );
  }
}

/* ------------------------- Compact table component ------------------------- */

function BlockTable({ title, rows, meta, activeProfile, adjustments, loopDeltaFor, input, redCell, yellowCell, setCell, blockKey }) {
  const corr = meta.correction || 0;
  const tol = meta.tolerance || 0;

  const styleFor = (sev) => (sev === "red" ? redCell : sev === "yellow" ? yellowCell : null);

  return (
    <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, overflow: "hidden", background: "#0e1018" }}>
      <div style={{ padding: 10, borderBottom: "1px solid #2a2f3f", fontWeight: 900 }}>{title} lines</div>

      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 620 }}>
          <thead>
            <tr style={{ color: "#aab1c3", fontSize: 12 }}>
              <Th>Line</Th>
              <Th>Group</Th>
              <Th align="right">Nominal</Th>
              <Th align="right">Measured L</Th>
              <Th align="right">Measured R</Th>
            </tr>
          </thead>

          <tbody>
            {!rows.length ? (
              <tr><Td colSpan={5} style={{ color: "#aab1c3" }}>No {title} rows found.</Td></tr>
            ) : (
              rows.map((b, idx) => {
                const groupName = groupForLine(activeProfile, b.line) || `${title}?`;

                const loopL = loopDeltaFor(b.line, "L");
                const loopR = loopDeltaFor(b.line, "R");

                const adjL = getAdjustment(adjustments, groupName, "L");
                const adjR = getAdjustment(adjustments, groupName, "R");

                const effL = b.measL == null ? null : b.measL + loopL;
                const effR = b.measR == null ? null : b.measR + loopR;

                const dL = deltaMm({ nominal: b.nominal, measured: effL, correction: corr, adjustment: adjL });
                const dR = deltaMm({ nominal: b.nominal, measured: effR, correction: corr, adjustment: adjR });

                const sevL = severity(dL, tol);
                const sevR = severity(dR, tol);

                return (
                  <tr key={`${b.line}-${idx}`} style={{ borderTop: "1px solid #2a2f3f" }}>
                    <Td><b>{b.line}</b></Td>
                    <Td style={{ color: "#aab1c3", fontSize: 12 }}>{groupName}</Td>
                    <Td align="right" style={{ fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                      {b.nominal ?? ""}
                    </Td>

                    <Td align="right">
                      <input
                        value={b.measL ?? ""}
                        onChange={(e) => setCell(b.rowIndex, blockKey, "measL", e.target.value)}
                        style={{
                          ...input,
                          ...(styleFor(sevL) || null),
                          width: 120,
                          textAlign: "right",
                          fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                        }}
                        inputMode="numeric"
                      />
                      <div style={{ color: "#aab1c3", fontSize: 11, marginTop: 6, fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                        loop {loopL > 0 ? `+${loopL}` : `${loopL}`} | Δ {Number.isFinite(dL) ? `${dL > 0 ? "+" : ""}${Math.round(dL)}mm` : "–"}
                      </div>
                    </Td>

                    <Td align="right">
                      <input
                        value={b.measR ?? ""}
                        onChange={(e) => setCell(b.rowIndex, blockKey, "measR", e.target.value)}
                        style={{
                          ...input,
                          ...(styleFor(sevR) || null),
                          width: 120,
                          textAlign: "right",
                          fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                        }}
                        inputMode="numeric"
                      />
                      <div style={{ color: "#aab1c3", fontSize: 11, marginTop: 6, fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                        loop {loopR > 0 ? `+${loopR}` : `${loopR}`} | Δ {Number.isFinite(dR) ? `${dR > 0 ? "+" : ""}${Math.round(dR)}mm` : "–"}
                      </div>
                    </Td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      <div style={{ padding: 10, color: "#aab1c3", fontSize: 12 }}>
        Yellow: within 3mm of tolerance. Red: at/over tolerance.
      </div>
    </div>
  );
}

/* ------------------------- Wing Profile Editor (friendly UI) ------------------------- */

function ProfileEditor({ profiles, profileKey, setProfileKey, profileJson, setProfileJson }) {
  const [open, setOpen] = React.useState(false);
  const [tab, setTab] = React.useState("builder"); // builder | json | help
  const [draftKey, setDraftKey] = React.useState(profileKey);
  const [draft, setDraft] = React.useState(() => deepClone(profiles[profileKey] || {}));
  const [message, setMessage] = React.useState("");

  React.useEffect(() => {
    setDraftKey(profileKey);
    setDraft(deepClone(profiles[profileKey] || {}));
  }, [profileKey]); // eslint-disable-line

  const allKeys = Object.keys(profiles || {});
  const current = profiles?.[draftKey] || null;

  function saveDraftToProfiles() {
    try {
      const next = { ...profiles };
      next[draftKey] = draft;
      const json = JSON.stringify(next, null, 2);
      setProfileJson(json);
      setProfileKey(draftKey);
      setMessage("Saved ✅");
      setTimeout(() => setMessage(""), 1500);
    } catch {
      setMessage("Could not save (invalid data).");
    }
  }

  function validateDraftClick() {
    const result = validateProfile(draftKey, draft);
    if (result.ok) setMessage("Looks good ✅");
    else setMessage(result.errors.join(" • "));
  }

  function addNewProfile() {
    const baseName = "New Wing Profile";
    let k = baseName;
    let i = 2;
    while ((profiles || {})[k]) k = `${baseName} ${i++}`;

    const nextDraft = { name: k, mmPerLoop: 10, mapping: { A: [], B: [], C: [], D: [] } };
    setDraftKey(k);
    setDraft(nextDraft);
    setMessage("Created a new blank profile. Add ranges below, then Save.");
  }

  function duplicateProfile() {
    const base = current || draft;
    const baseName = `${draftKey} (copy)`;
    let k = baseName;
    let i = 2;
    while ((profiles || {})[k]) k = `${baseName} ${i++}`;
    setDraftKey(k);
    setDraft(deepClone(base));
    setMessage("Duplicated profile. Rename it and edit ranges, then Save.");
  }

  function deleteProfile() {
    if (allKeys.length <= 1) {
      setMessage("You need at least 1 profile.");
      return;
    }
    const ok = confirm(`Delete profile "${draftKey}"?`);
    if (!ok) return;

    const next = { ...profiles };
    delete next[draftKey];
    const json = JSON.stringify(next, null, 2);
    setProfileJson(json);

    const remaining = Object.keys(next);
    setProfileKey(remaining[0]);
    setMessage("Deleted.");
  }

  function openPopout() {
    const w = window.open("", "_blank", "noopener,noreferrer");
    if (!w) {
      setMessage("Popup blocked. Allow popups then try again.");
      return;
    }

    const html = `
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Wing Profile JSON Editor</title>
  <style>
    body { font-family: system-ui, sans-serif; margin: 20px; background:#0b0c10; color:#eef1ff;}
    textarea { width:100%; min-height: 60vh; background:#0d0f16; color:#eef1ff; border:1px solid #2a2f3f; border-radius: 10px; padding: 12px; font-family: ui-monospace, Menlo, Consolas, monospace; font-size: 12px; outline:none;}
    .card { border:1px solid #2a2f3f; border-radius: 14px; background:#11131a; padding:12px; margin-bottom: 12px;}
    button { padding:10px 12px; border-radius: 10px; border:1px solid #2a2f3f; background:#0d0f16; color:#eef1ff; cursor:pointer; font-weight:700;}
    .muted { color:#aab1c3; font-size: 12px; line-height: 1.4;}
    code { color:#b084ff; }
  </style>
</head>
<body>
  <div class="card">
    <h2 style="margin:0 0 6px 0;">Wing Profiles JSON (Pop-out)</h2>
    <div class="muted">
      Tips:
      <ul>
        <li><code>mmPerLoop</code>: used for loop-size suggestions + target plan.</li>
        <li><code>mapping.A</code>: ranges mapping A line numbers to group labels (example: <code>[1,4,"AR1"]</code>).</li>
        <li>Ranges should not overlap within the same letter.</li>
      </ul>
      Workflow: edit JSON → click “Copy” → paste back in the app editor → Save.
    </div>
  </div>
  <textarea id="t"></textarea>
  <div style="display:flex; gap:10px; margin-top:10px; flex-wrap:wrap;">
    <button onclick="navigator.clipboard.writeText(document.getElementById('t').value)">Copy to clipboard</button>
    <button onclick="document.getElementById('t').value='';">Clear</button>
  </div>
  <script>
    const initial = ${JSON.stringify(profileJson)};
    document.getElementById('t').value = initial;
  </script>
</body>
</html>`;
    w.document.write(html);
    w.document.close();
  }

  const modalStyle = {
    position: "fixed",
    inset: 0,
    background: "rgba(0,0,0,0.55)",
    display: open ? "flex" : "none",
    alignItems: "center",
    justifyContent: "center",
    padding: 12,
    zIndex: 50,
  };

  const panelStyle = {
    width: "min(1100px, 100%)",
    maxHeight: "90vh",
    overflow: "auto",
    borderRadius: 16,
    border: "1px solid #2a2f3f",
    background: "#11131a",
    padding: 14,
  };

  const btn = {
    padding: "10px 12px",
    borderRadius: 10,
    border: "1px solid #2a2f3f",
    background: "#0d0f16",
    color: "#eef1ff",
    cursor: "pointer",
    fontWeight: 750,
    fontSize: 13,
  };

  const btnWarn = { ...btn, border: "1px solid rgba(255,214,102,0.65)", background: "rgba(255,214,102,0.12)" };
  const btnDanger = { ...btn, border: "1px solid rgba(255,107,107,0.55)", background: "rgba(255,107,107,0.12)" };

  const input = {
    width: "100%",
    borderRadius: 10,
    border: "1px solid #2a2f3f",
    background: "#0d0f16",
    color: "#eef1ff",
    padding: "10px 10px",
    outline: "none",
  };

  return (
    <>
      <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        <button style={btn} onClick={() => setOpen(true)}>Open Profile Editor</button>
        <button style={btn} onClick={openPopout}>Pop-out editor</button>
        <span style={{ color: "#aab1c3", fontSize: 12 }}>(Guided editor + JSON view + validation)</span>
      </div>

      <div style={modalStyle} onMouseDown={() => setOpen(false)}>
        <div style={panelStyle} onMouseDown={(e) => e.stopPropagation()}>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
            <div>
              <div style={{ fontWeight: 900, fontSize: 16 }}>Wing Profile Editor</div>
              <div style={{ color: "#aab1c3", fontSize: 12, marginTop: 4 }}>
                Edit how line IDs map into riser groups (AR1/BR2/CR4...). This affects grouping, target, and the 3D wing view.
              </div>
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button style={btn} onClick={validateDraftClick}>Validate</button>
              <button style={btnWarn} onClick={saveDraftToProfiles}>Save</button>
              <button style={btn} onClick={() => setOpen(false)}>Close</button>
            </div>
          </div>

          {message ? (
            <div style={{ marginTop: 10, padding: "10px 12px", borderRadius: 12, border: "1px solid #2a2f3f", background: "#0d0f16", color: "#eef1ff", fontSize: 12 }}>
              {message}
            </div>
          ) : null}

          <div style={{ display: "flex", gap: 8, marginTop: 12, flexWrap: "wrap" }}>
            <TabButton active={tab === "builder"} onClick={() => setTab("builder")}>Builder</TabButton>
            <TabButton active={tab === "json"} onClick={() => setTab("json")}>Raw JSON</TabButton>
            <TabButton active={tab === "help"} onClick={() => setTab("help")}>Hints & Tips</TabButton>
          </div>

          <div style={{ marginTop: 12, border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 0.9fr 0.9fr", gap: 10, alignItems: "end" }}>
              <div>
                <label style={{ color: "#aab1c3", fontSize: 12 }}>Profile</label>
                <select
                  value={draftKey}
                  onChange={(e) => {
                    const k = e.target.value;
                    setDraftKey(k);
                    setDraft(deepClone(profiles[k] || {}));
                    setMessage("");
                  }}
                  style={{ ...input, padding: "10px 10px" }}
                >
                  {allKeys.map((k) => (
                    <option key={k} value={k}>{k}</option>
                  ))}
                </select>
              </div>

              <div>
                <label style={{ color: "#aab1c3", fontSize: 12 }}>Display name</label>
                <input
                  value={draft?.name ?? draftKey}
                  onChange={(e) => setDraft((d) => ({ ...d, name: e.target.value }))}
                  style={input}
                />
              </div>

              <div>
                <label style={{ color: "#aab1c3", fontSize: 12 }}>mm per loop</label>
                <input
                  value={draft?.mmPerLoop ?? 10}
                  onChange={(e) => setDraft((d) => ({ ...d, mmPerLoop: n(e.target.value) ?? 10 }))}
                  style={input}
                  inputMode="numeric"
                />
              </div>
            </div>

            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 10 }}>
              <button style={btn} onClick={addNewProfile}>New</button>
              <button style={btn} onClick={duplicateProfile}>Duplicate</button>
              <button style={btnDanger} onClick={deleteProfile}>Delete</button>
            </div>
          </div>

          {tab === "builder" ? <ProfileBuilder draft={draft} setDraft={setDraft} /> : null}
          {tab === "json" ? <ProfileJsonView profileJson={profileJson} setProfileJson={setProfileJson} /> : null}
          {tab === "help" ? <ProfileHelp /> : null}
        </div>
      </div>
    </>
  );
}

function TabButton({ active, onClick, children }) {
  const style = {
    padding: "8px 10px",
    borderRadius: 10,
    border: "1px solid #2a2f3f",
    background: active ? "rgba(176,132,255,0.14)" : "#0d0f16",
    color: active ? "#eef1ff" : "#aab1c3",
    cursor: "pointer",
    fontWeight: 800,
    fontSize: 12,
  };
  return (
    <button style={style} onClick={onClick}>
      {children}
    </button>
  );
}

function ProfileBuilder({ draft, setDraft }) {
  const input = {
    width: "100%",
    borderRadius: 10,
    border: "1px solid #2a2f3f",
    background: "#0d0f16",
    color: "#eef1ff",
    padding: "10px 10px",
    outline: "none",
  };

  const mapping = draft?.mapping || { A: [], B: [], C: [], D: [] };

  function setLetter(letter, nextArr) {
    setDraft((d) => ({ ...d, mapping: { ...(d.mapping || {}), [letter]: nextArr } }));
  }

  function addRow(letter) {
    const arr = Array.isArray(mapping[letter]) ? mapping[letter] : [];
    setLetter(letter, [...arr, [1, 1, `${letter}R1`]]);
  }

  function updateRow(letter, idx, col, value) {
    const arr = Array.isArray(mapping[letter]) ? mapping[letter] : [];
    const next = arr.map((r, i) => (i === idx ? [...r] : r));
    const row = next[idx] || [1, 1, `${letter}R1`];

    if (col === "min" || col === "max") {
      const nv = n(value);
      row[col === "min" ? 0 : 1] = Number.isFinite(nv) ? nv : row[col === "min" ? 0 : 1];
    } else if (col === "group") {
      row[2] = value;
    }
    next[idx] = row;
    setLetter(letter, next);
  }

  function removeRow(letter, idx) {
    const arr = Array.isArray(mapping[letter]) ? mapping[letter] : [];
    setLetter(letter, arr.filter((_, i) => i !== idx));
  }

  const letters = ["A", "B", "C", "D"];

  return (
    <div style={{ marginTop: 12 }}>
      <div style={{ color: "#aab1c3", fontSize: 12, lineHeight: 1.5 }}>
        <b>Builder:</b> Add ranges like “A 1–4 → AR1”. The app maps line IDs like <b>A01</b>, <b>A2</b>, <b>A16</b>.
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 12, marginTop: 12 }}>
        {letters.map((L) => (
          <div key={L} style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
              <div style={{ fontWeight: 900 }}>{L} mapping</div>
              <button
                style={{ padding: "8px 10px", borderRadius: 10, border: "1px solid #2a2f3f", background: "#0d0f16", color: "#eef1ff", cursor: "pointer", fontWeight: 800, fontSize: 12 }}
                onClick={() => addRow(L)}
              >
                + Add range
              </button>
            </div>

            <div style={{ color: "#aab1c3", fontSize: 12, marginTop: 6 }}>
              Example: [{L}1–{L}4] → {L}R1 (group label should match your diagram like AR1/BR2).
            </div>

            <div style={{ marginTop: 10, display: "grid", gap: 8 }}>
              {(Array.isArray(mapping[L]) ? mapping[L] : []).length ? (
                (mapping[L] || []).map((r, idx) => (
                  <div key={idx} style={{ display: "grid", gridTemplateColumns: "0.8fr 0.8fr 1.2fr auto", gap: 8, alignItems: "center" }}>
                    <input style={input} value={r?.[0] ?? 1} onChange={(e) => updateRow(L, idx, "min", e.target.value)} inputMode="numeric" />
                    <input style={input} value={r?.[1] ?? 1} onChange={(e) => updateRow(L, idx, "max", e.target.value)} inputMode="numeric" />
                    <input style={input} value={r?.[2] ?? `${L}R1`} onChange={(e) => updateRow(L, idx, "group", e.target.value)} />
                    <button
                      style={{ padding: "8px 10px", borderRadius: 10, border: "1px solid rgba(255,107,107,0.55)", background: "rgba(255,107,107,0.12)", color: "#eef1ff", cursor: "pointer", fontWeight: 800, fontSize: 12 }}
                      onClick={() => removeRow(L, idx)}
                    >
                      ✕
                    </button>
                  </div>
                ))
              ) : (
                <div style={{ color: "#aab1c3", fontSize: 12 }}>No ranges yet. Click “Add range”.</div>
              )}
            </div>

            <div style={{ color: "#aab1c3", fontSize: 12, marginTop: 10 }}>
              Hint: Avoid overlaps (1–4 and 4–8). Gaps mean some lines won’t get grouped.
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function ProfileJsonView({ profileJson, setProfileJson }) {
  return (
    <div style={{ marginTop: 12 }}>
      <div style={{ color: "#aab1c3", fontSize: 12, lineHeight: 1.5 }}>
        Raw JSON view (advanced). If you edit here, use Validate then Save.
      </div>
      <textarea
        value={profileJson}
        onChange={(e) => setProfileJson(e.target.value)}
        style={{
          width: "100%",
          minHeight: 260,
          borderRadius: 12,
          border: "1px solid #2a2f3f",
          background: "#0d0f16",
          color: "#eef1ff",
          padding: 10,
          fontFamily: "ui-monospace, Menlo, Consolas, monospace",
          fontSize: 12,
          outline: "none",
          marginTop: 10,
        }}
      />
    </div>
  );
}

function ProfileHelp() {
  return (
    <div style={{ marginTop: 12, border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018", color: "#aab1c3", fontSize: 12, lineHeight: 1.55 }}>
      <div style={{ color: "#eef1ff", fontWeight: 900, marginBottom: 8 }}>Hints & tips</div>
      <ul style={{ margin: 0, paddingLeft: 18 }}>
        <li><b>Group labels must match your diagram</b> (AR1/BR2/CR4...).</li>
        <li><b>Ranges decide grouping</b>: [1,4,"AR1"] means A1–A4 belong to AR1.</li>
        <li><b>Avoid overlaps</b>: overlapping ranges make one line match multiple groups.</li>
        <li><b>mmPerLoop affects loop suggestions</b> in the target plan.</li>
      </ul>
    </div>
  );
}

function validateProfile(profileKey, profile) {
  const errors = [];
  if (!profile || typeof profile !== "object") errors.push("Profile is not an object.");
  if (!profile.name) errors.push("Profile is missing 'name'.");
  if (!Number.isFinite(profile.mmPerLoop)) errors.push("mmPerLoop should be a number.");
  if (!profile.mapping || typeof profile.mapping !== "object") errors.push("mapping is missing.");

  const letters = ["A", "B", "C", "D"];
  for (const L of letters) {
    const arr = profile?.mapping?.[L];
    if (!Array.isArray(arr)) {
      errors.push(`mapping.${L} should be an array.`);
      continue;
    }
    const rows = arr.map((r, idx) => {
      if (!Array.isArray(r) || r.length !== 3) errors.push(`mapping.${L}[${idx}] must be [min,max,"Group"].`);
      const min = n(r?.[0]);
      const max = n(r?.[1]);
      const g = r?.[2];
      if (!Number.isFinite(min) || !Number.isFinite(max)) errors.push(`mapping.${L}[${idx}] min/max must be numbers.`);
      if (Number.isFinite(min) && Number.isFinite(max) && min > max) errors.push(`mapping.${L}[${idx}] min cannot be > max.`);
      if (typeof g !== "string" || !g.trim()) errors.push(`mapping.${L}[${idx}] group label must be a non-empty string.`);
      return { min, max };
    });

    const sorted = rows
      .filter((r) => Number.isFinite(r.min) && Number.isFinite(r.max))
      .sort((a, b) => a.min - b.min);

    for (let i = 1; i < sorted.length; i++) {
      const prev = sorted[i - 1];
      const cur = sorted[i];
      if (cur.min <= prev.max) {
        errors.push(`Overlap in ${L}: range starting at ${cur.min} overlaps previous ending at ${prev.max}.`);
      }
    }
  }

  if (typeof profileKey !== "string" || !profileKey.trim()) errors.push("Profile key invalid.");
  return { ok: errors.length === 0, errors };
}

function deepClone(x) {
  try {
    return JSON.parse(JSON.stringify(x));
  } catch {
    return x;
  }
}

/* ------------------------- Table helpers ------------------------- */

function Th({ children, align }) {
  return <th style={{ padding: "10px 10px", textAlign: align || "left" }}>{children}</th>;
}
function Td({ children, colSpan, align, style }) {
  return (
    <td colSpan={colSpan} style={{ padding: "10px 10px", textAlign: align || "left", ...style, verticalAlign: "top" }}>
      {children}
    </td>
  );
}
