import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

const SITE_VERSION = "Trim Tuning • Step1–3 Sandbox • v1.4.3";

const ATTACHED_TEST_CSV = `Make ,Model,tolerance ,Korrektur,,,,,,,,,,,,
Ozone,Speedster3,10,-507,,,,,,,,,,,,
A,Soll,Ist L,Ist R,B,Soll,L,R,C,Soll,Ist L,Ist R,D,Soll,L,R
A1,6717,7220,7222,B1,6635,7142,7145,C1,6712,7219,7213,D1,6871,7379,7380
A2,6676,7184,7185,B2,6593,7100,7101,C2,6672,7177,7175,D2,6833,7334,7341
A3,6646,7151,7156,B3,6566,7071,7077,C3,6644,7149,7149,D3,6801,7309,7309
A4,6616,7119,7123,B4,6534,7038,7041,C4,6613,7116,7116,D4,6769,7276,7279
A5,6590,7090,7093,B5,6504,7007,7008,C5,6587,7088,7088,D5,6742,7247,7247
A6,6563,7060,7060,B6,6473,6978,6975,C6,6560,7059,7056,D6,6716,7217,7213
A7,6532,7026,7028,B7,6441,6944,6948,C7,6529,7027,7027,D7,6683,7184,7184
A8,6498,6991,6994,B8,6408,6909,6914,C8,6495,6991,6991,D8,6648,7149,7151
A9,6468,6958,6958,B9,6375,6879,6879,C9,6465,6956,6957,D9,6616,7118,7118
A10,6440,6929,6928,B10,6347,6850,6851,C10,6437,6927,6927,D10,6585,7088,7090
A11,6412,6900,6900,B11,6318,6820,6822,C11,6409,6898,6898,D11,6552,7053,7052
A12,6382,6868,6869,B12,6287,6788,6791,C12,6380,6867,6867,D12,6518,7016,7017
A13,6352,6839,6839,B13,6257,6760,6760,C13,6350,6838,6838,D13,6486,6987,6986
A14,6322,6808,6808,B14,6226,6729,6730,C14,6320,6808,6808,D14,6453,6952,6953
A15,6294,6779,6779,B15,6197,6702,6702,C15,6292,6777,6777,D15,6422,6922,6923
A16,6264,6749,6749,B16,6165,6670,6671,C16,6262,6748,6748,D16,6389,6887,6887`;

const theme = {
  bg: "#0f1115",
  panel: "rgba(255,255,255,0.05)",
  panel2: "rgba(0,0,0,0.35)",
  border: "rgba(255,255,255,0.14)",
  text: "rgba(255,255,255,0.92)",
  good: "rgba(34,197,94,0.95)",
  bad: "rgba(239,68,68,0.95)",
  warnBg: "rgba(245,158,11,0.10)",
  warnStroke: "rgba(245,158,11,0.55)",
};

const PALETTE = {
  A: { base: "#1e6eff", s2: "#2d7fff", s3: "#5aa6ff", s4: "#86c4ff" },
  B: { base: "#8b5cf6", s2: "#9d73ff", s3: "#b69bff", s4: "#cdbbff" },
  C: { base: "#ff9f43", s2: "#ffb76b", s3: "#ffd1a3", s4: "#ffe2c7" },
  D: { base: "#facc15", s2: "#fde047", s3: "#fef08a", s4: "#fff6bf" },
};

function clamp(n, a, b) {
  const x = Number(n);
  if (!Number.isFinite(x)) return a;
  return Math.max(a, Math.min(b, x));
}
function safeNum(v) {
  const n = Number(String(v ?? "").trim().replace(",", "."));
  return Number.isFinite(n) ? n : null;
}
function deepClone(obj) {
  try {
    return JSON.parse(JSON.stringify(obj));
  } catch {
    return obj;
  }
}
function chipColorFromLineId(lineId) {
  const first = String(lineId || "").trim().toUpperCase().charAt(0);
  return (PALETTE[first] || PALETTE.A).base;
}
function groupColor(letter, bucket) {
  const p = PALETTE[letter] || PALETTE.A;
  if (bucket === 1) return p.base;
  if (bucket === 2) return p.s2;
  if (bucket === 3) return p.s3;
  return p.s4;
}

// ---------------- Parsing ----------------
function rowsFromCSVText(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .filter((l) => l.trim().length > 0);

  const rows = [];
  for (const line of lines) {
    const out = [];
    let cur = "";
    let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        inQ = !inQ;
        continue;
      }
      if (ch === "," && !inQ) {
        out.push(cur);
        cur = "";
      } else {
        cur += ch;
      }
    }
    out.push(cur);
    rows.push(out);
  }
  return rows;
}
function rowsFromSheetAOA(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
}
function parseWideTableFromRows(rows) {
  const meta = { make: "", model: "", tolerance: 10, correction: 0 };
  const wideRows = [];

  if (!Array.isArray(rows) || rows.length < 4) return { meta, wideRows };

  const metaKeys = rows[0] || [];
  const metaVals = rows[1] || [];

  const findMeta = (needle) => {
    const idx = metaKeys.findIndex((k) => String(k || "").toLowerCase().includes(needle));
    return idx >= 0 ? metaVals[idx] : null;
  };

  meta.make = String(findMeta("make") ?? metaVals[0] ?? "").trim();
  meta.model = String(findMeta("model") ?? metaVals[1] ?? "").trim();

  const tolVal = findMeta("tolerance");
  const corrVal = findMeta("korrektur") ?? findMeta("correction");
  if (safeNum(tolVal) != null) meta.tolerance = safeNum(tolVal);
  if (safeNum(corrVal) != null) meta.correction = safeNum(corrVal);

  const header = (rows[2] || []).map((h) => String(h ?? "").trim());
  const letters = ["A", "B", "C", "D"];

  const letterCols = [];
  for (let c = 0; c < header.length; c++) {
    const h = header[c].toUpperCase();
    if (letters.includes(h)) letterCols.push({ letter: h, colStart: c });
  }
  if (letterCols.length === 0 && header.length >= 16) {
    letterCols.push({ letter: "A", colStart: 0 });
    letterCols.push({ letter: "B", colStart: 4 });
    letterCols.push({ letter: "C", colStart: 8 });
    letterCols.push({ letter: "D", colStart: 12 });
  }

  for (let r = 3; r < rows.length; r++) {
    const row = rows[r] || [];
    for (const blk of letterCols) {
      const c = blk.colStart;
      const base = String(row[c] ?? "").trim();
      if (!base) continue;

      const nominal = safeNum(row[c + 1]);
      const measuredL = safeNum(row[c + 2]);
      const measuredR = safeNum(row[c + 3]);

      const m = base.match(/^([A-Za-z])\s*([0-9]+)\s*$/);
      const letter = (m?.[1] || blk.letter || "").toUpperCase();
      const idx = m ? Number(m[2]) : safeNum(base.replace(/[^\d]/g, "")) || null;

      if (!["A", "B", "C", "D"].includes(letter) || idx == null) continue;
      wideRows.push({ letter, idx, lineBase: `${letter}${idx}`, nominal, measuredL, measuredR });
    }
  }

  return { meta, wideRows };
}

// ---------------- Step 2 grouping ----------------
function makeDefaultRanges(maxIdx, groupCount) {
  const m = Math.max(1, maxIdx);
  if (groupCount === 4) {
    return {
      1: { start: 1, end: Math.min(3, m) },
      2: { start: Math.min(4, m), end: Math.min(6, m) },
      3: { start: Math.min(7, m), end: Math.min(9, m) },
      4: { start: Math.min(10, m), end: m },
    };
  }
  return {
    1: { start: 1, end: Math.min(4, m) },
    2: { start: Math.min(5, m), end: Math.min(8, m) },
    3: { start: Math.min(9, m), end: m },
  };
}
function buildInitialLineToGroup({ maxByLetter, groupCountByLetter, prefixByLetter, rangesByLetter }) {
  const mapping = {};
  for (const letter of ["A", "B", "C", "D"]) {
    const maxIdx = Math.max(0, Number(maxByLetter?.[letter] || 0));
    const count = Number(groupCountByLetter?.[letter] || 3);
    const ranges = rangesByLetter?.[letter] || makeDefaultRanges(maxIdx, count);
    const prefix = prefixByLetter?.[letter] || `${letter}R`;

    for (let idx = 1; idx <= maxIdx; idx++) {
      let bucket = 1;
      for (let b = 1; b <= count; b++) {
        const s = clamp(ranges[b]?.start ?? 1, 1, maxIdx);
        const e = clamp(ranges[b]?.end ?? maxIdx, 1, maxIdx);
        if (idx >= s && idx <= e) {
          bucket = b;
          break;
        }
      }
      mapping[`${letter}${idx}L`] = `${prefix}${bucket}L`;
      mapping[`${letter}${idx}R`] = `${prefix}${bucket}R`;
    }
  }
  return mapping;
}
function getGroupOptions(prefixByLetter, groupCountByLetter) {
  const out = [];
  for (const letter of ["A", "B", "C", "D"]) {
    const prefix = prefixByLetter[letter] || `${letter}R`;
    const count = Number(groupCountByLetter[letter] || 3);
    for (let b = 1; b <= count; b++) {
      out.push(`${prefix}${b}L`);
      out.push(`${prefix}${b}R`);
    }
  }
  out.push("CUSTOM_L", "CUSTOM_R");
  return out;
}

// ---------------- Diagram ----------------
const DIAGRAM_SCALE = 0.85;
const DIAGRAM_BASE_W = 2400;
const DIAGRAM_BASE_H = 980;
const DIAGRAM_W = Math.round(DIAGRAM_BASE_W * DIAGRAM_SCALE);
const DIAGRAM_H = Math.round(DIAGRAM_BASE_H * DIAGRAM_SCALE);

function DiagramPreview({ lineToGroup, prefixByLetter, groupCountByLetter, showWingOutline, compactLayout, setLineToGroupFromDrag }) {
  const W = DIAGRAM_W;
  const H = DIAGRAM_H;
  const cx = W / 2;
  const [drag, setDrag] = useState({ active: false });

  const rowY = {
    A: Math.round(170 * DIAGRAM_SCALE),
    B: Math.round(375 * DIAGRAM_SCALE),
    C: Math.round(580 * DIAGRAM_SCALE),
    D: Math.round(785 * DIAGRAM_SCALE),
  };

  const centerGap = compactLayout ? Math.round(26 * DIAGRAM_SCALE) : Math.round(34 * DIAGRAM_SCALE);
  const bucketSpacing = compactLayout ? Math.round(8 * DIAGRAM_SCALE) : Math.round(12 * DIAGRAM_SCALE);

  const chip = compactLayout
    ? { w: Math.round(86 * DIAGRAM_SCALE), h: Math.round(32 * DIAGRAM_SCALE), gapX: Math.round(10 * DIAGRAM_SCALE), gapY: Math.round(8 * DIAGRAM_SCALE) }
    : { w: Math.round(90 * DIAGRAM_SCALE), h: Math.round(34 * DIAGRAM_SCALE), gapX: Math.round(10 * DIAGRAM_SCALE), gapY: Math.round(8 * DIAGRAM_SCALE) };

  const maxCols = compactLayout ? 7 : 8;
  function chooseCols(n) {
    if (n <= 1) return 1;
    const ideal = Math.ceil(Math.sqrt(n) * 1.15);
    return clamp(ideal, 2, maxCols);
  }
  function boxSizeForLines(lines) {
    const n = lines.length;
    const cols = n <= 1 ? 1 : chooseCols(n);
    const rows = Math.max(1, Math.ceil(n / cols));
    const innerW = cols * chip.w + (cols - 1) * chip.gapX;
    const innerH = rows * chip.h + (rows - 1) * chip.gapY;
    const padW = compactLayout ? Math.round(28 * DIAGRAM_SCALE) : Math.round(32 * DIAGRAM_SCALE);
    const padH = compactLayout ? Math.round(46 * DIAGRAM_SCALE) : Math.round(50 * DIAGRAM_SCALE);
    return { bw: innerW + padW, bh: innerH + padH, cols };
  }

  const layout = useMemo(() => {
    const zones = [];
    const blocks = [];
    for (const letter of ["A", "B", "C", "D"]) {
      const prefix = prefixByLetter[letter] || `${letter}R`;
      const count = Number(groupCountByLetter[letter] || 3);

      for (const side of ["L", "R"]) {
        const groups = [];
        for (let b = 1; b <= count; b++) {
          const groupId = `${prefix}${b}${side}`;
          const lines = Object.keys(lineToGroup || {}).filter((lineId) => lineToGroup[lineId] === groupId);

          const sorted = lines
            .slice()
            .sort((a, b2) => {
              const ai = Number(String(a).match(/(\d+)/)?.[1] || 0);
              const bi = Number(String(b2).match(/(\d+)/)?.[1] || 0);
              return side === "L" ? bi - ai : ai - bi;
            });

          groups.push({ bucket: b, groupId, lines: sorted });
        }

        const y = rowY[letter];
        const measured = groups.map((g) => ({ ...g, ...boxSizeForLines(g.lines) }));

        if (side === "R") {
          let x = cx + centerGap;
          for (const g of measured) {
            const boxX = x;
            const boxY = y - g.bh / 2;
            zones.push({ groupId: g.groupId, x: boxX, y: boxY, w: g.bw, h: g.bh });
            blocks.push({ letter, side, groupId: g.groupId, boxX, boxY, bw: g.bw, bh: g.bh, lines: g.lines, cols: g.cols, bucketNum: g.bucket });
            x = x + g.bw + bucketSpacing;
          }
        } else {
          let x = cx - centerGap;
          for (const g of measured) {
            const boxX = x - g.bw;
            const boxY = y - g.bh / 2;
            zones.push({ groupId: g.groupId, x: boxX, y: boxY, w: g.bw, h: g.bh });
            blocks.push({ letter, side, groupId: g.groupId, boxX, boxY, bw: g.bw, bh: g.bh, lines: g.lines, cols: g.cols, bucketNum: g.bucket });
            x = x - g.bw - bucketSpacing;
          }
        }
      }
    }
    return { zones, blocks };
  }, [lineToGroup, prefixByLetter, groupCountByLetter, compactLayout]);

  function screenToSvgPoint(e) {
    const svg = e.currentTarget.ownerSVGElement || e.currentTarget;
    const pt = svg.createSVGPoint();
    pt.x = e.clientX;
    pt.y = e.clientY;
    const m = svg.getScreenCTM();
    if (!m) return { x: 0, y: 0 };
    const p = pt.matrixTransform(m.inverse());
    return { x: p.x, y: p.y };
  }
  function findZoneAt(x, y) {
    return layout.zones.find((z) => x >= z.x && x <= z.x + z.w && y >= z.y && y <= z.y + z.h) || null;
  }

  function onPointerDownChip(e, payload) {
    e.preventDefault();
    e.stopPropagation();
    try {
      e.currentTarget.setPointerCapture(e.pointerId);
    } catch {}
    const { x, y } = screenToSvgPoint(e);
    setDrag({ active: true, ...payload, x, y });
  }
  function onPointerMove(e) {
    if (!drag.active) return;
    const { x, y } = screenToSvgPoint(e);
    setDrag((d) => ({ ...d, x, y }));
  }
  function onPointerUp(e) {
    if (!drag.active) return;
    const { x, y } = screenToSvgPoint(e);
    const zone = findZoneAt(x, y);
    if (zone?.groupId) setLineToGroupFromDrag(drag.lineId, zone.groupId);
    setDrag({ active: false });
  }

  const items = [];
  items.push(<rect key="bg" x={0} y={0} width={W} height={H} fill="rgba(0,0,0,0)" />);

  if (showWingOutline) {
    items.push(
      <path
        key="wing"
        d={`M ${W / 2} ${Math.round(62 * DIAGRAM_SCALE)}
           C ${W / 2 - Math.round(740 * DIAGRAM_SCALE)} ${Math.round(88 * DIAGRAM_SCALE)}, ${W / 2 - Math.round(1340 * DIAGRAM_SCALE)} ${Math.round(260 * DIAGRAM_SCALE)}, ${W / 2 - Math.round(1460 * DIAGRAM_SCALE)} ${Math.round(470 * DIAGRAM_SCALE)}
           C ${W / 2 - Math.round(1340 * DIAGRAM_SCALE)} ${Math.round(720 * DIAGRAM_SCALE)}, ${W / 2 - Math.round(740 * DIAGRAM_SCALE)} ${Math.round(860 * DIAGRAM_SCALE)}, ${W / 2} ${Math.round(900 * DIAGRAM_SCALE)}
           C ${W / 2 + Math.round(740 * DIAGRAM_SCALE)} ${Math.round(860 * DIAGRAM_SCALE)}, ${W / 2 + Math.round(1340 * DIAGRAM_SCALE)} ${Math.round(720 * DIAGRAM_SCALE)}, ${W / 2 + Math.round(1460 * DIAGRAM_SCALE)} ${Math.round(470 * DIAGRAM_SCALE)}
           C ${W / 2 + Math.round(1340 * DIAGRAM_SCALE)} ${Math.round(260 * DIAGRAM_SCALE)}, ${W / 2 + Math.round(740 * DIAGRAM_SCALE)} ${Math.round(88 * DIAGRAM_SCALE)}, ${W / 2} ${Math.round(62 * DIAGRAM_SCALE)}`}
        fill="none"
        stroke="rgba(255,255,255,0.18)"
        strokeWidth={Math.max(2, Math.round(4 * DIAGRAM_SCALE))}
      />
    );
  }

  items.push(<line key="center" x1={cx} y1={20} x2={cx} y2={H - 20} stroke="rgba(255,255,255,0.14)" strokeWidth={2} />);

  for (const b of layout.blocks) {
    const bucketNum = Number(String(b.groupId).match(/(\d+)/)?.[1] || b.bucketNum || 1);
    const bucketStroke = groupColor(b.letter, bucketNum);

    items.push(
      <rect
        key={`box-${b.groupId}-${b.side}-${b.boxX}`}
        x={b.boxX}
        y={b.boxY}
        width={b.bw}
        height={b.bh}
        rx={18}
        ry={18}
        fill={bucketStroke}
        opacity={0.09}
        stroke={bucketStroke}
        strokeOpacity={0.92}
        strokeWidth={2}
      />
    );

    const title = b.groupId;
    const pillW = Math.max(120, title.length * 10);
    const pillH = 30;
    const pillX = b.boxX + (b.bw - pillW) / 2;

    items.push(
      <g key={`title-${b.groupId}-${b.side}-${b.boxX}`}>
        <rect x={pillX} y={b.boxY + 8} width={pillW} height={pillH} rx={16} ry={16} fill="rgba(0,0,0,0.58)" stroke={bucketStroke} strokeOpacity={0.85} />
        <text x={pillX + pillW / 2} y={b.boxY + 30} textAnchor="middle" fill={bucketStroke} fontWeight={950} fontSize={18}>
          {title}
        </text>
      </g>
    );

    const innerX = b.boxX + 16;
    const innerY = b.boxY + 44;
    const cols = b.cols || 1;

    for (let i = 0; i < b.lines.length; i++) {
      const lineId = b.lines[i];
      const rr = Math.floor(i / cols);
      const ccRaw = i % cols;
      const cc = b.side === "L" ? cols - 1 - ccRaw : ccRaw;
      const x = innerX + cc * (chip.w + chip.gapX);
      const yy = innerY + rr * (chip.h + chip.gapY);

      const chipStroke = chipColorFromLineId(lineId);

      items.push(
        <g key={`chip-${b.groupId}-${lineId}`}>
          <rect x={x} y={yy} width={chip.w} height={chip.h} rx={12} ry={12} fill="rgba(255,255,255,0.10)" stroke={chipStroke} strokeOpacity={0.60} />
          <text x={x + chip.w / 2} y={yy + chip.h * 0.74} textAnchor="middle" fill="rgba(255,255,255,0.93)" fontSize={18} fontWeight={950}>
            {lineId}
          </text>
          <rect x={x} y={yy} width={chip.w} height={chip.h} rx={12} ry={12} fill="transparent" style={{ cursor: "grab" }} onPointerDown={(e) => onPointerDownChip(e, { lineId, color: chipStroke })} />
        </g>
      );
    }
  }

  return (
    <div style={{ width: "100%", border: `1px solid ${theme.border}`, borderRadius: 18, background: "rgba(0,0,0,0.38)", overflow: "hidden" }}>
      <svg
        width="100%"
        viewBox={`0 0 ${W} ${H}`}
        style={{ display: "block", touchAction: "none" }}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerUp}
        onPointerLeave={onPointerUp}
      >
        {items}
      </svg>
    </div>
  );
}

// ---------------- UI atoms ----------------
function Panel({ title, right, children, tint = false }) {
  return (
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 18, background: tint ? "linear-gradient(180deg, rgba(59,130,246,0.08), rgba(255,255,255,0.04))" : theme.panel, overflow: "hidden" }}>
      <div style={{ padding: "10px 12px", borderBottom: `1px solid ${theme.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", background: "rgba(255,255,255,0.035)" }}>
        <div style={{ fontWeight: 950, letterSpacing: -0.2, fontSize: 16 }}>{title}</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>{right}</div>
      </div>
      <div style={{ padding: 12 }}>{children}</div>
    </div>
  );
}
function NumInput({ value, onChange, min = -99999, max = 99999, step = 1, width = 120 }) {
  return (
    <input
      type="number"
      value={value}
      min={min}
      max={max}
      step={step}
      onChange={(e) => onChange?.(Number(e.target.value))}
      style={{
        width,
        padding: "7px 10px",
        borderRadius: 12,
        border: `1px solid ${theme.border}`,
        background: "rgba(0,0,0,0.68)",
        color: theme.text,
        outline: "none",
        fontWeight: 900,
        fontSize: 14,
      }}
    />
  );
}
function Select({ value, onChange, options, width = 140 }) {
  return (
    <select
      value={value}
      onChange={(e) => onChange?.(e.target.value)}
      style={{
        width,
        padding: "7px 10px",
        borderRadius: 12,
        border: `1px solid ${theme.border}`,
        background: "rgba(0,0,0,0.80)",
        color: theme.text,
        outline: "none",
        fontWeight: 950,
        fontSize: 14,
      }}
    >
      {options.map((o) => (
        <option key={o.value} value={o.value} style={{ color: "#111" }}>
          {o.label}
        </option>
      ))}
    </select>
  );
}
function Toggle({ value, onChange, label }) {
  return (
    <label style={{ display: "inline-flex", alignItems: "center", gap: 10, cursor: "pointer", userSelect: "none" }}>
      <span
        style={{
          width: 40,
          height: 24,
          borderRadius: 999,
          border: `1px solid ${theme.border}`,
          background: value ? "rgba(59,130,246,0.35)" : "rgba(255,255,255,0.08)",
          position: "relative",
        }}
        onClick={() => onChange?.(!value)}
      >
        <span
          style={{
            width: 18,
            height: 18,
            borderRadius: 999,
            background: value ? "rgba(255,255,255,0.92)" : "rgba(255,255,255,0.58)",
            position: "absolute",
            top: 2.5,
            left: value ? 19 : 3,
          }}
        />
      </span>
      <span style={{ opacity: 0.92, fontWeight: 900, fontSize: 14 }}>{label}</span>
    </label>
  );
}
function ImportStatusRadio({ loaded, label = "Import status" }) {
  const color = loaded ? theme.good : theme.bad;
  const text = loaded ? "File loaded" : "No file loaded";
  return (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 10, padding: "8px 12px", borderRadius: 999, border: `1px solid ${theme.border}`, background: "rgba(0,0,0,0.45)" }}>
      <span style={{ width: 14, height: 14, borderRadius: 999, background: color }} />
      <div style={{ display: "grid", lineHeight: 1.05 }}>
        <span style={{ fontSize: 12, opacity: 0.78, fontWeight: 900 }}>{label}</span>
        <span style={{ fontSize: 13, fontWeight: 950, color }}>{text}</span>
      </div>
    </div>
  );
}
function WarningBanner({ title, children }) {
  return (
    <div style={{ border: `1px solid ${theme.warnStroke}`, background: theme.warnBg, borderRadius: 16, padding: 10 }}>
      <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
        <div style={{ width: 10, height: 10, borderRadius: 999, background: theme.warnStroke, marginTop: 5 }} />
        <div style={{ display: "grid", gap: 6 }}>
          <div style={{ fontWeight: 950 }}>{title}</div>
          <div style={{ opacity: 0.82, fontWeight: 800, fontSize: 13 }}>{children}</div>
        </div>
      </div>
    </div>
  );
}
function SegTabs({ value, onChange, tabs }) {
  return (
    <div style={{ display: "inline-flex", border: `1px solid ${theme.border}`, borderRadius: 14, overflow: "hidden", background: "rgba(0,0,0,0.35)", flexWrap: "wrap" }}>
      {tabs.map((t) => {
        const active = value === t.value;
        return (
          <button
            key={t.value}
            onClick={() => onChange(t.value)}
            style={{
              padding: "9px 12px",
              border: "none",
              background: active ? "rgba(59,130,246,0.25)" : "transparent",
              color: theme.text,
              cursor: "pointer",
              fontWeight: 950,
              fontSize: 14,
              whiteSpace: "nowrap",
            }}
          >
            {t.label}
          </button>
        );
      })}
    </div>
  );
}

// JSON helpers
function downloadJSON(obj, filename) {
  const json = JSON.stringify(obj, null, 2);
  const blob = new Blob([json], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || "wing-profile.json";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2500);
}
async function readFileText(file) {
  return await file.text();
}

export default function App() {
  const [step, setStep] = useState(1);

  // Step 1
  const [meta, setMeta] = useState({ make: "", model: "", tolerance: 10, correction: 0 });
  const [wideRows, setWideRows] = useState([]);
  const [importStatus, setImportStatus] = useState({ ok: false, name: "", err: "" });
  const [tab, setTab] = useState("import");
  const fileInputRef = useRef(null);

  // Step 2
  const [prefixByLetter, setPrefixByLetter] = useState({ A: "AR", B: "BR", C: "CR", D: "DR" });
  const [groupCountByLetter, setGroupCountByLetter] = useState({ A: 3, B: 3, C: 3, D: 3 });
  const [maxByLetter, setMaxByLetter] = useState({ A: 0, B: 0, C: 0, D: 0 });
  const [rangesByLetter, setRangesByLetter] = useState({
    A: makeDefaultRanges(1, 3),
    B: makeDefaultRanges(1, 3),
    C: makeDefaultRanges(1, 3),
    D: makeDefaultRanges(1, 3),
  });
  const [lineToGroup, setLineToGroup] = useState({});
  const [rangeTab, setRangeTab] = useState("A");
  const [showOverrides, setShowOverrides] = useState(true);

  // Step 3 diagram + summary
  const [diagramZoom, setDiagramZoom] = useState(1.0);
  const [diagramWingOutline, setDiagramWingOutline] = useState(true);
  const [diagramCompact, setDiagramCompact] = useState(false);
  const diagramBoxRef = useRef(null);

  // Baseline mapping snapshot for “changes summary”
  const [defaultMappingSnapshot, setDefaultMappingSnapshot] = useState(null);

  // Profile JSON
  const [profileName, setProfileName] = useState("");
  const profileImportRef = useRef(null);

  function resetAll() {
    setStep(1);
    setMeta({ make: "", model: "", tolerance: 10, correction: 0 });
    setWideRows([]);
    setImportStatus({ ok: false, name: "", err: "" });

    setPrefixByLetter({ A: "AR", B: "BR", C: "CR", D: "DR" });
    setGroupCountByLetter({ A: 3, B: 3, C: 3, D: 3 });
    setMaxByLetter({ A: 0, B: 0, C: 0, D: 0 });
    setRangesByLetter({
      A: makeDefaultRanges(1, 3),
      B: makeDefaultRanges(1, 3),
      C: makeDefaultRanges(1, 3),
      D: makeDefaultRanges(1, 3),
    });
    setLineToGroup({});
    setRangeTab("A");
    setShowOverrides(true);

    setDiagramZoom(1.0);
    setDiagramWingOutline(true);
    setDiagramCompact(false);

    setDefaultMappingSnapshot(null);
    setProfileName("");
  }

  function applyParsedImport(parsed, name) {
    setMeta(parsed.meta);
    setWideRows(parsed.wideRows);
    setImportStatus({ ok: true, name, err: "" });

    const max = { A: 0, B: 0, C: 0, D: 0 };
    for (const r of parsed.wideRows) {
      if (!r?.letter || r.idx == null) continue;
      max[r.letter] = Math.max(max[r.letter] || 0, Number(r.idx || 0));
    }
    setMaxByLetter(max);

    const newRanges = {
      A: makeDefaultRanges(max.A || 1, groupCountByLetter.A || 3),
      B: makeDefaultRanges(max.B || 1, groupCountByLetter.B || 3),
      C: makeDefaultRanges(max.C || 1, groupCountByLetter.C || 3),
      D: makeDefaultRanges(max.D || 1, groupCountByLetter.D || 3),
    };
    setRangesByLetter(newRanges);

    const initMap = buildInitialLineToGroup({
      maxByLetter: max,
      groupCountByLetter,
      prefixByLetter,
      rangesByLetter: newRanges,
    });

    setLineToGroup(initMap);
    setDefaultMappingSnapshot(deepClone(initMap));
    setRangeTab("A");
    setStep(2);
  }

  async function handleFile(file) {
    if (!file) return;
    resetAll();

    const name = file.name || "file";
    const lower = name.toLowerCase();

    try {
      if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const rows = rowsFromSheetAOA(sheet);
        const parsed = parseWideTableFromRows(rows);
        applyParsedImport(parsed, name);
      } else {
        const text = await file.text();
        const rows = rowsFromCSVText(text);
        const parsed = parseWideTableFromRows(rows);
        applyParsedImport(parsed, name);
      }
      setTab("import");
    } catch (e) {
      console.error("Import error:", e);
      resetAll();
      setImportStatus({ ok: false, name, err: "Import failed. Check file format." });
      alert("Import failed. Please check the file format.");
    }
  }

  function loadAttachedTestData() {
    resetAll();
    const rows = rowsFromCSVText(ATTACHED_TEST_CSV);
    const parsed = parseWideTableFromRows(rows);
    applyParsedImport(parsed, "Speedster3 ML.csv (attached)");
    setTab("testdata");
  }

  const loaded = importStatus.ok && wideRows.length > 0;

  const summary = useMemo(() => {
    const max = { A: 0, B: 0, C: 0, D: 0 };
    for (const r of wideRows) {
      if (!r?.letter || r.idx == null) continue;
      max[r.letter] = Math.max(max[r.letter] || 0, Number(r.idx || 0));
    }
    const totalLines = (max.A + max.B + max.C + max.D) * 2;
    return { max, totalLines };
  }, [wideRows]);

  const groupOptionsForSelect = useMemo(() => {
    const opts = getGroupOptions(prefixByLetter, groupCountByLetter);
    return opts.map((g) => ({ value: g, label: g }));
  }, [prefixByLetter, groupCountByLetter]);

  const letterIdxRows = useMemo(() => {
    const out = [];
    for (const L of ["A", "B", "C", "D"]) {
      const m = Number(maxByLetter[L] || 0);
      for (let i = 1; i <= m; i++) out.push({ letter: L, idx: i });
    }
    return out;
  }, [maxByLetter]);

  function setRange(letter, bucket, field, value) {
    setRangesByLetter((prev) => {
      const next = { ...(prev || {}) };
      const count = Number(groupCountByLetter[letter] || 3);
      const cur = next[letter] || makeDefaultRanges(maxByLetter[letter] || 1, count);
      const b = { ...(cur[bucket] || {}) };
      b[field] = Number.isFinite(Number(value)) ? Number(value) : b[field];
      next[letter] = { ...cur, [bucket]: b };
      return next;
    });
  }

  function rebuildMappingFromRanges(resetOverridesToRanges = false) {
    const initMap = buildInitialLineToGroup({
      maxByLetter,
      groupCountByLetter,
      prefixByLetter,
      rangesByLetter,
    });

    if (resetOverridesToRanges) {
      setLineToGroup(initMap);
      setDefaultMappingSnapshot(deepClone(initMap));
      return;
    }

    setLineToGroup((prev) => {
      const next = { ...(prev || {}) };
      for (const L of ["A", "B", "C", "D"]) {
        const m = Number(maxByLetter[L] || 0);
        for (let i = 1; i <= m; i++) {
          const l = `${L}${i}L`;
          const r = `${L}${i}R`;
          if (!next[l]) next[l] = initMap[l] || "";
          if (!next[r]) next[r] = initMap[r] || "";
        }
      }
      return next;
    });
  }

  function setLineToGroupFromDrag(lineId, newGroupId) {
    if (!lineId || !newGroupId) return;
    setLineToGroup((prev) => ({ ...(prev || {}), [lineId]: newGroupId }));
  }

  function fitDiagramToScreen() {
    const el = diagramBoxRef.current;
    if (!el) return;
    const available = Math.max(320, el.clientWidth - 24);
    const z = clamp(available / DIAGRAM_W, 0.4, 1.8);
    setDiagramZoom(Number(z.toFixed(2)));
  }
  useEffect(() => {
    if (step !== 3) return;
    const t = setTimeout(() => fitDiagramToScreen(), 80);
    return () => clearTimeout(t);
  }, [step]);

  const changes = useMemo(() => {
    const base = defaultMappingSnapshot || {};
    const cur = lineToGroup || {};
    const keys = Object.keys(cur).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    const out = [];
    for (const k of keys) {
      const from = base[k] || "";
      const to = cur[k] || "";
      if (from && to && from !== to) out.push({ lineId: k, from, to });
    }
    return out;
  }, [defaultMappingSnapshot, lineToGroup]);

  function buildWingProfileJSON() {
    return {
      schema: "trim-tuning-wing-profile-v1",
      exportedAt: new Date().toISOString(),
      wing: { make: meta.make || "", model: meta.model || "" },
      step2: {
        prefixByLetter: deepClone(prefixByLetter),
        groupCountByLetter: deepClone(groupCountByLetter),
        rangesByLetter: deepClone(rangesByLetter),
        lineToGroup: deepClone(lineToGroup),
        defaultMappingSnapshot: deepClone(defaultMappingSnapshot),
      },
      diagram: { wingOutline: diagramWingOutline, compact: diagramCompact },
    };
  }
  function exportWingProfileJSON() {
    if (!loaded) return alert("Load a wing data file first.");
    const name = String(profileName || "").trim() || `${meta.make || "Wing"}-${meta.model || "Profile"}`.replace(/\s+/g, "-");
    downloadJSON(buildWingProfileJSON(), `${name}.json`);
  }
  async function importWingProfileJSON(file) {
    if (!file) return;
    if (!loaded) return alert("Import wing data first, then load a profile JSON.");
    try {
      const text = await readFileText(file);
      const parsed = JSON.parse(text);
      if (!parsed || parsed.schema !== "trim-tuning-wing-profile-v1") return alert("Not a valid wing profile JSON.");

      const s2 = parsed.step2 || {};
      if (s2.prefixByLetter) setPrefixByLetter((p) => ({ ...p, ...s2.prefixByLetter }));
      if (s2.groupCountByLetter) setGroupCountByLetter((p) => ({ ...p, ...s2.groupCountByLetter }));
      if (s2.rangesByLetter) setRangesByLetter((p) => ({ ...p, ...s2.rangesByLetter }));
      if (s2.lineToGroup) setLineToGroup((p) => ({ ...(p || {}), ...s2.lineToGroup }));
      if (s2.defaultMappingSnapshot) setDefaultMappingSnapshot(s2.defaultMappingSnapshot);

      if (parsed.diagram && typeof parsed.diagram === "object") {
        if (typeof parsed.diagram.wingOutline === "boolean") setDiagramWingOutline(parsed.diagram.wingOutline);
        if (typeof parsed.diagram.compact === "boolean") setDiagramCompact(parsed.diagram.compact);
      }

      setStep(2);
      alert("Profile loaded.");
    } catch (e) {
      console.error(e);
      alert("Failed to load profile JSON.");
    }
  }

  function PrefixTile({ letter }) {
    // Reduced overall bucket width + aligned left
    return (
      <div
        style={{
          padding: 8,
          borderRadius: 12,
          border: `1px solid ${theme.border}`,
          background: "rgba(255,255,255,0.03)",
          width: 160,
        }}
      >
        <div style={{ fontWeight: 950, fontSize: 12 }}>
          {letter} prefix <span style={{ opacity: 0.7, fontWeight: 850 }}>({(prefixByLetter[letter] || "") + "1L"} / {(prefixByLetter[letter] || "") + "1R"})</span>
        </div>
        <div style={{ marginTop: 6 }}>
          <input
            value={prefixByLetter[letter] || ""}
            onChange={(e) => setPrefixByLetter((p) => ({ ...p, [letter]: e.target.value.toUpperCase() }))}
            style={{
              width: 68,
              padding: "7px 9px",
              borderRadius: 12,
              border: `1px solid ${theme.border}`,
              background: "rgba(0,0,0,0.68)",
              color: theme.text,
              outline: "none",
              fontWeight: 950,
              textTransform: "uppercase",
              fontSize: 13,
            }}
          />
        </div>
      </div>
    );
  }

  function RangeEditor({ L }) {
    const count = Number(groupCountByLetter[L] || 3);
    const r = rangesByLetter[L] || makeDefaultRanges(maxByLetter[L] || 1, count);

    return (
      <div style={{ border: `1px solid ${theme.border}`, borderRadius: 14, padding: 10, background: "rgba(0,0,0,0.38)" }}>
        <div style={{ display: "flex", justifyContent: "flex-start", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <div style={{ fontWeight: 950, fontSize: 14 }}>
            {L} ranges <span style={{ opacity: 0.7, fontWeight: 850 }}>(max {maxByLetter[L] || 0})</span>
          </div>
          <Select
            value={String(count)}
            onChange={(v) => {
              const nextCount = Number(v) === 4 ? 4 : 3;
              setGroupCountByLetter((p) => ({ ...p, [L]: nextCount }));
              setRangesByLetter((prev) => ({ ...prev, [L]: makeDefaultRanges(maxByLetter[L] || 1, nextCount) }));
            }}
            options={[
              { value: "3", label: "3 groups" },
              { value: "4", label: "4 groups" },
            ]}
            width={120}
          />
        </div>

        <div style={{ marginTop: 8, display: "grid", gap: 8 }}>
          {Array.from({ length: count }, (_, i) => i + 1).map((bucket) => {
            const col = groupColor(L, bucket);
            const prefix = (prefixByLetter[L] || `${L}R`) + bucket;
            return (
              <div
                key={`${L}-${bucket}`}
                style={{
                  display: "grid",
                  gridTemplateColumns: "86px 110px 110px",
                  justifyContent: "start",
                  gap: 8,
                  alignItems: "center",
                  padding: 8,
                  borderRadius: 12,
                  border: `1px solid ${theme.border}`,
                  background: "rgba(255,255,255,0.03)",
                }}
              >
                <div style={{ display: "flex", alignItems: "center", gap: 8, justifySelf: "start" }}>
                  <span style={{ width: 9, height: 9, borderRadius: 999, background: col }} />
                  <span style={{ fontWeight: 950, color: col, fontSize: 13 }}>{prefix}</span>
                </div>

                <div style={{ display: "flex", gap: 8, alignItems: "center", justifySelf: "start" }}>
                  <span style={{ opacity: 0.72, fontWeight: 850, fontSize: 12 }}>S</span>
                  <NumInput value={r[bucket]?.start ?? 1} min={1} max={maxByLetter[L]} onChange={(vv) => setRange(L, bucket, "start", clamp(vv, 1, maxByLetter[L] || 1))} width={62} />
                </div>

                <div style={{ display: "flex", gap: 8, alignItems: "center", justifySelf: "start" }}>
                  <span style={{ opacity: 0.72, fontWeight: 850, fontSize: 12 }}>E</span>
                  <NumInput value={r[bucket]?.end ?? maxByLetter[L]} min={1} max={maxByLetter[L]} onChange={(vv) => setRange(L, bucket, "end", clamp(vv, 1, maxByLetter[L] || 1))} width={62} />
                </div>
              </div>
            );
          })}
        </div>

        <div style={{ marginTop: 8, opacity: 0.76, fontSize: 12, fontWeight: 850 }}>
          After changing ranges, click <b>Apply ranges</b>.
        </div>
      </div>
    );
  }

  function ColoredRangeTabs() {
    const tabs = ["A", "B", "C", "D"].map((L) => ({ L, color: PALETTE[L].base }));
    return (
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", justifyContent: "flex-start" }}>
        {tabs.map((t) => {
          const active = rangeTab === t.L;
          return (
            <button
              key={t.L}
              onClick={() => setRangeTab(t.L)}
              style={{
                padding: "8px 12px",
                borderRadius: 999,
                border: `1px solid ${theme.border}`,
                background: active ? "rgba(255,255,255,0.10)" : "rgba(0,0,0,0.30)",
                color: theme.text,
                cursor: "pointer",
                fontWeight: 950,
                display: "inline-flex",
                alignItems: "center",
                gap: 8,
                whiteSpace: "nowrap",
              }}
            >
              <span style={{ width: 10, height: 10, borderRadius: 999, background: t.color }} />
              {t.L} ranges
            </button>
          );
        })}
      </div>
    );
  }

  function DiagramScrollBox({ height, width }) {
    // Fixed height + fixed width + obvious scrollbars
    return (
      <div
        ref={diagramBoxRef}
        className="diagramScrollBox"
        style={{
          border: `2px solid rgba(255,255,255,0.18)`,
          borderRadius: 18,
          overflow: "scroll",
          height: height || 640,
          width: width || 980,
          maxWidth: "100%",
          background: "rgba(0,0,0,0.34)",
          boxShadow: "inset 0 0 0 1px rgba(255,255,255,0.10)",
        }}
      >
        <div
        style={{
          width: Math.max(DIAGRAM_W * diagramZoom, (width || 980) + 260),
          height: Math.max(DIAGRAM_H * diagramZoom, (height || 640) + 140),
        }}
      >
        <div
          style={{
            width: DIAGRAM_W,
            height: DIAGRAM_H,
            transform: `scale(${diagramZoom})`,
            transformOrigin: "top left",
          }}
        >
          <DiagramPreview
            lineToGroup={lineToGroup}
            prefixByLetter={prefixByLetter}
            groupCountByLetter={groupCountByLetter}
            showWingOutline={diagramWingOutline}
            compactLayout={diagramCompact}
            setLineToGroupFromDrag={setLineToGroupFromDrag}
          />
        </div>
      </div>
      </div>
    );
  }

  const stepTabs = [
    { value: 1, label: "Step 1" },
    { value: 2, label: "Step 2" },
    { value: 3, label: "Step 3 (Diagram)" },
  ];

  return (
    <div style={{ minHeight: "100vh", background: theme.bg, color: theme.text, padding: 16, fontFamily: 'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, "Helvetica Neue", Arial' }}>
      {/* scrollbars styling */}
      <style>{`
        .diagramScrollBox { scrollbar-width: auto; scrollbar-color: rgba(255,255,255,0.55) rgba(0,0,0,0.35); }
        .diagramScrollBox::-webkit-scrollbar { height: 16px; width: 16px; }
        .diagramScrollBox::-webkit-scrollbar-track { background: rgba(0,0,0,0.35); }
        .diagramScrollBox::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.35); border: 3px solid rgba(0,0,0,0.35); border-radius: 999px; }
        .diagramScrollBox::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.55); }
      `}</style>

      {/* reduced overall width to match overrides panel */}
      <div style={{ maxWidth: 980, margin: "0 auto", display: "grid", gap: 10 }}>
        {/* Header */}
        <div style={{ border: `1px solid ${theme.border}`, borderRadius: 22, padding: 14, background: "linear-gradient(180deg, rgba(59,130,246,0.16), rgba(255,255,255,0.03))" }}>
          <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
            <div>
              <div style={{ fontSize: 36, fontWeight: 950, letterSpacing: -0.9 }}>Paraglider Trim Tuning</div>
              <div style={{ marginTop: 6, opacity: 0.86, fontSize: 14, fontWeight: 900 }}>
                {SITE_VERSION} <span style={{ opacity: 0.7, fontWeight: 850 }}>• Step 1–3</span>
              </div>
            </div>

            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
              <ImportStatusRadio loaded={loaded} />
              {stepTabs.map((t) => {
                const disabled = t.value !== 1 && !loaded;
                return (
                  <button
                    key={t.value}
                    style={{
                      ...topBtn,
                      ...(step === t.value ? topBtnActive : {}),
                      opacity: disabled ? 0.55 : 1,
                      cursor: disabled ? "not-allowed" : "pointer",
                    }}
                    onClick={() => !disabled && setStep(t.value)}
                  >
                    {t.label}
                  </button>
                );
              })}
              <button style={{ ...topBtn, background: "rgba(239,68,68,0.16)" }} onClick={resetAll}>
                Reset all
              </button>
            </div>
          </div>

          {!loaded ? (
            <div style={{ marginTop: 10 }}>
              <WarningBanner title="No file loaded yet">
                Import a CSV/XLSX in the Speedster-style wide format. Or use <b>Test data</b> to load the embedded file.
              </WarningBanner>
            </div>
          ) : null}
        </div>

        {/* Step 1 */}
        {step === 1 ? (
          <Panel
            tint
            title={tab === "import" ? "Step 1 — Import CSV/XLSX" : "Step 1 — Test data (attached)"}
            right={<SegTabs value={tab} onChange={setTab} tabs={[{ value: "import", label: "Import" }, { value: "testdata", label: "Test data" }]} />}
          >
            <div style={{ display: "grid", gap: 10 }}>
              {tab === "import" ? (
                <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
                  <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={(e) => handleFile(e.target.files?.[0])} style={{ display: "none" }} />
                  <button style={chooseBtn} onClick={() => fileInputRef.current?.click()}>
                    Choose file…
                  </button>

                  <div style={{ opacity: 0.85, fontWeight: 900 }}>
                    Imported rows: <b>{wideRows.length}</b> • Lines total (L+R): <b>{summary.totalLines}</b>
                  </div>

                  {importStatus.ok ? (
                    <div style={{ opacity: 0.90, fontWeight: 900 }}>
                      File: <b>{importStatus.name}</b>
                    </div>
                  ) : importStatus.err ? (
                    <div style={{ color: theme.bad, fontWeight: 950 }}>{importStatus.err}</div>
                  ) : null}

                  {loaded ? (
                    <div style={{ marginLeft: "auto" }}>
                      <button style={topBtn} onClick={() => setStep(2)}>
                        Go to Step 2 →
                      </button>
                    </div>
                  ) : null}
                </div>
              ) : (
                <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
                  <button style={chooseBtn} onClick={loadAttachedTestData}>
                    Load attached test data
                  </button>
                  <div style={{ opacity: 0.86, fontWeight: 900 }}>
                    Uses embedded <b>Speedster3 ML.csv</b>.
                  </div>
                </div>
              )}

              <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 10 }}>
                <div style={card}>
                  <div style={cardLabel}>Make</div>
                  <div style={cardValue}>{meta.make || "—"}</div>
                </div>
                <div style={card}>
                  <div style={cardLabel}>Model</div>
                  <div style={cardValue}>{meta.model || "—"}</div>
                </div>
                <div style={card}>
                  <div style={cardLabel}>Tolerance (mm)</div>
                  <div style={{ marginTop: 8 }}>
                    <NumInput value={meta.tolerance} onChange={(v) => setMeta((m) => ({ ...m, tolerance: clamp(v, 0, 999) }))} width={110} />
                  </div>
                </div>
                <div style={card}>
                  <div style={cardLabel}>Correction</div>
                  <div style={{ marginTop: 8 }}>
                    <NumInput value={meta.correction} onChange={(v) => setMeta((m) => ({ ...m, correction: Number(v || 0) }))} width={110} />
                  </div>
                </div>
              </div>
            </div>
          </Panel>
        ) : null}

        {/* Step 2 */}
        {step === 2 ? (
          <Panel
            tint
            title="Step 2 — Map lines to maillon groups (setup)"
            right={
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                <button style={topBtn} onClick={() => setStep(1)}>
                  ← Back
                </button>
                <button style={{ ...topBtn, background: "rgba(59,130,246,0.20)" }} onClick={() => rebuildMappingFromRanges(false)} disabled={!loaded}>
                  Apply ranges
                </button>
                <button style={{ ...topBtn, background: "rgba(239,68,68,0.12)" }} onClick={() => rebuildMappingFromRanges(true)} disabled={!loaded}>
                  Reset to ranges
                </button>
                <button style={{ ...topBtn, background: "rgba(59,130,246,0.22)" }} onClick={() => setStep(3)} disabled={!loaded}>
                  Open diagram page →
                </button>
              </div>
            }
          >
            {!loaded ? (
              <WarningBanner title="No file loaded">Go back to Step 1 and import a file (or load test data) before using Step 2.</WarningBanner>
            ) : (
              <div style={{ display: "grid", gap: 10 }}>
                <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", justifyContent: "space-between" }}>
                    <div>
                      <div style={{ fontWeight: 950 }}>Wing profile — Export / Import (JSON)</div>
                      <div style={{ opacity: 0.76, fontSize: 13, marginTop: 4 }}>Export saves Step 2 mapping. Import applies it to the current file’s lines.</div>
                    </div>

                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                      <input
                        value={profileName}
                        onChange={(e) => setProfileName(e.target.value)}
                        placeholder="Profile file name…"
                        style={{ width: 240, padding: "7px 10px", borderRadius: 12, border: `1px solid ${theme.border}`, background: "rgba(0,0,0,0.68)", color: theme.text, outline: "none", fontWeight: 900, fontSize: 14 }}
                      />
                      <button style={{ ...topBtn, background: "rgba(59,130,246,0.20)" }} onClick={exportWingProfileJSON}>
                        Export JSON
                      </button>

                      <input ref={profileImportRef} type="file" accept=".json,application/json" style={{ display: "none" }} onChange={(e) => importWingProfileJSON(e.target.files?.[0])} />
                      <button style={topBtn} onClick={() => profileImportRef.current?.click()}>
                        Import JSON
                      </button>
                    </div>
                  </div>
                </div>

                <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
                  <div style={{ fontWeight: 950 }}>Defaults</div>

                  <div style={{ marginTop: 10 }}>
                    <div style={{ fontWeight: 950, marginBottom: 8 }}>Prefixes</div>
                    {/* Reduced bucket widths + moved left */}
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 160px)", gap: 8, justifyContent: "start", alignItems: "start" }}>
                      <PrefixTile letter="A" />
                      <PrefixTile letter="B" />
                      <PrefixTile letter="C" />
                      <PrefixTile letter="D" />
                    </div>
                  </div>

                  <div style={{ marginTop: 12 }}>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-start", gap: 10, flexWrap: "wrap" }}>
                      <div style={{ fontWeight: 950 }}>Ranges</div>
                      <ColoredRangeTabs />
                    </div>
                    <div style={{ marginTop: 10 }}>
                      <RangeEditor L={rangeTab} />
                    </div>

                    <div style={{ marginTop: 10, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", justifyContent: "flex-start" }}>
                      <button style={topBtn} onClick={() => setShowOverrides((v) => !v)}>
                        {showOverrides ? "Hide line grouping overrides" : "Show line grouping overrides"}
                      </button>
                      <div style={{ opacity: 0.78, fontWeight: 900, fontSize: 13 }}>Overrides = per-line dropdown assignments.</div>
                    </div>
                  </div>
                </div>

                {showOverrides ? (
                  <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
                    <div style={{ fontWeight: 950 }}>Line grouping overrides</div>

                    <div style={{ marginTop: 10, maxHeight: "60vh", overflow: "auto", border: `1px solid ${theme.border}`, borderRadius: 14 }}>
                      <table style={{ width: "100%", borderCollapse: "collapse" }}>
                        <thead>
                          <tr style={{ background: "rgba(255,255,255,0.05)" }}>
                            <th style={th}>Cascade</th>
                            <th style={th}>Line L</th>
                            <th style={th}>Group L</th>
                            <th style={th}>Line R</th>
                            <th style={th}>Group R</th>
                          </tr>
                        </thead>
                        <tbody>
                          {letterIdxRows.map(({ letter, idx }) => {
                            const lineL = `${letter}${idx}L`;
                            const lineR = `${letter}${idx}R`;
                            const col = chipColorFromLineId(lineL);

                            return (
                              <tr key={`${letter}-${idx}`} style={{ borderTop: `1px solid ${theme.border}` }}>
                                <td style={{ ...td, fontWeight: 950, color: col }}>
                                  {letter}
                                  {idx}
                                </td>
                                <td style={td}>{lineL}</td>
                                <td style={td}>
                                  <Select value={lineToGroup[lineL] || ""} onChange={(v) => setLineToGroup((p) => ({ ...(p || {}), [lineL]: v }))} options={groupOptionsForSelect} width={160} />
                                </td>
                                <td style={td}>{lineR}</td>
                                <td style={td}>
                                  <Select value={lineToGroup[lineR] || ""} onChange={(v) => setLineToGroup((p) => ({ ...(p || {}), [lineR]: v }))} options={groupOptionsForSelect} width={160} />
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ) : null}
              </div>
            )}
          </Panel>
        ) : null}

        {/* Step 3 */}
        {step === 3 ? (
          <Panel
            tint
            title="Step 3 — Diagram + changes summary"
            right={
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                <button style={topBtn} onClick={() => setStep(2)}>
                  ← Back to Step 2
                </button>
                <button style={topBtn} onClick={fitDiagramToScreen}>
                  Fit to screen
                </button>
                <button style={{ ...topBtn, background: "rgba(59,130,246,0.20)" }} onClick={() => rebuildMappingFromRanges(false)}>
                  Apply ranges
                </button>
                <button style={{ ...topBtn, background: "rgba(239,68,68,0.12)" }} onClick={() => rebuildMappingFromRanges(true)}>
                  Reset to ranges
                </button>
                <Toggle value={diagramWingOutline} onChange={setDiagramWingOutline} label="Wing outline" />
                <Toggle value={diagramCompact} onChange={setDiagramCompact} label="Compact" />
                <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 10px", border: `1px solid ${theme.border}`, borderRadius: 14, background: "rgba(0,0,0,0.35)" }}>
                  <button style={miniBtn} onClick={() => setDiagramZoom((z) => clamp(Number((z - 0.1).toFixed(2)), 0.4, 2.0))}>
                    −
                  </button>
                  <input type="range" min={0.4} max={2.0} step={0.01} value={diagramZoom} onChange={(e) => setDiagramZoom(Number(e.target.value))} style={{ width: 180 }} />
                  <button style={miniBtn} onClick={() => setDiagramZoom((z) => clamp(Number((z + 0.1).toFixed(2)), 0.4, 2.0))}>
                    +
                  </button>
                  <span style={{ opacity: 0.82, fontWeight: 900, fontSize: 13, minWidth: 54, textAlign: "right" }}>{(diagramZoom * 100).toFixed(0)}%</span>
                </div>
              </div>
            }
          >
            {!loaded ? (
              <WarningBanner title="No file loaded">Go back to Step 1 and import a file (or load test data).</WarningBanner>
            ) : (
              <div style={{ display: "grid", gap: 10 }}>
                {/* Step 3 includes Apply/Reset buttons (same as Step 2) */}
                <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                    <div>
                      <div style={{ fontWeight: 950 }}>Map lines to maillon groups (setup)</div>
                      <div style={{ opacity: 0.78, fontSize: 13, marginTop: 4 }}>Apply/Reset ranges here without leaving the diagram.</div>
                    </div>

                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                      <button style={{ ...topBtn, background: "rgba(59,130,246,0.20)" }} onClick={() => rebuildMappingFromRanges(false)}>
                        Apply ranges
                      </button>
                      <button style={{ ...topBtn, background: "rgba(239,68,68,0.12)" }} onClick={() => rebuildMappingFromRanges(true)}>
                        Reset to ranges
                      </button>
                      <button style={topBtn} onClick={() => setStep(2)}>
                        Edit ranges →
                      </button>
                    </div>
                  </div>
                </div>

                <div style={{ display: "grid", gap: 10 }}>
                  <div>
                    <div style={{ fontWeight: 950, marginBottom: 8 }}>Live grouping diagram (drag chips to move)</div>
                    {/* Fixed size + scrollbars */}
                    <DiagramScrollBox height={640} width={980} />
                    <div style={{ marginTop: 8, opacity: 0.78, fontSize: 13 }}>
                      Drag any line chip into another group bucket. Scrollbars are always visible (bottom + right).
                    </div>
                  </div>

                  <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
                    <div style={{ fontWeight: 950 }}>Changes summary</div>
                    <div style={{ opacity: 0.78, fontSize: 13, marginTop: 4 }}>
                      Shows lines moved compared to the <b>default mapping</b> created when you imported the file (or clicked “Reset to ranges”).
                    </div>

                    <div style={{ marginTop: 10, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                      <div style={{ opacity: 0.86, fontWeight: 900 }}>
                        Changed lines: <b>{changes.length}</b>
                      </div>
                      <button
                        style={{ ...topBtn, background: "rgba(59,130,246,0.20)" }}
                        onClick={() => {
                          const payload = { schema: "trim-tuning-step2-changes-v1", exportedAt: new Date().toISOString(), wing: { make: meta.make || "", model: meta.model || "" }, changes };
                          const name = `${meta.make || "Wing"}-${meta.model || "Changes"}`.replace(/\s+/g, "-");
                          downloadJSON(payload, `${name}-mapping-changes.json`);
                        }}
                        disabled={changes.length === 0}
                      >
                        Export changes JSON
                      </button>
                    </div>

                    <div style={{ marginTop: 10, maxHeight: 320, overflow: "auto", border: `1px solid ${theme.border}`, borderRadius: 14 }}>
                      {changes.length === 0 ? (
                        <div style={{ padding: 10, opacity: 0.78, fontWeight: 900 }}>No overrides yet. Drag a line into a new bucket or use dropdowns in Step 2.</div>
                      ) : (
                        <table style={{ width: "100%", borderCollapse: "collapse" }}>
                          <thead>
                            <tr style={{ background: "rgba(255,255,255,0.05)" }}>
                              <th style={th}>Line</th>
                              <th style={th}>From</th>
                              <th style={th}>To</th>
                            </tr>
                          </thead>
                          <tbody>
                            {changes.map((c) => (
                              <tr key={c.lineId} style={{ borderTop: `1px solid ${theme.border}` }}>
                                <td style={{ ...td, color: chipColorFromLineId(c.lineId), fontWeight: 950 }}>{c.lineId}</td>
                                <td style={td}>{c.from}</td>
                                <td style={td}>{c.to}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      )}
                    </div>

                    <div style={{ marginTop: 10 }}>
                      <WarningBanner title="Workflow tip">Once this diagram looks correct, we’ll plug mapping into loops + trimming next.</WarningBanner>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </Panel>
        ) : null}
      </div>
    </div>
  );
}

// ---------------- Styles ----------------
const th = { textAlign: "left", padding: "8px 8px", fontSize: 12, fontWeight: 950, color: "rgba(255,255,255,0.82)", whiteSpace: "nowrap" };
const td = { padding: "8px 8px", fontSize: 12, fontWeight: 850, color: "rgba(255,255,255,0.90)", whiteSpace: "nowrap" };

const card = { border: `1px solid ${theme.border}`, borderRadius: 16, background: "linear-gradient(180deg, rgba(0,0,0,0.38), rgba(255,255,255,0.03))", padding: 10 };
const cardLabel = { opacity: 0.78, fontWeight: 900, fontSize: 13 };
const cardValue = { marginTop: 4, fontWeight: 950, fontSize: 14 };

const topBtn = {
  padding: "9px 12px",
  borderRadius: 12,
  border: `1px solid ${theme.border}`,
  background: "rgba(255,255,255,0.08)",
  color: theme.text,
  cursor: "pointer",
  fontWeight: 950,
  fontSize: 14,
};
const topBtnActive = { background: "rgba(59,130,246,0.25)" };
const chooseBtn = {
  padding: "10px 14px",
  borderRadius: 12,
  border: `1px solid ${theme.border}`,
  background: "rgba(59,130,246,0.20)",
  color: theme.text,
  cursor: "pointer",
  fontWeight: 950,
  fontSize: 14,
};
const miniBtn = {
  width: 36,
  height: 34,
  borderRadius: 10,
  border: `1px solid ${theme.border}`,
  background: "rgba(255,255,255,0.08)",
  color: theme.text,
  cursor: "pointer",
  fontWeight: 950,
};