import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

const SITE_VERSION = "Trim Tuning v1.2";


// Step 3 – Loop sizes (mm) are wing-specific and must be set before baseline loops
const DEFAULT_LOOP_SIZES = {
  SL: 0,
  DL: -7,
  TL: -10,
  AS: -12,
  "AS+": -16,
  "AS++": -20,
  CUSTOM: 0,
};
const LOOP_TYPES = Object.keys(DEFAULT_LOOP_SIZES);


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
  bg: "#12151b",
  panel: "rgba(255,255,255,0.08)",
  panel2: "rgba(0,0,0,0.22)",
  // Shared dark pill background used by StatPill / ControlPill / TogglePill
  bg2: "rgba(0,0,0,0.22)",
  border: "rgba(255,255,255,0.14)",
  text: "rgba(255,255,255,0.92)",
  textSub: "rgba(170,177,195,0.85)",
  green: "rgba(34,197,94,0.95)",
  good: "rgba(34,197,94,0.95)",
  bad: "rgba(239,68,68,0.95)",
  warn: "rgba(245,158,11,0.95)",
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
  const s = String(v ?? "").trim();
  if (!s) return null; // treat blanks as missing, not 0
  const n = Number(s.replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

function median(values) {
  const nums = (values || []).filter((x) => typeof x === "number" && Number.isFinite(x));
  if (nums.length === 0) return null;
  nums.sort((a, b) => a - b);
  const mid = Math.floor(nums.length / 2);
  return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
}

function deepClone(obj) {
  try {
    return JSON.parse(JSON.stringify(obj));
  } catch (e) {
    // Fall back to structuredClone where available.
    try {
      // eslint-disable-next-line no-undef
      if (typeof structuredClone === "function") return structuredClone(obj);
    } catch (e2) {
      // ignore
    }
    return obj;
  }
}

function bandForDelta(delta, tolerance) {
  if (delta == null || !Number.isFinite(Number(delta))) return "na";
  const abs = Math.abs(Number(delta));
  const tol = Number.isFinite(Number(tolerance)) ? Number(tolerance) : 0;
  if (abs <= 4) return "good"; // GREEN within ±4mm
  if (abs < tol) return "warn"; // YELLOW (over 4mm but still within tolerance)
  return "bad"; // RED at/over tolerance
}

function severity(delta, tolerance) {
  // Returns "green" | "yellow" | "red" | "na" for chart coloring.
  if (delta == null || !Number.isFinite(Number(delta))) return "na";
  const abs = Math.abs(Number(delta));
  const tol = Number.isFinite(Number(tolerance)) ? Number(tolerance) : 0;
  if (abs <= 4) return "green";
  if (tol > 0 && abs >= tol) return "red";
  if (tol > 0) return "yellow";
  // If tolerance isn't set, fall back to: green ≤4, yellow >4.
  return "yellow";
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
  const letters = ["A", "B", "C", "D"]; // core riser cascades
  const BRAKE_LETTER = "BR";

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

  // Optional brake columns commonly labeled BRK/BRKL/BRKR in some line-check sheets.
  const headerU = header.map((h) => String(h || "").trim().toUpperCase());
  const brkStart = headerU.findIndex((h) => h === "BRK" || h === "BRAKE" || h === "BR");
  const brkLIdx = headerU.findIndex((h) => h === "BRKL" || h === "BRK L" || h === "BRAKEL" || h === "BRAKE L");
  const brkRIdx = headerU.findIndex((h) => h === "BRKR" || h === "BRK R" || h === "BRAKER" || h === "BRAKE R");

  for (let r = 3; r < rows.length; r++) {
    const row = rows[r] || [];
    let anyIdx = null;
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
      if (anyIdx == null && idx != null) anyIdx = idx;

      if (!["A", "B", "C", "D"].includes(letter) || idx == null) continue;
      wideRows.push({ letter, idx, lineBase: `${letter}${idx}`, nominal, measuredL, measuredR });

    }

    // Brake row parsing: attach per-index brake measurements if present.
    // Uses same idx as the current row's A/B/C/D index when available.
    if ((brkStart >= 0 || (brkLIdx >= 0 && brkRIdx >= 0))) {
      const idx = anyIdx != null ? Number(anyIdx) : null;
      if (idx != null && Number.isFinite(idx)) {
        const nominal =
          brkStart >= 0 ? safeNum(row[brkStart + 1]) :
          (brkLIdx > 0 && String(headerU[brkLIdx - 1] || "").includes("SOLL") ? safeNum(row[brkLIdx - 1]) : null);
        const nominalClean = nominal === 0 ? null : nominal; // treat 0 factory as missing for brakes

        const measuredL = brkStart >= 0 ? safeNum(row[brkStart + 2]) : safeNum(row[brkLIdx]);
        const measuredR = brkStart >= 0 ? safeNum(row[brkStart + 3]) : safeNum(row[brkRIdx]);

        // Only push if there is at least one numeric measurement.
        if (nominalClean != null || measuredL != null || measuredR != null) {
          wideRows.push({
            letter: BRAKE_LETTER,
            idx,
            lineBase: `${BRAKE_LETTER}${idx}`,
            nominal: nominalClean,
            measuredL,
            measuredR,
          });
        }
      }
    }
  }

  return { meta, wideRows };
}

// ---------------- Step 2 grouping ----------------
function makeDefaultRanges(maxIdx, groupCount) {
  var m = Math.max(1, Number(maxIdx || 1));
  var n = Math.max(1, Number(groupCount || 3));

  // Default paragliding convention:
  // - 3 groups: AR1=4 lines, AR2=4 lines, AR3=rest
  // - 4 groups: AR1=4 lines, AR2=4 lines, AR3=4 lines, AR4=rest
  // Fallback: if m is small, ranges are clamped and remain contiguous.
  var out = {};

  function setBucket(b, s, e) {
    var ss = Math.max(1, Math.min(m, s));
    var ee = Math.max(1, Math.min(m, e));
    if (ee < ss) ee = ss;
    out[b] = { start: ss, end: ee };
  }

  if (n === 3 || n === 4) {
    var s1 = 1;
    var e1 = Math.min(m, 4);
    setBucket(1, s1, e1);

    var s2 = e1 + 1;
    if (s2 > m) s2 = m;
    var e2 = Math.min(m, s2 + 4 - 1);
    setBucket(2, s2, e2);

    if (n === 3) {
      var s3 = e2 + 1;
      if (s3 > m) s3 = m;
      setBucket(3, s3, m);
      return out;
    }

    // n === 4
    var s3 = e2 + 1;
    if (s3 > m) s3 = m;
    var e3 = Math.min(m, s3 + 4 - 1);
    setBucket(3, s3, e3);

    var s4 = e3 + 1;
    if (s4 > m) s4 = m;
    setBucket(4, s4, m);
    return out;
  }

  // Fallback: evenly split 1..m into n contiguous buckets.
  var step = Math.ceil(m / n);
  var start = 1;
  for (var b = 1; b <= n; b++) {
    var s = start;
    if (s > m) s = m;
    var e = (b === n) ? m : Math.min(m, s + step - 1);
    if (e < s) e = s;
    out[b] = { start: s, end: e };
    start = e + 1;
  }
  return out;
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
const DIAGRAM_SCALE = 0.9;
const DIAGRAM_BASE_W = 2400;
const DIAGRAM_BASE_H = 980;
const DIAGRAM_W = Math.round(DIAGRAM_BASE_W * DIAGRAM_SCALE);
const DIAGRAM_H = Math.round(DIAGRAM_BASE_H * DIAGRAM_SCALE);

function DiagramPreview({ lineToGroup, prefixByLetter, groupCountByLetter, showWingOutline, compactLayout, setLineToGroupFromDrag, changedLineIds }) {
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
    var minBw = Math.round((compactLayout ? 240 : 280) * DIAGRAM_SCALE);
    return { bw: Math.max(innerW + padW, minBw), bh: innerH + padH, cols };
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
      const moved = !!changedLineIds?.has?.(lineId);
      const chipText = moved ? theme.bad : "rgba(255,255,255,0.93)";

      items.push(
        <g key={`chip-${b.groupId}-${lineId}`}>
          <rect x={x} y={yy} width={chip.w} height={chip.h} rx={12} ry={12} fill="rgba(255,255,255,0.10)" stroke={chipStroke} strokeOpacity={0.60} />
          <text x={x + chip.w / 2} y={yy + chip.h * 0.74} textAnchor="middle" fill={chipText} fontSize={18} fontWeight={950}>
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

        {null}

    </div>
  );
}

// ---------------- UI atoms ----------------

function BlockTable({ title, rows, theme, th, td, showCorrected, tolerance = 10, step4LineCorr, setStep4LineCorr }) {
  if (!Array.isArray(rows) || rows.length === 0) return null;

  const cellStyle = (band) => {
    if (band === "good") return { background: "rgba(34,197,94,0.14)", borderColor: "rgba(34,197,94,0.35)" };
    if (band === "warn") return { background: "rgba(234,179,8,0.14)", borderColor: "rgba(234,179,8,0.35)" };
    if (band === "bad") return { background: "rgba(239,68,68,0.14)", borderColor: "rgba(239,68,68,0.35)" };
    return {};
  };

  const fmt = (v, digits = 0) => (v == null || !Number.isFinite(Number(v)) ? "—" : Number(v).toFixed(digits));

  const bandFromDelta = (delta) => {
    const d = Number(delta);
    if (!Number.isFinite(d)) return "";
    const a = Math.abs(d);
    if (a <= 4) return "good";
    if (a < Number(tolerance)) return "warn";
    return "bad";
  };
  const textColorForBand = (band) => {
    if (band === "good") return "rgba(34,197,94,0.95)";
    if (band === "warn") return "rgba(234,179,8,0.95)";
    if (band === "bad") return "rgba(239,68,68,0.95)";
    return theme.text;
  };

  return (
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 10, flexWrap: "wrap" }}>
        <div style={{ fontWeight: 950 }}>{title}</div>
        <div style={{ opacity: 0.78, fontSize: 12 }}>
          Display: <b>{showCorrected ? "corrected" : "raw"}</b>
        </div>
      </div>

      <div style={{ marginTop: 10, overflow: "auto", width: "100%", maxWidth: "100%",
          minWidth: 0, border: `1px solid ${theme.border}`, borderRadius: 12 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 980 }}>
          <thead>
            <tr>
              <th style={{ ...th }}>Line</th>
              <th style={{ ...th }}>Nominal</th>

              <th style={{ ...th }}>L {showCorrected ? "Val" : "Raw"}</th>
              <th style={{ ...th }}>L Corr (mm)</th>
              <th style={{ ...th }}>L Before</th>
              <th style={{ ...th }}>L After</th>
              <th style={{ ...th }}>L Δ</th>

              <th style={{ ...th }}>R {showCorrected ? "Val" : "Raw"}</th>
              <th style={{ ...th }}>R Corr (mm)</th>
              <th style={{ ...th }}>R Before</th>
              <th style={{ ...th }}>R After</th>
              <th style={{ ...th }}>R Δ</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => {
              const L = r.L;
              const R = r.R;
              return (
                <tr key={r.lineBase}>
                  <td style={{ ...td, fontWeight: 900 }}>{r.lineBase}</td>
                  <td style={{ ...td }}>{fmt(r.nominal)}</td>

                  <td style={{ ...td }}>{fmt(L?.corrected)}</td>
                  <td style={{ ...td }}>
                    <input
                      type="number"
                      value={Number(step4LineCorr?.[`${r.lineBase}L`] ?? 0)}
                      onChange={(e) => {
                        const v = Number(e.target.value);
                        setStep4LineCorr?.((prev) => ({ ...prev, [`${r.lineBase}L`]: Number.isFinite(v) ? v : 0 }));
                      }}
                      style={{ width: 86, padding: "6px 8px", borderRadius: 10, border: `1px solid ${theme.border}`, background: theme.panel, color: theme.text, outline: "none" }}
                    />
                  </td>
                  <td style={{ ...td }}>{fmt(L?.before)}</td>
                  <td style={{ ...td }}>{fmt(L?.afterVal)}</td>
                  {(() => { const band = bandFromDelta(L?.delta); return (<td style={{ ...td, border: `1px solid ${theme.border}`, ...cellStyle(band), color: textColorForBand(band), fontWeight: 900 }}>{fmt(L?.delta)}</td>); })()}

                  <td style={{ ...td }}>{fmt(R?.corrected)}</td>
                  <td style={{ ...td }}>
                    <input
                      type="number"
                      value={Number(step4LineCorr?.[`${r.lineBase}R`] ?? 0)}
                      onChange={(e) => {
                        const v = Number(e.target.value);
                        setStep4LineCorr?.((prev) => ({ ...prev, [`${r.lineBase}R`]: Number.isFinite(v) ? v : 0 }));
                      }}
                      style={{ width: 86, padding: "6px 8px", borderRadius: 10, border: `1px solid ${theme.border}`, background: theme.panel, color: theme.text, outline: "none" }}
                    />
                  </td>
                  <td style={{ ...td }}>{fmt(R?.before)}</td>
                  <td style={{ ...td }}>{fmt(R?.afterVal)}</td>
                  {(() => { const band = bandFromDelta(R?.delta); return (<td style={{ ...td, border: `1px solid ${theme.border}`, ...cellStyle(band), color: textColorForBand(band), fontWeight: 900 }}>{fmt(R?.delta)}</td>); })()}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={{ marginTop: 8, opacity: 0.72, fontSize: 12 }}>
        Colors: <span style={{ color: "rgb(34,197,94)" }}>green</span> ≤ ±4mm,{" "}
        <span style={{ color: "rgb(234,179,8)" }}>yellow</span> &gt; ±4mm but within tolerance,{" "}
        <span style={{ color: "rgb(239,68,68)" }}>red</span> ≥ tolerance.
      </div>
    </div>
  );
}


function Panel({ title, right, children, tint = false }) {
  return (
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 18, background: tint ? "linear-gradient(180deg, rgba(59,130,246,0.08), rgba(255,255,255,0.04))" : theme.panel, overflow: "visible" }}>
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

function StatPill({ label, value, n }) {
  const v = Number.isFinite(value) ? Math.round(value) : null;
  return (
    <div
      style={{
        border: `1px solid ${theme.border}`,
        background: theme.bg2,
        borderRadius: 999,
        padding: "8px 10px",
        display: "flex",
        gap: 8,
        alignItems: "baseline",
      }}
    >
      <span style={{ fontSize: 12, opacity: 0.8 }}>{label}</span>
      <span style={{ fontWeight: 900 }}>{v == null ? "—" : `${v} mm`}</span>
      <span style={{ fontSize: 12, opacity: 0.6 }}>{typeof n === "number" ? `n=${n}` : ""}</span>
    </div>
  );
}



function ControlPill({ label, value, onChange, suffix = "mm", width = 90, step = 1, min, max, inputColor = theme.text }) {
  return (
    <div
      style={{
        border: `1px solid ${theme.border}`,
        background: theme.bg2,
        borderRadius: 999,
        padding: "8px 10px",
        display: "flex",
        gap: 8,
        alignItems: "baseline",
      }}
    >
      <span style={{ fontSize: 12, opacity: 0.8 }}>{label}</span>
      <input
        type="number"
        value={Number.isFinite(value) ? value : 0}
        step={step}
        min={min}
        max={max}
        onChange={(e) => onChange(Number(e.target.value || 0))}
        style={{
          width,
          background: "transparent",
          color: theme.text,
          border: `1px solid ${theme.border}`,
          borderRadius: 12,
          padding: "6px 8px",
          fontWeight: 900,
          outline: "none",
        }}
      />
      <span style={{ fontSize: 12, opacity: 0.75 }}>{suffix}</span>
    </div>
  );
}

function TogglePill({ label, checked, onChange }) {
  return (
    <button
      type="button"
      onClick={() => onChange(!checked)}
      style={{
        border: `1px solid ${theme.border}`,
        background: theme.bg2,
        borderRadius: 999,
        padding: "8px 10px",
        display: "flex",
        gap: 10,
        alignItems: "center",
        cursor: "pointer",
        color: theme.text,
      }}
      title={label}
    >
      <span style={{ fontSize: 12, opacity: 0.8 }}>{label}</span>
      <span
        style={{
          width: 36,
          height: 20,
          borderRadius: 999,
          border: `1px solid ${theme.border}`,
          background: checked ? "rgba(34,197,94,0.35)" : "rgba(148,163,184,0.18)",
          position: "relative",
        }}
      >
        <span
          style={{
            position: "absolute",
            top: 2,
            left: checked ? 18 : 2,
            width: 16,
            height: 16,
            borderRadius: 999,
            background: checked ? "rgba(34,197,94,0.95)" : "rgba(148,163,184,0.65)",
            transition: "left 120ms ease",
          }}
        />
      </span>
    </button>
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

  // Step 3 sub-page view (keeps Step 3 from getting too busy)
  const [step3View, setStep3View] = useState("diagram");

  // Step 3 baseline loops (installed on maillon groups)
  const [loopSizes, setLoopSizes] = useState(() => deepClone(DEFAULT_LOOP_SIZES));
  const [groupLoopSetup, setGroupLoopSetup] = useState({});

  // Step 4 (trim) — frozen snapshot of Step 3 baseline loops (set ONCE when entering Step 4)
  const [groupLoopBaseline, setGroupLoopBaseline] = useState(null); // Record<groupId, loopType> | null

  // Step 4 editable state (must NOT affect Step 3 baseline)
  const [groupLoopChange, setGroupLoopChange] = useState({}); // optional override: Record<groupId, loopType>
  const [groupAdjustments, setGroupAdjustments] = useState({}); // mm adjust: Record<groupId, number>

  // Step 4 per-line correction (mm). Applies to corrected value in Step 4 only.
  const [step4LineCorr, setStep4LineCorr] = useState({}); // Record<lineId, number>

  // Step 4 view options
  const [showCorrected, setShowCorrected] = useState(true);
  const [includeBrakeBlock, setIncludeBrakeBlock] = useState(true);
  const [showLoopModeCounts, setShowLoopModeCounts] = useState(false);
  const [groupPitchTol, setGroupPitchTol] = useState(4);
  const [autoLoopStatus, setAutoLoopStatus] = useState(null); // "factory" | "minimal" | null

  const [step4LetterFilter, setStep4LetterFilter] = useState({ A: true, B: true, C: true, D: true, BR: true });


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
    setStep3View("diagram");
    setLoopSizes(deepClone(DEFAULT_LOOP_SIZES));
    setGroupLoopSetup({});
    setGroupLoopBaseline(null);
    setGroupLoopChange({});
    setGroupAdjustments({});
    setStep4LineCorr({});
    setShowCorrected(true);
    setStep4LetterFilter({ A: true, B: true, C: true, D: true, BR: true });
    setRangeTab("A");
    setShowOverrides(true);

    setDiagramZoom(1.0);
    setDiagramWingOutline(true);
    setDiagramCompact(false);

    setDefaultMappingSnapshot(null);
    setProfileName("");
  }

  function confirmResetAll() {
    const ok = window.confirm(
      "This will reset EVERYTHING (imported data, mapping, baseline loops, frozen baseline, and all Step 4 adjustments). Continue?"
    );
    if (ok) resetAll();
  }


  function applyParsedImport(parsed, name) {
    setMeta(parsed.meta);
    setWideRows(parsed.wideRows);
    setImportStatus({ ok: true, name, err: "" });

    const maxCore = { A: 0, B: 0, C: 0, D: 0 };
    let maxBR = 0;
    for (const r of parsed.wideRows) {
      const L = String(r?.letter || "").toUpperCase();
      if (r?.idx == null) continue;
      const ix = Number(r.idx || 0);
      if (!Number.isFinite(ix)) continue;
      if (L === "BR") {
        maxBR = Math.max(maxBR, ix);
        continue;
      }
      if (!maxCore.hasOwnProperty(L)) continue;
      maxCore[L] = Math.max(maxCore[L] || 0, ix);
    }
    setMaxByLetter({ ...maxCore, BR: maxBR });

    const newRanges = {
      A: makeDefaultRanges(maxCore.A || 1, groupCountByLetter.A || 3),
      B: makeDefaultRanges(maxCore.B || 1, groupCountByLetter.B || 3),
      C: makeDefaultRanges(maxCore.C || 1, groupCountByLetter.C || 3),
      D: makeDefaultRanges(maxCore.D || 1, groupCountByLetter.D || 3),
    };
    setRangesByLetter(newRanges);

    const initMap = buildInitialLineToGroup({
      maxByLetter: maxCore,
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
    const maxCore = { A: 0, B: 0, C: 0, D: 0 };
    let maxBR = 0;
    for (const r of wideRows) {
      const L = String(r?.letter || "").toUpperCase();
      if (r?.idx == null) continue;
      const ix = Number(r.idx || 0);
      if (!Number.isFinite(ix)) continue;
      if (L === "BR") {
        maxBR = Math.max(maxBR, ix);
        continue;
      }
      if (!maxCore.hasOwnProperty(L)) continue;
      maxCore[L] = Math.max(maxCore[L] || 0, ix);
    }
    const totalLines = (maxCore.A + maxCore.B + maxCore.C + maxCore.D) * 2;
    return { max: { ...maxCore, BR: maxBR }, totalLines };
  }, [wideRows]);

  const groupsInUse = useMemo(() => {
    const vals = Object.values(lineToGroup || {}).filter(Boolean);
    const uniq = Array.from(new Set(vals));
    const parse = (s) => {
      const m = String(s || "").match(/^([A-Z]+)(\d+)([LR])$/i);
      if (!m) return { p: String(s || ""), n: 0, side: "" };
      return { p: m[1].toUpperCase(), n: Number(m[2] || 0), side: m[3].toUpperCase() };
    };
    uniq.sort((a, b) => {
      const A = parse(a);
      const B = parse(b);
      if (A.p !== B.p) return A.p.localeCompare(B.p);
      if (A.n !== B.n) return A.n - B.n;
      return (A.side || "").localeCompare(B.side || "");
    });
    return uniq;
  }, [lineToGroup]);

  const zeroingStats = useMemo(() => {
    const all = [];
    const left = [];
    const right = [];
    for (const r of wideRows || []) {
      const nominal = typeof r.nominal === "number" ? r.nominal : safeNum(r.nominal);
      const ml = typeof r.measuredL === "number" ? r.measuredL : safeNum(r.measuredL);
      const mr = typeof r.measuredR === "number" ? r.measuredR : safeNum(r.measuredR);
      if (Number.isFinite(nominal) && Number.isFinite(ml)) {
        const d = nominal - ml;
        all.push(d);
        left.push(d);
      }
      if (Number.isFinite(nominal) && Number.isFinite(mr)) {
        const d = nominal - mr;
        all.push(d);
        right.push(d);
      }
    }
    const wholeMedian = median(all);
    const leftMedian = median(left);
    const rightMedian = median(right);
    return {
      wholeMedian,
      leftMedian,
      rightMedian,
      nAll: all.length,
      nLeft: left.length,
      nRight: right.length,
    };
  }, [wideRows]);


  const groupOptionsForSelect = useMemo(() => {
    const opts = getGroupOptions(prefixByLetter, groupCountByLetter);
    return opts.map((g) => ({ value: g, label: g }));
  }, [prefixByLetter, groupCountByLetter]);

  const loopTypeOptions = useMemo(() => LOOP_TYPES.map((t) => ({ value: t, label: t })), []);

  // Step 4 baseline freeze: take a snapshot of Step 3 loops ONCE when Step 4 is first entered.
  useEffect(() => {
    if (step !== 4) return;
    if (groupLoopBaseline !== null) return;
    // Freeze exactly once per reset/import session.
    setGroupLoopBaseline(deepClone(groupLoopSetup || {}));
  }, [step, groupLoopBaseline, groupLoopSetup]);

  const letterIdxRows = useMemo(() => {
    const out = [];
    for (const L of ["A", "B", "C", "D"]) {
      const m = Number(maxByLetter[L] || 0);
      for (let i = 1; i <= m; i++) out.push({ letter: L, idx: i });
    }
    return out;
  }, [maxByLetter]);


  const step4LineRows = useMemo(() => {
    // Build per-line rows for the whole wing. Each side is a separate entity (A1L, A1R, ...).
    if (!Array.isArray(wideRows) || wideRows.length === 0) return [];
    const out = [];
    const tol = Number(meta?.tolerance ?? 0);
    const corr = Number(meta?.correction ?? 0);

    const loopMm = (t) => Number(loopSizes?.[t] ?? 0);

    for (const r of wideRows) {
      const letter = String(r?.letter || "").toUpperCase();
      if (!step4LetterFilter?.[letter]) continue;

      const idx = Number(r?.idx);
      if (!Number.isFinite(idx)) continue;

      const base = `${letter}${idx}`;
      const nominal = Number.isFinite(Number(r?.nominal)) ? Number(r.nominal) : null;

      for (const side of ["L", "R"]) {
        const lineId = `${base}${side}`;
        const measured = side === "L" ? r?.measuredL : r?.measuredR;
        const raw = Number.isFinite(Number(measured)) ? Number(measured) : null;

        const lineCorr = Number(step4LineCorr?.[lineId] ?? 0);

        const groupId = String(lineToGroup?.[lineId] || "");

        const baseLoop = groupLoopBaseline?.[groupId] || "SL";
        const override = groupLoopChange?.[groupId] || "";
        const afterLoop = override || baseLoop;

        const adj = Number(groupAdjustments?.[groupId] ?? 0);

        const corrected = raw == null ? null : raw + lineCorr + (showCorrected ? corr : 0);

        // IMPORTANT: Baseline loops are the *installed* state on the wing.
        // Step 4 must not change the underlying measurements when you set baseline loops in Step 3.
        // We apply ONLY the *difference* when an override loop is selected.
        const loopDelta = override ? (loopMm(override) - loopMm(baseLoop)) : 0;

        const before = corrected;
        const afterVal = corrected == null ? null : corrected + loopDelta + adj;

        const delta = nominal == null || afterVal == null ? null : afterVal - nominal;

        const band = bandForDelta(delta, tol);

        let sev = "na";
        if (delta != null) {
          const ad = Math.abs(delta);
          if (ad >= tol) sev = "bad";
          else if (ad <= 4) sev = "green";
          else if (ad >= Math.max(0, tol - 3)) sev = "warn";
          else sev = "good";
        }

        out.push({
          lineId,
          lineBase: base,
          letter,
          idx,
          side,
          groupId,
          nominal,
          raw,
          corrected,
          baseLoop,
          afterLoop,
          adj,
          before,
          after: afterVal,
          delta,
          sev,
        });
      }
    }

    // Sort: A..D, then idx ascending, then L before R
    out.sort((a, b) => {
      if (a.letter !== b.letter) return a.letter.localeCompare(b.letter);
      if (a.idx !== b.idx) return a.idx - b.idx;
      return a.side.localeCompare(b.side);
    });
    return out;
  }, [
    wideRows,
    lineToGroup,
    groupLoopBaseline,
    groupLoopChange,
    groupAdjustments,
    loopSizes,
    meta?.tolerance,
    meta?.correction,
    showCorrected,
    step4LetterFilter,
  ]);


  



// --- Chart data (uses ONLY Step 4 derived rows; never reads Step 3) ---
const chartPointsByLetter = useMemo(() => {
  const out = { A: [], B: [], C: [], D: [] };
  const maxIdx = {
    A: Number(maxByLetter?.A || 0),
    B: Number(maxByLetter?.B || 0),
    C: Number(maxByLetter?.C || 0),
    D: Number(maxByLetter?.D || 0),
  };

  const sevAfterFor = (sev) => {
    if (sev === 'bad') return 'red';
    if (sev === 'warn') return 'yellow';
    if (sev === 'green') return 'ok';
    if (sev === 'good') return 'ok';
    return 'na';
  };

  for (const r of step4LineRows) {
    const L = String(r.letter || '').toUpperCase();
    if (!out[L]) continue;
    const nominal = Number.isFinite(r.nominal) ? r.nominal : null;
    const beforeAbs = Number.isFinite(r.before) ? r.before : null;
    const afterDelta = Number.isFinite(r.delta) ? r.delta : null;
    const beforeDelta = nominal != null && beforeAbs != null ? beforeAbs - nominal : null;

    const m = maxIdx[L] || 0;
    // xIndex spans left tip -> right tip for each letter.
    const xIndex = r.side === 'L' ? (m - (r.idx || 0)) : (m + (r.idx || 0) - 1);

    out[L].push({
      id: r.lineId,
      line: r.lineId,
      side: r.side,
      xIndex,
      before: beforeDelta,
      after: afterDelta,
      sevAfter: sevAfterFor(r.sev),
    });
  }

  for (const k of Object.keys(out)) {
    out[k].sort((a, b) => a.xIndex - b.xIndex);
  }
  return out;
}, [step4LineRows, maxByLetter]);

const step4GroupStats = useMemo(() => {
  // Average Δ before/after per group + side.
  const acc = new Map();
  for (const r of step4LineRows) {
    const groupName = String(r.groupId || '').trim();
    if (!groupName) continue;
    const key = `${groupName}|${r.side}`;
    if (!acc.has(key)) acc.set(key, { groupName, side: r.side, n: 0, sumAfter: 0, sumBefore: 0, nBefore: 0 });
    const a = acc.get(key);
    if (Number.isFinite(r.delta)) {
      a.sumAfter += r.delta;
      a.n += 1;
    }
    const nominal = Number.isFinite(r.nominal) ? r.nominal : null;
    const beforeAbs = Number.isFinite(r.before) ? r.before : null;
    if (nominal != null && beforeAbs != null) {
      a.sumBefore += (beforeAbs - nominal);
      a.nBefore += 1;
    }
  }
  const out = [];
  for (const a of acc.values()) {
    out.push({
      groupName: a.groupName,
      side: a.side,
      before: a.nBefore ? a.sumBefore / a.nBefore : null,
      after: a.n ? a.sumAfter / a.n : null,
    });
  }
  out.sort((x, y) => {
    const gx = String(x.groupName);
    const gy = String(y.groupName);
    if (gx !== gy) return gx.localeCompare(gy);
    return String(x.side).localeCompare(String(y.side));
  });
  return out;
}, [step4LineRows]);



  const step4BlockRowsByLetter = useMemo(() => {
    const byBase = new Map();
    for (const r of step4LineRows) {
      const base = r?.lineBase;
      if (!base) continue;
      let rec = byBase.get(base);
      if (!rec) {
        rec = { letter: r.letter, idx: r.idx, lineBase: base, nominal: r.nominal ?? null, sides: {} };
        byBase.set(base, rec);
      }
      rec.sides[r.side] = r;
    }
    const out = { A: [], B: [], C: [], D: [] };
    for (const rec of byBase.values()) {
      const L = rec.sides.L || null;
      const R = rec.sides.R || null;
      const letter = String(rec.letter || "").toUpperCase();
      if (!out[letter]) continue;
      out[letter].push({ ...rec, L, R });
    }
    for (const k of Object.keys(out)) {
      out[k].sort((a, b) => (a.idx ?? 0) - (b.idx ?? 0));
    }
    return out;
  }, [step4LineRows]);

  
// Step 4 — Spreadsheet-style A/B/C/D summary (Factory / Left / Right / Δ / Sym)
const step4SheetByLetter = useMemo(() => {
  const letters = ["A", "B", "C", "D", "BR"];
  const byLetter = {};
  for (const L of letters) byLetter[L] = [];

  // Collect per (letter, idx) with left/right pivot
  const map = new Map(); // key = `${letter}:${idx}`
  for (const r of step4LineRows) {
    const letter = String(r?.letter || "").toUpperCase();
    if (!byLetter[letter]) continue;
    const idx = Number(r?.idx);
    if (!Number.isFinite(idx)) continue;

    const key = `${letter}:${idx}`;
    const cur = map.get(key) || {
      letter,
      idx,
      factory: Number.isFinite(Number(r?.nominal)) ? Number(r.nominal) : null,
      L: null,
      R: null,
      dL: null,
      dR: null,
    };

    if (r?.side === "L") {
      cur.L = Number.isFinite(Number(r?.after)) ? Number(r.after) : null;
      cur.dL = Number.isFinite(Number(r?.delta)) ? Number(r.delta) : null;
    } else if (r?.side === "R") {
      cur.R = Number.isFinite(Number(r?.after)) ? Number(r.after) : null;
      cur.dR = Number.isFinite(Number(r?.delta)) ? Number(r.delta) : null;
    }

    // Prefer the nominal from whichever side has it
    if (cur.factory == null && Number.isFinite(Number(r?.nominal))) cur.factory = Number(r.nominal);

    map.set(key, cur);
  }

  for (const v of map.values()) {
    const sym = v.L == null || v.R == null ? null : v.L - v.R;
    byLetter[v.letter].push({ ...v, sym });
  }

  for (const L of Object.keys(byLetter)) {
    byLetter[L].sort((a, b) => a.idx - b.idx);
  }

  // Max difference per letter for dL/dR/sym (spread = max - min)
  const maxDiff = {};
  const spread = (arr) => {
    const nums = arr.filter((x) => x != null && Number.isFinite(Number(x))).map(Number);
    if (nums.length === 0) return null;
    return Math.max(...nums) - Math.min(...nums);
  };

  for (const L of letters) {
    const rows = byLetter[L];
    maxDiff[L] = {
      dL: spread(rows.map((r) => r.dL)),
      dR: spread(rows.map((r) => r.dR)),
      sym: spread(rows.map((r) => r.sym)),
    };
  }

  return { byLetter, maxDiff };
}, [step4LineRows]);

var abcAverages = useMemo(() => {
  const letters = ["A", "B", "C", "D"];
  const out = {};
  for (const L of letters) {
    out[L] = { L: { avg: null, n: 0 }, R: { avg: null, n: 0 }, sym: null };
    for (const side of ["L", "R"]) {
      const vals = step4LineRows
        .filter((r) => r.letter === L && r.side === side && Number.isFinite(r.delta))
        .map((r) => Number(r.delta));
      const n = vals.length;
      const avg = n ? vals.reduce((a, b) => a + b, 0) / n : null;
      out[L][side] = { avg, n };
    }
    const aL = out[L].L.avg;
    const aR = out[L].R.avg;
    out[L].sym = Number.isFinite(aL) && Number.isFinite(aR) ? aL - aR : null;
  }
  return out;
}, [step4LineRows]);

const pitchStats = useMemo(() => {
  // "Pitch" here is derived from relative front-to-rear group deltas (After Δ vs nominal).
  // Larger + values generally indicate the front groups (A/B) are longer relative to rear (C/D) -> lower AoA / faster trim.
  const getAvg = (letter, side) => {
    const g = abcAverages && abcAverages[letter] ? abcAverages[letter] : null;
    const s = g && g[side] ? g[side] : null;
    const v = s && Number.isFinite(Number(s.avg)) ? Number(s.avg) : null;
    return v;
  };

  const mean2 = (a, b) => (Number.isFinite(a) && Number.isFinite(b) ? (a + b) / 2 : (Number.isFinite(a) ? a : (Number.isFinite(b) ? b : null)));

  const row = (letter) => {
    const L = getAvg(letter, "L");
    const R = getAvg(letter, "R");
    const both = mean2(L, R);
    return { letter, L, R, both };
  };

  const A = row("A");
  const B = row("B");
  const C = row("C");
  const D = row("D");

  // Whole-wing pitch proxy: mean(front) - mean(rear)
  const front = { L: mean2(A.L, B.L), R: mean2(A.R, B.R), both: mean2(A.both, B.both) };
  const rear = { L: mean2(C.L, D.L), R: mean2(C.R, D.R), both: mean2(C.both, D.both) };
  const pitchWhole = {
    L: Number.isFinite(front.L) && Number.isFinite(rear.L) ? front.L - rear.L : null,
    R: Number.isFinite(front.R) && Number.isFinite(rear.R) ? front.R - rear.R : null,
    both: Number.isFinite(front.both) && Number.isFinite(rear.both) ? front.both - rear.both : null,
  };

  // Per-line-group "slope" (adjacent differences)
  const seg = (x, y) => ({
    L: Number.isFinite(x.L) && Number.isFinite(y.L) ? x.L - y.L : null,
    R: Number.isFinite(x.R) && Number.isFinite(y.R) ? x.R - y.R : null,
    both: Number.isFinite(x.both) && Number.isFinite(y.both) ? x.both - y.both : null,
  });

  return {
    rows: [A, B, C, D],
    pitchWhole,
    segments: { AB: seg(A, B), BC: seg(B, C), CD: seg(C, D) },
    front,
    rear,
  };
}, [abcAverages]);



const step4HasLoopEdits = useMemo(() => {
  const changes = groupLoopChange || {};
  for (const [gid, v] of Object.entries(changes)) {
    if (v == null || v === "") continue;
    const base = groupLoopBaseline?.[gid];
    if (v !== base) return true;
  }
  return false;
}, [groupLoopChange, groupLoopBaseline]);


const abcLoopModeCounts = useMemo(() => {
  // Counts of loop types currently in effect per letter+side (A/B/C only).
  // "Currently in effect" means: groupLoopChange overrides baseline, else baseline.
  const letters = ["A", "B", "C", "D"];
  const out = {};
  for (const L of letters) {
    out[L] = { L: {}, R: {} };
    for (const side of ["L", "R"]) out[L][side] = {};
  }

  for (const r of step4LineRows) {
    const L = r.letter;
    const side = r.side;
    if (!out[L] || !out[L][side]) continue;
    const gid = String(r.groupId || "").trim();
    if (!gid) continue;
    const curLoop = step4HasLoopEdits
      ? (groupLoopChange?.[gid] || groupLoopBaseline?.[gid] || "SL")
      : (groupLoopBaseline?.[gid] || "SL");
    out[L][side][curLoop] = (out[L][side][curLoop] || 0) + 1;
  }
  return out;
}, [step4LineRows, groupLoopBaseline, groupLoopChange]);

const abcSuggestions = useMemo(() => {
  const letters = ["A", "B", "C", "D"];
  const tol = Number(meta?.tolerance ?? 0);

  const loopMm = (t) => Number(loopSizes?.[t] ?? 0);

  const loopTypesSorted = [...LOOP_TYPES].sort((a, b) => loopMm(a) - loopMm(b));

  const pickModeLoop = (countsObj) => {
    // Pick most frequent loop; tie-break on smallest mm.
    let best = "SL";
    let bestCount = -1;
    let bestMm = loopMm(best);
    for (const lt of Object.keys(countsObj || {})) {
      const c = Number(countsObj[lt] || 0);
      const mm = loopMm(lt);
      if (c > bestCount || (c === bestCount && mm < bestMm)) {
        best = lt;
        bestCount = c;
        bestMm = mm;
      }
    }
    return best;
  };

  const clampAdj = (v) => {
    if (!Number.isFinite(v)) return null;
    if (!Number.isFinite(tol)) return Math.round(v);
    return Math.round(clamp(v, -tol, tol));
  };

  const out = {};
  for (const L of letters) {
    out[L] = { L: null, R: null };
    for (const side of ["L", "R"]) {
      const avg = abcAverages?.[L]?.[side]?.avg;
      if (!Number.isFinite(avg)) {
        out[L][side] = null;
        continue;
      }

      const neededMm = -Number(avg);

      const counts = abcLoopModeCounts?.[L]?.[side] || {};
      const repLoop = pickModeLoop(counts);
      const repMm = loopMm(repLoop);

      let bestLoop = repLoop;
      let bestLoopDelta = 0;
      let bestErr = Infinity;

      for (const cand of loopTypesSorted) {
        const d = loopMm(cand) - repMm;
        const err = Math.abs(neededMm - d);
        if (err < bestErr) {
          bestErr = err;
          bestLoop = cand;
          bestLoopDelta = d;
        }
      }

      const residual = neededMm - bestLoopDelta;
      const suggestedAdjMm = clampAdj(residual);
      const residualAfterAdj = Number.isFinite(suggestedAdjMm) ? residual - suggestedAdjMm : residual;

      out[L][side] = {
        avgAfterDelta: avg,
        neededMm,
        repLoop,
        bestLoop,
        loopDeltaMm: bestLoopDelta,
        suggestedAdjMm,
        residualAfterAdj,
      };
    }
  }
  return out;
}, [abcAverages, abcLoopModeCounts, loopSizes, meta?.tolerance]);

  function applyAutoLoopPlan(kind) {
    // kind: "factory" (closest to factory) or "minimal" (within tolerance with least loop changes)
    if (step !== 4) return;
    if (!groupLoopBaseline) return;

    var tol = Number(meta && meta.tolerance != null ? meta.tolerance : 0);
    if (!isFinite(tol)) tol = 0;

    // Build avg delta per maillon group using BASELINE loops only (no overrides, no fine adjust).
    var sums = {};
    var counts = {};

    var i;
    for (i = 0; i < (step4LineRows || []).length; i++) {
      var r = step4LineRows[i];
      if (!r) continue;
      var gid = String(r.groupId || "").trim();
      if (!gid) continue;

      var nominal = r.nominal;
      var corrected = r.corrected;
      if (nominal == null || !isFinite(Number(nominal))) continue;
      if (corrected == null || !isFinite(Number(corrected))) continue;

      var baseLoop = (groupLoopBaseline && groupLoopBaseline[gid]) ? groupLoopBaseline[gid] : (r.baseLoop || "SL");
      var baseMm = Number(loopSizes && loopSizes[baseLoop] != null ? loopSizes[baseLoop] : 0);
      if (!isFinite(baseMm)) baseMm = 0;

      var after0 = Number(corrected) + baseMm;
      var d0 = after0 - Number(nominal);
      if (!isFinite(d0)) continue;

      sums[gid] = (sums[gid] || 0) + d0;
      counts[gid] = (counts[gid] || 0) + 1;
    }

    var changes = {};
    var gids = Object.keys(sums);

    for (i = 0; i < gids.length; i++) {
      var gid2 = gids[i];
      var n = counts[gid2] || 0;
      if (!n) continue;

      var avgDelta = sums[gid2] / n;
      if (!isFinite(avgDelta)) continue;

      var baseLoop2 = (groupLoopBaseline && groupLoopBaseline[gid2]) ? groupLoopBaseline[gid2] : "SL";
      var baseMm2 = Number(loopSizes && loopSizes[baseLoop2] != null ? loopSizes[baseLoop2] : 0);
      if (!isFinite(baseMm2)) baseMm2 = 0;

      var best = baseLoop2;
      var bestAbs = Math.abs(avgDelta);
      var bestShiftAbs = 0;

      // For "minimal": if already within tolerance, keep baseline.
      if (kind === "minimal" && tol > 0 && Math.abs(avgDelta) <= tol) {
        best = baseLoop2;
      } else {
        var j;
        var foundWithin = false;
        for (j = 0; j < (LOOP_TYPES || []).length; j++) {
          var cand = LOOP_TYPES[j];
          var candMm = Number(loopSizes && loopSizes[cand] != null ? loopSizes[cand] : 0);
          if (!isFinite(candMm)) candMm = 0;
          var shift = candMm - baseMm2;
          var afterDelta = avgDelta + shift;
          var absAfter = Math.abs(afterDelta);
          var shiftAbs = Math.abs(shift);

          if (kind === "minimal" && tol > 0) {
            if (absAfter <= tol) {
              if (!foundWithin || shiftAbs < bestShiftAbs || (shiftAbs === bestShiftAbs && absAfter < bestAbs)) {
                foundWithin = true;
                best = cand;
                bestAbs = absAfter;
                bestShiftAbs = shiftAbs;
              }
            } else if (!foundWithin) {
              // If none within tol so far, track the best improvement (closest) with smallest shift.
              if (absAfter < bestAbs || (absAfter === bestAbs && shiftAbs < bestShiftAbs)) {
                best = cand;
                bestAbs = absAfter;
                bestShiftAbs = shiftAbs;
              }
            }
          } else {
            // "factory": just minimize residual.
            if (absAfter < bestAbs || (absAfter === bestAbs && shiftAbs < bestShiftAbs)) {
              best = cand;
              bestAbs = absAfter;
              bestShiftAbs = shiftAbs;
            }
          }
        }
      }

      if (best && best !== baseLoop2) {
        changes[gid2] = best;
      }
    }

    // Apply: loops only. Do NOT write fine adjust values here.
    setGroupAdjustments({});
    setGroupLoopChange(changes);
    setAutoLoopStatus(kind === "minimal" ? "minimal" : "factory");
  }


function setRange(letter, bucket, field, value) {
    setRangesByLetter((prev) => {
      const next = { ...(prev || {}) };
      const maxIdx = Math.max(1, Number(maxByLetter[letter] || 1));
      const count = Number(groupCountByLetter[letter] || 3);
      const cur = next[letter] || makeDefaultRanges(maxIdx, count);

      const b = Number(bucket);
      const v = Number(value);
      if (!Number.isFinite(v)) return prev;

      // Always enforce contiguous ranges:
      // - Bucket 1 always starts at 1
      // - Bucket b>1 always starts at prev.end + 1
      // - Last bucket always ends at maxIdx
      const out = { ...cur };

      // Helper to read bucket values with sane defaults
      const getS = (bb) => clamp(Number((out[bb] && out[bb].start) != null ? out[bb].start : 1), 1, maxIdx);
      const getE = (bb) => clamp(Number((out[bb] && out[bb].end) != null ? out[bb].end : maxIdx), 1, maxIdx);

      // Ensure bucket 1 start fixed
      if (!out[1]) out[1] = { start: 1, end: Math.min(maxIdx, 4) };
      out[1] = { ...out[1], start: 1 };

      // Apply user edit: only allow editing END for any bucket except last,
      // and allow editing END for last only if you want, but we clamp it to maxIdx anyway.
      if (field === "end") {
        const s = (b === 1) ? 1 : (getE(b - 1) + 1);
        let e = clamp(v, 1, maxIdx);
        if (e < s) e = s;
        out[b] = { start: s, end: e };
      } else {
        // If they try to edit "start", we just re-enforce the rule.
        const s = (b === 1) ? 1 : (getE(b - 1) + 1);
        const e = getE(b);
        out[b] = { start: s, end: Math.max(s, e) };
      }

      // Now reflow all buckets to keep them contiguous
      for (let bb = 2; bb <= count; bb++) {
        const prevEnd = getE(bb - 1);
        const s = clamp(prevEnd + 1, 1, maxIdx);
        let e = getE(bb);

        // Last bucket always ends at maxIdx (rest)
        if (bb === count) e = maxIdx;

        if (e < s) e = s;
        out[bb] = { start: s, end: e };
      }

      next[letter] = out;
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

    // "Reset to ranges" overwrites everything and sets the new baseline snapshot.
    if (resetOverridesToRanges) {
      setLineToGroup(initMap);
      setDefaultMappingSnapshot(deepClone(initMap));
      return;
    }

    // "Apply ranges" updates lines that are currently empty OR still equal to the previous baseline.
    // This keeps any manual overrides you already made, but makes Step 3 + overrides reflect your new ranges.
    setLineToGroup((prev) => {
      const cur = prev || {};
      const base = defaultMappingSnapshot || {};
      const next = { ...cur };

      Object.keys(initMap).forEach((k) => {
        const curV = cur[k];
        const baseV = base[k];
        if (!curV || curV === "" || (baseV && curV === baseV)) {
          next[k] = initMap[k] || "";
        }
      });

      return next;
    });

    // Update the baseline snapshot to the new ranges layout.
    setDefaultMappingSnapshot(deepClone(initMap));
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

  const changedLineIdSet = useMemo(() => new Set((changes || []).map((c) => c.lineId)), [changes]);

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
          // snug: avoid fixed widths that can overlap on smaller screens
          width: "fit-content",
          minWidth: 140,
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

        {count === 4 && (
          <div style={{ marginTop: 8, padding: "8px 10px", borderRadius: 14, border: `1px solid ${theme.border}`, background: "rgba(250,204,21,0.14)", color: "rgba(255,255,255,0.92)", fontWeight: 900, fontSize: 12 }}>
            Remember to <b>Apply ranges</b> before moving to <b>Step 3</b>.
          </div>
        )}

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
                  <NumInput value={bucket === 1 ? 1 : (Number(r[bucket - 1]?.end ?? 0) + 1)} min={1} max={maxByLetter[L]} disabled={bucket !== 1} onChange={(vv) => setRange(L, bucket, "start", vv)} width={62} />
                </div>

                <div style={{ display: "flex", gap: 8, alignItems: "center", justifySelf: "start" }}>
                  <span style={{ opacity: 0.72, fontWeight: 850, fontSize: 12 }}>E</span>
                  <NumInput value={bucket === count ? (maxByLetter[L] || 1) : (r[bucket]?.end ?? (maxByLetter[L] || 1))} min={1} max={maxByLetter[L]} disabled={bucket === count} onChange={(vv) => setRange(L, bucket, "end", vv)} width={62} />
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
    // Fixed height + responsive width + obvious scrollbars
    return (
      <div
        ref={diagramBoxRef}
        className="diagramScrollBox"
        style={{
          border: `2px solid rgba(255,255,255,0.18)`,
          borderRadius: 18,
          overflow: "scroll",
          height: height || 720,
          width: width || "100%",
          maxWidth: "100%",
          minWidth: 0,
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
            changedLineIds={changedLineIdSet}
          />
        </div>
      </div>
      </div>
    );
  }

  const stepTabs = [
    { value: 1, label: "Step 1" },
    { value: 2, label: "Step 2" },
    { value: 3, label: "Step 3" },
    { value: 4, label: "Step 4 (Trim)" },
  ];

  return (
    <div style={{ minHeight: "100vh", overflowX: "hidden", background: theme.bg, color: theme.text, padding: 16, fontFamily: 'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, "Helvetica Neue", Arial' }}>
      {/* scrollbars styling */}
      <style>{`
        .diagramScrollBox { scrollbar-width: auto; scrollbar-color: rgba(255,255,255,0.55) rgba(0,0,0,0.35); }
        .diagramScrollBox::-webkit-scrollbar { height: 16px; width: 16px; }
        .diagramScrollBox::-webkit-scrollbar-track { background: rgba(0,0,0,0.35); }
        .diagramScrollBox::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.35); border: 3px solid rgba(0,0,0,0.35); border-radius: 999px; }
        .diagramScrollBox::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.55); }
      
        .scrollBox { overflow: scroll; scrollbar-width: auto; scrollbar-color: rgba(255,255,255,0.55) rgba(0,0,0,0.35); scrollbar-gutter: stable both-edges; }
        .scrollBox::-webkit-scrollbar { height: 16px; width: 16px; }
        .scrollBox::-webkit-scrollbar-track { background: rgba(0,0,0,0.35); }
        .scrollBox::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.35); border: 3px solid rgba(0,0,0,0.35); border-radius: 999px; }
        .scrollBox::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.55); }
`}</style>

      {/* reduced overall width to match overrides panel */}
      <div style={step !== 4 ? { width: "100%", maxWidth: 1100, margin: "0 auto", paddingLeft: 12, paddingRight: 12, display: "grid", gap: 10 } : { width: "100%", paddingLeft: 12, paddingRight: 12, display: "grid", gap: 10 }}>
        {/* Header */}
        <div style={{ border: `1px solid ${theme.border}`, borderRadius: 22, padding: 14, background: "linear-gradient(180deg, rgba(59,130,246,0.16), rgba(255,255,255,0.03))" }}>
          <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
            <div>
              <div style={{ fontSize: 36, fontWeight: 950, letterSpacing: -0.9 }}>Paraglider Trim Tuning</div>
              <div style={{ marginTop: 6, opacity: 0.86, fontSize: 14, fontWeight: 900 }}>
                {SITE_VERSION} <span style={{ opacity: 0.7, fontWeight: 850 }}>• Step 1–4</span>
              </div>
            </div>

            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
              <button style={{ ...topBtn, opacity: 0, padding: "2px 6px", fontSize: 10, minHeight: 0 }} onClick={() => setStep(2)}>← Back to Step 2</button>
              <ImportStatusRadio loaded={loaded} />
              {/* Step navigation buttons removed from the header (non-destructive). */}
              {false ? (
                stepTabs.map((t) => {
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
                })
              ) : null}
              <button style={{ ...topBtn, background: "rgba(239,68,68,0.16)" }} onClick={confirmResetAll}>
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

        {step !== 4 ? (
          <div style={step123Wrap}>
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
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-start", alignItems: "flex-start" }}>
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
                      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        <button style={{ ...topBtn, background: "rgba(59,130,246,0.22)" }} onClick={() => rebuildMappingFromRanges(false)} disabled={!loaded}>
                          Apply ranges
                        </button>
                        <button style={{ ...topBtn, background: "rgba(239,68,68,0.12)" }} onClick={() => rebuildMappingFromRanges(true)} disabled={!loaded}>
                          Reset ranges
                        </button>
                      </div>
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

                    <div style={{ marginTop: 10, maxHeight: "60vh", overflow: "auto", width: "100%", maxWidth: "100%",
          minWidth: 0, border: `1px solid ${theme.border}`, borderRadius: 14 }}>
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
            title="Step 3 — Mapping + baseline loops"
            right={
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                <div style={{ display: "flex", gap: 8, padding: "6px 8px", borderRadius: 14, border: `1px solid ${theme.border}`, background: theme.panel2 }}>
                  <button style={{ ...topBtn, ...(step3View === "diagram" ? { background: "rgba(59,130,246,0.25)" } : null) }} onClick={() => setStep3View("diagram")}>
                    Diagram
                  </button>
                  <button style={{ ...topBtn, ...(step3View === "baseline" ? { background: "rgba(59,130,246,0.25)" } : null) }} onClick={() => setStep3View("baseline")}>
                    Baseline loops
                  </button>
                </div>

                <button style={topBtn} onClick={() => setStep(2)}>
                  ← Back to Step 2
                </button>
                <div style={{ padding: "6px 10px", borderRadius: 14, border: `1px solid ${theme.border}`, background: "rgba(245,158,11,0.14)", color: theme.text, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                  <div style={{ fontWeight: 950, opacity: 0.95 }}>Baseline loops required</div>
                  <div style={{ fontSize: 12, opacity: 0.82, lineHeight: 1.15 }}>
                    Set all installed baseline loops first. Step 4 freezes baseline and you can’t return to Step 3 unless you reset.
                  </div>
                  <button
                    type="button"
                    style={{ ...topBtn, padding: "6px 10px", background: "rgba(255,255,255,0.08)" }}
                    onClick={() => {
                      setStep3View("baseline");
                      try {
                        setTimeout(() => {
                          const el = document.getElementById("baseline-loops-panel");
                          if (el && el.scrollIntoView) el.scrollIntoView({ behavior: "smooth", block: "start" });
                        }, 50);
                      } catch (e) {}
                    }}
                    title="Jump to baseline loop selection"
                  >
                    Review installed loops
                  </button>
                </div>

                <button style={{ ...topBtn, background: "rgba(34,197,94,0.18)" }} onClick={() => setStep(4)}>
                  Go to Step 4 →
                </button>
                <button style={topBtn} onClick={fitDiagramToScreen}>
                  Fit to screen
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

                <WarningBanner title="Set baseline loops before Step 4">
                  Step 4 freezes a snapshot of your installed baseline loops. After you enter Step 4, you can’t return to Step 3 without using the full Reset.
                  <div style={{ marginTop: 10, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                    <button
                      style={{ ...topBtn, background: "rgba(245,158,11,0.20)", borderColor: "rgba(245,158,11,0.35)" }}
                      onClick={() => {
                        setStep3View("baseline");
                        setTimeout(() => {
                          var el = document.getElementById("step3-baseline-loops");
                          if (el && el.scrollIntoView) {
                            el.scrollIntoView({ behavior: "smooth", block: "start" });
                          }
                        }, 0);
                      }}
                      title="Jump to Installed loops per maillon group (baseline)"
                    >
                      Go to baseline loops
                    </button>
                    <div style={{ opacity: 0.78, fontSize: 12 }}>
                      Make sure every group has an installed loop selected (Left and Right) before freezing Step 4.
                    </div>
                  </div>
                </WarningBanner>
                {step3View === "baseline" ? (
                  <div id="baseline-loops-panel" style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                      <div>
                        <div style={{ fontWeight: 950 }}>Installed loops per maillon group (baseline)</div>
                        <div style={{ opacity: 0.78, fontSize: 13, marginTop: 4 }}>Set loop sizes (mm), then select what is currently installed on the wing. Step 4 will freeze this baseline.</div>
                      </div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "360px 1fr", gap: 10, marginTop: 10 }}>
                      <div style={{ border: `1px solid ${theme.border}`, borderRadius: 14, background: theme.panel, padding: 10 }}>
                        <div style={{ fontWeight: 900, marginBottom: 8 }}>Loop sizes (mm)</div>
                        <div style={{ display: "grid", gap: 8 }}>
                          {LOOP_TYPES.map((lt) => (
                            <div key={lt} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
                              <div style={{ fontWeight: 900 }}>{lt}</div>
                              <input type="number" value={Number(loopSizes?.[lt] ?? 0)} onChange={(e) => setLoopSizes((prev) => ({ ...(prev || {}), [lt]: Number(e.target.value || 0) }))} style={{ width: 90, padding: "8px 10px", borderRadius: 10, border: `1px solid ${theme.border}`, background: "rgba(0,0,0,0.45)", color: theme.text, fontWeight: 900 }} />
                            </div>
                          ))}
                        </div>
                      </div>
                      <div style={{ border: `1px solid ${theme.border}`, borderRadius: 14, background: theme.panel, padding: 10, minWidth: 0 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                          <div style={{ fontWeight: 900 }}>Baseline installed loop by group</div>
                          <div style={{ opacity: 0.75, fontSize: 12 }}>Groups shown: <b>{groupsInUse.length}</b></div>
                        </div>
                        {groupsInUse.length === 0 ? (
                          <div style={{ marginTop: 10, opacity: 0.75 }}>No groups detected yet. Complete Step 2 mapping first (Apply ranges or drag chips).</div>
                        ) : (
                          <div style={{ marginTop: 10, maxHeight: 520, overflow: "auto", width: "100%", maxWidth: "100%",
          minWidth: 0, border: `1px solid ${theme.border}`, borderRadius: 12 }}>
                            <table style={{ width: "100%", borderCollapse: "collapse" }}>
                              <thead>
                                <tr>
                                  <th style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 1 }} rowSpan={2}>Group</th>
                                  <th style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 1 }} colSpan={2}>Left</th>
                                  <th style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 1 }} colSpan={2}>Right</th>
                                </tr>
                                <tr>
                                  <th style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 1 }}>Installed loop</th>
                                  <th style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 1 }}>mm</th>
                                  <th style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 1 }}>Installed loop</th>
                                  <th style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 1 }}>mm</th>
                                </tr>
                              </thead>
                              <tbody>
                                {(() => {
                                  const parseGroup = (g) => {
                                    const m = /^([A-Za-z]+\d+)([LR])$/.exec(String(g || ""));
                                    if (!m) return { base: String(g || ""), side: "" };
                                    return { base: m[1], side: m[2] };
                                  };

                                  const byBase = {};
                                  const order = [];

                                  for (let i = 0; i < groupsInUse.length; i++) {
                                    const g = groupsInUse[i];
                                    const info = parseGroup(g);
                                    const base = info.base;
                                    const side = info.side;

                                    if (!byBase[base]) {
                                      byBase[base] = { base: base, L: null, R: null };
                                      order.push(base);
                                    }

                                    if (side === "L") byBase[base].L = g;
                                    else if (side === "R") byBase[base].R = g;
                                    else {
                                      // Unpaired group id (no L/R suffix) -> treat as Left-only display
                                      if (!byBase[base].L) byBase[base].L = g;
                                    }
                                  }

                                  order.sort((a, b) => {
                                    // Sort like AR1, AR2, ... if possible
                                    const ma = /^([A-Za-z]+)(\d+)$/.exec(a);
                                    const mb = /^([A-Za-z]+)(\d+)$/.exec(b);
                                    if (ma && mb) {
                                      const pa = ma[1].toUpperCase();
                                      const pb = mb[1].toUpperCase();
                                      if (pa < pb) return -1;
                                      if (pa > pb) return 1;
                                      const na = Number(ma[2] || 0);
                                      const nb = Number(mb[2] || 0);
                                      return na - nb;
                                    }
                                    return String(a).localeCompare(String(b));
                                  });

                                  return order.map((base) => {
                                    const row = byBase[base];

                                    const gL = row && row.L ? row.L : null;
                                    const gR = row && row.R ? row.R : null;

                                    const ltL = gL ? (groupLoopSetup?.[gL] || "SL") : "";
                                    const ltR = gR ? (groupLoopSetup?.[gR] || "SL") : "";

                                    const mmL = gL ? Number(loopSizes?.[ltL] ?? 0) : null;
                                    const mmR = gR ? Number(loopSizes?.[ltR] ?? 0) : null;

                                    return (
                                      <tr key={base} style={{ borderTop: `1px solid ${theme.border}` }}>
                                        <td style={td}>
                                          <div style={{ fontWeight: 950 }}>{base}</div>
                                          <div style={{ opacity: 0.72, fontSize: 12, marginTop: 2 }}>
                                            {gL ? gL : ""}{gL && gR ? " • " : ""}{gR ? gR : ""}
                                          </div>
                                        </td>

                                        <td style={td}>
                                          {gL ? (
                                            <Select
                                              value={ltL}
                                              onChange={(v) => setGroupLoopSetup((prev) => ({ ...(prev || {}), [gL]: v || "SL" }))}
                                              options={loopTypeOptions}
                                              width={150}
                                            />
                                          ) : (
                                            <div style={{ opacity: 0.45 }}>—</div>
                                          )}
                                        </td>
                                        <td style={{ ...td, fontWeight: 950, opacity: 0.9 }}>{gL ? mmL : "—"}</td>

                                        <td style={td}>
                                          {gR ? (
                                            <Select
                                              value={ltR}
                                              onChange={(v) => setGroupLoopSetup((prev) => ({ ...(prev || {}), [gR]: v || "SL" }))}
                                              options={loopTypeOptions}
                                              width={150}
                                            />
                                          ) : (
                                            <div style={{ opacity: 0.45 }}>—</div>
                                          )}
                                        </td>
                                        <td style={{ ...td, fontWeight: 950, opacity: 0.9 }}>{gR ? mmR : "—"}</td>
                                      </tr>
                                    );
                                  });
                                })()}
                              </tbody>
                            </table>
                          </div>
                        )}
                        <div style={{ marginTop: 10, opacity: 0.75, fontSize: 12 }}>Tip: Use <b>CUSTOM</b> for wings with non-standard loop lengths.</div>
                      </div>
                    </div>
                  </div>
				) : (
				  <>
				  {/* Step 3 includes Apply/Reset buttons (same as Step 2) */}
                <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                    <div>
                      <div style={{ fontWeight: 950 }}>Map lines to maillon groups (setup)</div>
                      <div style={{ opacity: 0.78, fontSize: 13, marginTop: 4 }}>Apply/Reset ranges here without leaving the diagram.</div>
                    </div>

                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
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

                    <div style={{ marginTop: 10, maxHeight: 320, overflow: "auto", width: "100%", maxWidth: "100%",
          minWidth: 0, border: `1px solid ${theme.border}`, borderRadius: 14 }}>
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
				  </>
				)}
              </div>
            )}
          </Panel>
        ) : null}
          </div>
        ) : null}

        {/* Step 4 */}
        {step === 4 ? (
          <div style={{ display: "flex", justifyContent: "center", width: "100%" }}>
            <div style={{ width: "100%", maxWidth: 1600 }}>
              <Panel
            tint
            title="Step 4 — Trim (frozen baseline)"
            right={
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                <button
                  style={{ ...topBtn, background: "rgba(239,68,68,0.12)" }}
                  onClick={() => {
                    setGroupLoopChange({});
                    setGroupAdjustments({});
    setStep4LineCorr({});
                  }}
                >
                  Reset Step 4 adjustments
                </button>
              </div>
            }
          >
            {!loaded ? (
              <WarningBanner title="No file loaded">Go back to Step 1 and import a file (or load test data).</WarningBanner>
            ) : groupLoopBaseline === null ? (
              <WarningBanner title="Baseline not frozen yet">
                Step 4 must freeze a snapshot of Step 3 <b>exactly once</b>. Click “Freeze baseline now” to continue.
                <div style={{ marginTop: 10 }}>
                  <button
                    style={{ ...topBtn, background: "rgba(34,197,94,0.18)" }}
                    onClick={() => setGroupLoopBaseline(deepClone(groupLoopSetup || {}))}
                  >
                    Freeze baseline now
                  </button>
                </div>
              </WarningBanner>
            ) : (
              <div style={{ display: "grid", gap: 10 }}>
                <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 10 }}>
                    <div>
                      <div style={{ fontWeight: 950 }}>Meta controls</div>
                      <div style={{ opacity: 0.78, fontSize: 13, marginTop: 4 }}>
                        Adjust tolerance and correction used for all Step 4 tables. (Does not change Step 3 baseline.)
                      </div>
                    </div>
                    <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap", justifyContent: "flex-start" }}>
                      <ControlPill
                        label="Tolerance"
                        value={meta.tolerance ?? 6}
                        onChange={(v) => setMeta((p) => ({ ...p, tolerance: v }))}
                        suffix="mm"
                        width={90}
                        min={0}
                      />
                      <ControlPill
                        label="Correction"
                        value={meta.correction ?? 0}
                        onChange={(v) => setMeta((p) => ({ ...p, correction: v }))}
                        suffix="mm"
                        width={110}
                      />
                      <TogglePill label="Show corrected" checked={!!showCorrected} onChange={setShowCorrected} />
                      <TogglePill label="Brake" checked={!!includeBrakeBlock} onChange={setIncludeBrakeBlock} />
                    </div>
                  </div>

                  <div style={{ marginTop: 10, paddingTop: 10, borderTop: `1px solid ${theme.border}` }}>
                    <div style={{ fontWeight: 850, marginBottom: 6 }}>Zeroing wizard (auto-suggest correction)</div>
                    <div style={{ fontSize: 13, opacity: 0.82 }}>
                      Suggested correction uses the <b>median</b> of <b>(Soll − Ist)</b> across all valid measurements.
                    </div>

                    <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginTop: 10, alignItems: "center" }}>
                      <StatPill label="Whole wing median" value={zeroingStats.wholeMedian} n={zeroingStats.nAll} />
                      <StatPill label="Left median" value={zeroingStats.leftMedian} n={zeroingStats.nLeft} />
                      <StatPill label="Right median" value={zeroingStats.rightMedian} n={zeroingStats.nRight} />

                      <button
                        style={{ ...topBtn, background: "rgba(99,102,241,0.20)" }}
                        disabled={!Number.isFinite(zeroingStats.wholeMedian)}
                        onClick={() => {
                          const v = zeroingStats.wholeMedian;
                          if (!Number.isFinite(v)) return;
                          setMeta((p) => ({ ...p, correction: Math.round(v) }));
                        }}
                        title="Apply whole wing median as correction"
                      >
                        Apply suggested correction
                      </button>
                    </div>
                  </div>
                </div>
                <div style={{ marginTop: 10, display: "flex", gap: 12, alignItems: "flex-start", flexWrap: "wrap" }}>
                  <div style={{ flex: "1 1 720px", minWidth: 0 }}>
                    <RearViewChart
                      rows={step4LineRows}
                      tolerance={Number(meta && meta.tolerance != null ? meta.tolerance : 0)}
                      height={520}
                      loopTypes={LOOP_TYPES}
                      groupLoopChange={groupLoopChange}
                      setGroupLoopChange={setGroupLoopChange}
                    />
                  </div>

                  <div style={{ flex: "0 0 auto" }}>
                    <div
                      style={{
                        border: `1px solid ${theme.border}`,
                        background: theme.bg2,
                        borderRadius: 999,
                        padding: "8px 10px",
                        display: "flex",
                        gap: 8,
                        alignItems: "center",
                      }}
                      title="Correction (mm)"
                    >
                      <span style={{ fontSize: 12, opacity: 0.8 }}>Correction</span>

                      <div style={{ display: "grid", gridTemplateRows: "1fr 1fr", gap: 4 }}>
                        <button
                          type="button"
                          style={{ width: 28, height: 18, borderRadius: 8, border: `1px solid ${theme.border}`, background: "rgba(255,255,255,0.08)", color: theme.text, cursor: "pointer", fontWeight: 950, lineHeight: 1 }}
                          onClick={() => {
                            setMeta((p) => {
                              var vRaw = (p && p.correction != null) ? p.correction : 0;
                              var v = Number(vRaw);
                              if (!isFinite(v)) v = 0;

                              // Nudge by exactly +1mm from the current displayed value.
                              // Do NOT clamp here; users may temporarily be far outside tolerance.
                              var next = v + 1;
                              next = Math.round(next);
                              return Object.assign({}, p, { correction: next });
                            });
                          }}
                        >
                          ▲
                        </button>
                        <button
                          type="button"
                          style={{ width: 28, height: 18, borderRadius: 8, border: `1px solid ${theme.border}`, background: "rgba(255,255,255,0.08)", color: theme.text, cursor: "pointer", fontWeight: 950, lineHeight: 1 }}
                          onClick={() => {
                            setMeta((p) => {
                              var vRaw = (p && p.correction != null) ? p.correction : 0;
                              var v = Number(vRaw);
                              if (!isFinite(v)) v = 0;

                              // Nudge by exactly -1mm from the current displayed value.
                              // Do NOT clamp here; users may temporarily be far outside tolerance.
                              var next = v - 1;
                              next = Math.round(next);
                              return Object.assign({}, p, { correction: next });
                            });
                          }}
                        >
                          ▼
                        </button>
                      </div>

                      <input
                        type="number"
                        value={Number(meta && meta.correction != null ? meta.correction : 0)}
                        onChange={(e) => {
                          var v = Number(e.target.value || 0);
                          if (!isFinite(v)) v = 0;
                          // Do NOT clamp here; allow any value and let other parts clamp if needed.
                          setMeta((p) => ({ ...p, correction: v }));
                        }}
                        style={{
                          width: 90,
                          background: "transparent",
                          color: theme.text,
                          border: `1px solid ${theme.border}`,
                          borderRadius: 12,
                          padding: "6px 8px",
                          fontWeight: 900,
                          outline: "none",
                        }}
                      />
                      <span style={{ fontSize: 12, opacity: 0.75 }}>mm</span>
                    </div>
                  
	                    <button
	                      type="button"
	                      title="Reset all loop overrides"
	                      onClick={() => {
	                        // Reset BOTH loop overrides and fine adjustments so the wing returns
	                        // to the baseline loop set recorded in Step 3.
	                        setGroupLoopChange({});
	                        setGroupAdjustments({});
	                      }}
                      style={{
                        marginTop: 10,
                        width: "100%",
                        border: "1px solid rgba(255,220,80,0.55)",
                        background: "rgba(0,0,0,0.45)",
                        color: "rgba(255,220,80,0.95)",
                        borderRadius: 999,
                        padding: "8px 10px",
                        fontWeight: 950,
                        fontSize: 12,
                        cursor: "pointer",
                        whiteSpace: "nowrap",
                      }}
                    >
                      Reset all loops
                    </button>
</div>
                </div>

                <div
                  style={{
                    border: `1px solid ${theme.border}`,
                    borderRadius: 16,
                    background: theme.panel2,
                    padding: 10,
                    minWidth: 0,
                    overflow: "auto",
                  }}
                >
                  <div style={{ fontWeight: 950 }}>A/B/C/D — Factory vs Left/Right (Step 4 data)</div>
                  <div style={{ opacity: 0.78, fontSize: 13, marginTop: 4 }}>
                    Mirrors the spreadsheet layout: <b>Factory</b>, <b>Left</b>, <b>Right</b>, <b>Δ Left</b>, <b>Δ Right</b>, <b>Sym</b>. Values come from Step 4 “after” and “delta”.
                  </div>

                  {(() => {
                    const letters = includeBrakeBlock ? ["A", "B", "C", "D", "BR"] : ["A", "B", "C", "D"];
                    const byLetter = step4SheetByLetter?.byLetter || {};
                    const maxDiff = step4SheetByLetter?.maxDiff || {};
                    const letterMaps = {};
                    let maxIdx = 0;

                    for (const L of letters) {
                      const rows = Array.isArray(byLetter[L]) ? byLetter[L] : [];
                      const m = new Map();
                      for (const r of rows) {
                        m.set(r.idx, r);
                        if (Number.isFinite(r.idx)) maxIdx = Math.max(maxIdx, r.idx);
                      }
                      letterMaps[L] = m;
                    }

                    const headerCell = { padding: "6px 8px", borderBottom: `1px solid ${theme.border}`, whiteSpace: "nowrap" };
                    const cell = { padding: "6px 8px", borderBottom: `1px solid ${theme.border}`, textAlign: "center", whiteSpace: "nowrap" };

                    const fmt = (v, digits = 0) => (v == null || !Number.isFinite(Number(v)) ? "—" : Number(v).toFixed(digits));

                    const avgOf = (arr) => {
                      const xs = (arr || []).map(Number).filter((n) => Number.isFinite(n));
                      if (xs.length === 0) return null;
                      const sum = xs.reduce((a, b) => a + b, 0);
                      return sum / xs.length;
                    };

                    const bandFromDelta = (delta) => {
                      const d = Number(delta);
                      if (!Number.isFinite(d)) return "";
                      const a = Math.abs(d);
                      const tol = Number(meta?.tolerance ?? 10);
                      if (a <= 4) return "good";
                      if (a < tol) return "warn";
                      return "bad";
                    };
                    const bgForBand = (band) => {
                      if (band === "good") return "rgba(34,197,94,0.14)";
                      if (band === "warn") return "rgba(234,179,8,0.14)";
                      if (band === "bad") return "rgba(239,68,68,0.14)";
                      return "transparent";
                    };
                    const colorForBand = (band) => {
                      if (band === "good") return "rgba(34,197,94,0.95)";
                      if (band === "warn") return "rgba(234,179,8,0.95)";
                      if (band === "bad") return "rgba(239,68,68,0.95)";
                      return theme.text;
                    };

                    return (
                      <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 10, fontSize: 13 }}>
                        <thead>
                          <tr>
                            <th style={{ ...headerCell, textAlign: "left", position: "sticky", left: 0, background: theme.panel2, zIndex: 2 }}>#</th>
                            {letters.map((L) => (
                              <th key={L} colSpan={6} style={{ ...headerCell, textAlign: "center", background: "rgba(255,255,255,0.03)" }}>
                                {L}
                              </th>
                            ))}
                          </tr>
                          <tr>
                            <th style={{ ...headerCell, textAlign: "left", position: "sticky", left: 0, background: theme.panel2, zIndex: 2 }} />
                            {letters.map((L) => (
                              <React.Fragment key={`${L}-sub`}>
                                <th style={headerCell}>Factory</th>
                                <th style={headerCell}>Left</th>
                                <th style={headerCell}>Right</th>
                                <th style={headerCell}>Δ Left</th>
                                <th style={headerCell}>Δ Right</th>
                                <th style={headerCell}>Sym</th>
                              </React.Fragment>
                            ))}
                          </tr>
                        </thead>

                        <tbody>
                          {Array.from({ length: Math.max(0, maxIdx) }, (_, i) => i + 1).map((idx) => (
                            <tr key={`row-${idx}`}>
                              <td style={{ ...cell, textAlign: "left", position: "sticky", left: 0, background: theme.panel2, zIndex: 1, fontWeight: 700 }}>{idx}</td>
                              {letters.map((L) => {
                                const r = letterMaps[L].get(idx);
                                return (
                                  <React.Fragment key={`${L}-${idx}`}>
                                    <td style={cell}>{fmt(r?.factory, 0)}</td>
                                    <td style={cell}>{fmt(r?.L, 0)}</td>
                                    <td style={cell}>{fmt(r?.R, 0)}</td>
                                    <td style={{ ...cell, background: bgForBand(bandFromDelta(r?.dL)), color: colorForBand(bandFromDelta(r?.dL)), fontWeight: 950 }}>{fmt(r?.dL, 0)}</td>
                                    <td style={{ ...cell, background: bgForBand(bandFromDelta(r?.dR)), color: colorForBand(bandFromDelta(r?.dR)), fontWeight: 950 }}>{fmt(r?.dR, 0)}</td>
                                    <td style={{ ...cell, background: bgForBand(bandFromDelta(r?.sym)), color: colorForBand(bandFromDelta(r?.sym)), fontWeight: 950 }}>{fmt(r?.sym, 0)}</td>
                                  </React.Fragment>
                                );
                              })}
                            </tr>
                          ))}

                          <tr>
                            <td style={{ ...cell, textAlign: "left", position: "sticky", left: 0, background: theme.panel2, zIndex: 1, fontWeight: 900 }}>
                              Averages
                            </td>
                            {letters.map((L) => {
                              const rows = Array.isArray(byLetter[L]) ? byLetter[L] : [];
                              const aDL = avgOf(rows.map((r) => r?.dL));
                              const aDR = avgOf(rows.map((r) => r?.dR));
                              const aSY = avgOf(rows.map((r) => r?.sym));
                              return (
                                <React.Fragment key={`avg-${L}`}>
                                  <td style={cell} />
                                  <td style={cell} />
                                  <td style={cell} />
                                  <td style={{ ...cell, background: bgForBand(bandFromDelta(aDL)), color: colorForBand(bandFromDelta(aDL)), fontWeight: 950 }}>{fmt(aDL, 1)}</td>
                                  <td style={{ ...cell, background: bgForBand(bandFromDelta(aDR)), color: colorForBand(bandFromDelta(aDR)), fontWeight: 950 }}>{fmt(aDR, 1)}</td>
                                  <td style={{ ...cell, background: bgForBand(bandFromDelta(aSY)), color: colorForBand(bandFromDelta(aSY)), fontWeight: 950 }}>{fmt(aSY, 1)}</td>
                                </React.Fragment>
                              );
                            })}
                          </tr>

                          <tr>
                            <td style={{ ...cell, textAlign: "left", position: "sticky", left: 0, background: theme.panel2, zIndex: 1, fontWeight: 900 }}>
                              <span title="Maximum Difference (peak-to-peak) for Δ Left, Δ Right, and Sym">Diffs</span>
                            </td>
                            {letters.map((L) => (
                              <React.Fragment key={`max-${L}`}>
                                <td style={cell} />
                                <td style={cell} />
                                <td style={cell} />
                                <td style={cell}>{fmt(maxDiff?.[L]?.dL, 0)}</td>
                                <td style={cell}>{fmt(maxDiff?.[L]?.dR, 0)}</td>
                                <td style={cell}>{fmt(maxDiff?.[L]?.sym, 0)}</td>
                              </React.Fragment>
                            ))}
                          </tr>
                        </tbody>
                      </table>
                    );
                  })()}
                </div>




                <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
                  <div style={{ fontWeight: 950 }}>Trim adjustments per maillon group (mm)</div>
                  <div style={{ opacity: 0.78, fontSize: 13, marginTop: 4 }}>
                    Uses <b>frozen baseline</b> from Step 3. Step 3 edits will not affect this page until you Reset all.
                  </div>

                  {groupsInUse.length === 0 ? (
                    <div style={{ marginTop: 10, opacity: 0.75 }}>No groups detected. Complete Step 2 mapping first.</div>
                  ) : (() => { try { return (
                    <div style={{ marginTop: 10, overflow: "auto", width: "100%", maxWidth: "100%",
          minWidth: 0, border: `1px solid ${theme.border}`, borderRadius: 12 }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1100 }}>
                        <thead>
                          <tr>
                            <th rowSpan={2} style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 2 }}>Group</th>
                            <th colSpan={5} style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 2, textAlign: "center" }}>Left</th>
                            <th colSpan={5} style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 2, textAlign: "center" }}>Right</th>
                          </tr>
                          <tr>
                            {["Baseline loop", "Override loop", "Loop Δ (mm)", "Adjust (mm)", "Total Δ (mm)"].map((h) => (
                              <th key={"L-"+h} style={{ ...th, position: "sticky", top: 34, background: theme.panel2, zIndex: 2 }}>{h}</th>
                            ))}
                            {["Baseline loop", "Override loop", "Loop Δ (mm)", "Adjust (mm)", "Total Δ (mm)"].map((h) => (
                              <th key={"R-"+h} style={{ ...th, position: "sticky", top: 34, background: theme.panel2, zIndex: 2 }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {(() => {
                            const parse = (s) => {
                              const m = String(s || "").match(/^([A-Z]+)(\d+)([LR])$/i);
                              if (!m) return null;
                              return { p: m[1].toUpperCase(), n: Number(m[2] || 0), side: m[3].toUpperCase() };
                            };

                            const byKey = new Map();
                            (groupsInUse || []).forEach((g) => {
                              const p = parse(g);
                              if (!p) {
                                byKey.set(g, { key: g, L: null, R: null, single: g, labelL: g, labelR: "" });
                                return;
                              }
                              const key = `${p.p}${p.n}`;
                              const cur = byKey.get(key) || { key, L: null, R: null, single: null, labelL: `${key}L`, labelR: `${key}R` };
                              if (p.side === "L") cur.L = g;
                              if (p.side === "R") cur.R = g;
                              byKey.set(key, cur);
                            });

                            const keys = Array.from(byKey.values());
                            keys.sort((a, b) => {
                              const ap = parse(a.labelL) || { p: a.key, n: 0, side: "" };
                              const bp = parse(b.labelL) || { p: b.key, n: 0, side: "" };
                              if (ap.p !== bp.p) return ap.p.localeCompare(bp.p);
                              if (ap.n !== bp.n) return ap.n - bp.n;
                              return 0;
                            });

                            const cellFor = (gid) => {
                              if (!gid) {
                                return {
                                  baseLoop: "",
                                  override: "",
                                  loopDelta: "",
                                  adj: "",
                                  total: "",
                                  totalColor: "transparent",
                                };
                              }
                              const baseLoop = groupLoopBaseline?.[gid] || "SL";
                              const override = groupLoopChange?.[gid] || "";
                              const afterLoop = override || baseLoop;

                              const baseMm = Number(loopSizes?.[baseLoop] ?? 0);
                              const afterMm = Number(loopSizes?.[afterLoop] ?? 0);
                              const loopDelta = afterMm - baseMm;

                              const adj = Number(groupAdjustments?.[gid] ?? 0);
                              const total = loopDelta + adj;

                              const tol = Number((meta && meta.tolerance != null) ? meta.tolerance : 0);
                              const totalColor =
                                Math.abs(total) >= tol ? theme.bad : Math.abs(total) >= Math.max(0, tol - 3) ? theme.warn : theme.good;

                              return { baseLoop, override, loopDelta, adj, total, totalColor };
                            };

                            return keys.map((row) => {
                              const L = cellFor(row.L);
                              const R = cellFor(row.R);

                              return (
                                <tr key={row.key} style={{ borderTop: `1px solid ${theme.border}` }}>
                                  <td style={{ ...td, whiteSpace: "nowrap" }}>
                                    <div style={{ fontWeight: 900 }}>{row.L || row.single || row.labelL}</div>
                                    {row.R ? <div style={{ opacity: 0.85, marginTop: 2 }}>{row.R}</div> : null}
                                  </td>

                                  {/* Left */}
                                  <td style={td}>{L.baseLoop}</td>
                                  <td style={td}>
                                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                                      <select
                                        style={{ ...miniInput, width: 86, padding: "4px 8px", background: theme.panel2, color: theme.text }}
                                        disabled={!row.L}
                                        value={(row.L && groupLoopChange && groupLoopChange[row.L]) ? groupLoopChange[row.L] : ""}
                                        onChange={(e) => {
                                          if (!row.L) return;
                                          const v = e.target.value;
                                          if (!v) {
                                            setGroupLoopChange((p) => {
                                              const n = { ...(p || {}) };
                                              delete n[row.L];
                                              return n;
                                            });
                                          } else {
                                            setGroupLoopChange((p) => ({ ...(p || {}), [row.L]: v }));
                                          }
                                        }}
                                      >
                                        <option value="" style={{ background: theme.panel2, color: theme.text }}>(baseline)</option>
                                        {LOOP_TYPES.map((lt) => (
                                          <option key={`L-opt-${lt}`} value={lt} style={{ background: theme.panel2, color: theme.text }}>{lt}</option>
                                        ))}
                                      </select>
                                      <button
                                        key={"L-"+row.key+"-reset"}
                                        title="Reset to baseline"
                                        style={{
                                          padding: "4px 8px",
                                          borderRadius: 999,
                                          border: `1px solid ${theme.border}`,
                                          background: "rgba(0,0,0,0.18)",
                                          color: theme.text,
                                          fontWeight: 900,
                                          cursor: row.L ? "pointer" : "not-allowed",
                                          opacity: row.L ? 0.9 : 0.5,
                                        }}
                                        disabled={!row.L}
                                        onClick={() => {
                                          if (!row.L) return;
                                          setGroupLoopChange((p) => {
                                            const n = { ...(p || {}) };
                                            delete n[row.L];
                                            return n;
                                          });
                                        }}
                                      >
                                        ↺
                                      </button>
                                    </div>
                                  </td>
                                  <td style={td}>{Number.isFinite(L.loopDelta) ? Math.round(L.loopDelta) : ""}</td>
                                  <td style={td}>
                                    <input
                                      style={miniInput}
                                      value={Number.isFinite(L.adj) ? String(L.adj) : ""}
                                      onChange={(e) => {
                                        const v = Number(e.target.value);
                                        if (!row.L) return;
                                        setGroupAdjustments((p) => ({ ...p, [row.L]: Number.isFinite(v) ? v : 0 }));
                                      }}
                                    />
                                  </td>
                                  <td style={{ ...td, textAlign: "center" }}><div style={{ display: "inline-block", minWidth: 46, padding: "4px 10px", borderRadius: 999, border: `1px solid ${theme.border}`, background: L.totalColor, fontWeight: 950, lineHeight: 1 }}>{Number.isFinite(L.total) ? Math.round(L.total) : ""}</div></td>

                                  {/* Right */}
                                  <td style={td}>{R.baseLoop}</td>
                                  <td style={td}>
                                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                                      <select
                                        style={{ ...miniInput, width: 86, padding: "4px 8px", background: theme.panel2, color: theme.text }}
                                        disabled={!row.R}
                                        value={(row.R && groupLoopChange && groupLoopChange[row.R]) ? groupLoopChange[row.R] : ""}
                                        onChange={(e) => {
                                          if (!row.R) return;
                                          const v = e.target.value;
                                          if (!v) {
                                            setGroupLoopChange((p) => {
                                              const n = { ...(p || {}) };
                                              delete n[row.R];
                                              return n;
                                            });
                                          } else {
                                            setGroupLoopChange((p) => ({ ...(p || {}), [row.R]: v }));
                                          }
                                        }}
                                      >
                                        <option value="" style={{ background: theme.panel2, color: theme.text }}>(baseline)</option>
                                        {LOOP_TYPES.map((lt) => (
                                          <option key={`R-opt-${lt}`} value={lt} style={{ background: theme.panel2, color: theme.text }}>{lt}</option>
                                        ))}
                                      </select>
                                      <button
                                        key={"R-"+row.key+"-reset"}
                                        title="Reset to baseline"
                                        style={{
                                          padding: "4px 8px",
                                          borderRadius: 999,
                                          border: `1px solid ${theme.border}`,
                                          background: "rgba(0,0,0,0.18)",
                                          color: theme.text,
                                          fontWeight: 900,
                                          cursor: row.R ? "pointer" : "not-allowed",
                                          opacity: row.R ? 0.9 : 0.5,
                                        }}
                                        disabled={!row.R}
                                        onClick={() => {
                                          if (!row.R) return;
                                          setGroupLoopChange((p) => {
                                            const n = { ...(p || {}) };
                                            delete n[row.R];
                                            return n;
                                          });
                                        }}
                                      >
                                        ↺
                                      </button>
                                    </div>
                                  </td>
                                  <td style={td}>{Number.isFinite(R.loopDelta) ? Math.round(R.loopDelta) : ""}</td>
                                  <td style={td}>
                                    <input
                                      style={miniInput}
                                      value={Number.isFinite(R.adj) ? String(R.adj) : ""}
                                      onChange={(e) => {
                                        const v = Number(e.target.value);
                                        if (!row.R) return;
                                        setGroupAdjustments((p) => ({ ...p, [row.R]: Number.isFinite(v) ? v : 0 }));
                                      }}
                                    />
                                  </td>
                                  <td style={{ ...td, textAlign: "center" }}><div style={{ display: "inline-block", minWidth: 46, padding: "4px 10px", borderRadius: 999, border: `1px solid ${theme.border}`, background: R.totalColor, fontWeight: 950, lineHeight: 1 }}>{Number.isFinite(R.total) ? Math.round(R.total) : ""}</div></td>
                                </tr>
                              );
                            });
                          })()}
                        </tbody>
                      </table>
                    </div>); } catch (e) { return (
                    <div style={{ marginTop: 10, color: theme.bad, fontWeight: 900 }}>
                      Group table render error: {String((e && e.message) ? e.message : e)}
                    </div>
                  ); } })()}
                </div>




{/* Group averages / maillon loop advisory (A/B/C/D) */}
<div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
    <div>
      <div style={{ fontWeight: 950 }}>Group averages + loop suggestions (A/B/C/D)</div>
      <div style={{ opacity: 0.78, fontSize: 13, marginTop: 4 }}>
        Uses Step 4 <b>After</b> deltas (Δ vs nominal). Suggestions are advisory only — they won’t change any group settings automatically.
      </div>
    </div>

    <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
      <button
        style={{ ...topBtn, background: "rgba(255,255,255,0.06)" }}
        onClick={async () => {
          try {
            const payload = { schema: "abc-loop-suggestions-v1", exportedAt: new Date().toISOString(), wing: { make: meta.make || "", model: meta.model || "" }, suggestions: abcSuggestions, averages: abcAverages };
            await navigator.clipboard.writeText(JSON.stringify(payload, null, 2));
            alert("Copied suggestions JSON to clipboard.");
          } catch (e) {
            alert("Clipboard copy failed (browser permission).");
          }
        }}
      >
        Copy suggestions
      </button>

      <ControlPill label="Manual pitch tol" value={groupPitchTol} onChange={setGroupPitchTol} suffix="mm" width={110} step={1} min={0} max={20} />

      <button
        style={{ ...topBtn, background: showLoopModeCounts ? "rgba(99,102,241,0.25)" : "rgba(255,255,255,0.06)" }}
        onClick={() => setShowLoopModeCounts((v) => !v)}
      >
        {showLoopModeCounts ? "Hide loop counts" : "Show loop counts"}
      </button>
    </div>
  </div>

  <div style={{ marginTop: 10, display: "grid", gap: 10 }}>
    {/* Averages */}
    



<div className="abcGrid" style={{ display: "grid", gridTemplateColumns: "1fr", gap: 10 }}>
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 14, background: theme.panel, padding: 10, minWidth: 0, width: "100%" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
        <div style={{ fontWeight: 950 }}>Suggested loop change + fine adjust (advisory)</div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <button
            type="button"
            style={{ padding: "6px 10px", borderRadius: 999, border: `1px solid ${theme.border}`, background: "rgba(255,255,255,0.06)", color: theme.text, fontWeight: 950, cursor: "pointer" }}
            onClick={() => applyAutoLoopPlan("factory")}
            title="Choose the closest achievable loop configuration using discrete loops (no fine-adjust)."
          >
            Auto: closest factory loops
          </button>

          <button
            type="button"
            style={{ padding: "6px 10px", borderRadius: 999, border: `1px solid ${theme.border}`, background: "rgba(255,255,255,0.06)", color: theme.text, fontWeight: 950, cursor: "pointer" }}
            onClick={() => applyAutoLoopPlan("minimal")}
            title="Bring the wing within tolerance with the least loop change, using discrete loops only (no fine-adjust)."
          >
            Auto: minimal loop changes
          </button>

          {autoLoopStatus ? (
            <div style={{ padding: "6px 10px", borderRadius: 999, border: `1px solid ${theme.border}`, background: "rgba(99,102,241,0.20)", color: theme.text, fontWeight: 950, fontSize: 12 }}>
              Auto applied
            </div>
          ) : null}
        </div>
      </div>
      <div style={{ overflow: "auto", border: `1px solid ${theme.border}`, borderRadius: 12 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 760 }}>
          <thead>
            <tr style={{ background: "rgba(255,255,255,0.05)" }}>
              <th style={th} rowSpan={2}>Group</th>
              <th style={th} colSpan={5}>Left</th>
              <th style={th} colSpan={5}>Right</th>
            </tr>
            <tr style={{ background: "rgba(255,255,255,0.05)" }}>
              {["Rep", "Suggest", "Loop Δ", "Adj", "Residual"].map((h) => (
                <th key={`L-${h}`} style={th}>{h}</th>
              ))}
              {["Rep", "Suggest", "Loop Δ", "Adj", "Residual"].map((h) => (
                <th key={`R-${h}`} style={th}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {["A", "B", "C", "D"].map((L) => {
              const sL = abcSuggestions?.[L]?.L || null;
              const sR = abcSuggestions?.[L]?.R || null;

              const fmt1 = (n) => (n == null || !Number.isFinite(Number(n)) ? "—" : Number(n).toFixed(1));
              const fmti = (n) => (n == null || !Number.isFinite(Number(n)) ? "—" : Math.round(Number(n)));

              return (
                <tr key={L} style={{ borderTop: `1px solid ${theme.border}` }}>
                  <td style={{ ...td, fontWeight: 950, color: (PALETTE[L] || PALETTE.A).base }}>{L}</td>

                  <td style={td}>{sL ? sL.repLoop : "—"}</td>
                  <td style={td}>{sL ? sL.bestLoop : "—"}</td>
                  <td style={td}>{sL ? fmti(sL.loopDeltaMm) : "—"}</td>
                  <td style={td}>{sL ? fmti(sL.suggestedAdjMm) : "—"}</td>
                  <td style={{ ...td, fontWeight: 950 }}>{sL ? fmt1(sL.residualAfterAdj) : "—"}</td>

                  <td style={td}>{sR ? sR.repLoop : "—"}</td>
                  <td style={td}>{sR ? sR.bestLoop : "—"}</td>
                  <td style={td}>{sR ? fmti(sR.loopDeltaMm) : "—"}</td>
                  <td style={td}>{sR ? fmti(sR.suggestedAdjMm) : "—"}</td>
                  <td style={{ ...td, fontWeight: 950 }}>{sR ? fmt1(sR.residualAfterAdj) : "—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={{ marginTop: 8, opacity: 0.78, fontSize: 12 }}>
        Fine adjust is clamped to ±Tolerance (<b>{Number(meta?.tolerance ?? 0)}mm</b>). Suggestions use the most common current loop type in each A/B/C side as the “representative” baseline.
      </div>

      <div style={{ marginTop: 10, borderTop: `1px solid ${theme.border}`, paddingTop: 10 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 10, flexWrap: "wrap" }}>
          <div style={{ fontWeight: 950 }}>Pitch summary (A/B vs C/D)</div>
          <div style={{ opacity: 0.78, fontSize: 12 }}>
            Pitch proxy = avg Δ(front) − avg Δ(rear). Default tolerance ±{Number.isFinite(Number(groupPitchTol)) ? Number(groupPitchTol) : 4}mm.
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 10, marginTop: 10 }}>
          <div style={{ border: `1px solid ${theme.border}`, borderRadius: 12, background: "rgba(0,0,0,0.22)", padding: 10, overflow: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 520 }}>
              <thead>
                <tr style={{ background: "rgba(255,255,255,0.05)" }}>
                  <th style={th}>Metric</th>
                  <th style={th}>Left</th>
                  <th style={th}>Right</th>
                  <th style={th}>Whole wing</th>
                </tr>
              </thead>
              <tbody>
                {[
                  { key: "whole", label: "Whole pitch (front − rear)", v: pitchStats && pitchStats.pitchWhole ? pitchStats.pitchWhole : null },
                  { key: "ab", label: "A − B", v: pitchStats && pitchStats.segments ? pitchStats.segments.AB : null },
                  { key: "bc", label: "B − C", v: pitchStats && pitchStats.segments ? pitchStats.segments.BC : null },
                  { key: "cd", label: "C − D", v: pitchStats && pitchStats.segments ? pitchStats.segments.CD : null },
                ].map((row) => {
                  const f1 = (n) => (n == null || !Number.isFinite(Number(n)) ? "—" : Number(n).toFixed(1));
                  const vL = row.v ? row.v.L : null;
                  const vR = row.v ? row.v.R : null;
                  const vB = row.v ? row.v.both : null;

                  const sevL = severity(vL, groupPitchTol);
                  const sevR = severity(vR, groupPitchTol);
                  const sevB = severity(vB, groupPitchTol);

                  const colFor = (sev) => {
                    if (sev === "red") return "rgba(255,90,90,1)";
                    if (sev === "yellow") return "rgba(255,215,90,1)";
                    return "rgba(140,255,190,1)";
                  };

                  return (
                    <tr key={row.key} style={{ borderTop: `1px solid ${theme.border}` }}>
                      <td style={{ ...td, fontWeight: 950 }}>{row.label}</td>
                      <td style={{ ...td, fontWeight: 950, color: colFor(sevL) }}>{f1(vL)}</td>
                      <td style={{ ...td, fontWeight: 950, color: colFor(sevR) }}>{f1(vR)}</td>
                      <td style={{ ...td, fontWeight: 950, color: colFor(sevB) }}>{f1(vB)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>

            <div style={{ marginTop: 8, opacity: 0.78, fontSize: 12 }}>
              Interpretation tip: if A/B averages are relatively longer than C/D, the wing tends to fly “faster / lower AoA”; if A/B are shorter relative to C/D, it tends to fly “slower / higher AoA”.
            </div>
          </div>

          <div style={{ border: `1px solid ${theme.border}`, borderRadius: 12, background: "rgba(0,0,0,0.22)", padding: 10 }}>
            <div style={{ fontWeight: 950, marginBottom: 6 }}>Wing pitch profile (concept)</div>
            <div style={{ opacity: 0.75, fontSize: 12, marginBottom: 8 }}>
              A simple chord-line visual with your computed whole-wing pitch overlaid (not a flight dynamics simulator).
            </div>

            {pitchStats && pitchStats.pitchWhole ? (
              <WingPitchViz pitchMm={pitchStats.pitchWhole.both} tolMm={groupPitchTol} />
            ) : null}
          </div>
        </div>

        <div style={{ marginTop: 10 }}>
          <div style={{ fontWeight: 950, marginBottom: 6 }}>Row averages overlaid (A/B/C/D)</div>
          <PitchTrimChart rows={step4LineRows} tolerance={groupPitchTol} height={200} />
        </div>
      </div>

      {showLoopModeCounts ? (
        <div style={{ marginTop: 10, borderTop: `1px solid ${theme.border}`, paddingTop: 10 }}>
          <div style={{ fontWeight: 950, marginBottom: 6 }}>Loop counts used to pick “Rep”</div>
          <div style={{ display: "grid", gap: 8 }}>
            {["A", "B", "C", "D"].map((L) => (
              <div key={`counts-${L}`} style={{ border: `1px solid ${theme.border}`, borderRadius: 12, padding: 8, background: "rgba(0,0,0,0.25)" }}>
                <div style={{ fontWeight: 950, marginBottom: 6, color: (PALETTE[L] || PALETTE.A).base }}>{L}</div>
                {["L", "R"].map((side) => {
                  const counts = abcLoopModeCounts?.[L]?.[side] || {};
                  const entries = Object.entries(counts).sort((a, b) => (Number(b[1]) - Number(a[1])) || String(a[0]).localeCompare(String(b[0])));
                  return (
                    <div key={`${L}-${side}`} style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 6 }}>
                      <span style={{ opacity: 0.8, fontWeight: 950, width: 42 }}>{side === "L" ? "Left" : "Right"}</span>
                      {entries.length === 0 ? (
                        <span style={{ opacity: 0.7 }}>—</span>
                      ) : (
                        entries.map(([lt, c]) => (
                          <span key={`${L}-${side}-${lt}`} style={{ padding: "3px 8px", borderRadius: 999, border: `1px solid ${theme.border}`, background: "rgba(255,255,255,0.06)", fontWeight: 950, fontSize: 12 }}>
                            {lt}: {c}
                          </span>
                        ))
                      )}
                    </div>
                  );
                })}
              </div>
            ))}
          </div>
        </div>
      ) : null}
    </div>
  </div>
</div>

  <style>{`
    @media (max-width: 860px) {
      .abcGrid { grid-template-columns: 1fr !important; }
    }
  `}</style>
</div>
<div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 10 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                    <div>
                      <div style={{ fontWeight: 950 }}>Whole wing — per line lengths</div>
                      <div style={{ opacity: 0.78, fontSize: 13, marginTop: 4 }}>
                        Each line side (L/R) is treated as a separate entity. Values use the <b>frozen baseline</b> + Step 4 overrides/adjustments.
                      </div>
                    </div>

                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                      <button
                        style={{ ...topBtn, ...(showCorrected ? topBtnActive : null) }}
                        onClick={() => setShowCorrected((v) => !v)}
                        title="Toggle whether correction is applied to measured values"
                      >
                        {showCorrected ? "Corrected: ON" : "Corrected: OFF"}
                      </button>

                      {["A", "B", "C", "D"].map((L) => (
                        <button
                          key={L}
                          style={{
                            ...topBtn,
                            background: step4LetterFilter?.[L] ? "rgba(255,255,255,0.10)" : "rgba(255,255,255,0.04)",
                            borderColor: step4LetterFilter?.[L] ? "rgba(255,255,255,0.22)" : theme.border,
                            color: step4LetterFilter?.[L] ? theme.text : "rgba(255,255,255,0.65)",
                          }}
                          onClick={() => setStep4LetterFilter((p) => ({ ...(p || {}), [L]: !p?.[L] }))}
                        >
                          {L}
                        </button>
                      ))}
                    </div>
                  </div>

                  {step4LineRows.length === 0 ? (
                    <div style={{ marginTop: 10, opacity: 0.75 }}>No per-line data available yet. Import data in Step 1 and map lines in Step 2.</div>
                  ) : (
                    <div
                      style={{
                        marginTop: 10,
                        overflow: "auto", width: "100%", maxWidth: "100%",
          minWidth: 0, border: `1px solid ${theme.border}`,
                        borderRadius: 12,
                        maxHeight: 520,
                      }}
                    >
                      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1200 }}>
                        <thead>
                          <tr>
                            {[
                              "Line",
                              "Group",
                              "Nominal",
                              "Raw",
                              "Corrected",
                              "Baseline loop",
                              "After loop",
                              "Adj (mm)",
                              "Before",
                              "After",
                              "Δ vs nominal",
                            ].map((h) => (
                              <th key={h} style={{ ...th, position: "sticky", top: 0, background: theme.panel2, zIndex: 1 }}>
                                {h}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {step4LineRows.map((r) => {
                            const sevColor = r.sev === "bad" ? theme.bad : r.sev === "warn" ? theme.warn : r.sev === "green" ? theme.green : r.sev === "good" ? theme.good : "rgba(255,255,255,0.65)";
                            const rowBg =
                              r.sev === "bad" ? "rgba(239,68,68,0.08)" : r.sev === "warn" ? "rgba(245,158,11,0.08)" : "transparent";

                            const fmt = (n) => (n == null || !Number.isFinite(Number(n)) ? "—" : Math.round(Number(n)));

                            return (
                              <tr key={r.lineId} style={{ borderTop: `1px solid ${theme.border}`, background: rowBg }}>
                                <td style={{ ...td, fontWeight: 950, color: chipColorFromLineId(r.lineId) }}>{r.lineId}</td>
                                <td style={td}>{r.groupId || "—"}</td>
                                <td style={td}>{fmt(r.nominal)}</td>
                                <td style={td}>{fmt(r.raw)}</td>
                                <td style={td}>{showCorrected ? fmt(r.corrected) : "—"}</td>
                                <td style={td}>{r.baseLoop}</td>
                                <td style={td}>{r.afterLoop}</td>
                                <td style={td}>{fmt(r.adj)}</td>
                                <td style={td}>{fmt(r.before)}</td>
                                <td style={td}>{fmt(r.after)}</td>
                                <td style={{ ...td, fontWeight: 950, color: sevColor }}>{r.delta == null ? "—" : fmt(r.delta)}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}

                  <div style={{ marginTop: 10, opacity: 0.78, fontSize: 13 }}>
                    Tolerance: <b>{Number(meta?.tolerance ?? 0)}mm</b> • Yellow within <b>3mm</b> of tolerance • Red at/over tolerance
                  </div>
                </div>

                
                
{/* Charts (Step 4 only; uses frozen baseline-derived step4LineRows) */}
<div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 12 }}>
  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
    <div style={{ fontWeight: 950, fontSize: 16 }}>Charts</div>
    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", opacity: 0.9 }}>
      <span style={{ fontSize: 12, opacity: 0.85 }}>Showing:</span>
      {(["A", "B", "C", "D"]).map((L) => (
        <span
          key={L}
          style={{
            padding: "3px 8px",
            borderRadius: 999,
            border: `1px solid ${theme.border}`,
            background: step4LetterFilter?.[L] ? "rgba(255,255,255,0.08)" : "rgba(255,255,255,0.03)",
            color: step4LetterFilter?.[L] ? "rgba(255,255,255,0.9)" : "rgba(255,255,255,0.45)",
            fontWeight: 950,
            fontSize: 12,
          }}
        >
          {L}
        </span>
      ))}
    </div>
  </div>

  <PitchTrimChart rows={step4LineRows} tolerance={Number(meta?.tolerance ?? 0)} height={220} />

  {/* Group averages (After Δ) — placed under Pitch trim */}
  <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel2, padding: 12, width: "100%" }}>
    <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 12, marginBottom: 8, flexWrap: "wrap" }}>
      <div style={{ fontWeight: 950 }}>Group averages (After Δ)</div>
      <div style={{ opacity: 0.75, fontSize: 12 }}>A/B/C/D averages — L, R, and symmetry</div>
    </div>

    <div style={{ overflow: "auto", border: `1px solid ${theme.border}`, borderRadius: 12, width: "100%" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 520 }}>
        <thead>
          <tr style={{ background: "rgba(255,255,255,0.05)" }}>
            <th style={th}>Group</th>
            <th style={th}>Left avg Δ</th>
            <th style={th}>n</th>
            <th style={th}>Right avg Δ</th>
            <th style={th}>n</th>
            <th style={th}>Sym avg (L−R)</th>
          </tr>
        </thead>
        <tbody>
          {(["A", "B", "C", "D"]).map((L) => {
            const aL = abcAverages?.[L]?.L?.avg;
            const nL = abcAverages?.[L]?.L?.n ?? 0;
            const aR = abcAverages?.[L]?.R?.avg;
            const nR = abcAverages?.[L]?.R?.n ?? 0;
            const sym = abcAverages?.[L]?.sym;

            const f1 = (n) => (n == null || !Number.isFinite(Number(n)) ? "—" : Number(n).toFixed(1));

            return (
              <tr key={L} style={{ borderTop: `1px solid ${theme.border}` }}>
                <td style={{ ...td, fontWeight: 950, color: (PALETTE[L] || PALETTE.A).base }}>{L}</td>
                <td style={td}>{f1(aL)}</td>
                <td style={{ ...td, opacity: 0.85 }}>{nL || "—"}</td>
                <td style={td}>{f1(aR)}</td>
                <td style={{ ...td, opacity: 0.85 }}>{nR || "—"}</td>
                <td style={{ ...td, fontWeight: 950 }}>{f1(sym)}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>

    <div style={{ marginTop: 8, opacity: 0.78, fontSize: 12 }}>
      Tip: Positive Δ means “long” vs nominal; negative means “short”. Sym is based on averages only.
    </div>
  </div>

  <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 12 }}>
    {(["A", "B", "C", "D"]).map((L) =>
      step4LetterFilter?.[L] && (chartPointsByLetter?.[L]?.length || 0) > 0 ? (
        <DeltaLineChart
          key={L}
          title={`Δ per line — ${L} rows (After vs Before)  `}
          points={chartPointsByLetter[L]}
          tolerance={Number(meta?.tolerance ?? 0)}
          height={240}
        />
      ) : null
    )}
  </div>

  <WingProfileChart groupStats={step4GroupStats} tolerance={Number(meta?.tolerance ?? 0)} height={260} />
</div>

{/* Close Step 4 content grid */}
</div>

            )}
          </Panel>
            </div>
          </div>
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

// Small input style used in compact Step 4 group table
const miniInput = {
  width: 74,
  padding: "6px 8px",
  borderRadius: 10,
  border: `1px solid ${theme.border}`,
  background: "rgba(255,255,255,0.08)",
  color: theme.text,
  fontSize: 12,
  textAlign: "center",
};


const step123Wrap = {
  width: "100%",
  maxWidth: 1100,
  margin: "0 auto",
};


function RearViewChart({ rows, tolerance, height, loopTypes, groupLoopChange, setGroupLoopChange }) {
  const width = 1240;
  const heightPx = Number.isFinite(Number(height)) ? Number(height) : 460;
  const pad = 24;
  const tol = Number.isFinite(Number(tolerance)) ? Number(tolerance) : 0;

  const [hover, setHover] = React.useState(null);
  const [showWingOutline, setShowWingOutline] = React.useState(true);
  const [showBeforePoints, setShowBeforePoints] = React.useState(false);
  const [showGroupCuts, setShowGroupCuts] = React.useState(true);
  const [spanMode, setSpanMode] = React.useState("real");
  const [pickedGroupId, setPickedGroupId] = React.useState(null);

  const baselineLoopByGroupKey = React.useMemo(() => {
    const out = {};
    if (!Array.isArray(rows)) return out;
    for (let i = 0; i < rows.length; i++) {
      const rr = rows[i];
      const raw = rr && (rr.groupId || rr.group) ? String(rr.groupId || rr.group) : "";
      const key = raw ? (raw.split("|")[0] || "") : "";
      if (!key) continue;
      const base = rr && rr.baseLoop ? String(rr.baseLoop) : "";
      if (base && !out[key]) out[key] = base;
    }
    return out;
  }, [rows]);

  // Build per-cascade points (A/B/C/D rows) from Step 4 computed per-line rows.
  // IMPORTANT: This chart ONLY uses Step 4 computed rows (frozen baseline + overrides + adjustments),
  // never Step 3 live state.
  const data = React.useMemo(() => {
    if (!Array.isArray(rows) || rows.length === 0) return null;

    const byKey = new Map();
    for (const rr of rows) {
      const letter = String(rr?.letter || "").toUpperCase();
      if (!["A", "B", "C", "D"].includes(letter)) continue;
      const idx = Number(rr?.idx);
      if (!Number.isFinite(idx)) continue;

      const key = `${letter}|${idx}`;
      let p = byKey.get(key);
      if (!p) {
        p = {
          letter,
          idx,
          lineId: `${letter}${idx}`,
          groupNameL: "—",
          groupNameR: "—",
          beforeL: null,
          beforeR: null,
          afterL: null,
          afterR: null,
          lineIdL: null,
          lineIdR: null,
        };
        byKey.set(key, p);
      }

      const side = String(rr?.side || "").toUpperCase();
      const groupName = String(rr?.groupId || rr?.group || "—").split("|")[0] || "—";
      if (side === "L") p.groupNameL = groupName;
      else if (side === "R") p.groupNameR = groupName;
      const nominal = rr?.nominal;
      const beforeAbs = rr?.before;
      const afterDelta = rr?.delta;

      const beforeDelta = nominal == null || beforeAbs == null ? null : beforeAbs - nominal;

      if (side === "L") {
        p.lineIdL = rr?.lineId || p.lineIdL;
        p.beforeL = beforeDelta;
        p.afterL = afterDelta;
      } else if (side === "R") {
        p.lineIdR = rr?.lineId || p.lineIdR;
        p.beforeR = beforeDelta;
        p.afterR = afterDelta;
      }
    }

    const points = Array.from(byKey.values()).sort((a, b) => {
      const la = a.letter.localeCompare(b.letter);
      if (la) return la;
      return a.idx - b.idx;
    });

    const byLetter = { A: [], B: [], C: [], D: [] };
    for (const p of points) {
      if (byLetter[p.letter]) byLetter[p.letter].push(p);
    }
    return { points, ...byLetter };
  }, [rows]);



  if (!data) {
    return (
      <div
        style={{
          padding: 12,
          border: "1px solid #2a2f3f",
          borderRadius: 14,
          background: "#0e1018",
          color: "#aab1c3",
          fontSize: 12,
        }}
      >
        Rear view chart will appear after importing a file.
      </div>
    );
  }

  const bands = {
    A: { y0: pad + 74, y1: pad + 74 + 85 },
    B: { y0: pad + 74 + 95, y1: pad + 74 + 180 },
    C: { y0: pad + 74 + 190, y1: pad + 74 + 275 },
    D: { y0: pad + 74 + 285, y1: pad + 74 + 370 },
  };

  function sevColor(sev) {
    if (sev === "red") return "rgba(255,90,90,1)";
    if (sev === "yellow") return "rgba(255,215,90,1)";
    return "rgba(140,255,190,1)";
  }

  function bandY(letter, v) {
    const b = bands[letter];
    const range = Math.max(30, tol > 0 ? tol * 2.2 : 50);
    const mid = (b.y0 + b.y1) / 2;
    const pxPerMm = (b.y1 - b.y0) / (range * 2);
    return mid - v * pxPerMm;
  }

  function spanScale(t) {
    if (spanMode === "linear") return t;
    const gamma = 0.75; // <1 expands inner, compresses tips
    return Math.pow(t, gamma);
  }

  function xFor(side, i, count) {
    const center = width / 2;
    const halfSpan = (width - pad * 2) / 2 - 20;
    const centerGap = 18;

    const countFixed = 25;

    const t = countFixed <= 1 ? 0 : i / (countFixed - 1);
    const ts = spanScale(t);
    const dx = ts * halfSpan + centerGap;

    return side === "L" ? center - dx : center + dx;
  }

  
function groupCuts(letter) {
  if (!showGroupCuts) return [];
  const arr = data[letter] || [];
  const out = [];
  let last = null;
  for (let i = 0; i < arr.length; i++) {
    const g = (arr[i] && (arr[i].groupNameL || arr[i].groupNameR)) || "";
    if (i === 0) {
      last = g;
      continue;
    }
    if (g !== last) {
      out.push({ idx: i - 0.5, from: last, to: g });
      last = g;
    }
  }
  return out;
}

function groupBands(letter, side) {
  if (!showGroupCuts) return [];
  const arr = data[letter] || [];
  const key = side === "L" ? "groupNameL" : "groupNameR";
  const out = [];
  let start = 0;
  let current = null;
  for (let i = 0; i < arr.length; i++) {
    const g = (arr[i] && arr[i][key]) || "—";
    if (i === 0) {
      current = g;
      start = 0;
      continue;
    }
    if (g !== current) {
      out.push({ start, end: i - 1, name: current });
      current = g;
      start = i;
    }
  }
  if (arr.length > 0) out.push({ start, end: arr.length - 1, name: current || "—" });
  return out;
}

  return (
    <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start", flexWrap: "wrap" }}>
        <div>
          <div style={{ fontWeight: 900, marginBottom: 6 }}>Rear view wing shape (A/B/C/D rows)</div>
          <div style={{ color: "#aab1c3", fontSize: 12, lineHeight: 1.5 }}>
            Symmetric about the centreline. Points are <b>After</b> (severity color). Dashed = Before.
          </div>
        </div>

        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <label style={{ color: "#aab1c3", fontSize: 12, display: "flex", gap: 8, alignItems: "center" }}>
            Span spacing
            <select
              value={spanMode}
              onChange={(e) => setSpanMode(e.target.value)}
              style={{
                borderRadius: 10,
                border: "1px solid #2a2f3f",
                background: "#0d0f16",
                color: "#eef1ff",
                padding: "6px 10px",
                outline: "none",
                fontSize: 12,
              }}
            >
              <option value="real">Realistic</option>
              <option value="linear">Linear</option>
            </select>
          </label>

          <label style={{ color: "#aab1c3", fontSize: 12, display: "flex", gap: 8, alignItems: "center" }}>
            <input type="checkbox" checked={showGroupCuts} onChange={(e) => setShowGroupCuts(e.target.checked)} />
            Group boundaries
          </label>

          <label style={{ color: "#aab1c3", fontSize: 12, display: "flex", gap: 8, alignItems: "center" }}>
            <input type="checkbox" checked={showBeforePoints} onChange={(e) => setShowBeforePoints(e.target.checked)} />
            Before points
          </label>
        </div>
      </div>

      <div style={{ height: 10 }} />

      <div style={{ overflowX: "auto", maxWidth: "100%" }}>
        <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} style={{ display: "block" }}>
          {/* Top labels */}
          <text x={pad} y={pad + 16} fill="rgba(170,177,195,0.9)" fontSize="12" fontFamily="ui-monospace, Menlo, Consolas, monospace">
            LEFT
          </text>
          <text
            x={width - pad}
            y={pad + 16}
            textAnchor="end"
            fill="rgba(170,177,195,0.9)"
            fontSize="12"
            fontFamily="ui-monospace, Menlo, Consolas, monospace"
          >
            RIGHT
          </text>
          <text
            x={width / 2}
            y={pad + 16}
            textAnchor="middle"
            fill="rgba(170,177,195,0.9)"
            fontSize="12"
            fontFamily="ui-monospace, Menlo, Consolas, monospace"
          >
            CENTRE
          </text>

          {/* Centreline */}
          <line x1={width / 2} y1={pad + 24} x2={width / 2} y2={height - pad} stroke="rgba(42,47,63,0.85)" strokeDasharray="6 6" />

          {/* Span ticks + MID→TIP labels */}
          {(() => {
            const y = pad + 30;
            const center = width / 2;
            const halfSpan = (width - pad * 2) / 2 - 20;
            const centerGap = 18;
            const ticks = [
              { t: 0.0, label: "MID" },
              { t: 0.25, label: "25%" },
              { t: 0.5, label: "50%" },
              { t: 0.75, label: "75%" },
              { t: 1.0, label: "TIP" },
            ];
            const scaleT = (t) => (spanMode === "linear" ? t : Math.pow(t, 0.75));

            return (
              <g>
                {ticks.map((tk) => {
                  const dx = scaleT(tk.t) * halfSpan + centerGap;
                  const xL = center - dx;
                  const xR = center + dx;
                  return (
                    <g key={`ticks-${tk.t}`}>
                      <line x1={xL} y1={y} x2={xL} y2={y + 8} stroke="rgba(255,255,255,0.10)" />
                      <line x1={xR} y1={y} x2={xR} y2={y + 8} stroke="rgba(255,255,255,0.10)" />
                      <text x={xL} y={y + 22} textAnchor="middle" fill="rgba(170,177,195,0.85)" fontSize="11" fontFamily="ui-monospace, Menlo, Consolas, monospace">
                        {tk.label}
                      </text>
                      <text x={xR} y={y + 22} textAnchor="middle" fill="rgba(170,177,195,0.85)" fontSize="11" fontFamily="ui-monospace, Menlo, Consolas, monospace">
                        {tk.label}
                      </text>
                    </g>
                  );
                })}
              </g>
            );
          })()}

          {/* Subtle wing outline arc (background) */}
          {(() => {
            const left = pad + 20;
            const right = width - pad - 20;
            const top = pad + 66;
            const bottom = height - pad - 18;
            const midX = width / 2;
            const ctrlY = top - 26;

            const d = `
              M ${midX} ${top}
              C ${midX - 180} ${ctrlY}, ${left + 60} ${ctrlY + 10}, ${left} ${top + 18}
              L ${left} ${bottom}
              L ${right} ${bottom}
              L ${right} ${top + 18}
              C ${right - 60} ${ctrlY + 10}, ${midX + 180} ${ctrlY}, ${midX} ${top}
              Z
            `;

            return <path d={d} fill="rgba(255,255,255,0.015)" stroke="rgba(255,255,255,0.06)" strokeWidth="2" />;
          })()}

          {/* Bands + 0mm guides + riser labels */}
          {["A", "B", "C", "D"].map((L) => {
            const b = bands[L];
            const yMid = (b.y0 + b.y1) / 2;

            return (
              <g key={`band-${L}`}>
                <rect x={pad} y={b.y0} width={width - pad * 2} height={b.y1 - b.y0} fill="rgba(255,255,255,0.02)" />
                <line x1={pad} y1={b.y0} x2={width - pad} y2={b.y0} stroke="rgba(42,47,63,0.85)" />
                <line x1={pad} y1={b.y1} x2={width - pad} y2={b.y1} stroke="rgba(42,47,63,0.85)" />

                {/* 0mm guide (target) */}
                <line x1={pad} y1={yMid} x2={width - pad} y2={yMid} stroke="rgba(255,90,90,0.85)" strokeDasharray="4 6" />

                {/* Row label */}
                <text x={pad + 8} y={b.y0 + 18} fill="rgba(170,177,195,0.85)" fontSize="12" fontFamily="ui-monospace, Menlo, Consolas, monospace">
                  {L}-row
                </text>

                {/* Riser label at centreline */}
                <text
                  x={width / 2}
                  y={b.y0 + 18}
                  textAnchor="middle"
                  fill="rgba(238,241,255,0.85)"
                  fontSize="12"
                  fontFamily="ui-monospace, Menlo, Consolas, monospace"
                >
                  {L}
                </text>
              </g>
            );
          })}

          {/* Plots */}
          {["A", "B", "C", "D"].map((L) => {
            const arr = data[L] || [];
            const count = arr.length || 1;

            const buildPath = (side, which) => {
              let d = "";
              for (let i = 0; i < arr.length; i++) {
                const p = arr[i];
                const v =
                  side === "L"
                    ? which === "before"
                      ? p.beforeL
                      : p.afterL
                    : which === "before"
                    ? p.beforeR
                    : p.afterR;

                if (!Number.isFinite(v)) continue;
                const x = xFor(side, i, count);
                const y = bandY(L, v);
                d += d ? ` L ${x} ${y}` : `M ${x} ${y}`;
              }
              return d;
            };

            const cuts = groupCuts(L);

            return (
              <g key={`plot-${L}`}>
                {/* group boundary lines (both sides) */}
                {cuts.map((c, idx) => {
                  const b = bands[L];
                  const xL = xFor("L", c.idx, count);
                  const xR = xFor("R", c.idx, count);
                  return (
                    <g key={`cut-${L}-${idx}`}>
                      <line x1={xL} y1={b.y0 + 2} x2={xL} y2={b.y1 - 2} stroke="rgba(255,220,80,0.85)" />
                      <line x1={xR} y1={b.y0 + 2} x2={xR} y2={b.y1 - 2} stroke="rgba(255,220,80,0.85)" />
                    </g>
                  );
                })}


                {/* Group labels (midpoint of each grouping) */}
                {showGroupCuts &&
                  (() => {
                    const b = bands[L];
                    const bandsL = groupBands(L, "L");
                    const bandsR = groupBands(L, "R");
                    const yText = ((b.y0 + b.y1) / 2) + 38;
                    const renderBand = (side, band, i) => {
                      if (!band || !band.name || band.name === "—") return null;
                      const mid = (band.start + band.end) / 2;
                      const x = xFor(side, mid, count);
                      const key = String((band && band.name) || "");
                      const displayName = key.replace(/([LR])$/, "");
                      const baseLoop = baselineLoopByGroupKey && baselineLoopByGroupKey[key] ? baselineLoopByGroupKey[key] : "";
                      // Cosmetic tooltip: show the grouping key AND the Step 3 baseline loop (reference only).
                      // Use newlines so it reads clearly in the native title tooltip.
                      const titleText = key + "\nBaseline: " + (baseLoop ? baseLoop : "—");
                      const cur = (groupLoopChange && groupLoopChange[key]) ? groupLoopChange[key] : "";
                      const w = 68;
                      const h = 22;
                      return (
                        <foreignObject
                          key={`glabel-${L}-${side}-${i}`}
                          x={x - w / 2}
                          y={yText - h + 6}
                          width={w}
                          height={h}
                        >
                          <div
                            xmlns="http://www.w3.org/1999/xhtml"
                            style={{
                              width: "100%",
                              height: "100%",
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "center",
                              borderRadius: 999,
                              border: "1px solid rgba(255,220,80,0.55)",
                              background: "rgba(0,0,0,0.45)",
                              boxSizing: "border-box",
                            }}
                            title={titleText}
                          >
                            <select
                              value={cur}
                              onChange={(e) => {
                                const v = e.target.value || "";
                                if (!setGroupLoopChange) return;
                                setGroupLoopChange(function (prev) {
                                  const next = Object.assign({}, prev || {});
                                  if (v) next[key] = v;
                                  else delete next[key];
                                  return next;
                                });
                              }}
                              style={{
                                width: "100%",
                                height: "100%",
                                border: "none",
                                outline: "none",
                                background: "transparent",
                                color: (cur ? "rgba(120,200,255,0.95)" : "rgba(255,220,80,0.95)"),
                                fontSize: 11,
                                fontWeight: 950,
                                fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                                textAlignLast: "center",
                                cursor: "pointer",
                                paddingLeft: 6,
                                paddingRight: 6,
                                appearance: "none",
                                WebkitAppearance: "none",
                                MozAppearance: "none",
                              }}
                            >
                              <option value="">{displayName}</option>
                              {((loopTypes && loopTypes.length) ? loopTypes : LOOP_TYPES).map((lt) => (
                                <option key={lt} value={lt}>
                                  {lt}
                                </option>
                              ))}
                            </select>
                          </div>
                        </foreignObject>
                      );
                    };
                    return (
                      <g>
                        {bandsL.map((band, i) => renderBand("L", band, i))}
                        {bandsR.map((band, i) => renderBand("R", band, i))}
                      </g>
                    );
                  })()}
                {/* Before dashed paths */}
                <path d={buildPath("L", "before")} fill="none" stroke="rgba(176,132,255,0.65)" strokeWidth="2" strokeDasharray="6 6" />
                <path d={buildPath("R", "before")} fill="none" stroke="rgba(102,204,255,0.65)" strokeWidth="2" strokeDasharray="6 6" />

                {/* After solid paths */}
                <path d={buildPath("L", "after")} fill="none" stroke="rgba(176,132,255,1)" strokeWidth="3" />
                <path d={buildPath("R", "after")} fill="none" stroke="rgba(102,204,255,1)" strokeWidth="3" />

                {/* Points */}
                {arr.map((p, i) => {
                  const pts = [
                    { side: "L", before: p.beforeL, after: p.afterL },
                    { side: "R", before: p.beforeR, after: p.afterR },
                  ];

                  return pts.map((it) => {
                    const x = xFor(it.side, i, count);

                    // BEFORE points (small hollow circles)
                    const beforeNode =
                      showBeforePoints && Number.isFinite(it.before) ? (
                        <circle
                          key={`${p.lineId}-${it.side}-before`}
                          cx={x}
                          cy={bandY(L, it.before)}
                          r={4}
                          fill="transparent"
                          stroke="rgba(255,255,255,0.30)"
                          strokeWidth="2"
                        />
                      ) : null;

                    // AFTER points (colored)
                    const afterNode = Number.isFinite(it.after) ? (
                      <circle
                        key={`${p.lineId}-${it.side}-after`}
                        cx={x}
                        cy={bandY(L, it.after)}
                        r={5}
                        fill={sevColor(severity(it.after, tol))}
                        stroke="rgba(10,12,16,0.9)"
                        strokeWidth="2"
                        onMouseEnter={() =>
                          setHover({
                            letter: L,
                            lineId: p.lineId,
                            groupName: it.side === "L" ? p.groupNameL : p.groupNameR,
                            side: it.side,
                            before: it.before,
                            after: it.after,
                            sev: severity(it.after, tol),
                            x,
                            y: bandY(L, it.after),
                          })
                        }
                        onMouseLeave={() => setHover(null)}
                      />
                    ) : null;

                    return (
                      <g key={`${p.lineId}-${it.side}`}>
                        {beforeNode}
                        {afterNode}
                      </g>
                    );
                  });
                })}
              </g>
            );
          })}

          {/* Tooltip */}
          {hover ? (
            <g>
              <rect
                x={Math.min(width - 330, Math.max(10, hover.x + 12))}
                y={Math.max(10, hover.y - 80)}
                width={320}
                height={70}
                rx={10}
                ry={10}
                fill="rgba(12,14,22,0.95)"
                stroke="rgba(42,47,63,1)"
              />
              <text
                x={Math.min(width - 312, Math.max(20, hover.x + 22))}
                y={Math.max(28, hover.y - 52)}
                fill="#eef1ff"
                fontSize="12"
                fontFamily="ui-monospace, Menlo, Consolas, monospace"
              >
                {`${hover.lineId} (${hover.side})  group: ${hover.groupName}`}
              </text>
              <text
                x={Math.min(width - 312, Math.max(20, hover.x + 22))}
                y={Math.max(48, hover.y - 32)}
                fill="rgba(170,177,195,0.95)"
                fontSize="12"
                fontFamily="ui-monospace, Menlo, Consolas, monospace"
              >
                {`Before: ${Number.isFinite(hover.before) ? Math.round(hover.before) : "—"}mm   After: ${
                  Number.isFinite(hover.after) ? Math.round(hover.after) : "—"
                }mm   Sev: ${hover.sev}`}
              </text>
            </g>
          ) : null}
        </svg>
      </div>

      <div style={{ color: "#aab1c3", fontSize: 12, marginTop: 8, display: "flex", gap: 14, flexWrap: "wrap" }}>
        <span>Solid = After</span>
        <span>Dashed = Before</span>
        <span>Target (0mm) = dotted line</span>
        {tol > 0 ? (
          <>
            <span>Yellow = within 3mm of tolerance</span>
            <span>Red = outside tolerance</span>
          </>
        ) : null}
      </div>
    </div>
  );
}




function WingPitchViz({ pitchMm, tolMm, height = 160 }) {
  const w = 980;
  const h = height;
  const xPad = 80;
  const yMid = h / 2;

  const p = Number.isFinite(Number(pitchMm)) ? Number(pitchMm) : 0;
  const tol = Number.isFinite(Number(tolMm)) ? Number(tolMm) : 0;

  // Map millimetres to a small visual rotation (purely illustrative).
  const clamp = (v, lo, hi) => (v < lo ? lo : (v > hi ? hi : v));
  const pClamped = clamp(p, -20, 20);
  const deg = (pClamped / 20) * 8; // ±20mm -> ±8°

  const sev = severity(p, tol);
  const col = sev === "red" ? "rgba(255,90,90,1)" : (sev === "yellow" ? "rgba(255,215,90,1)" : "rgba(140,255,190,1)");

  const cx = w / 2;
  const cy = yMid;

  const chordLen = w - xPad * 2;
  const x1 = cx - chordLen / 2;
  const x2 = cx + chordLen / 2;

  return (
    <div style={{ width: "100%" }}>
      <svg width="100%" viewBox={`0 0 ${w} ${h}`} style={{ display: "block" }}>
        {/* Reference horizon */}
        <line x1={xPad} y1={yMid} x2={w - xPad} y2={yMid} stroke="rgba(148,163,184,0.25)" />

        {/* Wing chord line (rotated) */}
        <g transform={`rotate(${deg} ${cx} ${cy})`}>
          <line x1={x1} y1={cy} x2={x2} y2={cy} stroke={col} strokeWidth="10" strokeLinecap="round" opacity="0.95" />
          {/* Leading edge marker */}
          <circle cx={x1} cy={cy} r="10" fill="rgba(255,255,255,0.9)" />
          <text x={x1 - 6} y={cy + 34} fontSize="24" fill="rgba(255,255,255,0.75)">LE</text>
          {/* Trailing edge marker */}
          <circle cx={x2} cy={cy} r="10" fill="rgba(255,255,255,0.5)" />
          <text x={x2 - 12} y={cy + 34} fontSize="24" fill="rgba(255,255,255,0.75)">TE</text>
        </g>

        {/* Readout */}
        <text x={xPad} y={28} fontSize="26" fill="rgba(255,255,255,0.92)" fontWeight="900">
          Pitch: {Number.isFinite(p) ? p.toFixed(1) : "—"} mm
        </text>
        <text x={xPad} y={58} fontSize="20" fill="rgba(255,255,255,0.7)">
          Visual rotation: {Number.isFinite(deg) ? deg.toFixed(1) : "0.0"}°
        </text>

        {tol > 0 ? (
          <text x={w - xPad} y={28} fontSize="20" fill="rgba(255,255,255,0.7)" textAnchor="end">
            Tolerance: ±{tol.toFixed(0)}mm
          </text>
        ) : null}
      </svg>

      <div style={{ display: "flex", justifyContent: "space-between", gap: 8, flexWrap: "wrap", marginTop: 4 }}>
        <div style={{ opacity: 0.75, fontSize: 12 }}>Green/yellow/red reflects the tolerance threshold.</div>
        <div style={{ opacity: 0.75, fontSize: 12 }}>This graphic is illustrative only.</div>
      </div>
    </div>
  );
}

function PitchTrimChart({ rows, tolerance, height = 220 }) {
  const safeTol = Number.isFinite(Number(tolerance)) ? Number(tolerance) : 0;
  const w = 980;
  const h = height;

  const stats = useMemo(() => {
    const out = [];
    const list = Array.isArray(rows) ? rows : [];
    for (const letter of ["A", "B", "C", "D"]) {
      const L = list.filter((r) => r?.letter === letter && r?.side === "L" && Number.isFinite(Number(r?.delta)));
      const R = list.filter((r) => r?.letter === letter && r?.side === "R" && Number.isFinite(Number(r?.delta)));
      const mean = (arr) => (arr.length ? arr.reduce((s, r) => s + Number(r.delta), 0) / arr.length : 0);
      out.push({ letter, left: mean(L), right: mean(R), nL: L.length, nR: R.length });
    }
    return out;
  }, [rows]);

  const maxAbs = useMemo(() => {
    let m = 5;
    for (const s of stats) m = Math.max(m, Math.abs(s.left), Math.abs(s.right), safeTol || 0);
    return m;
  }, [stats, safeTol]);

  const xPad = 60;
  const yPad = 30;
  const xStep = (w - 2 * xPad) / Math.max(1, stats.length);
  const yMid = h / 2;
  const yScale = (h * 0.38) / Math.max(10, maxAbs);

  const yFor = (v) => yMid - v * yScale;

  const barW = Math.max(10, xStep * 0.22);

  const bandFill = (delta) => {
    const b = bandForDelta(delta, safeTol);
    if (b === "good") return "rgba(34,197,94,0.75)";
    if (b === "warn") return "rgba(245,158,11,0.75)";
    if (b === "bad") return "rgba(239,68,68,0.78)";
    return "rgba(148,163,184,0.55)";
  };

  return (
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.bg2, padding: 12 }}>
      <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 12, marginBottom: 8 }}>
        <div style={{ fontWeight: 950 }}>Pitch trim (avg Δ after vs nominal)</div>
        <div style={{ opacity: 0.75, fontSize: 12 }}>Per row average — L and R</div>
      </div>

      <svg width="100%" viewBox={`0 0 ${w} ${h}`} style={{ display: "block" }}>
        {/* Midline */}
        <line x1={xPad} y1={yMid} x2={w - xPad} y2={yMid} stroke="rgba(148,163,184,0.25)" />
        {/* Tolerance bands */}
        {safeTol > 0 && (
          <>
            <line x1={xPad} y1={yFor(+safeTol)} x2={w - xPad} y2={yFor(+safeTol)} stroke="rgba(239,68,68,0.18)" />
            <line x1={xPad} y1={yFor(-safeTol)} x2={w - xPad} y2={yFor(-safeTol)} stroke="rgba(239,68,68,0.18)" />
            <line x1={xPad} y1={yFor(+4)} x2={w - xPad} y2={yFor(+4)} stroke="rgba(34,197,94,0.18)" />
            <line x1={xPad} y1={yFor(-4)} x2={w - xPad} y2={yFor(-4)} stroke="rgba(34,197,94,0.18)" />
          </>
        )}

        {stats.map((s, i) => {
          const cx = xPad + xStep * (i + 0.5);
          const yL = yFor(s.left);
          const yR = yFor(s.right);

          const barTopL = Math.min(yMid, yL);
          const barH_L = Math.abs(yL - yMid);
          const barTopR = Math.min(yMid, yR);
          const barH_R = Math.abs(yR - yMid);

          return (
            <g key={s.letter}>
              {/* label */}
              <text x={cx} y={h - 8} textAnchor="middle" fontSize="14" fill={theme.text} style={{ fontWeight: 950 }}>
                {s.letter}
              </text>

              {/* Left bar */}
              <rect
                x={cx - barW - 6}
                y={barTopL}
                width={barW}
                height={Math.max(0.5, barH_L)}
                rx="6"
                fill={bandFill(s.left)}
              />
              <text x={cx - barW / 2 - 6} y={barTopL - 6} textAnchor="middle" fontSize="12" fill={theme.textSub}>
                L {s.left.toFixed(1)}
              </text>

              {/* Right bar */}
              <rect
                x={cx + 6}
                y={barTopR}
                width={barW}
                height={Math.max(0.5, barH_R)}
                rx="6"
                fill={bandFill(s.right)}
              />
              <text x={cx + 6 + barW / 2} y={barTopR - 6} textAnchor="middle" fontSize="12" fill={theme.textSub}>
                R {s.right.toFixed(1)}
              </text>
            </g>
          );
        })}
      </svg>
    </div>
  );
}
;
// ---------------- Additional charts ported from Reference app good.jsx ----------------
function DeltaLineChart({ title, points, tolerance, height = 240 }) {
  const w = 980;
  const pad = { l: 38, r: 16, t: 26, b: 28 };

  const safeTol = Number.isFinite(tolerance) ? Math.max(0, tolerance) : 0;
  const yMax = Math.max(safeTol + 6, 25);

  const xMin = 0;
  const xMax = Math.max(1, ...points.map((p) => Number(p.xIndex) || 0));

  const xScale = (x) => pad.l + ((x - xMin) / (xMax - xMin || 1)) * (w - pad.l - pad.r);
  const yScale = (y) => pad.t + ((yMax - y) / (2 * yMax)) * (height - pad.t - pad.b);

  const sevColor = (sev) => {
    if (sev === "red") return "rgba(239,68,68,0.95)";
    if (sev === "yellow") return "rgba(245,158,11,0.95)";
    if (sev === "ok") return "rgba(34,197,94,0.95)";
    return "rgba(255,255,255,0.55)";
  };

  const pathFor = (kind) => {
    const pts = points
      .filter((p) => Number.isFinite(p[kind]))
      .sort((a, b) => (a.xIndex || 0) - (b.xIndex || 0));
    if (!pts.length) return "";
    return pts
      .map((p, i) => {
        const x = xScale(p.xIndex || 0);
        const y = yScale(p[kind]);
        return `${i === 0 ? "M" : "L"} ${x.toFixed(2)} ${y.toFixed(2)}`;
      })
      .join(" ");
  };

  const gridLines = [];
  for (let i = -yMax; i <= yMax; i += 5) {
    gridLines.push(i);
  }

  return (
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel, overflow: "hidden" }}>
      <div style={{ padding: 10, borderBottom: `1px solid ${theme.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ fontWeight: 950 }}>{title}</div>
        <div style={{ opacity: 0.75, fontSize: 12 }}>Δ(mm) vs nominal • green ≤ 4mm • yellow &gt; 4mm • red ≥ tolerance</div>
      </div>
      <div style={{ padding: 10, overflowX: "auto", maxWidth: "100%" }}>
        <svg viewBox={`0 0 ${w} ${height}`} style={{ width: "100%", height: "auto", display: "block" }}>
          {/* Grid */}
          {gridLines.map((g) => (
            <line
              key={g}
              x1={pad.l}
              x2={w - pad.r}
              y1={yScale(g)}
              y2={yScale(g)}
              stroke="rgba(255,255,255,0.07)"
              strokeWidth={1}
            />
          ))}
          {/* Axis */}
          <line x1={pad.l} x2={w - pad.r} y1={yScale(0)} y2={yScale(0)} stroke="rgba(255,255,255,0.22)" strokeWidth={1} />
          <line x1={pad.l} x2={pad.l} y1={pad.t} y2={height - pad.b} stroke="rgba(255,255,255,0.18)" strokeWidth={1} />

          {/* Tolerance bands */}
          {safeTol > 0 ? (
            <>
              <rect
                x={pad.l}
                y={yScale(safeTol)}
                width={w - pad.l - pad.r}
                height={yScale(4) - yScale(safeTol)}
                fill="rgba(245,158,11,0.08)"
              />
              <rect
                x={pad.l}
                y={yScale(-4)}
                width={w - pad.l - pad.r}
                height={yScale(-safeTol) - yScale(-4)}
                fill="rgba(245,158,11,0.08)"
              />
              <rect
                x={pad.l}
                y={yScale(4)}
                width={w - pad.l - pad.r}
                height={yScale(-4) - yScale(4)}
                fill="rgba(34,197,94,0.06)"
              />
            </>
          ) : (
            <rect x={pad.l} y={yScale(4)} width={w - pad.l - pad.r} height={yScale(-4) - yScale(4)} fill="rgba(34,197,94,0.06)" />
          )}

          {/* Lines */}
          <path d={pathFor("before")} fill="none" stroke="rgba(148,163,184,0.85)" strokeWidth={2} strokeDasharray="6 4" />
          <path d={pathFor("after")} fill="none" stroke="rgba(59,130,246,0.90)" strokeWidth={2.5} />

          {/* Points */}
          {points
            .filter((p) => Number.isFinite(p.after))
            .map((p) => (
              <circle
                key={`${p.id}-a`}
                cx={xScale(p.xIndex || 0)}
                cy={yScale(p.after)}
                r={4}
                fill={sevColor(p.sevAfter)}
                stroke="rgba(0,0,0,0.4)"
                strokeWidth={1}
              >
                <title>{`${p.line}: after Δ=${Math.round(p.after)}mm`}</title>
              </circle>
            ))}

          {/* Labels */}
          <text x={8} y={14} fill="rgba(255,255,255,0.70)" fontSize={11} fontWeight={900}>
            Δ(mm)
          </text>
        </svg>
      </div>
    </div>
  );
}

function WingProfileChart({ groupStats, tolerance, height = 260 }) {
  const w = 980;
  const pad = { l: 40, r: 16, t: 26, b: 28 };

  const safeTol = Number.isFinite(tolerance) ? Math.max(0, tolerance) : 0;
  const yMax = Math.max(safeTol + 6, 25);

  const groups = Array.from(new Set(groupStats.map((g) => g.groupName))).sort((a, b) => String(a).localeCompare(String(b)));
  const xMin = 0;
  const xMax = Math.max(1, groups.length - 1);

  const xScale = (i) => pad.l + ((i - xMin) / (xMax - xMin || 1)) * (w - pad.l - pad.r);
  const yScale = (y) => pad.t + ((yMax - y) / (2 * yMax)) * (height - pad.t - pad.b);

  const sevColor = (delta) => {
    if (!Number.isFinite(delta)) return "rgba(255,255,255,0.55)";
    const ad = Math.abs(delta);
    if (safeTol > 0 && ad >= safeTol) return "rgba(239,68,68,0.95)";
    if (ad > 4) return "rgba(245,158,11,0.95)";
    return "rgba(34,197,94,0.95)";
  };

  const seriesFor = (side) => {
    const pts = [];
    groups.forEach((gName, i) => {
      const rec = groupStats.find((r) => r.groupName === gName && r.side === side);
      pts.push({ i, y: rec?.after ?? null, before: rec?.before ?? null, groupName: gName });
    });
    return pts;
  };

  const pathFor = (pts, key) => {
    const filtered = pts.filter((p) => Number.isFinite(p[key]));
    if (!filtered.length) return "";
    return filtered
      .map((p, idx) => {
        const x = xScale(p.i);
        const y = yScale(p[key]);
        return `${idx === 0 ? "M" : "L"} ${x.toFixed(2)} ${y.toFixed(2)}`;
      })
      .join(" ");
  };

  const L = seriesFor("L");
  const R = seriesFor("R");

  return (
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel, overflow: "hidden" }}>
      <div style={{ padding: 10, borderBottom: `1px solid ${theme.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ fontWeight: 950 }}>Δ per maillon group (After)</div>
        <div style={{ opacity: 0.75, fontSize: 12 }}>Green ≤ 4mm • Yellow &gt; 4mm • Red ≥ tolerance</div>
      </div>
      <div style={{ padding: 10, overflowX: "auto", maxWidth: "100%" }}>
        <svg viewBox={`0 0 ${w} ${height}`} style={{ width: "100%", height: "auto", display: "block" }}>
          {/* Axis */}
          <line x1={pad.l} x2={w - pad.r} y1={yScale(0)} y2={yScale(0)} stroke="rgba(255,255,255,0.22)" strokeWidth={1} />
          <line x1={pad.l} x2={pad.l} y1={pad.t} y2={height - pad.b} stroke="rgba(255,255,255,0.18)" strokeWidth={1} />

          {/* Bands */}
          {safeTol > 0 ? (
            <>
              <rect x={pad.l} y={yScale(4)} width={w - pad.l - pad.r} height={yScale(-4) - yScale(4)} fill="rgba(34,197,94,0.06)" />
              <rect x={pad.l} y={yScale(safeTol)} width={w - pad.l - pad.r} height={yScale(4) - yScale(safeTol)} fill="rgba(245,158,11,0.08)" />
              <rect x={pad.l} y={yScale(-4)} width={w - pad.l - pad.r} height={yScale(-safeTol) - yScale(-4)} fill="rgba(245,158,11,0.08)" />
            </>
          ) : (
            <rect x={pad.l} y={yScale(4)} width={w - pad.l - pad.r} height={yScale(-4) - yScale(4)} fill="rgba(34,197,94,0.06)" />
          )}

          {/* Lines */}
          <path d={pathFor(L, "after")} fill="none" stroke="rgba(59,130,246,0.90)" strokeWidth={2.5} />
          <path d={pathFor(R, "after")} fill="none" stroke="rgba(168,85,247,0.90)" strokeWidth={2.5} />

          {/* Points + labels */}
          {groups.map((gName, i) => {
            const x = xScale(i);
            const lRec = groupStats.find((r) => r.groupName === gName && r.side === "L");
            const rRec = groupStats.find((r) => r.groupName === gName && r.side === "R");
            const yL = Number.isFinite(lRec?.after) ? yScale(lRec.after) : null;
            const yR = Number.isFinite(rRec?.after) ? yScale(rRec.after) : null;

            return (
              <g key={gName}>
                <text x={x} y={height - 10} textAnchor="middle" fill="rgba(255,255,255,0.60)" fontSize={10} fontWeight={900}>
                  {gName}
                </text>
                {yL != null ? (
                  <circle cx={x - 6} cy={yL} r={4.2} fill={sevColor(lRec.after)} stroke="rgba(0,0,0,0.45)" strokeWidth={1}>
                    <title>{`${gName} L: Δ=${Math.round(lRec.after)}mm`}</title>
                  </circle>
                ) : null}
                {yR != null ? (
                  <circle cx={x + 6} cy={yR} r={4.2} fill={sevColor(rRec.after)} stroke="rgba(0,0,0,0.45)" strokeWidth={1}>
                    <title>{`${gName} R: Δ=${Math.round(rRec.after)}mm`}</title>
                  </circle>
                ) : null}
              </g>
            );
          })}

          <text x={8} y={14} fill="rgba(255,255,255,0.70)" fontSize={11} fontWeight={900}>
            Δ(mm)
          </text>
          <text x={pad.l + 6} y={pad.t + 14} fill="rgba(59,130,246,0.90)" fontSize={11} fontWeight={900}>
            L
          </text>
          <text x={pad.l + 22} y={pad.t + 14} fill="rgba(168,85,247,0.90)" fontSize={11} fontWeight={900}>
            R
          </text>
        </svg>
      </div>
    </div>
  );
}