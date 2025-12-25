import BUILTIN_PROFILES from "./wingProfiles.json";

import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";


/**
 * Paraglider Trim Tuning — stable “patch A” build
 * - Step 3 uses GROUP loops (AR1 affects A1–A4 etc)
 * - Step 4 compact measurement tables
 * - Keeps legacy per-line loops as fallback (won’t break older saved sessions)
 */

const APP_VERSION = "0.2.2-patchE";

/* ------------------------- Built-in profiles ------------------------- */


/* ------------------------- Helpers ------------------------- */

function n(x) {
  const v = parseFloat(String(x ?? "").replace(",", "."));
  return Number.isFinite(v) ? v : null;
}

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

function isWideFormat(_grid) {
  // Keep for backwards compatibility, but don't block imports anymore.
  return true;
}

function parseWideFlexible(grid) {
  // 1) Try to detect meta header row (optional)
  let headerRow = -1;
  let inputCol = 0;
  let tolCol = -1;
  let corrCol = -1;

  const maxScan = Math.min(20, grid.length);
  for (let r = 0; r < maxScan; r++) {
    const row = grid[r] || [];
    for (let c = 0; c < row.length; c++) {
      const t = String(row[c] ?? "").toLowerCase();
      if (!t) continue;

      if (t.includes("eingabe") || t.includes("input")) {
        headerRow = r;
        inputCol = c;
      }
      if (t.includes("toleranz") || t.includes("tolerance")) {
        headerRow = r;
        tolCol = c;
      }
      if (t.includes("korrektur") || t.includes("correction")) {
        headerRow = r;
        corrCol = c;
      }
    }
    // if we found at least tolerance or correction headers, good enough
    if (headerRow >= 0 && (tolCol >= 0 || corrCol >= 0)) break;
  }

  const metaRow = headerRow >= 0 ? headerRow + 1 : 1;
  const metaValues = grid[metaRow] || [];

  const meta = {
    input1: String(metaValues[inputCol] ?? ""),
    input2: String(metaValues[inputCol + 1] ?? ""),
    tolerance: tolCol >= 0 ? (n(metaValues[tolCol]) ?? 0) : 0,
    correction: corrCol >= 0 ? (n(metaValues[corrCol]) ?? 0) : 0,
  };

  // 2) Parse rows by scanning for line IDs like A1, B12, C03, D7
  const rows = [];
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    const entry = { A: null, B: null, C: null, D: null };

    for (let c = 0; c <= row.length - 4; c++) {
      const cell = String(row[c] ?? "").trim();
      const m = cell.match(/^([A-Da-d])\s*0*([0-9]+)$/);
      if (!m) continue;

      const letter = m[1].toUpperCase();
      const line = `${letter}${parseInt(m[2], 10)}`;

      const nominal = n(row[c + 1]);
      const measL = n(row[c + 2]);
      const measR = n(row[c + 3]);

      entry[letter] = { line, nominal, measL, measR };

      // Skip forward a bit; typical layout is 4-wide blocks
      c += 3;
    }

    if (entry.A || entry.B || entry.C || entry.D) rows.push(entry);
  }

  return { meta, rows };
}


// Wide format parser (A/B/C/D blocks)
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
      entry[b.k] = {
        line,
        nominal: n(row[b.i + 1]),
        measL: n(row[b.i + 2]),
        measR: n(row[b.i + 3]),
      };
    }
    if (entry.A || entry.B || entry.C || entry.D) rows.push(entry);
  }

  return { meta, rows };
}

function makeProfileNameFromMeta(meta) {
  const a = String(meta?.input1 || "").trim();
  const b = String(meta?.input2 || "").trim();
  const combined = `${a} ${b}`.trim().replace(/\s+/g, " ");
  return combined || "Imported Wing";
}

function parseLineId(lineId) {
  const m = String(lineId || "")
    .trim()
    .match(/^([A-Za-z])\s*0*([0-9]+)$/);
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
  // fallback to mapping if nothing in file
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

function deltaMm({ nominal, measured, correction, adjustment }) {
  if (nominal == null || measured == null) return null;
  return measured + (correction || 0) + (adjustment || 0) - nominal;
}

function severity(delta, tolerance) {
  if (!Number.isFinite(delta)) return "none";
  const a = Math.abs(delta);
  const tol = tolerance || 0;
  if (tol <= 0) return "ok";
  const warnBand = Math.max(0, tol - 3); // yellow band within 3mm of tolerance
  if (a >= tol) return "red";
  if (a >= warnBand) return "yellow";
  return "ok";
}

function avg(nums) {
  const v = nums.filter((x) => Number.isFinite(x));
  if (!v.length) return null;
  return v.reduce((a, b) => a + b, 0) / v.length;
}

function getAdjustment(adjustments, groupName, side) {
  const key = `${groupName}|${side}`;
  return Number.isFinite(adjustments[key]) ? adjustments[key] : 0;
}

/* ------------------------- App ------------------------- */

export default function App() {
  const [step, setStep] = useState(() => {
    const s = localStorage.getItem("workflowStep");
    const v = parseInt(s || "1", 10);
    return Number.isFinite(v) ? Math.min(4, Math.max(1, v)) : 1;
  });
  useEffect(() => localStorage.setItem("workflowStep", String(step)), [step]);

  const [meta, setMeta] = useState({ input1: "", input2: "", tolerance: 0, correction: 0 });
  const [wideRows, setWideRows] = useState([]);

  // Profiles JSON (persisted)
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

 const [profileKey, setProfileKey] = useState(() => Object.keys(BUILTIN_PROFILES)[0] || "");

  const activeProfile =
    profiles[profileKey] || Object.values(profiles)[0] || Object.values(BUILTIN_PROFILES)[0];

  // Adjustments (per group)
  const [adjustments, setAdjustments] = useState(() => {
    try {
      const s = localStorage.getItem("groupAdjustments");
      return s ? JSON.parse(s) : {};
    } catch {
      return {};
    }
  });
  function persistAdjustments(next) {
    setAdjustments(next);
    localStorage.setItem("groupAdjustments", JSON.stringify(next));
  }

  // Loop types
  const [loopTypes, setLoopTypes] = useState(() => {
    try {
      const s = localStorage.getItem("loopTypes");
      return s
        ? JSON.parse(s)
        : { SL: 0, DL: -7, AS: -10, "AS+": -16, PH: -18, "LF++": -23 };
    } catch {
      return { SL: 0, DL: -7, AS: -10, "AS+": -16, PH: -18, "LF++": -23 };
    }
  });
  function persistLoopTypes(next) {
    setLoopTypes(next);
    localStorage.setItem("loopTypes", JSON.stringify(next));
  }

  // Legacy per-line loop setup (kept so old sessions don’t break)
  const [loopSetup, setLoopSetup] = useState(() => {
    try {
      const s = localStorage.getItem("loopSetup");
      return s ? JSON.parse(s) : {};
    } catch {
      return {};
    }
  });
  function persistLoopSetup(next) {
    setLoopSetup(next);
    localStorage.setItem("loopSetup", JSON.stringify(next));
  }

  // NEW: group loop setup (AR1|L -> "SL")
  const [groupLoopSetup, setGroupLoopSetup] = useState(() => {
    try {
      const s = localStorage.getItem("groupLoopSetup");
      return s ? JSON.parse(s) : {};
    } catch {
      return {};
    }
  });
  function persistGroupLoopSetup(next) {
    setGroupLoopSetup(next);
    localStorage.setItem("groupLoopSetup", JSON.stringify(next));
  }

  const fileInputRef = useRef(null);
  const [selectedFileName, setSelectedFileName] = useState("");

  const hasCSV = wideRows.length > 0;

  // Derived: lines + groups
  const allLines = useMemo(() => getAllLinesFromWide(wideRows), [wideRows]);
  const allGroupNames = useMemo(() => extractGroupNames(wideRows, activeProfile), [wideRows, activeProfile]);

  // Build group -> lines list
  const groupToLines = useMemo(() => {
    const map = new Map();
    for (const { lineId } of allLines) {
      const g = groupForLine(activeProfile, lineId);
      if (!g) continue;
      if (!map.has(g)) map.set(g, []);
      map.get(g).push(lineId);
    }
    for (const [k, arr] of map.entries()) {
      arr.sort((a, b) => {
        const pa = parseLineId(a);
        const pb = parseLineId(b);
        if (!pa || !pb) return a.localeCompare(b);
        if (pa.prefix !== pb.prefix) return pa.prefix.localeCompare(pb.prefix);
        return pa.num - pb.num;
      });
      map.set(k, arr);
    }
    return map;
  }, [allLines, activeProfile]);

  // Auto profile name from CSV (A2 + B2)
  const csvProfileName = useMemo(() => makeProfileNameFromMeta(meta), [meta]);

  // Ensure profile exists for imported name
  function ensureProfileExistsByName(name) {
    const key = String(name || "").trim();
    if (!key) return;

    if (profiles[key]) {
      setProfileKey(key);
      return;
    }

    const nextProfiles = { ...profiles };
    const base = profiles[profileKey] || activeProfile || Object.values(BUILTIN_PROFILES)[0];
    const clone = JSON.parse(JSON.stringify(base));
    clone.name = key;
    nextProfiles[key] = clone;

    const json = JSON.stringify(nextProfiles, null, 2);
    setProfileJson(json);
    localStorage.setItem("wingProfilesJson", json);
    setProfileKey(key);
  }

function onImportFile(file) {
  const name = (file?.name || "").toLowerCase();

  // XLSX (Excel)
if (file.name.toLowerCase().endsWith(".xlsx")) {
  const reader = new FileReader();

  reader.onload = () => {
    try {
      const data = reader.result;

      // Read workbook
      const workbook = XLSX.read(data, { type: "array" });

      // Use first sheet
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // Convert to 2D grid (rows x columns)
      const grid = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        raw: false,
        defval: "",
      });

      // FLEXIBLE parsing (no header assumptions)
      const w = parseWideFlexible(grid);

      if (!w.rows.length) {
        alert(
          "Excel imported, but no line rows were detected.\n\n" +
          "Check that line IDs look like A1, B12, C03 etc."
        );
        return;
      }

      setMeta(w.meta);
      setWideRows(w.rows);

      const importName = makeProfileNameFromMeta(w.meta);
      ensureProfileExistsByName(importName);

      setSelectedFileName(file.name);
      setStep(2);
    } catch (err) {
      console.error(err);
      alert(
        "Failed to read Excel file.\n\n" +
        "Make sure it is a .xlsx file in the same layout as the CSV."
      );
    }
  };

  reader.readAsArrayBuffer(file);
  return;
}


// CSV (existing)
const reader = new FileReader();
reader.onload = () => {
  const text = String(reader.result || "");
  const parsed = parseDelimited(text);

  const w = parseWideFlexible(parsed.grid);
  if (!w.rows.length) {
    alert("File imported, but no line rows were detected. Make sure line IDs look like A1, B12, C03 etc.");
    return;
  }

  setMeta(w.meta);
  setWideRows(w.rows);

  const importName = makeProfileNameFromMeta(w.meta);
  ensureProfileExistsByName(importName);

  setSelectedFileName(file.name);
  setStep(2);
};
reader.readAsText(file);
}



  // Group-based loop delta
  function loopDeltaFor(lineId, side) {
    // Prefer GROUP loop setup if possible
    const g = groupForLine(activeProfile, lineId);
    if (g) {
      const key = `${g}|${side}`;
      const type = groupLoopSetup[key] || "SL";
      const v = loopTypes?.[type];
      return Number.isFinite(v) ? v : 0;
    }

    // Fallback legacy per-line loop setup
    const legacyKey = `${lineId}|${side}`;
    const legacyType = loopSetup?.[legacyKey] || "SL";
    const lv = loopTypes?.[legacyType];
    return Number.isFinite(lv) ? lv : 0;
  }

  // Bulk tools now operate on GROUP loops
  function applyAllSL() {
    const next = {};
    for (const g of allGroupNames) {
      next[`${g}|L`] = "SL";
      next[`${g}|R`] = "SL";
    }
    persistGroupLoopSetup(next);
  }

  function mirrorLtoR() {
    const next = { ...groupLoopSetup };
    for (const g of allGroupNames) {
      next[`${g}|R`] = next[`${g}|L`] || "SL";
    }
    persistGroupLoopSetup(next);
  }

  function mirrorRtoL() {
    const next = { ...groupLoopSetup };
    for (const g of allGroupNames) {
      next[`${g}|L`] = next[`${g}|R`] || "SL";
    }
    persistGroupLoopSetup(next);
  }

  function resetAdjustments() {
    persistAdjustments({});
  }

  // Measurement table blocks (compact)
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

  // Suggestions / group stats (after loops + adjustments)
  const computed = useMemo(() => {
    const corr = meta.correction || 0;

    const bucket = new Map(); // group|side -> [delta]
    for (const r of wideRows) {
      for (const letter of ["A", "B", "C", "D"]) {
        const b = r[letter];
        if (!b || !b.line || b.nominal == null) continue;
        const g = groupForLine(activeProfile, b.line) || `${letter}?`;

        const loopL = loopDeltaFor(b.line, "L");
        const loopR = loopDeltaFor(b.line, "R");

        const effL = b.measL == null ? null : b.measL + loopL;
        const effR = b.measR == null ? null : b.measR + loopR;

        const adjL = getAdjustment(adjustments, g, "L");
        const adjR = getAdjustment(adjustments, g, "R");

        const dL = deltaMm({ nominal: b.nominal, measured: effL, correction: corr, adjustment: adjL });
        const dR = deltaMm({ nominal: b.nominal, measured: effR, correction: corr, adjustment: adjR });

        if (Number.isFinite(dL)) {
          const k = `${g}|L`;
          if (!bucket.has(k)) bucket.set(k, []);
          bucket.get(k).push(dL);
        }
        if (Number.isFinite(dR)) {
          const k = `${g}|R`;
          if (!bucket.has(k)) bucket.set(k, []);
          bucket.get(k).push(dR);
        }
      }
    }

    const groupStats = [];
    for (const [key, arr] of bucket.entries()) {
      const [groupName, side] = key.split("|");
      const mean = avg(arr);
      if (!Number.isFinite(mean)) continue;
      groupStats.push({ groupName, side, meanDelta: mean });
    }
    groupStats.sort((a, b) =>
      (groupSortKey(a.groupName) + a.side).localeCompare(groupSortKey(b.groupName) + b.side)
    );

    return { groupStats };
  }, [wideRows, meta.correction, activeProfile, adjustments, groupLoopSetup, loopSetup, loopTypes]);

  // UI styling
  const page = {
    minHeight: "100vh",
    background: "#0b0c10",
    color: "#eef1ff",
    fontFamily: "system-ui, sans-serif",
  };
  const wrap = { maxWidth: 1200, margin: "0 auto", padding: 16, display: "flex", flexDirection: "column", gap: 12 };
  const card = { border: "1px solid #2a2f3f", borderRadius: 14, background: "#11131a", padding: 12 };
  const muted = { color: "#aab1c3" };
  const btn = {
    padding: "10px 12px",
    borderRadius: 10,
    border: "1px solid #2a2f3f",
    background: "#0d0f16",
    color: "#eef1ff",
    cursor: "pointer",
    fontWeight: 650,
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
  const redCell = { border: "1px solid rgba(255,107,107,0.85)", background: "rgba(255,107,107,0.14)" };
  const yellowCell = { border: "1px solid rgba(255,214,102,0.95)", background: "rgba(255,214,102,0.14)" };

  // Step guard
  useEffect(() => {
    if (step > 1 && !hasCSV) setStep(1);
  }, [step, hasCSV]);

  return (
    <div style={page}>
      <div style={wrap}>
        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
          <div>
            <h1 style={{ margin: 0, fontSize: 22 }}>
              Paraglider Trim Tuning{" "}
              <span style={{ ...muted, fontSize: 12, fontWeight: 700 }}>v{APP_VERSION}</span>
            </h1>
            <div style={{ ...muted, fontSize: 12, marginTop: 6 }}>
              Red: |Δ| ≥ tolerance. Yellow: within 3mm of tolerance.
            </div>
          </div>
          <div style={{ ...muted, fontSize: 12 }}>
            Profile name (from CSV A2+B2):{" "}
            <b style={{ color: "#eef1ff" }}>{csvProfileName}</b>
          </div>
        </div>

        {/* Safety */}
        <div style={{ ...card, borderColor: "rgba(255,204,102,0.5)", background: "rgba(255,204,102,0.08)" }}>
          <b>Safety notice:</b> This is an analysis/simulation tool. Trimming can be dangerous and may invalidate certification.
          Always follow manufacturer/check-center procedures and re-measure after any change.
        </div>

        {/* Workflow Stepper */}
        <div style={card}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", justifyContent: "space-between" }}>
            <div style={{ fontWeight: 900 }}>Workflow</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <StepButton current={step} num={1} setStep={setStep} enabled={true} label="1) Import CSV" />
              <StepButton current={step} num={2} setStep={setStep} enabled={hasCSV} label="2) Wing layout" />
              <StepButton current={step} num={3} setStep={setStep} enabled={hasCSV} label="3) Loops setup" />
              <StepButton current={step} num={4} setStep={setStep} enabled={hasCSV && allGroupNames.length > 0} label="4) Trim tables" />
            </div>
          </div>
          <div style={{ ...muted, fontSize: 12, marginTop: 10 }}>
            Tip: do Step 2–3 before trimming so “before” matches the real baseline.
          </div>
        </div>

        {/* STEP 1 */}
        {step === 1 ? (
          <div style={card}>
            <div style={{ fontWeight: 900, marginBottom: 8 }}>Step 1 — Import measurement CSV</div>
            <div style={{ ...muted, fontSize: 12, lineHeight: 1.5 }}>
              Upload your measurement file (wide layout with A/B/C/D blocks). The paraglider name is read from cells <b>A2</b> + <b>B2</b>.
            </div>

            <div style={{ height: 10 }} />

            <input
              ref={fileInputRef}
              type="file"
              accept=".csv,.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              style={{ display: "none" }}
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) {
                  setSelectedFileName(f.name);
                  onImportFile(f);
                }
                e.target.value = "";
              }}
            />

            <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
              <button style={btnWarn} onClick={() => fileInputRef.current?.click()}>
                Choose CSV…
              </button>

              <div style={{ ...muted, fontSize: 12 }}>
                {selectedFileName ? (
                  <>
                    Selected: <b style={{ color: "#eef1ff" }}>{selectedFileName}</b>
                  </>
                ) : (
                  <>No file selected.</>
                )}
              </div>
            </div>
          </div>
        ) : null}

        {/* STEP 2 */}
        {step === 2 ? (
          <div style={card}>
            <div style={{ fontWeight: 900, marginBottom: 8 }}>Step 2 — Wing layout (profile mapping)</div>
            <div style={{ ...muted, fontSize: 12, lineHeight: 1.5 }}>
              Choose the wing profile mapping so the app understands your diagram groupings (AR1/BR2/etc).
              The imported CSV name creates/chooses a matching profile automatically.
            </div>

            <div style={{ height: 10 }} />

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
                <div style={{ fontWeight: 850, marginBottom: 10 }}>Select profile</div>

                <label style={{ ...muted, fontSize: 12 }}>Profile</label>
                <select
                  value={profileKey}
                  onChange={(e) => setProfileKey(e.target.value)}
                  style={{ ...input, padding: "10px 10px", marginTop: 6 }}
                >
                  {Object.keys(profiles).sort((a, b) => a.localeCompare(b)).map((k) => (
                    <option key={k} value={k}>
                      {k}
                    </option>
                  ))}
                </select>

                <div style={{ height: 10 }} />
                <label style={{ ...muted, fontSize: 12 }}>mm per loop (target step size)</label>
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
                  style={{ ...input, marginTop: 6 }}
                  inputMode="numeric"
                />

                <div style={{ height: 10 }} />
                <div style={{ ...muted, fontSize: 12 }}>
                  Groups detected: <b style={{ color: "#eef1ff" }}>{allGroupNames.length}</b>
                </div>
              </div>

              <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
                <div style={{ fontWeight: 850, marginBottom: 8 }}>Edit / add wing profiles (JSON)</div>
                <div style={{ ...muted, fontSize: 12, marginBottom: 10 }}>
                  Advanced: edit mappings here if your wing has different group ranges.
                </div>
                <textarea
                  value={profileJson}
                  onChange={(e) => {
                    setProfileJson(e.target.value);
                    localStorage.setItem("wingProfilesJson", e.target.value);
                  }}
                  style={{
                    width: "100%",
                    minHeight: 240,
                    borderRadius: 12,
                    border: "1px solid #2a2f3f",
                    background: "#0d0f16",
                    color: "#eef1ff",
                    padding: 10,
                    fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                    fontSize: 12,
                    outline: "none",
                  }}
                />
                <div style={{ ...muted, fontSize: 12, marginTop: 10 }}>
                  Hint: mapping ranges should match the diagram labels. Example: A 1–4 → AR1.
                </div>
              </div>
            </div>

            <div style={{ marginTop: 12, display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button onClick={() => setStep(3)} style={btnWarn} disabled={!hasCSV}>
                Continue to Step 3 (Loops)
              </button>
              <button onClick={() => setStep(1)} style={btn}>
                Back
              </button>
            </div>
          </div>
        ) : null}

        {/* STEP 3 */}
        {step === 3 ? (
          <div style={card}>
            <div style={{ fontWeight: 900, marginBottom: 8 }}>Step 3 — Maillon loop setup (baseline)</div>
            <div style={{ ...muted, fontSize: 12, lineHeight: 1.5 }}>
              Set which loop type is installed on each <b>line group</b> maillon (Left/Right). Changing AR1 affects A1–A4 etc.
              This defines the real “Before trimming” baseline.
            </div>

            <div style={{ height: 10 }} />

            {/* Loop types */}
            <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
              <div style={{ fontWeight: 850, marginBottom: 8 }}>Loop types (editable)</div>
              <div style={{ color: "#aab1c3", fontSize: 12, marginBottom: 10 }}>
                Negative means the loop reduces line length (effective measured = measured + loopDelta).
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 10 }}>
                {Object.entries(loopTypes).map(([name, mm]) => (
                  <div
                    key={name}
                    style={{
                      border: "1px solid #2a2f3f",
                      borderRadius: 14,
                      padding: 10,
                      background: "#0d0f16",
                    }}
                  >
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

              <div style={{ height: 12 }} />

              <div style={{ padding: 12, borderRadius: 14, border: "1px solid #2a2f3f", background: "#0b0c10" }}>
                <div style={{ fontWeight: 850, marginBottom: 8 }}>Quick tools</div>

                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <button onClick={applyAllSL} style={btn}>All SL</button>
                  <button onClick={mirrorLtoR} style={btn}>Mirror L → R</button>
                  <button onClick={mirrorRtoL} style={btn}>Mirror R → L</button>
                </div>
              </div>
            </div>

            <div style={{ height: 12 }} />

            {/* GROUP loop setup table */}
            <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
              <div style={{ fontWeight: 850, marginBottom: 8 }}>Loops installed per line group</div>

              {!allGroupNames.length ? (
                <div style={{ ...muted, fontSize: 12 }}>No groups found. Check Step 2 mapping.</div>
              ) : (
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
                    <thead>
                      <tr style={{ color: "#aab1c3", fontSize: 12 }}>
                        <th style={{ textAlign: "left", padding: "8px 8px" }}>Group</th>
                        <th style={{ textAlign: "left", padding: "8px 8px" }}>Lines included</th>
                        <th style={{ textAlign: "left", padding: "8px 8px" }}>Left loop</th>
                        <th style={{ textAlign: "right", padding: "8px 8px" }}>Δ (mm)</th>
                        <th style={{ textAlign: "left", padding: "8px 8px" }}>Right loop</th>
                        <th style={{ textAlign: "right", padding: "8px 8px" }}>Δ (mm)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {allGroupNames.map((g) => {
                        const kL = `${g}|L`;
                        const kR = `${g}|R`;
                        const tL = groupLoopSetup[kL] || "SL";
                        const tR = groupLoopSetup[kR] || "SL";
                        const dL = Number.isFinite(loopTypes?.[tL]) ? loopTypes[tL] : 0;
                        const dR = Number.isFinite(loopTypes?.[tR]) ? loopTypes[tR] : 0;
                        const lines = groupToLines.get(g) || [];

                        return (
                          <tr key={g} style={{ borderTop: "1px solid rgba(42,47,63,0.9)" }}>
                            <td style={{ padding: "8px 8px", fontWeight: 900 }}>{g}</td>

                            <td style={{ padding: "8px 8px", color: "#aab1c3", fontSize: 12 }}>
                              {lines.length ? lines.join(", ") : "—"}
                            </td>

                            <td style={{ padding: "8px 8px" }}>
                              <select
                                value={tL}
                                onChange={(e) =>
                                  persistGroupLoopSetup({ ...groupLoopSetup, [kL]: e.target.value })
                                }
                                style={{
                                  width: 140,
                                  borderRadius: 10,
                                  border: "1px solid #2a2f3f",
                                  background: "#0d0f16",
                                  color: "#eef1ff",
                                  padding: "8px 10px",
                                  outline: "none",
                                }}
                              >
                                {Object.keys(loopTypes).map((name) => (
                                  <option key={name} value={name}>
                                    {name}
                                  </option>
                                ))}
                              </select>
                            </td>

                            <td
                              style={{
                                padding: "8px 8px",
                                textAlign: "right",
                                fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                                color: "#aab1c3",
                              }}
                            >
                              {dL > 0 ? `+${dL}` : `${dL}`}
                            </td>

                            <td style={{ padding: "8px 8px" }}>
                              <select
                                value={tR}
                                onChange={(e) =>
                                  persistGroupLoopSetup({ ...groupLoopSetup, [kR]: e.target.value })
                                }
                                style={{
                                  width: 140,
                                  borderRadius: 10,
                                  border: "1px solid #2a2f3f",
                                  background: "#0d0f16",
                                  color: "#eef1ff",
                                  padding: "8px 10px",
                                  outline: "none",
                                }}
                              >
                                {Object.keys(loopTypes).map((name) => (
                                  <option key={name} value={name}>
                                    {name}
                                  </option>
                                ))}
                              </select>
                            </td>

                            <td
                              style={{
                                padding: "8px 8px",
                                textAlign: "right",
                                fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                                color: "#aab1c3",
                              }}
                            >
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

            <div style={{ marginTop: 12, display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button onClick={() => setStep(4)} style={btnWarn} disabled={!allGroupNames.length}>
                Continue to Step 4 (Tables)
              </button>
              <button onClick={() => setStep(2)} style={btn}>
                Back
              </button>
            </div>
          </div>
        ) : null}

        {/* STEP 4 */}
        {step === 4 ? (
          <div style={card}>
            <div style={{ fontWeight: 900, marginBottom: 8 }}>Step 4 — Measurement tables (compact)</div>
            <div style={{ ...muted, fontSize: 12, marginBottom: 10 }}>
              Inputs are Measured L/R from the CSV. Table shows Δ after group loops + adjustments.
              Columns are compact to avoid wasted space.
            </div>

            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 10 }}>
              <button onClick={resetAdjustments} style={btnDanger}>
                Reset all adjustments
              </button>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 12 }}>
              <BlockTable
                title="A"
                rows={compactBlocks.A}
                meta={meta}
                activeProfile={activeProfile}
                adjustments={adjustments}
                loopDeltaFor={loopDeltaFor}
                input={input}
                redCell={redCell}
                yellowCell={yellowCell}
                setCell={setCell}
                blockKey="A"
              />
              <BlockTable
                title="B"
                rows={compactBlocks.B}
                meta={meta}
                activeProfile={activeProfile}
                adjustments={adjustments}
                loopDeltaFor={loopDeltaFor}
                input={input}
                redCell={redCell}
                yellowCell={yellowCell}
                setCell={setCell}
                blockKey="B"
              />
              <BlockTable
                title="C"
                rows={compactBlocks.C}
                meta={meta}
                activeProfile={activeProfile}
                adjustments={adjustments}
                loopDeltaFor={loopDeltaFor}
                input={input}
                redCell={redCell}
                yellowCell={yellowCell}
                setCell={setCell}
                blockKey="C"
              />
              <BlockTable
                title="D"
                rows={compactBlocks.D}
                meta={meta}
                activeProfile={activeProfile}
                adjustments={adjustments}
                loopDeltaFor={loopDeltaFor}
                input={input}
                redCell={redCell}
                yellowCell={yellowCell}
                setCell={setCell}
                blockKey="D"
              />
            </div>

            <div style={{ height: 12 }} />

            <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
              <div style={{ fontWeight: 850, marginBottom: 6 }}>Group average Δ (after loops + adjustments)</div>
              <div style={{ ...muted, fontSize: 12, marginBottom: 10 }}>
                Useful for deciding which group to change at the risers/maillons.
              </div>

              {!computed.groupStats.length ? (
                <div style={{ ...muted, fontSize: 12 }}>No stats yet.</div>
              ) : (
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 700 }}>
                    <thead>
                      <tr style={{ color: "#aab1c3", fontSize: 12 }}>
                        <th style={{ textAlign: "left", padding: "6px 8px" }}>Group</th>
                        <th style={{ textAlign: "left", padding: "6px 8px" }}>Side</th>
                        <th style={{ textAlign: "right", padding: "6px 8px" }}>Mean Δ (mm)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {computed.groupStats.map((s) => (
                        <tr key={`${s.groupName}|${s.side}`} style={{ borderTop: "1px solid rgba(42,47,63,0.9)" }}>
                          <td style={{ padding: "6px 8px", fontWeight: 900 }}>{s.groupName}</td>
                          <td style={{ padding: "6px 8px" }}>{s.side}</td>
                          <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                            {Math.round(s.meanDelta)}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            <div style={{ marginTop: 12, display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button onClick={() => setStep(3)} style={btn}>
                Back to Step 3 (Loops)
              </button>
              <button onClick={() => setStep(2)} style={btn}>
                Back to Step 2 (Layout)
              </button>
            </div>
          </div>
        ) : null}
      </div>
    </div>
  );

  function StepButton({ current, num, setStep, enabled, label }) {
    const active = current === num;
    return (
      <button
        onClick={() => enabled && setStep(num)}
        disabled={!enabled}
        style={{
          padding: "8px 10px",
          borderRadius: 10,
          border: "1px solid #2a2f3f",
          background: active ? "rgba(176,132,255,0.14)" : "#0d0f16",
          color: active ? "#eef1ff" : enabled ? "#aab1c3" : "rgba(170,177,195,0.4)",
          cursor: enabled ? "pointer" : "not-allowed",
          fontWeight: 800,
          fontSize: 12,
        }}
        title={!enabled ? "Complete previous steps first" : ""}
      >
        {label}
      </button>
    );
  }
}

/* ------------------------- Compact measurement table ------------------------- */

function BlockTable({ title, rows, meta, activeProfile, adjustments, loopDeltaFor, input, redCell, yellowCell, setCell, blockKey }) {
  const corr = meta.correction || 0;
  const tol = meta.tolerance || 0;

  const styleFor = (sev) => (sev === "red" ? redCell : sev === "yellow" ? yellowCell : null);

  return (
    <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, overflow: "hidden", background: "#0e1018" }}>
      <div style={{ padding: 10, borderBottom: "1px solid #2a2f3f", fontWeight: 900 }}>{title} lines</div>

      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 520 }}>
          <thead>
            <tr style={{ color: "#aab1c3", fontSize: 12 }}>
              <th style={{ textAlign: "left", padding: "6px 8px" }}>Line</th>
              <th style={{ textAlign: "left", padding: "6px 8px" }}>Group</th>
              <th style={{ textAlign: "right", padding: "6px 8px" }}>Nom</th>
              <th style={{ textAlign: "right", padding: "6px 8px" }}>Meas L</th>
              <th style={{ textAlign: "right", padding: "6px 8px" }}>Meas R</th>
            </tr>
          </thead>

          <tbody>
            {!rows.length ? (
              <tr>
                <td colSpan={5} style={{ padding: "6px 8px", color: "#aab1c3" }}>
                  No {title} rows found.
                </td>
              </tr>
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
                    <td style={{ padding: "6px 8px" }}>
                      <b>{b.line}</b>
                    </td>
                    <td style={{ padding: "6px 8px", color: "#aab1c3", fontSize: 12 }}>{groupName}</td>
                    <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                      {b.nominal ?? ""}
                    </td>

                    <td style={{ padding: "6px 8px", textAlign: "right" }}>
                      <input
                        value={b.measL ?? ""}
                        onChange={(e) => setCell(b.rowIndex, blockKey, "measL", e.target.value)}
                        style={{
                          ...input,
                          ...(styleFor(sevL) || null),
                          width: 86,
                          padding: "6px 8px",
                          textAlign: "right",
                          fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                        }}
                        inputMode="numeric"
                      />
                      <div style={{ color: "#aab1c3", fontSize: 10, marginTop: 4, fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                        loop {loopL > 0 ? `+${loopL}` : `${loopL}`} | Δ {Number.isFinite(dL) ? `${dL > 0 ? "+" : ""}${Math.round(dL)}mm` : "–"}
                      </div>
                    </td>

                    <td style={{ padding: "6px 8px", textAlign: "right" }}>
                      <input
                        value={b.measR ?? ""}
                        onChange={(e) => setCell(b.rowIndex, blockKey, "measR", e.target.value)}
                        style={{
                          ...input,
                          ...(styleFor(sevR) || null),
                          width: 86,
                          padding: "6px 8px",
                          textAlign: "right",
                          fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                        }}
                        inputMode="numeric"
                      />
                      <div style={{ color: "#aab1c3", fontSize: 10, marginTop: 4, fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                        loop {loopR > 0 ? `+${loopR}` : `${loopR}`} | Δ {Number.isFinite(dR) ? `${dR > 0 ? "+" : ""}${Math.round(dR)}mm` : "–"}
                      </div>
                    </td>
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
