import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import BUILTIN_PROFILES from "./wingProfiles.json";

/**
 * Paraglider Trim Tuning
 * - CSV + XLSX import (flexible scan)
 * - Step 2: Guided profile editor (mapping)
 * - Step 3: Loops per LINE GROUP (AR1 affects A1–A4, etc.)
 * - Step 4: Compact measurement tables + Raw/Corrected toggle
 * - Graphs:
 *    (1) Δ line chart: Before vs After overlay + points + hover tooltip
 *    (2) Wing profile chart: Group averages laid out left/right (updates live)
 *
 * Math:
 *   corrected = rawMeasured + correction
 *   baseline  = corrected + loopDelta
 *   after     = baseline + adjustment
 *   delta     = after - nominal
 */

const APP_VERSION = "0.3.2Stable without front chart";

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

function makeProfileNameFromMeta(meta) {
  const a = String(meta?.input1 || "").trim();
  const b = String(meta?.input2 || "").trim();
  const combined = `${a} ${b}`.trim().replace(/\s+/g, " ");
  return combined || "Imported Wing";
}
function deltaMm({ nominal, measured, correction, adjustment }) {
  if (nominal == null || measured == null) return null;
  const corr = Number.isFinite(correction) ? correction : 0;
  const adj = Number.isFinite(adjustment) ? adjustment : 0;
  return measured + corr + adj - nominal;
}

function severity(delta, tolerance) {
  if (!Number.isFinite(delta)) return "none";
  const a = Math.abs(delta);
  const tol = Number.isFinite(tolerance) ? tolerance : 0;
  if (tol <= 0) return "ok";
  const warnBand = Math.max(0, tol - 3); // yellow within 3mm of tolerance
  if (a >= tol) return "red";
  if (a >= warnBand) return "yellow";
  return "ok";
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


function avg(nums) {
  const v = nums.filter((x) => Number.isFinite(x));
  if (!v.length) return null;
  return v.reduce((a, b) => a + b, 0) / v.length;
}

function getAdjustment(adjustments, groupName, side) {
  const key = `${groupName}|${side}`;
  return Number.isFinite(adjustments[key]) ? adjustments[key] : 0;
}

function downloadTextFile(filename, text) {
  const blob = new Blob([text], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function safeParseProfilesJson(text) {
  const obj = JSON.parse(text);
  if (!obj || typeof obj !== "object" || Array.isArray(obj)) {
    throw new Error("Profiles JSON must be an object of { profileName: { ... } }");
  }
  return obj;
}

/**
 * Flexible parser for BOTH CSV + XLSX:
 * - Attempts to read meta (Eingabe/Input + Toleranz/Tolerance + Korrektur/Correction) if present
 * - Reads measurement rows by scanning for line IDs like A1, B12, C03, D7 and reading next 3 cells:
 *      [LineId] [Soll] [Ist L] [Ist R]
 */
function parseWideFlexible(grid) {
  // 1) Meta header detection (optional)
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

  // 2) Measurement rows by scanning for line IDs
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
      c += 3;
    }

    if (entry.A || entry.B || entry.C || entry.D) rows.push(entry);
  }

  return { meta, rows };
}

/* -------- Zeroing wizard: median correction suggestion -------- */

function median(values) {
  const v = values.filter((x) => Number.isFinite(x)).slice().sort((a, b) => a - b);
  if (!v.length) return null;
  const mid = Math.floor(v.length / 2);
  if (v.length % 2) return v[mid];
  return (v[mid - 1] + v[mid]) / 2;
}

function suggestCorrectionFromWideRows(wideRows) {
  const offsets = [];
  for (const r of wideRows || []) {
    for (const letter of ["A", "B", "C", "D"]) {
      const b = r?.[letter];
      if (!b || b.nominal == null) continue;
      if (b.measL != null) offsets.push(b.nominal - b.measL);
      if (b.measR != null) offsets.push(b.nominal - b.measR);
    }
  }
  const med = median(offsets);
  return med == null ? null : Math.round(med);
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

  const [showCorrected, setShowCorrected] = useState(() => {
    const s = localStorage.getItem("showCorrected");
    return s ? s === "1" : true;
  });
  useEffect(() => localStorage.setItem("showCorrected", showCorrected ? "1" : "0"), [showCorrected]);

 
  // Profiles JSON (persisted)
  const [profileJson, setProfileJson] = useState(() => {
    const saved = localStorage.getItem("wingProfilesJson");
    return saved || JSON.stringify({ ...BUILTIN_PROFILES }, null, 2);
  });

  /* ===============================
     Loop → adjustment helper
     =============================== */

  function loopTypeFromAdjustment(mm) {
    if (!Number.isFinite(mm)) return "";
    for (const [name, val] of Object.entries(loopTypes)) {
      if (Number.isFinite(val) && val === mm) return name;
    }
    return ""; // Custom / manual value
  }

  /* ===============================
     rest of App logic
     =============================== */

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

  // Guided Profile Editor state
  const [isProfileEditorOpen, setIsProfileEditorOpen] = useState(false);
  const [draftProfileKey, setDraftProfileKey] = useState("");
  const [draftProfile, setDraftProfile] = useState({});
  const [showAdvancedJson, setShowAdvancedJson] = useState(false);

  useEffect(() => {
    setDraftProfileKey(profileKey || "");
    setDraftProfile(JSON.parse(JSON.stringify(profiles[profileKey] || activeProfile || {})));
  }, [profileKey, profileJson]); // refresh after edits/import

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

  // Group loop setup (AR1|L -> "SL")
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
  const profilesImportRef = useRef(null);
  const [selectedFileName, setSelectedFileName] = useState("");

  const hasCSV = wideRows.length > 0;

  const allLines = useMemo(() => getAllLinesFromWide(wideRows), [wideRows]);
  const allGroupNames = useMemo(() => extractGroupNames(wideRows, activeProfile), [wideRows, activeProfile]);

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
	
	  /* ===============================
     Pitch Trim calculation (A − D)
     =============================== */
  const pitchTrim = useMemo(() => {
    // If you later add filters, these will pick them up automatically.
    const rowIncluded = (L) =>
      typeof includedRows === "object" ? !!includedRows[L] : true;

    const groupIncluded = (g) => {
      if (!g) return false;
      if (typeof includedGroups !== "object") return true;
      const keys = Object.keys(includedGroups || {});
      if (keys.length === 0) return true; // treat empty map as "all selected"
      return !!includedGroups[g];
    };

    // If you have a "Show corrected" toggle, use it.
    const corr =
      typeof showCorrected === "undefined" || showCorrected
        ? meta?.correction ?? 0
        : 0;

    // Collect AFTER deltas by row (A/B/C/D)
    const perRow = { A: [], B: [], C: [], D: [] };

    for (const r of wideRows || []) {
      for (const letter of ["A", "B", "C", "D"]) {
        if (!rowIncluded(letter)) continue;

        const b = r?.[letter];
        if (!b?.line) continue;

        const lineId = b.line;
        const nominal = b.nominal;
        if (!Number.isFinite(nominal)) continue;

        const groupName = groupForLine(activeProfile, lineId);
        if (!groupIncluded(groupName)) continue;

        // LEFT side
        if (Number.isFinite(b.measL)) {
          const loopType = groupLoopSetup?.[`${groupName}|L`] || "SL";
          const loopDelta =
            Number.isFinite(loopTypes?.[loopType]) ? loopTypes[loopType] : 0;
          const adj = getAdjustment(adjustments, groupName, "L") || 0;

          const corrected = b.measL + corr;
          const afterDelta = corrected + loopDelta + adj - nominal;
          if (Number.isFinite(afterDelta)) perRow[letter].push(afterDelta);
        }

        // RIGHT side
        if (Number.isFinite(b.measR)) {
          const loopType = groupLoopSetup?.[`${groupName}|R`] || "SL";
          const loopDelta =
            Number.isFinite(loopTypes?.[loopType]) ? loopTypes[loopType] : 0;
          const adj = getAdjustment(adjustments, groupName, "R") || 0;

          const corrected = b.measR + corr;
          const afterDelta = corrected + loopDelta + adj - nominal;
          if (Number.isFinite(afterDelta)) perRow[letter].push(afterDelta);
        }
      }
    }

    const avg = (arr) => {
      const v = (arr || []).filter((x) => Number.isFinite(x));
      if (!v.length) return null;
      return v.reduce((a, b) => a + b, 0) / v.length;
    };

    const A = avg(perRow.A);
    const B = avg(perRow.B);
    const C = avg(perRow.C);
    const D = avg(perRow.D);

    const pitch = Number.isFinite(A) && Number.isFinite(D) ? A - D : null;

    return {
      A,
      B,
      C,
      D,
      pitch,
      count: {
        A: perRow.A.length,
        B: perRow.B.length,
        C: perRow.C.length,
        D: perRow.D.length,
      },
    };
  }, [
    wideRows,
    meta?.correction,
    showCorrected,
    activeProfile,
    adjustments,
    groupLoopSetup,
    loopTypes,
    // If you don't have these filters yet, it's fine; React will ignore them.
    includedRows,
    includedGroups,
  ]);

    return map;
  }, [allLines, activeProfile]);

  const csvProfileName = useMemo(() => makeProfileNameFromMeta(meta), [meta]);

  function setProfilesObject(nextProfiles) {
    const json = JSON.stringify(nextProfiles, null, 2);
    setProfileJson(json);
    localStorage.setItem("wingProfilesJson", json);
  }

  function exportAllProfiles() {
    const filename = `wing-profiles-${new Date().toISOString().slice(0, 10)}.json`;
    downloadTextFile(filename, JSON.stringify(profiles, null, 2));
  }

  function exportCurrentProfileOnly() {
    const p = profiles[profileKey];
    if (!p) return alert("No profile selected.");
    const filename = `${(profileKey || "profile").replace(/[^\w\- ]+/g, "")}.json`;
    downloadTextFile(filename, JSON.stringify({ [profileKey]: p }, null, 2));
  }

  function resetProfilesToBuiltIn() {
    localStorage.removeItem("wingProfilesJson");
    setProfileJson(JSON.stringify({ ...BUILTIN_PROFILES }, null, 2));
    const first = Object.keys(BUILTIN_PROFILES)[0] || "";
    setProfileKey(first);
  }

  async function importProfilesFromFile(file) {
    try {
      const text = await file.text();
      const incoming = safeParseProfilesJson(text);

      const merged = { ...profiles, ...incoming };
      for (const [k, v] of Object.entries(merged)) {
        if (v && typeof v === "object" && !v.name) v.name = k;
      }

      setProfilesObject(merged);

      const keys = Object.keys(incoming);
      if (keys.length === 1) setProfileKey(keys[0]);

      alert(`Imported ${Object.keys(incoming).length} profile(s).`);
    } catch (e) {
      console.error(e);
      alert("Could not import profiles JSON. Make sure it is valid JSON exported from this app.");
    }
  }

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

    setProfilesObject(nextProfiles);
    setProfileKey(key);
  }

  function onImportFile(file) {
    const name = (file?.name || "").toLowerCase();

    // XLSX
    if (name.endsWith(".xlsx")) {
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const data = reader.result;
          const wb = XLSX.read(data, { type: "array" });

          const sheetName = wb.SheetNames[0];
          const ws = wb.Sheets[sheetName];

          const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });

          const w = parseWideFlexible(grid);
          if (!w.rows.length) {
            alert("Excel imported, but no line rows were detected.\n\nCheck that line IDs look like A1, B12, C03 etc.");
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
          alert("Failed to read Excel file. Please check it is a .xlsx in the same layout as the CSV.");
        }
      };
      reader.readAsArrayBuffer(file);
      return;
    }

    // CSV
    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result || "");
      const parsed = parseDelimited(text);

      const w = parseWideFlexible(parsed.grid);
      if (!w.rows.length) {
        alert("File imported, but no line rows were detected.\n\nCheck that line IDs look like A1, B12, C03 etc.");
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

  function loopDeltaFor(lineId, side) {
    const g = groupForLine(activeProfile, lineId);
    if (g) {
      const key = `${g}|${side}`;
      const type = groupLoopSetup[key] || "SL";
      const v = loopTypes?.[type];
      return Number.isFinite(v) ? v : 0;
    }
    return 0;
  }

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
    for (const g of allGroupNames) next[`${g}|R`] = next[`${g}|L`] || "SL";
    persistGroupLoopSetup(next);
  }
  function mirrorRtoL() {
    const next = { ...groupLoopSetup };
    for (const g of allGroupNames) next[`${g}|L`] = next[`${g}|R`] || "SL";
    persistGroupLoopSetup(next);
  }
  function resetAdjustments() {
    persistAdjustments({});
  }

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

  // Group average deltas (before vs after)
  const groupStats = useMemo(() => {
    const corr = meta.correction || 0;

    const bucketBefore = new Map(); // group|side -> [delta]
    const bucketAfter = new Map();

    for (const r of wideRows) {
      for (const letter of ["A", "B", "C", "D"]) {
        const b = r[letter];
        if (!b || !b.line || b.nominal == null) continue;

        const g = groupForLine(activeProfile, b.line) || `${letter}?`;

        const loopL = loopDeltaFor(b.line, "L");
        const loopR = loopDeltaFor(b.line, "R");

        const adjL = getAdjustment(adjustments, g, "L");
        const adjR = getAdjustment(adjustments, g, "R");

        const baseL = b.measL == null ? null : b.measL + corr + loopL;
        const baseR = b.measR == null ? null : b.measR + corr + loopR;

        const afterL = baseL == null ? null : baseL + adjL;
        const afterR = baseR == null ? null : baseR + adjR;

        const dL_before = baseL == null ? null : baseL - b.nominal;
        const dR_before = baseR == null ? null : baseR - b.nominal;

        const dL_after = afterL == null ? null : afterL - b.nominal;
        const dR_after = afterR == null ? null : afterR - b.nominal;

        if (Number.isFinite(dL_before)) {
          const k = `${g}|L`;
          if (!bucketBefore.has(k)) bucketBefore.set(k, []);
          bucketBefore.get(k).push(dL_before);
        }
        if (Number.isFinite(dR_before)) {
          const k = `${g}|R`;
          if (!bucketBefore.has(k)) bucketBefore.set(k, []);
          bucketBefore.get(k).push(dR_before);
        }
        if (Number.isFinite(dL_after)) {
          const k = `${g}|L`;
          if (!bucketAfter.has(k)) bucketAfter.set(k, []);
          bucketAfter.get(k).push(dL_after);
        }
        if (Number.isFinite(dR_after)) {
          const k = `${g}|R`;
          if (!bucketAfter.has(k)) bucketAfter.set(k, []);
          bucketAfter.get(k).push(dR_after);
        }
      }
    }

    const out = [];
    const keys = new Set([...bucketBefore.keys(), ...bucketAfter.keys()]);
    for (const key of keys) {
      const [groupName, side] = key.split("|");
      const before = avg(bucketBefore.get(key) || []);
      const after = avg(bucketAfter.get(key) || []);
      if (!Number.isFinite(before) && !Number.isFinite(after)) continue;
      out.push({ groupName, side, before, after });
    }

    out.sort((a, b) =>
      (groupSortKey(a.groupName) + a.side).localeCompare(groupSortKey(b.groupName) + b.side)
    );
    return out;
  }, [wideRows, meta.correction, activeProfile, adjustments, groupLoopSetup, loopTypes]);

  // Chart toggles (A/B/C/D)
  const [chartLetters, setChartLetters] = useState(() => {
    try {
      const s = localStorage.getItem("chartLetters");
      return s ? JSON.parse(s) : { A: true, B: true, C: false, D: false };
    } catch {
      return { A: true, B: true, C: false, D: false };
    }
  });
  useEffect(() => localStorage.setItem("chartLetters", JSON.stringify(chartLetters)), [chartLetters]);

  const chartPoints = useMemo(() => {
    const corr = meta.correction || 0;
    const tol = meta.tolerance || 0;

    const points = []; // {id, xIndex, letter, side, before, after, severityAfter}
    let x = 0;

    // Stable order by A/B/C/D then numeric
    const all = [];
    for (const r of wideRows) {
      for (const letter of ["A", "B", "C", "D"]) {
        const b = r[letter];
        if (!b?.line || b.nominal == null) continue;
        all.push({ letter, ...b });
      }
    }
    all.sort((u, v) => {
      const pu = parseLineId(u.line);
      const pv = parseLineId(v.line);
      const lu = pu?.prefix || u.letter;
      const lv = pv?.prefix || v.letter;
      if (lu !== lv) return lu.localeCompare(lv);
      return (pu?.num ?? 0) - (pv?.num ?? 0);
    });

    for (const b of all) {
      if (!chartLetters[b.letter]) continue;

      const g = groupForLine(activeProfile, b.line) || `${b.letter}?`;

      const loopL = loopDeltaFor(b.line, "L");
      const loopR = loopDeltaFor(b.line, "R");

      const adjL = getAdjustment(adjustments, g, "L");
      const adjR = getAdjustment(adjustments, g, "R");

      const baseL = b.measL == null ? null : b.measL + corr + loopL;
      const baseR = b.measR == null ? null : b.measR + corr + loopR;

      const afterL = baseL == null ? null : baseL + adjL;
      const afterR = baseR == null ? null : baseR + adjR;

      const dL_before = baseL == null ? null : baseL - b.nominal;
      const dR_before = baseR == null ? null : baseR - b.nominal;

      const dL_after = afterL == null ? null : afterL - b.nominal;
      const dR_after = afterR == null ? null : afterR - b.nominal;

      points.push({
        id: `${b.line}-L`,
        xIndex: x,
        line: b.line,
        letter: b.letter,
        side: "L",
        before: dL_before,
        after: dL_after,
        sevAfter: severity(dL_after, tol),
      });
      points.push({
        id: `${b.line}-R`,
        xIndex: x,
        line: b.line,
        letter: b.letter,
        side: "R",
        before: dR_before,
        after: dR_after,
        sevAfter: severity(dR_after, tol),
      });

      x += 1;
    }

    return points;
  }, [wideRows, meta.correction, meta.tolerance, activeProfile, adjustments, chartLetters, groupLoopSetup, loopTypes]);

  // Styles
  const page = {
    minHeight: "100vh",
    background: "#0b0c10",
    color: "#eef1ff",
    fontFamily: "system-ui, sans-serif",
  };
  const wrap = { maxWidth: 1250, margin: "0 auto", padding: 16, display: "flex", flexDirection: "column", gap: 12 };
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

  // Guided editor helpers
  function deepClone(x) {
    return JSON.parse(JSON.stringify(x));
  }
  function openProfileEditor() {
    setDraftProfileKey(profileKey);
    setDraftProfile(deepClone(profiles[profileKey] || activeProfile || {}));
    setShowAdvancedJson(false);
    setIsProfileEditorOpen(true);
  }
  function saveDraftProfile() {
    const nextProfiles = { ...profiles };
    const key = String(draftProfileKey || "").trim();
    if (!key) return alert("Profile name cannot be empty.");

    const p = deepClone(draftProfile || {});
    p.name = key;
    p.mmPerLoop = Number.isFinite(n(p.mmPerLoop)) ? n(p.mmPerLoop) : 10;
    p.mapping = p.mapping && typeof p.mapping === "object" ? p.mapping : { A: [], B: [], C: [], D: [] };

    nextProfiles[key] = p;
    setProfilesObject(nextProfiles);
    setProfileKey(key);
    setIsProfileEditorOpen(false);
  }
  function newProfileFromCurrent() {
    const base = deepClone(profiles[profileKey] || activeProfile || {});
    const name = prompt("New profile name?", `${profileKey} (copy)`);
    if (!name) return;
    base.name = name;
    setDraftProfileKey(name);
    setDraftProfile(base);
    setShowAdvancedJson(false);
    setIsProfileEditorOpen(true);
  }
  function deleteSelectedProfile() {
    if (!confirm(`Delete profile "${profileKey}"? This cannot be undone.`)) return;
    const nextProfiles = { ...profiles };
    delete nextProfiles[profileKey];

    setProfilesObject(nextProfiles);

    const first = Object.keys(nextProfiles)[0] || Object.keys(BUILTIN_PROFILES)[0] || "";
    setProfileKey(first);
  }

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
            Profile (from CSV/XLSX meta): <b style={{ color: "#eef1ff" }}>{csvProfileName}</b>
          </div>
        </div>

        {/* Safety */}
        <div style={{ ...card, borderColor: "rgba(255,204,102,0.5)", background: "rgba(255,204,102,0.08)" }}>
          <b>Safety notice:</b> This is an analysis/simulation tool. Trimming can be dangerous and may invalidate certification.
          Always follow manufacturer/check-center procedures and re-measure after any change.
        </div>

        {/* Workflow */}
        <div style={card}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", justifyContent: "space-between" }}>
            <div style={{ fontWeight: 900 }}>Workflow</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <StepButton current={step} num={1} setStep={setStep} enabled={true} label="1) Import" />
              <StepButton current={step} num={2} setStep={setStep} enabled={hasCSV} label="2) Wing layout" />
              <StepButton current={step} num={3} setStep={setStep} enabled={hasCSV} label="3) Loops setup" />
              <StepButton current={step} num={4} setStep={setStep} enabled={hasCSV && allGroupNames.length > 0} label="4) Trim tables + graphs" />
            </div>
          </div>
          <div style={{ ...muted, fontSize: 12, marginTop: 10 }}>
            Tip: set Step 2–3 before trimming so “before” matches the real baseline.
          </div>
        </div>

        {/* STEP 1 */}
        {step === 1 ? (
          <div style={card}>
            <div style={{ fontWeight: 900, marginBottom: 8 }}>Step 1 — Import measurement CSV / Excel</div>
            <div style={{ ...muted, fontSize: 12, lineHeight: 1.5 }}>
              Upload your measurement file. CSV or XLSX is supported.
              If the file contains a correction value (Korrektur), it is applied as:
              <br />
              <b>corrected = rawMeasured + correction</b> (e.g. 7220 + (-507) = 6713)
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
                Choose file…
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
              Imported wing name will auto-create a profile if it doesn’t exist.
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
                  {Object.keys(profiles)
                    .sort((a, b) => a.localeCompare(b))
                    .map((k) => (
                      <option key={k} value={k}>
                        {k}
                      </option>
                    ))}
                </select>

                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 10 }}>
                  <button style={btnWarn} onClick={openProfileEditor}>Edit selected profile…</button>
                  <button style={btn} onClick={newProfileFromCurrent}>New profile (copy)…</button>
                  <button style={btnDanger} onClick={deleteSelectedProfile}>Delete selected</button>
                </div>

                <div style={{ height: 10 }} />
                <div style={{ ...muted, fontSize: 12 }}>
                  Groups detected: <b style={{ color: "#eef1ff" }}>{allGroupNames.length}</b>
                </div>

                <div style={{ height: 10 }} />
                <div style={{ ...muted, fontSize: 12 }}>
                  Built-in profiles: <b>src/wingProfiles.json</b>. Your edits/custom profiles are saved in this browser.
                </div>
              </div>

              <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
                <div style={{ fontWeight: 850, marginBottom: 8 }}>Profile library (backup & share)</div>

                <input
                  ref={profilesImportRef}
                  type="file"
                  accept="application/json,.json"
                  style={{ display: "none" }}
                  onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) importProfilesFromFile(f);
                    e.target.value = "";
                  }}
                />

                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
                  <button onClick={exportAllProfiles} style={btn}>Export profiles</button>
                  <button onClick={exportCurrentProfileOnly} style={btn}>Export selected</button>
                  <button onClick={() => profilesImportRef.current?.click()} style={btnWarn}>Import profiles JSON…</button>
                  <button onClick={resetProfilesToBuiltIn} style={btnDanger}>Reset to built-in</button>
                </div>

                <details>
                  <summary style={{ cursor: "pointer", color: "#aab1c3", fontSize: 12 }}>
                    Advanced: Raw profiles JSON (power users)
                  </summary>
                  <div style={{ height: 8 }} />
                  <textarea
                    value={profileJson}
                    onChange={(e) => {
                      setProfileJson(e.target.value);
                      localStorage.setItem("wingProfilesJson", e.target.value);
                    }}
                    style={{
                      width: "100%",
                      minHeight: 220,
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
                </details>
              </div>
            </div>

            <div style={{ marginTop: 12, display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button onClick={() => setStep(3)} style={btnWarn} disabled={!hasCSV}>
                Continue to Step 3 (Loops)
              </button>
              <button onClick={() => setStep(1)} style={btn}>Back</button>
            </div>
          </div>
        ) : null}

        {/* STEP 3 */}
        {step === 3 ? (
          <div style={card}>
            <div style={{ fontWeight: 900, marginBottom: 8 }}>Step 3 — Maillon loop setup (baseline)</div>
            <div style={{ ...muted, fontSize: 12, lineHeight: 1.5 }}>
              Set which loop type is installed on each <b>line group</b> maillon (Left/Right). Changing AR1 affects A1–A4 etc.
            </div>

            <div style={{ height: 10 }} />

          {/* Loop types (editable) */}
<div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
  <div style={{ fontWeight: 850, marginBottom: 10 }}>Loop types (editable)</div>

  <div
    style={{
      display: "grid",
      gridTemplateColumns: "minmax(320px, 1fr) minmax(360px, 520px)",
      gap: 12,
      alignItems: "start",
    }}
  >
    {/* Left: text */}
    <div style={{ color: "#aab1c3", fontSize: 12, lineHeight: 1.55 }}>
      Negative means the loop reduces line length.
      <div style={{ height: 6 }} />
      Baseline uses:
      <div style={{ height: 6 }} />
      <div style={{ fontFamily: "ui-monospace, Menlo, Consolas, monospace", color: "#eef1ff" }}>
        baseline = (rawMeasured + correction) + loopDelta
      </div>
      <div style={{ height: 8 }} />
      Tip: keep these numbers matching your real loop set. Only one loop type per group side.
    </div>

    {/* Right: compact 2-column editor */}
<div
  style={{
    display: "grid",
    gap: 8,
    alignItems: "stretch",
  }}
>
  {(() => {
    const entries = Object.entries(loopTypes);
    const rows = [];
    for (let i = 0; i < entries.length; i += 2) {
      rows.push([entries[i], entries[i + 1] || null]);
    }

    const Cell = (entry) => {
      if (!entry) return <div />;
      const [name, mm] = entry;

      return (
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "max-content 86px",
            gap: 8,
            alignItems: "center",
            minWidth: 0,
          }}
        >
          <div style={{ fontWeight: 800, fontSize: 12, whiteSpace: "nowrap" }}>{name}</div>
          <input
            value={mm}
            onChange={(e) => {
              const v = n(e.target.value);
              const next = { ...loopTypes, [name]: Number.isFinite(v) ? v : 0 };
              persistLoopTypes(next);
            }}
            style={{
              width: "100%",
              minWidth: 0,
              borderRadius: 10,
              border: "1px solid #2a2f3f",
              background: "#0b0c10",
              color: "#eef1ff",
              padding: "6px 8px",
              outline: "none",
              fontFamily: "ui-monospace, Menlo, Consolas, monospace",
              textAlign: "right",
              fontSize: 12,
            }}
            inputMode="numeric"
            aria-label={`${name} mm`}
          />
        </div>
      );
    };

    return rows.map((pair, idx) => (
      <div
        key={`pairrow-${idx}`}
        style={{
          display: "grid",
          gridTemplateColumns: "1fr 1fr", // ✅ two loop-types per row
          gap: 12,
          alignItems: "center",
        }}
      >
        {Cell(pair[0])}
        {Cell(pair[1])}
      </div>
    ));
  })()}
</div>

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
                        <th style={{ textAlign: "right", padding: "8px 8px" }}>Δ(mm)</th>
                        <th style={{ textAlign: "left", padding: "8px 8px" }}>Right loop</th>
                        <th style={{ textAlign: "right", padding: "8px 8px" }}>Δ(mm)</th>
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
                                onChange={(e) => persistGroupLoopSetup({ ...groupLoopSetup, [kL]: e.target.value })}
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
                                onChange={(e) => persistGroupLoopSetup({ ...groupLoopSetup, [kR]: e.target.value })}
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

            <div style={{ marginTop: 12, display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button onClick={() => setStep(4)} style={btnWarn} disabled={!allGroupNames.length}>
                Continue to Step 4 (Tables + Graphs)
              </button>
              <button onClick={() => setStep(2)} style={btn}>Back</button>
            </div>
          </div>
        ) : null}

        {/* STEP 4 */}
        {step === 4 ? (
          <div style={card}>
            <div style={{ fontWeight: 900, marginBottom: 8 }}>Step 4 — Tables + Graphs</div>

            {/* Meta controls */}
            <div style={{ ...card, background: "#0e1018" }}>
              <div style={{ fontWeight: 850, marginBottom: 8 }}>Meta controls</div>
              <div style={{ ...muted, fontSize: 12, lineHeight: 1.5 }}>
                Correction is applied to raw measured values (Ist):
                <br />
                <b>corrected = rawMeasured + correction</b> (e.g. 7220 + (-507) = 6713)
              </div>

              <div style={{ height: 10 }} />

              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
                <div style={{ minWidth: 220 }}>
                  <div style={{ ...muted, fontSize: 12, marginBottom: 6 }}>Tolerance (mm)</div>
                  <input
                    value={meta.tolerance ?? 0}
                    onChange={(e) => setMeta((m) => ({ ...m, tolerance: n(e.target.value) ?? 0 }))}
                    style={{ ...input, textAlign: "right", fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}
                    inputMode="numeric"
                  />
                </div>

                <div style={{ minWidth: 220 }}>
                  <div style={{ ...muted, fontSize: 12, marginBottom: 6 }}>Correction (mm)</div>
                  <input
                    value={meta.correction ?? 0}
                    onChange={(e) => setMeta((m) => ({ ...m, correction: n(e.target.value) ?? 0 }))}
                    style={{ ...input, textAlign: "right", fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}
                    inputMode="numeric"
                  />
                </div>

                <label style={{ display: "flex", gap: 10, alignItems: "center", ...muted, fontSize: 12, paddingBottom: 8 }}>
                  <input
                    type="checkbox"
                    checked={showCorrected}
                    onChange={(e) => setShowCorrected(e.target.checked)}
                  />
                  Show corrected values (Ist + Korrektur)
                </label>

                <button onClick={resetAdjustments} style={btnDanger}>Reset all adjustments</button>
              </div>

              {/* Zeroing wizard */}
              <div style={{ marginTop: 10, padding: 12, borderRadius: 14, border: "1px solid #2a2f3f", background: "#0b0c10" }}>
                <div style={{ fontWeight: 850, marginBottom: 6 }}>Zeroing wizard (auto-suggest correction)</div>
                <div style={{ color: "#aab1c3", fontSize: 12, lineHeight: 1.5 }}>
                  Suggests a correction using the <b>median</b> of (Soll − Ist) across all lines. This removes a consistent offset
                  (e.g. ≈ -507mm).
                </div>

                <div style={{ height: 10 }} />

                <button
                  style={btnWarn}
                  onClick={() => {
                    const s = suggestCorrectionFromWideRows(wideRows);
                    if (s == null) return alert("Not enough data to suggest a correction.");
                    const ok = confirm(`Suggested correction: ${s}mm\n\nApply this to Correction now?`);
                    if (!ok) return;
                    setMeta((m) => ({ ...m, correction: s }));
                  }}
                >
                  Suggest correction (median)
                </button>
              </div>
            </div>

            <div style={{ height: 12 }} />

{/* Adjustment UI */}
<div style={{ ...card, background: "#0e1018" }}>
  <div style={{ fontWeight: 850, marginBottom: 8 }}>Trim adjustments per line group (mm)</div>
  <div style={{ ...muted, fontSize: 12, lineHeight: 1.5 }}>
    These simulate adding/removing length at the risers/maillons for each group.
    Positive = longer; negative = shorter.
    <div style={{ height: 6 }} />
    Use the dropdowns to auto-fill the adjustment with a known loop type (SL/DL/AS/etc). You can still type any custom value.
  </div>

  <div style={{ height: 10 }} />

  {!allGroupNames.length ? (
    <div style={{ ...muted, fontSize: 12 }}>No groups found. Check Step 2 mapping.</div>
  ) : (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 980 }}>
        <thead>
          <tr style={{ color: "#aab1c3", fontSize: 12 }}>
            <th style={{ textAlign: "left", padding: "6px 8px" }}>Group</th>
            <th style={{ textAlign: "right", padding: "6px 8px" }}>Adjust L (mm)</th>
            <th style={{ textAlign: "right", padding: "6px 8px" }}>Adjust R (mm)</th>
            <th style={{ textAlign: "right", padding: "6px 8px" }}>Avg Δ before</th>
            <th style={{ textAlign: "right", padding: "6px 8px" }}>Avg Δ after</th>
          </tr>
        </thead>

        <tbody>
          {allGroupNames.map((g) => {
            const kL = `${g}|L`;
            const kR = `${g}|R`;
            const aL = getAdjustment(adjustments, g, "L");
            const aR = getAdjustment(adjustments, g, "R");

            const statL = groupStats.find((s) => s.groupName === g && s.side === "L");
            const statR = groupStats.find((s) => s.groupName === g && s.side === "R");
            const beforeAvg = avg([statL?.before, statR?.before].filter((x) => Number.isFinite(x)));
            const afterAvg = avg([statL?.after, statR?.after].filter((x) => Number.isFinite(x)));

            const tol = meta.tolerance || 0;
            const sevAfter = severity(afterAvg, tol);

            return (
              <tr key={g} style={{ borderTop: "1px solid rgba(42,47,63,0.9)" }}>
                <td style={{ padding: "6px 8px", fontWeight: 900 }}>{g}</td>

                {/* Adjust L (dropdown + input) */}
                <td style={{ padding: "6px 8px", textAlign: "right" }}>
                  <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", alignItems: "center", flexWrap: "wrap" }}>
                    <select
                      value={loopTypeFromAdjustment(aL)}
                      onChange={(e) => {
                        const t = e.target.value;
                        if (!t) return; // Custom selected: do nothing (user can type)
                        const v = loopTypes[t];
                        persistAdjustments({ ...adjustments, [kL]: Number.isFinite(v) ? v : 0 });
                      }}
                      style={{
                        borderRadius: 10,
                        border: "1px solid #2a2f3f",
                        background: "#0d0f16",
                        color: "#eef1ff",
                        padding: "6px 8px",
                        outline: "none",
                        fontSize: 12,
                      }}
                      title="Select a loop type to auto-fill Adjust L"
                    >
                      <option value="">Custom</option>
                      {Object.keys(loopTypes).map((name) => (
                        <option key={name} value={name}>
                          {name} ({loopTypes[name] > 0 ? `+${loopTypes[name]}` : `${loopTypes[name]}`}mm)
                        </option>
                      ))}
                    </select>

                    <input
                      value={aL}
                      onChange={(e) => persistAdjustments({ ...adjustments, [kL]: n(e.target.value) ?? 0 })}
                      style={{
                        ...input,
                        width: 110,
                        padding: "6px 8px",
                        textAlign: "right",
                        fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                      }}
                      inputMode="numeric"
                      title="Manual override (mm)"
                    />
                  </div>
                </td>

                {/* Adjust R (dropdown + input) */}
                <td style={{ padding: "6px 8px", textAlign: "right" }}>
                  <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", alignItems: "center", flexWrap: "wrap" }}>
                    <select
                      value={loopTypeFromAdjustment(aR)}
                      onChange={(e) => {
                        const t = e.target.value;
                        if (!t) return; // Custom selected: do nothing
                        const v = loopTypes[t];
                        persistAdjustments({ ...adjustments, [kR]: Number.isFinite(v) ? v : 0 });
                      }}
                      style={{
                        borderRadius: 10,
                        border: "1px solid #2a2f3f",
                        background: "#0d0f16",
                        color: "#eef1ff",
                        padding: "6px 8px",
                        outline: "none",
                        fontSize: 12,
                      }}
                      title="Select a loop type to auto-fill Adjust R"
                    >
                      <option value="">Custom</option>
                      {Object.keys(loopTypes).map((name) => (
                        <option key={name} value={name}>
                          {name} ({loopTypes[name] > 0 ? `+${loopTypes[name]}` : `${loopTypes[name]}`}mm)
                        </option>
                      ))}
                    </select>

                    <input
                      value={aR}
                      onChange={(e) => persistAdjustments({ ...adjustments, [kR]: n(e.target.value) ?? 0 })}
                      style={{
                        ...input,
                        width: 110,
                        padding: "6px 8px",
                        textAlign: "right",
                        fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                      }}
                      inputMode="numeric"
                      title="Manual override (mm)"
                    />
                  </div>
                </td>

                <td
                  style={{
                    padding: "6px 8px",
                    textAlign: "right",
                    color: "#aab1c3",
                    fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                  }}
                >
                  {Number.isFinite(beforeAvg) ? Math.round(beforeAvg) : "—"}
                </td>

                <td
                  style={{
                    padding: "6px 8px",
                    textAlign: "right",
                    fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                    ...(sevAfter === "red" ? redCell : sevAfter === "yellow" ? yellowCell : null),
                  }}
                >
                  {Number.isFinite(afterAvg) ? Math.round(afterAvg) : "—"}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  )}
</div>

            <div style={{ height: 12 }} />

            {/* Graph controls */}
            <div style={{ ...card, background: "#0e1018" }}>
              <div style={{ fontWeight: 850, marginBottom: 8 }}>Graphs</div>
              <div style={{ ...muted, fontSize: 12, marginBottom: 10 }}>
                Before vs After overlay uses Δ = (after - nominal). Target is 0mm (factory trim).
              </div>

              <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                {["A", "B", "C", "D"].map((L) => (
                  <label key={L} style={{ display: "flex", gap: 8, alignItems: "center", ...muted, fontSize: 12 }}>
                    <input
                      type="checkbox"
                      checked={!!chartLetters[L]}
                      onChange={(e) => setChartLetters({ ...chartLetters, [L]: e.target.checked })}
                    />
                    {L}
                  </label>
                ))}
              </div>

              <div style={{ height: 10 }} />

              <DeltaLineChart
                title="Δ Line chart (Before vs After) — hover points"
                points={chartPoints}
                tolerance={meta.tolerance || 0}
              />

              <div style={{ height: 12 }} />

              <WingProfileChart
                title="Wing profile (Group average Δ)"
                groupStats={groupStats}
                tolerance={meta.tolerance || 0}
              />

              <div style={{ height: 12 }} />

              <RearViewWingChart
                wideRows={wideRows}
                activeProfile={activeProfile}
                tolerance={meta.tolerance || 0}
                showCorrected={showCorrected}
                correction={meta.correction || 0}
                adjustments={adjustments}
                loopTypes={loopTypes}
                groupLoopSetup={groupLoopSetup}
              />
            </div>

            <div style={{ height: 12 }} />

            {/* Tables */}

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
                showCorrected={showCorrected}
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
                showCorrected={showCorrected}
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
                showCorrected={showCorrected}
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
                showCorrected={showCorrected}
              />
            </div>

            <div style={{ marginTop: 12, display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button onClick={() => setStep(3)} style={btn}>Back to Step 3</button>
              <button onClick={() => setStep(2)} style={btn}>Back to Step 2</button>
            </div>
          </div>
        ) : null}

        {/* Guided Profile Editor Modal */}
        {isProfileEditorOpen ? (
          <div
            style={{
              position: "fixed",
              inset: 0,
              background: "rgba(0,0,0,0.6)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              padding: 16,
              zIndex: 9999,
            }}
            onMouseDown={(e) => {
              if (e.target === e.currentTarget) setIsProfileEditorOpen(false);
            }}
          >
            <div
              style={{
                width: "min(1100px, 100%)",
                maxHeight: "92vh",
                overflow: "auto",
                borderRadius: 16,
                border: "1px solid #2a2f3f",
                background: "#11131a",
                padding: 12,
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" }}>
                <div style={{ fontWeight: 950, fontSize: 16 }}>Wing Profile Editor</div>
                <button style={btn} onClick={() => setIsProfileEditorOpen(false)}>Close</button>
              </div>

              <div style={{ height: 10 }} />

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
                  <div style={{ fontWeight: 850, marginBottom: 8 }}>Profile basics</div>

                  <label style={{ color: "#aab1c3", fontSize: 12 }}>Profile name</label>
                  <input
                    value={draftProfileKey}
                    onChange={(e) => setDraftProfileKey(e.target.value)}
                    style={{ ...input, marginTop: 6 }}
                  />

                  <div style={{ height: 10 }} />

                  <label style={{ color: "#aab1c3", fontSize: 12 }}>mm per loop (step size)</label>
                  <input
                    value={draftProfile?.mmPerLoop ?? 10}
                    onChange={(e) => setDraftProfile({ ...draftProfile, mmPerLoop: n(e.target.value) ?? 10 })}
                    style={{ ...input, marginTop: 6 }}
                    inputMode="numeric"
                  />

                  <div style={{ height: 10 }} />
                  <div style={{ color: "#aab1c3", fontSize: 12, lineHeight: 1.5 }}>
                    Hints:
                    <ul style={{ margin: "8px 0 0 18px" }}>
                      <li>Ranges should match your rigging diagram groupings.</li>
                      <li>Example: A1–A4 → AR1 means changes on AR1 affect all A1..A4 lines.</li>
                      <li>Keep ranges non-overlapping for best results.</li>
                    </ul>
                  </div>

                  <div style={{ height: 12 }} />
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <button style={btnWarn} onClick={saveDraftProfile}>Save profile</button>
                    <button style={btn} onClick={() => setIsProfileEditorOpen(false)}>Cancel</button>
                  </div>
                </div>

                <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                    <div style={{ fontWeight: 850 }}>Advanced (JSON)</div>
                    <label style={{ display: "flex", gap: 8, alignItems: "center", color: "#aab1c3", fontSize: 12 }}>
                      <input
                        type="checkbox"
                        checked={showAdvancedJson}
                        onChange={(e) => setShowAdvancedJson(e.target.checked)}
                      />
                      Show JSON
                    </label>
                  </div>

                  {showAdvancedJson ? (
                    <textarea
                      value={JSON.stringify(draftProfile || {}, null, 2)}
                      onChange={(e) => {
                        try {
                          const obj = JSON.parse(e.target.value);
                          setDraftProfile(obj);
                        } catch {
                          // ignore while typing invalid JSON
                        }
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
                        marginTop: 10,
                      }}
                    />
                  ) : (
                    <div style={{ color: "#aab1c3", fontSize: 12, marginTop: 10 }}>
                      Use the guided table editor. Enable JSON only for edge cases.
                    </div>
                  )}
                </div>
              </div>

              <div style={{ height: 12 }} />
              <MappingEditor draftProfile={draftProfile} setDraftProfile={setDraftProfile} btn={btn} />
            </div>
          </div>
        ) : null}
      </div>
    </div>
  );
}

/* ------------------------- Guided Mapping Editor ------------------------- */

function MappingEditor({ draftProfile, setDraftProfile, btn }) {
  const mapping = draftProfile.mapping || { A: [], B: [], C: [], D: [] };
  const letters = ["A", "B", "C", "D"];

  function setRows(letter, rows) {
    const next = { ...draftProfile, mapping: { ...mapping, [letter]: rows } };
    setDraftProfile(next);
  }

  function addRow(letter) {
    const rows = (mapping[letter] || []).slice();
    rows.push([1, 1, `${letter}R1`]);
    setRows(letter, rows);
  }

  function updateCell(letter, idx, col, value) {
    const rows = (mapping[letter] || []).slice();
    const r = rows[idx] ? rows[idx].slice() : [1, 1, `${letter}R1`];
    if (col === 0 || col === 1) {
      const v = parseInt(String(value || "0"), 10);
      r[col] = Number.isFinite(v) ? v : r[col];
    } else {
      r[col] = String(value || "");
    }
    rows[idx] = r;
    setRows(letter, rows);
  }

  function removeRow(letter, idx) {
    const rows = (mapping[letter] || []).slice();
    rows.splice(idx, 1);
    setRows(letter, rows);
  }

  function sortRows(letter) {
    const rows = (mapping[letter] || []).slice().sort((a, b) => (a?.[0] ?? 0) - (b?.[0] ?? 0));
    setRows(letter, rows);
  }

  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 12 }}>
      {letters.map((L) => (
        <div key={L} style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" }}>
            <div style={{ fontWeight: 900 }}>{L} mapping</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button style={btn} onClick={() => addRow(L)}>Add row</button>
              <button style={btn} onClick={() => sortRows(L)}>Sort</button>
            </div>
          </div>

          <div style={{ height: 10 }} />

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 420 }}>
              <thead>
                <tr style={{ color: "#aab1c3", fontSize: 12 }}>
                  <th style={{ textAlign: "right", padding: "6px 8px" }}>From</th>
                  <th style={{ textAlign: "right", padding: "6px 8px" }}>To</th>
                  <th style={{ textAlign: "left", padding: "6px 8px" }}>Group</th>
                  <th style={{ padding: "6px 8px" }}></th>
                </tr>
              </thead>
              <tbody>
                {(mapping[L] || []).map((row, idx) => (
                  <tr key={idx} style={{ borderTop: "1px solid rgba(42,47,63,0.9)" }}>
<td style={{ padding: "6px 8px", textAlign: "right" }}>
  <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", alignItems: "center", flexWrap: "wrap" }}>
    <select
      value={loopTypeFromAdjustment(aL)}
      onChange={(e) => {
        const t = e.target.value;
        if (!t) return; // Custom: leave number as-is
        const v = loopTypes[t];
        persistAdjustments({ ...adjustments, [kL]: Number.isFinite(v) ? v : 0 });
      }}
      style={{
        borderRadius: 10,
        border: "1px solid #2a2f3f",
        background: "#0d0f16",
        color: "#eef1ff",
        padding: "6px 8px",
        outline: "none",
        fontSize: 12,
      }}
      title="Pick a loop type to auto-fill Adjust L"
    >
      <option value="">Custom</option>
      {Object.keys(loopTypes).map((name) => (
        <option key={name} value={name}>
          {name} ({loopTypes[name] > 0 ? `+${loopTypes[name]}` : `${loopTypes[name]}`}mm)
        </option>
      ))}
    </select>

    <input
      value={aL}
      onChange={(e) => persistAdjustments({ ...adjustments, [kL]: n(e.target.value) ?? 0 })}
      style={{
        ...input,
        width: 110,
        padding: "6px 8px",
        textAlign: "right",
        fontFamily: "ui-monospace, Menlo, Consolas, monospace",
      }}
      inputMode="numeric"
      title="Manual override (mm)"
    />
  </div>
</td>

<td style={{ padding: "6px 8px", textAlign: "right" }}>
  <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", alignItems: "center", flexWrap: "wrap" }}>
    <select
      value={loopTypeFromAdjustment(aR)}
      onChange={(e) => {
        const t = e.target.value;
        if (!t) return; // Custom: leave number as-is
        const v = loopTypes[t];
        persistAdjustments({ ...adjustments, [kR]: Number.isFinite(v) ? v : 0 });
      }}
      style={{
        borderRadius: 10,
        border: "1px solid #2a2f3f",
        background: "#0d0f16",
        color: "#eef1ff",
        padding: "6px 8px",
        outline: "none",
        fontSize: 12,
      }}
      title="Pick a loop type to auto-fill Adjust R"
    >
      <option value="">Custom</option>
      {Object.keys(loopTypes).map((name) => (
        <option key={name} value={name}>
          {name} ({loopTypes[name] > 0 ? `+${loopTypes[name]}` : `${loopTypes[name]}`}mm)
        </option>
      ))}
    </select>

    <input
      value={aR}
      onChange={(e) => persistAdjustments({ ...adjustments, [kR]: n(e.target.value) ?? 0 })}
      style={{
        ...input,
        width: 110,
        padding: "6px 8px",
        textAlign: "right",
        fontFamily: "ui-monospace, Menlo, Consolas, monospace",
      }}
      inputMode="numeric"
      title="Manual override (mm)"
    />
  </div>
</td>

                    <td style={{ padding: "6px 8px" }}>
                      <input
                        value={row?.[2] ?? ""}
                        onChange={(e) => updateCell(L, idx, 2, e.target.value)}
                        style={{
                          width: "100%",
                          padding: "6px 8px",
                          borderRadius: 10,
                          border: "1px solid #2a2f3f",
                          background: "#0d0f16",
                          color: "#eef1ff",
                        }}
                      />
                    </td>
                    <td style={{ padding: "6px 8px", textAlign: "right" }}>
                      <button style={btn} onClick={() => removeRow(L, idx)}>Delete</button>
                    </td>
                  </tr>
                ))}
                {!mapping[L] || mapping[L].length === 0 ? (
                  <tr>
                    <td colSpan={4} style={{ padding: "8px 8px", color: "#aab1c3", fontSize: 12 }}>
                      No ranges yet. Click “Add row”.
                    </td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </div>

          <div style={{ color: "#aab1c3", fontSize: 12, marginTop: 10 }}>
            Tip: If your diagram says AR1 controls A1–A4, set A: 1 → 4 = AR1.
          </div>
        </div>
      ))}
    </div>
  );
}

/* ------------------------- Charts ------------------------- */


function DeltaLineChart({ title, points, tolerance }) {
  const width = 1100;
  const height = 250;
  const pad = 26;

  const [hover, setHover] = useState(null); // {x,y, p}

  const series = useMemo(() => {
    const bySide = { L: [], R: [] };
    for (const p of points) {
      if (!Number.isFinite(p.before) && !Number.isFinite(p.after)) continue;
      bySide[p.side].push(p);
    }
    bySide.L.sort((a, b) => a.xIndex - b.xIndex);
    bySide.R.sort((a, b) => a.xIndex - b.xIndex);
    return bySide;
  }, [points]);

  const allValues = useMemo(() => {
    const v = [];
    for (const p of points) {
      if (Number.isFinite(p.before)) v.push(p.before);
      if (Number.isFinite(p.after)) v.push(p.after);
    }
    v.push(0);
    if ((tolerance || 0) > 0) v.push(tolerance, -tolerance);
    return v;
  }, [points, tolerance]);

  const { minY, maxY } = useMemo(() => {
    if (!allValues.length) return { minY: -10, maxY: 10 };
    let mn = Math.min(...allValues);
    let mx = Math.max(...allValues);
    if (mn === mx) {
      mn -= 1;
      mx += 1;
    }
    const span = mx - mn;
    return { minY: mn - span * 0.12, maxY: mx + span * 0.12 };
  }, [allValues]);

  const maxX = useMemo(() => {
    const mx = Math.max(0, ...points.map((p) => p.xIndex));
    return mx <= 0 ? 1 : mx;
  }, [points]);

  function xScale(x) {
    return pad + (x / maxX) * (width - pad * 2);
  }
  function yScale(y) {
    return pad + ((maxY - y) / (maxY - minY)) * (height - pad * 2);
  }

  function polyPath(ps, field) {
    const pts = ps
      .filter((p) => Number.isFinite(p[field]))
      .map((p) => `${xScale(p.xIndex)},${yScale(p[field])}`);
    if (!pts.length) return "";
    return `M ${pts[0]} ` + pts.slice(1).map((s) => `L ${s}`).join(" ");
  }

  const tol = tolerance || 0;
  const y0 = yScale(0);
  const yTolP = tol > 0 ? yScale(tol) : null;
  const yTolN = tol > 0 ? yScale(-tol) : null;
  const yWarnP = tol > 3 ? yScale(Math.max(0, tol - 3)) : null;
  const yWarnN = tol > 3 ? yScale(-Math.max(0, tol - 3)) : null;

  function sevColor(sev) {
    if (sev === "red") return "rgba(255,107,107,1)";
    if (sev === "yellow") return "rgba(255,214,102,1)";
    return "rgba(170,177,195,0.9)";
  }

  return (
    <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0d0f16" }}>
      <div style={{ fontWeight: 850, marginBottom: 8 }}>{title}</div>

      {!points.length ? (
        <div style={{ color: "#aab1c3", fontSize: 12 }}>No chart data yet.</div>
      ) : (
        <div style={{ position: "relative" }}>
          <svg
            width="100%"
            viewBox={`0 0 ${width} ${height}`}
            style={{ display: "block" }}
            onMouseLeave={() => setHover(null)}
          >
            {/* Target line */}
            <line x1={pad} x2={width - pad} y1={y0} y2={y0} stroke="rgba(170,177,195,0.35)" strokeWidth="2" />

            {/* Tolerance / Warning bands */}
            {tol > 0 ? (
              <>
                <line x1={pad} x2={width - pad} y1={yTolP} y2={yTolP} stroke="rgba(255,107,107,0.45)" strokeWidth="2" />
                <line x1={pad} x2={width - pad} y1={yTolN} y2={yTolN} stroke="rgba(255,107,107,0.45)" strokeWidth="2" />
              </>
            ) : null}
            {tol > 3 ? (
              <>
                <line x1={pad} x2={width - pad} y1={yWarnP} y2={yWarnP} stroke="rgba(255,214,102,0.55)" strokeWidth="2" />
                <line x1={pad} x2={width - pad} y1={yWarnN} y2={yWarnN} stroke="rgba(255,214,102,0.55)" strokeWidth="2" />
              </>
            ) : null}

            {/* Before (dashed) */}
            <path d={polyPath(series.L, "before")} fill="none" stroke="rgba(176,132,255,0.75)" strokeWidth="2" strokeDasharray="6 6" />
            <path d={polyPath(series.R, "before")} fill="none" stroke="rgba(102,204,255,0.75)" strokeWidth="2" strokeDasharray="6 6" />

            {/* After (solid) */}
            <path d={polyPath(series.L, "after")} fill="none" stroke="rgba(176,132,255,1)" strokeWidth="3" />
            <path d={polyPath(series.R, "after")} fill="none" stroke="rgba(102,204,255,1)" strokeWidth="3" />

            {/* Points (After) */}
            {points.map((p) => {
              if (!Number.isFinite(p.after)) return null;
              const cx = xScale(p.xIndex);
              const cy = yScale(p.after);
              return (
                <circle
                  key={`pt-${p.id}`}
                  cx={cx}
                  cy={cy}
                  r={4.2}
                  fill={sevColor(p.sevAfter)}
                  stroke="rgba(0,0,0,0.35)"
                  strokeWidth="1"
                  onMouseEnter={() => setHover({ x: cx, y: cy, p })}
                  onMouseMove={() => setHover({ x: cx, y: cy, p })}
                />
              );
            })}
          </svg>

          {/* Tooltip */}
          {hover ? (
            <div
              style={{
                position: "absolute",
                left: `${(hover.x / width) * 100}%`,
                top: `${(hover.y / height) * 100}%`,
                transform: "translate(12px, -12px)",
                pointerEvents: "none",
                background: "#0b0c10",
                border: "1px solid #2a2f3f",
                borderRadius: 12,
                padding: 10,
                minWidth: 220,
                boxShadow: "0 10px 30px rgba(0,0,0,0.35)",
                color: "#eef1ff",
                fontSize: 12,
              }}
            >
              <div style={{ fontWeight: 900, marginBottom: 6 }}>
                {hover.p.line} ({hover.p.side})
              </div>
              <div style={{ color: "#aab1c3", fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                Δ before:{" "}
                {Number.isFinite(hover.p.before)
                  ? `${hover.p.before > 0 ? "+" : ""}${Math.round(hover.p.before)}mm`
                  : "—"}
                <br />
                Δ after:{" "}
                {Number.isFinite(hover.p.after)
                  ? `${hover.p.after > 0 ? "+" : ""}${Math.round(hover.p.after)}mm`
                  : "—"}
                <br />
                Severity: <b>{hover.p.sevAfter}</b>
              </div>
            </div>
          ) : null}
        </div>
      )}

      <div style={{ color: "#aab1c3", fontSize: 12, marginTop: 8, display: "flex", gap: 14, flexWrap: "wrap" }}>
        <span>Solid = After</span>
        <span>Dashed = Before</span>
        <span>Points = After (hover)</span>
        <span>Target = 0mm</span>
      </div>
    </div>
  );
}



function WingProfileChart({ title, groupStats, tolerance }) {
  const width = 1100;
  const height = 260;
  const pad = 24;

  const tol = tolerance || 0;

  const rows = useMemo(() => {
    const byGroup = new Map();
    for (const s of groupStats) {
      if (!byGroup.has(s.groupName)) byGroup.set(s.groupName, { group: s.groupName });
      const obj = byGroup.get(s.groupName);
      obj[s.side] = s;
    }
    const out = Array.from(byGroup.values());
    out.sort((a, b) => groupSortKey(a.group).localeCompare(groupSortKey(b.group)));
    return out;
  }, [groupStats]);

  const values = useMemo(() => {
    const v = [0];
    for (const r of rows) {
      if (Number.isFinite(r.L?.after)) v.push(r.L.after);
      if (Number.isFinite(r.R?.after)) v.push(r.R.after);
      if (Number.isFinite(r.L?.before)) v.push(r.L.before);
      if (Number.isFinite(r.R?.before)) v.push(r.R.before);
    }
    if (tol > 0) v.push(tol, -tol);
    return v;
  }, [rows, tol]);

  const { minY, maxY } = useMemo(() => {
    if (!values.length) return { minY: -10, maxY: 10 };
    let mn = Math.min(...values);
    let mx = Math.max(...values);
    if (mn === mx) {
      mn -= 1;
      mx += 1;
    }
    const span = mx - mn;
    return { minY: mn - span * 0.12, maxY: mx + span * 0.12 };
  }, [values]);

  function yScale(y) {
    return pad + ((maxY - y) / (maxY - minY)) * (height - pad * 2);
  }

  const xCount = Math.max(1, rows.length);
  function xScale(i) {
    return pad + (i / (xCount - 1 || 1)) * (width - pad * 2);
  }

  const y0 = yScale(0);
  const yTolP = tol > 0 ? yScale(tol) : null;
  const yTolN = tol > 0 ? yScale(-tol) : null;

  const barW = Math.max(10, Math.min(18, (width - pad * 2) / (rows.length * 4)));

  return (
    <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0d0f16" }}>
      <div style={{ fontWeight: 850, marginBottom: 8 }}>{title}</div>
      {!rows.length ? (
        <div style={{ color: "#aab1c3", fontSize: 12 }}>No group data yet.</div>
      ) : (
        <svg width="100%" viewBox={`0 0 ${width} ${height}`} style={{ display: "block" }}>
          <line x1={pad} x2={width - pad} y1={y0} y2={y0} stroke="rgba(170,177,195,0.35)" strokeWidth="2" />
          {tol > 0 ? (
            <>
              <line x1={pad} x2={width - pad} y1={yTolP} y2={yTolP} stroke="rgba(255,107,107,0.45)" strokeWidth="2" />
              <line x1={pad} x2={width - pad} y1={yTolN} y2={yTolN} stroke="rgba(255,107,107,0.45)" strokeWidth="2" />
            </>
          ) : null}

          {rows.map((r, i) => {
            const x = xScale(i);

            const L_after = r.L?.after;
            const R_after = r.R?.after;
            const L_before = r.L?.before;
            const R_before = r.R?.before;

            const yL = Number.isFinite(L_after) ? yScale(L_after) : null;
            const yR = Number.isFinite(R_after) ? yScale(R_after) : null;

            const yLb = Number.isFinite(L_before) ? yScale(L_before) : null;
            const yRb = Number.isFinite(R_before) ? yScale(R_before) : null;

            return (
              <g key={r.group}>
                {yL != null ? (
                  <rect
                    x={x - barW - 3}
                    y={Math.min(y0, yL)}
                    width={barW}
                    height={Math.abs(y0 - yL)}
                    fill="rgba(176,132,255,0.9)"
                  />
                ) : null}
                {yR != null ? (
                  <rect
                    x={x + 3}
                    y={Math.min(y0, yR)}
                    width={barW}
                    height={Math.abs(y0 - yR)}
                    fill="rgba(102,204,255,0.9)"
                  />
                ) : null}

                {yLb != null ? (
                  <line x1={x - barW - 6} x2={x - 2} y1={yLb} y2={yLb} stroke="rgba(176,132,255,0.6)" strokeWidth="3" />
                ) : null}
                {yRb != null ? (
                  <line x1={x + 2} x2={x + barW + 6} y1={yRb} y2={yRb} stroke="rgba(102,204,255,0.6)" strokeWidth="3" />
                ) : null}

                {rows.length <= 18 || i % 2 === 0 ? (
                  <text x={x} y={height - 8} textAnchor="middle" fontSize="10" fill="rgba(170,177,195,0.9)">
                    {r.group}
                  </text>
                ) : null}
              </g>
            );
          })}
        </svg>
      )}

      <div style={{ color: "#aab1c3", fontSize: 12, marginTop: 8, display: "flex", gap: 14, flexWrap: "wrap" }}>
        <span>Purple = Left (after)</span>
        <span>Cyan = Right (after)</span>
        <span>Small tick = Before</span>
      </div>
    </div>
  );
}

/* ------------------------- Compact measurement table ------------------------- */

function BlockTable({
  title,
  rows,
  meta,
  activeProfile,
  adjustments,
  loopDeltaFor,
  input,
  redCell,
  yellowCell,
  setCell,
  blockKey,
  showCorrected,
}) {
  const corr = meta.correction || 0;
  const tol = meta.tolerance || 0;

  const styleFor = (sev) => (sev === "red" ? redCell : sev === "yellow" ? yellowCell : null);

  return (
    <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, overflow: "hidden", background: "#0e1018" }}>
      <div style={{ padding: 10, borderBottom: "1px solid #2a2f3f", fontWeight: 900 }}>
        {title} lines{" "}
        <span style={{ color: "#aab1c3", fontSize: 12, fontWeight: 700 }}>
          ({showCorrected ? "showing corrected" : "showing raw"})
        </span>
      </div>

      <div style={{ overflowX: "auto" }}>
	  // changed from 720 to 420
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 380 }}>
          <thead>
            <tr style={{ color: "#aab1c3", fontSize: 12 }}>
              <th style={{ textAlign: "left", padding: "6px 8px" }}>Line</th>
              <th style={{ textAlign: "left", padding: "6px 8px" }}>Group</th>
              <th style={{ textAlign: "right", padding: "6px 8px" }}>Soll</th>
              <th style={{ textAlign: "right", padding: "6px 8px" }}>{showCorrected ? "Ist L (corr)" : "Ist L (raw)"}</th>
              <th style={{ textAlign: "right", padding: "6px 8px" }}>{showCorrected ? "Ist R (corr)" : "Ist R (raw)"}</th>
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

                const correctedL = b.measL == null ? null : b.measL + corr;
                const correctedR = b.measR == null ? null : b.measR + corr;

                const baseL = correctedL == null ? null : correctedL + loopL;
                const baseR = correctedR == null ? null : correctedR + loopR;

                const afterL = baseL == null ? null : baseL + adjL;
                const afterR = baseR == null ? null : baseR + adjR;

                const dL_before = baseL == null || b.nominal == null ? null : baseL - b.nominal;
                const dR_before = baseR == null || b.nominal == null ? null : baseR - b.nominal;

                const dL_after = afterL == null || b.nominal == null ? null : afterL - b.nominal;
                const dR_after = afterR == null || b.nominal == null ? null : afterR - b.nominal;

                const sevL = severity(dL_after, tol);
                const sevR = severity(dR_after, tol);

                const displayL = showCorrected ? correctedL : b.measL;
                const displayR = showCorrected ? correctedR : b.measR;

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
                        title="Edit raw measured (Ist). Correction/loops/adjustments are applied automatically."
                      />
                      <div style={{ color: "#aab1c3", fontSize: 10, marginTop: 4, fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                        show {displayL == null ? "—" : Math.round(displayL)} | loop {loopL > 0 ? `+${loopL}` : `${loopL}`} | adj{" "}
                        {adjL > 0 ? `+${adjL}` : `${adjL}`}
                      </div>
                      <div style={{ color: "#aab1c3", fontSize: 10, marginTop: 2, fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                        Δ(before){" "}
                        {Number.isFinite(dL_before) ? `${dL_before > 0 ? "+" : ""}${Math.round(dL_before)}mm` : "–"} → Δ(after){" "}
                        {Number.isFinite(dL_after) ? `${dL_after > 0 ? "+" : ""}${Math.round(dL_after)}mm` : "–"}
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
                        title="Edit raw measured (Ist). Correction/loops/adjustments are applied automatically."
                      />
                      <div style={{ color: "#aab1c3", fontSize: 10, marginTop: 4, fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                        show {displayR == null ? "—" : Math.round(displayR)} | loop {loopR > 0 ? `+${loopR}` : `${loopR}`} | adj{" "}
                        {adjR > 0 ? `+${adjR}` : `${adjR}`}
                      </div>
                      <div style={{ color: "#aab1c3", fontSize: 10, marginTop: 2, fontFamily: "ui-monospace, Menlo, Consolas, monospace" }}>
                        Δ(before){" "}
                        {Number.isFinite(dR_before) ? `${dR_before > 0 ? "+" : ""}${Math.round(dR_before)}mm` : "–"} → Δ(after){" "}
                        {Number.isFinite(dR_after) ? `${dR_after > 0 ? "+" : ""}${Math.round(dR_after)}mm` : "–"}
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
        Yellow: within 3mm of tolerance. Red: at/over tolerance. Target is 0mm (factory trim).
      </div>
    </div>
  );
}

function RearViewWingChart({
  wideRows,
  activeProfile,
  tolerance,
  showCorrected,
  correction,
  adjustments,
  loopTypes,
  groupLoopSetup,
}) {
  const width = 1100;
  const height = 460;
  const pad = 24;
  const tol = Number.isFinite(tolerance) ? tolerance : 0;

  const [hover, setHover] = React.useState(null);
  const [spanMode, setSpanMode] = React.useState("real"); // "linear" | "real"
  const [showGroupCuts, setShowGroupCuts] = React.useState(true);
  const [showBeforePoints, setShowBeforePoints] = React.useState(false);

  const data = React.useMemo(() => {
    if (!wideRows?.length) return null;

    const points = [];
    for (const r of wideRows) {
      for (const letter of ["A", "B", "C", "D"]) {
        const b = r?.[letter];
        if (!b?.line) continue;

        const lineId = b.line;
        const groupName = groupForLine(activeProfile, lineId) || "—";

        const kL = `${groupName}|L`;
        const kR = `${groupName}|R`;

        const loopNameL = groupLoopSetup?.[kL] || "SL";
        const loopNameR = groupLoopSetup?.[kR] || "SL";

        const loopDeltaL = Number.isFinite(loopTypes?.[loopNameL]) ? loopTypes[loopNameL] : 0;
        const loopDeltaR = Number.isFinite(loopTypes?.[loopNameR]) ? loopTypes[loopNameR] : 0;

        const corr = showCorrected ? (Number.isFinite(correction) ? correction : 0) : 0;

        const beforeL = deltaMm({
          nominal: b.nominal,
          measured: b.measL,
          correction: corr,
          adjustment: loopDeltaL,
        });
        const beforeR = deltaMm({
          nominal: b.nominal,
          measured: b.measR,
          correction: corr,
          adjustment: loopDeltaR,
        });

        const adjL = getAdjustment(adjustments || {}, groupName, "L");
        const adjR = getAdjustment(adjustments || {}, groupName, "R");

        const afterL = deltaMm({
          nominal: b.nominal,
          measured: b.measL,
          correction: corr,
          adjustment: loopDeltaL + adjL,
        });
        const afterR = deltaMm({
          nominal: b.nominal,
          measured: b.measR,
          correction: corr,
          adjustment: loopDeltaR + adjR,
        });

        points.push({ letter, lineId, groupName, beforeL, beforeR, afterL, afterR });
      }
    }

    const byLetter = { A: [], B: [], C: [], D: [] };
    for (const p of points) byLetter[p.letter].push(p);

    for (const L of ["A", "B", "C", "D"]) {
      byLetter[L].sort((p1, p2) => {
        const a = parseLineId(p1.lineId);
        const b = parseLineId(p2.lineId);
        return (a?.num ?? 0) - (b?.num ?? 0);
      });
    }

    return byLetter;
  }, [wideRows, activeProfile, showCorrected, correction, adjustments, loopTypes, groupLoopSetup]);

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
    const halfSpan = (width - pad * 2) / 2 - 40;
    const centerGap = 18;

    const t = count <= 1 ? 0 : i / (count - 1);
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
      const g = arr[i]?.groupName || "";
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

      <div style={{ overflowX: "auto" }}>
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
            const halfSpan = (width - pad * 2) / 2 - 40;
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
                <line x1={pad} y1={yMid} x2={width - pad} y2={yMid} stroke="rgba(255,255,255,0.10)" strokeDasharray="4 6" />

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
                      <line x1={xL} y1={b.y0 + 2} x2={xL} y2={b.y1 - 2} stroke="rgba(255,255,255,0.08)" />
                      <line x1={xR} y1={b.y0 + 2} x2={xR} y2={b.y1 - 2} stroke="rgba(255,255,255,0.08)" />
                    </g>
                  );
                })}

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
                            groupName: p.groupName,
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

