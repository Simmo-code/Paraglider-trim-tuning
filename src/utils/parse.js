import * as XLSX from "xlsx";
import { safeNum } from "./math.js";

/**
 * Parse raw CSV text into a 2-D array of strings.
 * Handles quoted fields containing commas.
 */
export function rowsFromCSVText(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .filter((l) => l.trim().length > 0);

  return lines.map((line) => {
    const out = [];
    let cur = "";
    let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { inQ = !inQ; continue; }
      if (ch === "," && !inQ) { out.push(cur); cur = ""; }
      else { cur += ch; }
    }
    out.push(cur);
    return out;
  });
}

/**
 * Convert an XLSX worksheet to a 2-D array of strings
 * (same shape as rowsFromCSVText output).
 */
export function rowsFromSheetAOA(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
}

/**
 * Parse the wide-format line-check table used by this app.
 *
 * Expected layout:
 *   Row 0 – meta key headers  (Make, Model, Tolerance, Korrektur …)
 *   Row 1 – meta values
 *   Row 2 – column headers    (A, Soll, Ist L, Ist R, B, …)
 *   Row 3+ – data rows
 *
 * Returns { meta, wideRows } where each wideRow is:
 *   { letter, idx, lineBase, nominal, measuredL, measuredR }
 */
export function parseWideTableFromRows(rows) {
  const meta = {
    make: "", model: "", size: "", serial: "",
    checkedBy: "", date: "", tolerance: 10, correction: 0,
  };
  const wideRows = [];

  if (!Array.isArray(rows) || rows.length < 4) return { meta, wideRows };

  const metaKeys = rows[0] || [];
  const metaVals = rows[1] || [];

  const findMeta = (needle) => {
    const idx = metaKeys.findIndex(
      (k) => String(k || "").toLowerCase().includes(needle)
    );
    return idx >= 0 ? metaVals[idx] : null;
  };

  meta.make  = String(findMeta("make")  || metaVals[0] || "").trim();
  meta.model = String(findMeta("model") || metaVals[1] || "").trim();

  const tolVal  = findMeta("tolerance");
  const corrVal = findMeta("korrektur") || findMeta("correction");
  if (safeNum(tolVal)  != null) meta.tolerance  = safeNum(tolVal);
  if (safeNum(corrVal) != null) meta.correction = safeNum(corrVal);

  const header  = (rows[2] || []).map((h) => String(h || "").trim());
  const headerU = header.map((h) => h.toUpperCase());
  const letters = ["A", "B", "C", "D"];
  const BRAKE_LETTER = "BR";

  // Detect riser cascade column starts.
  const letterCols = [];
  for (let c = 0; c < header.length; c++) {
    if (letters.includes(headerU[c])) letterCols.push({ letter: headerU[c], colStart: c });
  }
  // Fallback: assume 4×4 layout if no explicit headers found.
  if (letterCols.length === 0 && header.length >= 16) {
    letterCols.push(
      { letter: "A", colStart: 0 },
      { letter: "B", colStart: 4 },
      { letter: "C", colStart: 8 },
      { letter: "D", colStart: 12 },
    );
  }

  // Optional brake columns.
  const brkStart = headerU.findIndex((h) => ["BRK", "BRAKE", "BR"].includes(h));
  const brkLIdx  = headerU.findIndex((h) => ["BRKL", "BRK L", "BRAKEL", "BRAKE L"].includes(h));
  const brkRIdx  = headerU.findIndex((h) => ["BRKR", "BRK R", "BRAKER", "BRAKE R"].includes(h));

  for (let r = 3; r < rows.length; r++) {
    const row = rows[r] || [];
    let anyIdx = null;

    for (const blk of letterCols) {
      const c    = blk.colStart;
      const base = String(row[c] || "").trim();
      if (!base) continue;

      const nominal   = safeNum(row[c + 1]);
      const measuredL = safeNum(row[c + 2]);
      const measuredR = safeNum(row[c + 3]);

      const m      = base.match(/^([A-Za-z])\s*([0-9]+)\s*$/);
      const letter = ((m && m[1]) || blk.letter || "").toUpperCase();
      const idx    = m ? Number(m[2]) : (safeNum(base.replace(/[^\d]/g, "")) ?? null);

      if (anyIdx == null && idx != null) anyIdx = idx;
      if (!"ABCD".includes(letter) || idx == null) continue;

      wideRows.push({ letter, idx, lineBase: `${letter}${idx}`, nominal, measuredL, measuredR });
    }

    // Brake rows.
    if (brkStart >= 0 || (brkLIdx >= 0 && brkRIdx >= 0)) {
      const idx = anyIdx != null ? Number(anyIdx) : null;
      if (idx != null && Number.isFinite(idx)) {
        const nominal =
          brkStart >= 0
            ? safeNum(row[brkStart + 1])
            : (brkLIdx > 0 && headerU[brkLIdx - 1]?.includes("SOLL")
                ? safeNum(row[brkLIdx - 1])
                : null);
        const nominalClean = nominal === 0 ? null : nominal;
        const measuredL = brkStart >= 0 ? safeNum(row[brkStart + 2]) : safeNum(row[brkLIdx]);
        const measuredR = brkStart >= 0 ? safeNum(row[brkStart + 3]) : safeNum(row[brkRIdx]);

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
