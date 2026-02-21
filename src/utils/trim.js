import { PALETTE } from "./constants.js";

/**
 * Returns "good" | "warn" | "bad" | "na" for a delta value.
 *   good  = green  (|Δ| ≤ 4 mm)
 *   warn  = yellow (> 4 mm but still within tolerance)
 *   bad   = red    (at or over tolerance)
 */
export function bandForDelta(delta, tolerance) {
  if (delta == null || !Number.isFinite(Number(delta))) return "na";
  const abs = Math.abs(Number(delta));
  const tol = Number.isFinite(Number(tolerance)) ? Number(tolerance) : 0;
  if (abs <= 4) return "good";
  if (abs < tol) return "warn";
  return "bad";
}

/**
 * Returns "green" | "yellow" | "red" | "na" for SVG chart colouring.
 *   green  = within active tolerance (default ±4 mm)
 *   yellow = outside ±4 mm but within a looser manual tolerance
 *   red    = exceeds manual tolerance
 */
export function severity(delta, tolerance) {
  if (delta == null || !Number.isFinite(Number(delta))) return "na";
  const abs = Math.abs(Number(delta));
  const tol = Number.isFinite(Number(tolerance)) ? Number(tolerance) : 0;
  const t = tol > 0 ? tol : 4;

  if (abs <= t) return "green";
  if (tol > 0 && abs > tol) return "red";
  if (tol > 0) return "yellow";
  return "yellow";
}

/** Base colour for a line chip derived from the riser letter (A/B/C/D). */
export function chipColorFromLineId(lineId) {
  const first = String(lineId || "").trim().toUpperCase().charAt(0);
  return (PALETTE[first] || PALETTE.A).base;
}

/** Shade of a riser group colour based on bucket depth (1 = darkest). */
export function groupColor(letter, bucket) {
  const p = PALETTE[letter] || PALETTE.A;
  if (bucket === 1) return p.base;
  if (bucket === 2) return p.s2;
  if (bucket === 3) return p.s3;
  return p.s4;
}
