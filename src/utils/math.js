/** Clamp n between a and b. Returns a if n is non-finite. */
export function clamp(n, a, b) {
  const x = Number(n);
  if (!Number.isFinite(x)) return a;
  return Math.max(a, Math.min(b, x));
}

/**
 * Parse a cell value to a finite number or null.
 * Accepts comma-as-decimal (European format) and trims whitespace.
 * Returns null for blank or non-numeric values (not 0).
 */
export function safeNum(v) {
  const s = String(v || "").trim();
  if (!s) return null;
  const n = Number(s.replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

/** Median of an array of numbers. Returns null for empty input. */
export function median(values) {
  const nums = (values || []).filter((x) => typeof x === "number" && Number.isFinite(x));
  if (nums.length === 0) return null;
  nums.sort((a, b) => a - b);
  const mid = Math.floor(nums.length / 2);
  return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
}

/** Deep-clone an object via structuredClone (with JSON fallback). */
export function deepClone(obj) {
  if (typeof structuredClone === "function") return structuredClone(obj);
  return JSON.parse(JSON.stringify(obj));
}
