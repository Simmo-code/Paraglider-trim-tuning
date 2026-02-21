import { clamp } from "./math.js";

/**
 * Build default index-range buckets for a given line count and group count.
 *
 * Paragliding convention:
 *   3 groups → AR1=lines 1-4, AR2=lines 5-8, AR3=rest
 *   4 groups → AR1=1-4, AR2=5-8, AR3=9-12, AR4=rest
 *   Other    → evenly split
 *
 * Returns { [bucket]: { start, end } }
 */
export function makeDefaultRanges(maxIdx, groupCount) {
  const m = Math.max(1, Number(maxIdx   || 1));
  const n = Math.max(1, Number(groupCount || 3));
  const out = {};

  const setBucket = (b, s, e) => {
    out[b] = {
      start: Math.max(1, Math.min(m, s)),
      end:   Math.max(1, Math.min(m, e)),
    };
    if (out[b].end < out[b].start) out[b].end = out[b].start;
  };

  if (n === 3 || n === 4) {
    const e1 = Math.min(m, 4);
    setBucket(1, 1, e1);

    const s2 = Math.min(m, e1 + 1);
    const e2 = Math.min(m, s2 + 3);
    setBucket(2, s2, e2);

    if (n === 3) {
      setBucket(3, Math.min(m, e2 + 1), m);
      return out;
    }

    const s3 = Math.min(m, e2 + 1);
    const e3 = Math.min(m, s3 + 3);
    setBucket(3, s3, e3);
    setBucket(4, Math.min(m, e3 + 1), m);
    return out;
  }

  // Generic evenly-split fallback.
  const step  = Math.ceil(m / n);
  let   start = 1;
  for (let b = 1; b <= n; b++) {
    const s = Math.min(start, m);
    const e = b === n ? m : Math.min(m, s + step - 1);
    out[b] = { start: s, end: Math.max(s, e) };
    start  = e + 1;
  }
  return out;
}

/**
 * Build the full lineId → groupId mapping for all riser letters.
 *
 * e.g. "A3L" → "AR1L", "B7R" → "BR2R"
 */
export function buildInitialLineToGroup({
  maxByLetter,
  groupCountByLetter,
  prefixByLetter,
  rangesByLetter,
}) {
  const mapping = {};

  for (const letter of ["A", "B", "C", "D"]) {
    const maxIdx = Math.max(0, Number(maxByLetter[letter]      || 0));
    const count  = Number(groupCountByLetter[letter]           || 3);
    const ranges = rangesByLetter[letter] || makeDefaultRanges(maxIdx, count);
    const prefix = prefixByLetter[letter]                      || `${letter}R`;

    for (let idx = 1; idx <= maxIdx; idx++) {
      let bucket = 1;
      for (let b = 1; b <= count; b++) {
        const s = clamp(ranges[b].start || 1,    1, maxIdx);
        const e = clamp(ranges[b].end   || maxIdx, 1, maxIdx);
        if (idx >= s && idx <= e) { bucket = b; break; }
      }
      mapping[`${letter}${idx}L`] = `${prefix}${bucket}L`;
      mapping[`${letter}${idx}R`] = `${prefix}${bucket}R`;
    }
  }

  return mapping;
}

/**
 * Return all valid groupId strings for the current configuration,
 * plus "CUSTOM_L" / "CUSTOM_R" at the end.
 */
export function getGroupOptions(prefixByLetter, groupCountByLetter) {
  const out = [];
  for (const letter of ["A", "B", "C", "D"]) {
    const prefix = prefixByLetter[letter] || `${letter}R`;
    const count  = Number(groupCountByLetter[letter] || 3);
    for (let b = 1; b <= count; b++) {
      out.push(`${prefix}${b}L`, `${prefix}${b}R`);
    }
  }
  out.push("CUSTOM_L", "CUSTOM_R");
  return out;
}
