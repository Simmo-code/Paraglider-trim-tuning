export function bandForDelta(delta, tolerance) {
  if (delta == null || !Number.isFinite(Number(delta))) return "na";
  const abs = Math.abs(Number(delta));
  const tol = Number.isFinite(Number(tolerance)) && Number(tolerance) > 0 ? Number(tolerance) : 10;
  if (abs <= 4) return "good";
  if (abs < tol) return "warn";
  return "bad";
}

export function severity(delta, tolerance) {
  if (delta == null || !Number.isFinite(Number(delta))) return "na";
  const abs = Math.abs(Number(delta));
  const tol = Number.isFinite(Number(tolerance)) && Number(tolerance) > 0 ? Number(tolerance) : 10;

  // Green: within 60% of tolerance
  // Yellow: 60%–100% of tolerance (warning zone)
  // Red: at or beyond tolerance
  if (abs >= tol) return "red";
  if (abs >= tol * 0.6) return "yellow";
  return "green";
}
