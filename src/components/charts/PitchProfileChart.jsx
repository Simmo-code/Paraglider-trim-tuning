import React from "react";

/**
 * PitchProfileChart — bar chart showing average Δ per line group (A/B/C/D),
 * split Left (top) / Right (bottom), colour-coded by severity.
 */
export function PitchProfileChart({ pitchStats, tolerance = 6, height = 260 }) {
  if (!pitchStats || !pitchStats.rows) return null;

  const w = 560;
  const h = height;
  const padL = 36;   // left label
  const padR = 16;
  const padTop = 32;
  const padBot = 32;
  const innerH = h - padTop - padBot;
  const halfH = innerH / 2;  // each side gets half
  const barAreaW = w - padL - padR;

  const letters = ["A", "B", "C", "D"];
  const rows = pitchStats.rows; // [{ letter, L, R, both }]
  const byLetter = {};
  for (const r of rows) byLetter[r.letter] = r;

  // Collect all finite values to compute scale
  const allVals = [];
  for (const L of letters) {
    const r = byLetter[L];
    if (!r) continue;
    if (Number.isFinite(r.L)) allVals.push(Math.abs(r.L));
    if (Number.isFinite(r.R)) allVals.push(Math.abs(r.R));
  }
  const maxAbs = Math.max(tolerance * 1.5, ...allVals, 1);

  const colFor = (val) => {
    if (!Number.isFinite(val)) return "rgba(255,255,255,0.18)";
    const a = Math.abs(val);
    if (a <= 4) return "rgba(34,197,94,0.85)";
    if (a < tolerance) return "rgba(234,179,8,0.85)";
    return "rgba(239,68,68,0.85)";
  };

  // Map a mm value to x pixel (0 = centre line)
  const toX = (mm) => {
    if (!Number.isFinite(mm)) return padL + barAreaW / 2;
    return padL + barAreaW / 2 + (mm / maxAbs) * (barAreaW / 2);
  };
  const centreX = padL + barAreaW / 2;

  const barH = Math.min(18, (halfH / letters.length) - 6);
  const groupH = innerH / letters.length;

  // Tolerance lines x positions
  const tolPosX = toX(tolerance);
  const tolNegX = toX(-tolerance);

  return (
    <div style={{ width: "100%" }}>
      <svg width="100%" viewBox={`0 0 ${w} ${h}`} style={{ display: "block" }}>

        {/* Background bands */}
        <rect x={tolNegX} y={padTop} width={tolPosX - tolNegX} height={innerH}
          fill="rgba(34,197,94,0.05)" />
        <rect x={padL} y={padTop} width={tolNegX - padL} height={innerH}
          fill="rgba(239,68,68,0.05)" />
        <rect x={tolPosX} y={padTop} width={w - padR - tolPosX} height={innerH}
          fill="rgba(239,68,68,0.05)" />

        {/* Tolerance lines */}
        <line x1={tolPosX} y1={padTop} x2={tolPosX} y2={padTop + innerH}
          stroke="rgba(239,68,68,0.35)" strokeDasharray="4 3" />
        <line x1={tolNegX} y1={padTop} x2={tolNegX} y2={padTop + innerH}
          stroke="rgba(239,68,68,0.35)" strokeDasharray="4 3" />

        {/* Centre line */}
        <line x1={centreX} y1={padTop - 8} x2={centreX} y2={padTop + innerH + 8}
          stroke="rgba(255,255,255,0.25)" strokeWidth={1} />

        {/* Bars per letter */}
        {letters.map((L, i) => {
          const r = byLetter[L];
          const groupY = padTop + i * groupH;
          const midY = groupY + groupH / 2;
          const Lval = r ? r.L : null;
          const Rval = r ? r.R : null;

          const barY_L = midY - barH - 2;
          const barY_R = midY + 2;

          const drawBar = (val, barY, label) => {
            if (!Number.isFinite(val)) return null;
            const x1 = val >= 0 ? centreX : toX(val);
            const x2 = val >= 0 ? toX(val) : centreX;
            const bw = Math.max(2, x2 - x1);
            const col = colFor(val);
            return (
              <g key={label}>
                <rect x={x1} y={barY} width={bw} height={barH} fill={col} rx={3} />
                <text
                  x={val >= 0 ? toX(val) + 4 : toX(val) - 4}
                  y={barY + barH * 0.72}
                  fontSize={10}
                  fill={col}
                  fontWeight={900}
                  textAnchor={val >= 0 ? "start" : "end"}
                >
                  {val > 0 ? "+" : ""}{val.toFixed(1)}
                </text>
              </g>
            );
          };

          return (
            <g key={L}>
              {/* Letter label */}
              <text x={padL - 6} y={midY + 4} fontSize={13} fill="rgba(255,255,255,0.7)"
                fontWeight={950} textAnchor="end">{L}</text>

              {/* Divider */}
              {i > 0 && (
                <line x1={padL} y1={groupY} x2={w - padR} y2={groupY}
                  stroke="rgba(255,255,255,0.07)" />
              )}

              {drawBar(Lval, barY_L, `${L}L`)}
              {drawBar(Rval, barY_R, `${L}R`)}

              {/* L/R mini labels */}
              <text x={centreX - 3} y={barY_L + barH * 0.76} fontSize={8}
                fill="rgba(255,255,255,0.4)" textAnchor="end">L</text>
              <text x={centreX - 3} y={barY_R + barH * 0.76} fontSize={8}
                fill="rgba(255,255,255,0.4)" textAnchor="end">R</text>
            </g>
          );
        })}

        {/* X axis labels */}
        <text x={centreX} y={h - 4} fontSize={10} fill="rgba(255,255,255,0.45)"
          textAnchor="middle">0</text>
        <text x={tolPosX} y={h - 4} fontSize={10} fill="rgba(239,68,68,0.65)"
          textAnchor="middle">+{tolerance}</text>
        <text x={tolNegX} y={h - 4} fontSize={10} fill="rgba(239,68,68,0.65)"
          textAnchor="middle">-{tolerance}</text>

        {/* Title */}
        <text x={padL} y={18} fontSize={11} fill="rgba(255,255,255,0.55)" fontWeight={900}>
          Avg Δ per line group (mm) — Left / Right
        </text>

      </svg>
      <div style={{ display: "flex", gap: 14, marginTop: 4, fontSize: 11, opacity: 0.65, flexWrap: "wrap" }}>
        <span>🟢 ≤4mm</span>
        <span>🟡 4mm–tol</span>
        <span>🔴 &gt;tol</span>
        <span style={{ marginLeft: "auto" }}>Top bar = Left · Bottom bar = Right</span>
      </div>
    </div>
  );
}
