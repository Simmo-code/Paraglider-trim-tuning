import React, { useMemo } from "react";
import { theme } from "../../utils/constants.js";
import { bandForDelta } from "../../utils/trim.js";
import { WingPitchViz } from "./WingPitchViz.jsx";

export function PitchTrimChart({ rows, tolerance, height = 220 }) {
  const safeTol = Number.isFinite(Number(tolerance)) ? Number(tolerance) : 0;
  const w = 980;
  const h = height;

  const stats = useMemo(() => {
    const out = [];
    const list = Array.isArray(rows) ? rows : [];
    for (const letter of ["A", "B", "C", "D"]) {
      const L = list.filter((r) => r.letter === letter && r.side === "L" && Number.isFinite(Number(r.delta)));
      const R = list.filter((r) => r.letter === letter && r.side === "R" && Number.isFinite(Number(r.delta)));
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

  var bandFill = (delta) => {
    const b = bandForDelta(delta, safeTol);
    if (b === "good") return "rgba(34,197,94,0.75)";
    if (b === "warn") return "rgba(245,158,11,0.75)";
    if (b === "bad") return "rgba(239,68,68,0.78)";
    return "rgba(148,163,184,0.55)";
  };

  return (
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: "rgba(0,0,0,0.55)", padding: 12 }}>
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
