import React from "react";
import { clamp } from "../../utils/math.js";
import { severity } from "../../utils/trim.js";

export function WingPitchViz({ pitchMm, tolMm, height = 160 }) {
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

