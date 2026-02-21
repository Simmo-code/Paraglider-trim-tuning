import React from "react";
import { theme } from "../../utils/constants.js";
import { severity } from "../../utils/trim.js";

export function DeltaLineChart({ title, points, tolerance, height = 240 }) {
  const w = 980;
  const pad = { l: 38, r: 16, t: 26, b: 28 };

  const safeTol = Number.isFinite(tolerance) ? Math.max(0, tolerance) : 0;
  const yMax = Math.max(safeTol + 6, 25);

  const xMin = 0;
  const xMax = Math.max(1, ...points.map((p) => Number(p.xIndex) || 0));

  const xScale = (x) => pad.l + ((x - xMin) / (xMax - xMin || 1)) * (w - pad.l - pad.r);
  const yScale = (y) => pad.t + ((yMax - y) / (2 * yMax)) * (height - pad.t - pad.b);

  const sevColor = (sev) => {
    if (sev === "red") return "rgba(239,68,68,0.95)";
    if (sev === "yellow") return "rgba(245,158,11,0.95)";
    if (sev === "ok") return "rgba(34,197,94,0.95)";
    return "rgba(255,255,255,0.55)";
  };

  const pathFor = (kind) => {
    const pts = points
      .filter((p) => Number.isFinite(p[kind]))
      .sort((a, b) => (a.xIndex || 0) - (b.xIndex || 0));
    if (!pts.length) return "";
    return pts
      .map((p, i) => {
        const x = xScale(p.xIndex || 0);
        const y = yScale(p[kind]);
        return `${i === 0 ? "M" : "L"} ${x.toFixed(2)} ${y.toFixed(2)}`;
      })
      .join(" ");
  };

  const gridLines = [];
  for (let i = -yMax; i <= yMax; i += 5) {
    gridLines.push(i);
  }

  return (
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel, overflow: "hidden" }}>
      <div style={{ padding: 8, borderBottom: `1px solid ${theme.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ fontWeight: 950 }}>{title}</div>
        <div style={{ opacity: 0.75, fontSize: 12 }}>Δ(mm) vs nominal • green ≤ 4mm • yellow &gt; 4mm • red ≥ tolerance</div>
      </div>
      <div style={{ padding: 8, overflowX: "auto", maxWidth: "100%" }}>
        <svg viewBox={`0 0 ${w} ${height}`} style={{ width: "100%", height: "auto", display: "block" }}>
          {/* Grid */}
          {gridLines.map((g) => (
            <line
              key={g}
              x1={pad.l}
              x2={w - pad.r}
              y1={yScale(g)}
              y2={yScale(g)}
              stroke="rgba(255,255,255,0.07)"
              strokeWidth={1}
            />
          ))}
          {/* Axis */}
          <line x1={pad.l} x2={w - pad.r} y1={yScale(0)} y2={yScale(0)} stroke="rgba(255,255,255,0.22)" strokeWidth={1} />
          <line x1={pad.l} x2={pad.l} y1={pad.t} y2={height - pad.b} stroke="rgba(255,255,255,0.18)" strokeWidth={1} />

          {/* Tolerance bands */}
          {safeTol > 0 ? (
            <>
              <rect
                x={pad.l}
                y={yScale(safeTol)}
                width={w - pad.l - pad.r}
                height={yScale(4) - yScale(safeTol)}
                fill="rgba(245,158,11,0.08)"
              />
              <rect
                x={pad.l}
                y={yScale(-4)}
                width={w - pad.l - pad.r}
                height={yScale(-safeTol) - yScale(-4)}
                fill="rgba(245,158,11,0.08)"
              />
              <rect
                x={pad.l}
                y={yScale(4)}
                width={w - pad.l - pad.r}
                height={yScale(-4) - yScale(4)}
                fill="rgba(34,197,94,0.06)"
              />
            </>
          ) : (
            <rect x={pad.l} y={yScale(4)} width={w - pad.l - pad.r} height={yScale(-4) - yScale(4)} fill="rgba(34,197,94,0.06)" />
          )}

          {/* Lines */}
          <path d={pathFor("before")} fill="none" stroke="rgba(148,163,184,0.85)" strokeWidth={2} strokeDasharray="6 4" />
          <path d={pathFor("after")} fill="none" stroke="rgba(59,130,246,0.90)" strokeWidth={2.5} />

          {/* Points */}
          {points
            .filter((p) => Number.isFinite(p.after))
            .map((p) => (
              <circle
                key={`${p.id}-a`}
                cx={xScale(p.xIndex || 0)}
                cy={yScale(p.after)}
                r={4}
                fill={sevColor(p.sevAfter)}
                stroke="rgba(0,0,0,0.4)"
                strokeWidth={1}
              >
                <title>{`${p.line}: after Δ=${Math.round(p.after)}mm`}</title>
              </circle>
            ))}

          {/* Labels */}
          <text x={8} y={14} fill="rgba(255,255,255,0.70)" fontSize={11} fontWeight={900}>
            Δ(mm)
          </text>
        </svg>
      </div>
    </div>
  );
}

