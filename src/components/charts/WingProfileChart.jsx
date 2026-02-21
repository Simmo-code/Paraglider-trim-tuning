import React from "react";
import { theme } from "../../utils/constants.js";
import { severity } from "../../utils/trim.js";

export function WingProfileChart({ groupStats, tolerance, height = 260 }) {
  const w = 980;
  const pad = { l: 40, r: 16, t: 26, b: 28 };

  const safeTol = Number.isFinite(tolerance) ? Math.max(0, tolerance) : 0;
  const yMax = Math.max(safeTol + 6, 25);

  const groups = Array.from(new Set(groupStats.map((g) => g.groupName))).sort((a, b) => String(a).localeCompare(String(b)));
  const xMin = 0;
  const xMax = Math.max(1, groups.length - 1);

  const xScale = (i) => pad.l + ((i - xMin) / (xMax - xMin || 1)) * (w - pad.l - pad.r);
  const yScale = (y) => pad.t + ((yMax - y) / (2 * yMax)) * (height - pad.t - pad.b);

  const sevColor = (delta) => {
    if (!Number.isFinite(delta)) return "rgba(255,255,255,0.55)";
    const ad = Math.abs(delta);
    if (safeTol > 0 && ad >= safeTol) return "rgba(239,68,68,0.95)";
    if (ad > 4) return "rgba(245,158,11,0.95)";
    return "rgba(34,197,94,0.95)";
  };

  const seriesFor = (side) => {
    const pts = [];
    groups.forEach((gName, i) => {
      const rec = groupStats.find((r) => r.groupName === gName && r.side === side);
      pts.push({ i: i, y: rec && rec.after != null ? rec.after : null, before: rec && rec.before != null ? rec.before : null, groupName: gName });
    });
    return pts;
  };

  const pathFor = (pts, key) => {
    const filtered = pts.filter((p) => Number.isFinite(p[key]));
    if (!filtered.length) return "";
    return filtered
      .map((p, idx) => {
        const x = xScale(p.i);
        const y = yScale(p[key]);
        return `${idx === 0 ? "M" : "L"} ${x.toFixed(2)} ${y.toFixed(2)}`;
      })
      .join(" ");
  };

  const L = seriesFor("L");
  const R = seriesFor("R");

  return (
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, background: theme.panel, overflow: "hidden" }}>
      <div style={{ padding: 8, borderBottom: `1px solid ${theme.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ fontWeight: 950 }}>Δ per maillon group (After)</div>
        <div style={{ opacity: 0.75, fontSize: 12 }}>Green ≤ 4mm • Yellow &gt; 4mm • Red ≥ tolerance</div>
      </div>
      <div style={{ padding: 8, overflowX: "auto", maxWidth: "100%" }}>
        <svg viewBox={`0 0 ${w} ${height}`} style={{ width: "100%", height: "auto", display: "block" }}>
          {/* Axis */}
          <line x1={pad.l} x2={w - pad.r} y1={yScale(0)} y2={yScale(0)} stroke="rgba(255,255,255,0.22)" strokeWidth={1} />
          <line x1={pad.l} x2={pad.l} y1={pad.t} y2={height - pad.b} stroke="rgba(255,255,255,0.18)" strokeWidth={1} />

          {/* Bands */}
          {safeTol > 0 ? (
            <>
              <rect x={pad.l} y={yScale(4)} width={w - pad.l - pad.r} height={yScale(-4) - yScale(4)} fill="rgba(34,197,94,0.06)" />
              <rect x={pad.l} y={yScale(safeTol)} width={w - pad.l - pad.r} height={yScale(4) - yScale(safeTol)} fill="rgba(245,158,11,0.08)" />
              <rect x={pad.l} y={yScale(-4)} width={w - pad.l - pad.r} height={yScale(-safeTol) - yScale(-4)} fill="rgba(245,158,11,0.08)" />
            </>
          ) : (
            <rect x={pad.l} y={yScale(4)} width={w - pad.l - pad.r} height={yScale(-4) - yScale(4)} fill="rgba(34,197,94,0.06)" />
          )}

          {/* Lines */}
          <path d={pathFor(L, "after")} fill="none" stroke="rgba(59,130,246,0.90)" strokeWidth={2.5} />
          <path d={pathFor(R, "after")} fill="none" stroke="rgba(168,85,247,0.90)" strokeWidth={2.5} />

          {/* Points + labels */}
          {groups.map(function (gName, i) {
            var x = xScale(i);
            var lRec = groupStats.find(function (r) { return r && r.groupName === gName && r.side === "L"; });
            var rRec = groupStats.find(function (r) { return r && r.groupName === gName && r.side === "R"; });
            var lAfter = lRec && Number.isFinite(lRec.after) ? lRec.after : null;
            var rAfter = rRec && Number.isFinite(rRec.after) ? rRec.after : null;
            var yL = lAfter != null ? yScale(lAfter) : null;
            var yR = rAfter != null ? yScale(rAfter) : null;

            return (
              <g key={gName}>
                <text x={x} y={height - 10} textAnchor="middle" fill="rgba(255,255,255,0.60)" fontSize={10} fontWeight={900}>
                  {gName}
                </text>
                {yL != null ? (
                  <circle cx={x - 6} cy={yL} r={4.2} fill={sevColor(lAfter)} stroke="rgba(0,0,0,0.45)" strokeWidth={1}>
                    <title>{"".concat(gName, " L: Δ=").concat(Math.round(lAfter), "mm")}</title>
                  </circle>
                ) : null}
                {yR != null ? (
                  <circle cx={x + 6} cy={yR} r={4.2} fill={sevColor(rAfter)} stroke="rgba(0,0,0,0.45)" strokeWidth={1}>
                    <title>{"".concat(gName, " R: Δ=").concat(Math.round(rAfter), "mm")}</title>
                  </circle>
                ) : null}
              </g>
            );
          })}

          <text x={8} y={14} fill="rgba(255,255,255,0.70)" fontSize={11} fontWeight={900}>
            Δ(mm)
          </text>
          <text x={pad.l + 6} y={pad.t + 14} fill="rgba(59,130,246,0.90)" fontSize={11} fontWeight={900}>
            L
          </text>
          <text x={pad.l + 22} y={pad.t + 14} fill="rgba(168,85,247,0.90)" fontSize={11} fontWeight={900}>
            R
          </text>
        </svg>
      </div>
    </div>
  );
}