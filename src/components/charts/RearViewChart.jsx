import React from "react";
import { LOOP_TYPES } from "../../utils/constants.js";
import { severity } from "../../utils/trim.js";

export function RearViewChart({ rows, tolerance, height, loopTypes, groupLoopChange, setGroupLoopChange }) {
  const width = 1240;
  const heightPx = Number.isFinite(Number(height)) ? Number(height) : 460;
  const pad = 24;
  const tol = Number.isFinite(Number(tolerance)) ? Number(tolerance) : 0;

  const [hover, setHover] = React.useState(null);
  const [showWingOutline, setShowWingOutline] = React.useState(true);
  const [showBeforePoints, setShowBeforePoints] = React.useState(false);
  const [showGroupCuts, setShowGroupCuts] = React.useState(true);
  const [spanMode, setSpanMode] = React.useState("real");
  const [pickedGroupId, setPickedGroupId] = React.useState(null);

  const baselineLoopByGroupKey = React.useMemo(() => {
    const out = {};
    if (!Array.isArray(rows)) return out;
    for (let i = 0; i < rows.length; i++) {
      const rr = rows[i];
      const raw = rr && (rr.groupId || rr.group) ? String(rr.groupId || rr.group) : "";
      const key = raw ? (raw.split("|")[0] || "") : "";
      if (!key) continue;
      const base = rr && rr.baseLoop ? String(rr.baseLoop) : "";
      if (base && !out[key]) out[key] = base;
    }
    return out;
  }, [rows]);

  // Build per-cascade points (A/B/C/D rows) from Step 4 computed per-line rows.
  // IMPORTANT: This chart ONLY uses Step 4 computed rows (frozen baseline + overrides + adjustments),
  // never Step 3 live state.
  const data = React.useMemo(() => {
    if (!Array.isArray(rows) || rows.length === 0) return null;

    const byKey = new Map();
    for (const rr of rows) {
      const letter = String(rr.letter || "").toUpperCase();
      if ((["A", "B", "C", "D"].indexOf(letter) === -1)) continue;
      const idx = Number(rr.idx);
      if (!Number.isFinite(idx)) continue;

      const key = `${letter}|${idx}`;
      let p = byKey.get(key);
      if (!p) {
        p = {
          letter,
          idx,
          lineId: `${letter}${idx}`,
          groupNameL: "—",
          groupNameR: "—",
          beforeL: null,
          beforeR: null,
          afterL: null,
          afterR: null,
          lineIdL: null,
          lineIdR: null,
        };
        byKey.set(key, p);
      }

      const side = String(rr.side || "").toUpperCase();
      const groupName = String(rr.groupId || rr.group || "—").split("|")[0] || "—";
      if (side === "L") p.groupNameL = groupName;
      else if (side === "R") p.groupNameR = groupName;
      const toNumOrNull = (v) => {
        if (v === null || v === undefined) return null;
        const s = String(v).trim();
        if (s === "") return null;
        const n = Number(s);
        return Number.isFinite(n) ? n : null;
      };

      const nominal = (() => {
        const n = toNumOrNull(rr.nominal);
        return n === 0 ? null : n;
      })();
      const beforeAbs = (() => {
        const n = toNumOrNull(rr.before);
        return n === 0 ? null : n;
      })();
      const afterAbs = (() => {
        const n = toNumOrNull(rr.after);
        return n === 0 ? null : n;
      })();

      const beforeDelta = nominal == null || beforeAbs == null ? null : beforeAbs - nominal;
      const afterDelta = nominal == null || afterAbs == null ? null : afterAbs - nominal;

      if (side === "L") {
        p.lineIdL = rr.lineId || p.lineIdL;
        p.beforeL = beforeDelta;
        p.afterL = afterDelta;
      } else if (side === "R") {
        p.lineIdR = rr.lineId || p.lineIdR;
        p.beforeR = beforeDelta;
        p.afterR = afterDelta;
      }
    }

    const points = Array.from(byKey.values()).sort((a, b) => {
      const la = a.letter.localeCompare(b.letter);
      if (la) return la;
      return a.idx - b.idx;
    });

    const byLetter = { A: [], B: [], C: [], D: [] };
    for (const p of points) {
      if (byLetter[p.letter]) byLetter[p.letter].push(p);
    }
    return { points, ...byLetter };
  }, [rows]);



  if (!data) {
    return (
      <div
        style={{
          padding: 12,
          border: "1px solid #2a2f3f",
          borderRadius: 14,
          background: "#0e1018",
          color: "#aab1c3",
          fontSize: 12,
        }}
      >
        Rear view chart will appear after importing a file.
      </div>
    );
  }

  var bands = {
    A: { y0: pad + 74, y1: pad + 74 + 85 },
    B: { y0: pad + 74 + 95, y1: pad + 74 + 180 },
    C: { y0: pad + 74 + 190, y1: pad + 74 + 275 },
    D: { y0: pad + 74 + 285, y1: pad + 74 + 370 },
  };

  function sevColor(sev) {
    if (sev === "red") return "rgba(255,90,90,1)";
    if (sev === "yellow") return "rgba(255,215,90,1)";
    return "rgba(140,255,190,1)";
  }

  function bandY(letter, v) {
    const b = bands[letter];
    const range = Math.max(30, tol > 0 ? tol * 2.2 : 50);
    const mid = (b.y0 + b.y1) / 2;
    const pxPerMm = (b.y1 - b.y0) / (range * 2);
    return mid - v * pxPerMm;
  }

  function spanScale(t) {
    if (spanMode === "linear") return t;
    const gamma = 0.75; // <1 expands inner, compresses tips
    return Math.pow(t, gamma);
  }

  function xFor(side, i, count) {
    const center = width / 2;
    const halfSpan = (width - pad * 2) / 2 - 20;
    const centerGap = 18;

    const countFixed = 25;

    const t = countFixed <= 1 ? 0 : i / (countFixed - 1);
    const ts = spanScale(t);
    const dx = ts * halfSpan + centerGap;

    return side === "L" ? center - dx : center + dx;
  }

  
function groupCuts(letter) {
  if (!showGroupCuts) return [];
  const arr = data[letter] || [];
  const out = [];
  let last = null;
  for (let i = 0; i < arr.length; i++) {
    const g = (arr[i] && (arr[i].groupNameL || arr[i].groupNameR)) || "";
    if (i === 0) {
      last = g;
      continue;
    }
    if (g !== last) {
      out.push({ idx: i - 0.5, from: last, to: g });
      last = g;
    }
  }
  return out;
}

function groupBands(letter, side) {
  if (!showGroupCuts) return [];
  const arr = data[letter] || [];
  const key = side === "L" ? "groupNameL" : "groupNameR";
  const out = [];
  let start = 0;
  let current = null;
  for (let i = 0; i < arr.length; i++) {
    const g = (arr[i] && arr[i][key]) || "—";
    if (i === 0) {
      current = g;
      start = 0;
      continue;
    }
    if (g !== current) {
      out.push({ start, end: i - 1, name: current });
      current = g;
      start = i;
    }
  }
  if (arr.length > 0) out.push({ start, end: arr.length - 1, name: current || "—" });
  return out;
}

  return (
    <div style={{ border: "1px solid #2a2f3f", borderRadius: 14, padding: 12, background: "#0e1018" }}>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start", flexWrap: "wrap" }}>
        <div>
          <div style={{ fontWeight: 900, marginBottom: 6 }}>Rear view wing shape (A/B/C/D rows)</div>
          <div style={{ color: "#aab1c3", fontSize: 12, lineHeight: 1.5 }}>
            Symmetric about the centreline. Points are <b>After</b> (severity color). Dashed = Before.
          </div>
        </div>

        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <label style={{ color: "#aab1c3", fontSize: 12, display: "flex", gap: 8, alignItems: "center" }}>
            Span spacing
            <select
              value={spanMode}
              onChange={(e) => setSpanMode(e.target.value)}
              style={{
                borderRadius: 10,
                border: "1px solid #2a2f3f",
                background: "#0d0f16",
                color: "#eef1ff",
                padding: "6px 10px",
                outline: "none",
                                              userSelect: "auto",
                fontSize: 12,
              }}
            >
              <option value="real">Realistic</option>
              <option value="linear">Linear</option>
            </select>
          </label>

          <label style={{ color: "#aab1c3", fontSize: 12, display: "flex", gap: 8, alignItems: "center" }}>
            <input type="checkbox" checked={showGroupCuts} onChange={(e) => setShowGroupCuts(e.target.checked)} />
            Group boundaries
          </label>

          <label style={{ color: "#aab1c3", fontSize: 12, display: "flex", gap: 8, alignItems: "center" }}>
            <input type="checkbox" checked={showBeforePoints} onChange={(e) => setShowBeforePoints(e.target.checked)} />
            Before points
          </label>
        </div>
      </div>

      <div style={{ height: 10 }} />

      <div style={{ overflowX: "auto", maxWidth: "100%" }}>
        <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} style={{ display: "block" }}>
          {/* Top labels */}
          <text x={pad} y={pad + 16} fill="rgba(170,177,195,0.9)" fontSize="12" fontFamily="ui-monospace, Menlo, Consolas, monospace">
            LEFT
          </text>
          <text
            x={width - pad}
            y={pad + 16}
            textAnchor="end"
            fill="rgba(170,177,195,0.9)"
            fontSize="12"
            fontFamily="ui-monospace, Menlo, Consolas, monospace"
          >
            RIGHT
          </text>
          <text
            x={width / 2}
            y={pad + 16}
            textAnchor="middle"
            fill="rgba(170,177,195,0.9)"
            fontSize="12"
            fontFamily="ui-monospace, Menlo, Consolas, monospace"
          >
            CENTRE
          </text>

          {/* Centreline */}
          <line x1={width / 2} y1={pad + 24} x2={width / 2} y2={height - pad} stroke="rgba(42,47,63,0.85)" strokeDasharray="6 6" />

          {/* Span ticks + MID→TIP labels */}
          {(function () {
            const y = pad + 30;
            const center = width / 2;
            const halfSpan = (width - pad * 2) / 2 - 20;
            const centerGap = 18;
            const ticks = [
              { t: 0.0, label: "MID" },
              { t: 0.25, label: "25%" },
              { t: 0.5, label: "50%" },
              { t: 0.75, label: "75%" },
              { t: 1.0, label: "TIP" },
            ];
            const scaleT = (t) => (spanMode === "linear" ? t : Math.pow(t, 0.75));

            return (
              <g>
                {ticks.map((tk) => {
                  const dx = scaleT(tk.t) * halfSpan + centerGap;
                  const xL = center - dx;
                  const xR = center + dx;
                  return (
                    <g key={`ticks-${tk.t}`}>
                      <line x1={xL} y1={y} x2={xL} y2={y + 8} stroke="rgba(255,255,255,0.10)" />
                      <line x1={xR} y1={y} x2={xR} y2={y + 8} stroke="rgba(255,255,255,0.10)" />
                      <text x={xL} y={y + 22} textAnchor="middle" fill="rgba(170,177,195,0.85)" fontSize="11" fontFamily="ui-monospace, Menlo, Consolas, monospace">
                        {tk.label}
                      </text>
                      <text x={xR} y={y + 22} textAnchor="middle" fill="rgba(170,177,195,0.85)" fontSize="11" fontFamily="ui-monospace, Menlo, Consolas, monospace">
                        {tk.label}
                      </text>
                    </g>
                  );
                })}
              </g>
            );
          })()}

          {/* Subtle wing outline arc (background) */}
          {(function () {
            const left = pad + 20;
            const right = width - pad - 20;
            const top = pad + 66;
            const bottom = height - pad - 18;
            const midX = width / 2;
            const ctrlY = top - 26;

            const d = `
              M ${midX} ${top}
              C ${midX - 180} ${ctrlY}, ${left + 60} ${ctrlY + 10}, ${left} ${top + 18}
              L ${left} ${bottom}
              L ${right} ${bottom}
              L ${right} ${top + 18}
              C ${right - 60} ${ctrlY + 10}, ${midX + 180} ${ctrlY}, ${midX} ${top}
              Z
            `;

            return <path d={d} fill="rgba(255,255,255,0.015)" stroke="rgba(255,255,255,0.06)" strokeWidth="2" />;
          })()}

          {/* Bands + 0mm guides + riser labels */}
          {["A", "B", "C", "D"].map((L) => {
            const b = bands[L];
            const yMid = (b.y0 + b.y1) / 2;

            return (
              <g key={`band-${L}`}>
                <rect x={pad} y={b.y0} width={width - pad * 2} height={b.y1 - b.y0} fill="rgba(255,255,255,0.02)" />
                <line x1={pad} y1={b.y0} x2={width - pad} y2={b.y0} stroke="rgba(42,47,63,0.85)" />
                <line x1={pad} y1={b.y1} x2={width - pad} y2={b.y1} stroke="rgba(42,47,63,0.85)" />

                {/* 0mm guide (target) */}
                <line x1={pad} y1={yMid} x2={width - pad} y2={yMid} stroke="rgba(255,90,90,0.85)" strokeDasharray="4 6" />

                {/* Row label */}
                <text x={pad + 8} y={b.y0 + 18} fill="rgba(170,177,195,0.85)" fontSize="12" fontFamily="ui-monospace, Menlo, Consolas, monospace">
                  {L}-row
                </text>

                {/* Riser label at centreline */}
                <text
                  x={width / 2}
                  y={b.y0 + 18}
                  textAnchor="middle"
                  fill="rgba(238,241,255,0.85)"
                  fontSize="12"
                  fontFamily="ui-monospace, Menlo, Consolas, monospace"
                >
                  {L}
                </text>
              </g>
            );
          })}

          {/* Plots */}
          {["A", "B", "C", "D"].map((L) => {
            const arr = data[L] || [];
            const count = arr.length || 1;

            const buildPath = (side, which) => {
              let d = "";
              for (let i = 0; i < arr.length; i++) {
                const p = arr[i];
                const v =
                  side === "L"
                    ? which === "before"
                      ? p.beforeL
                      : p.afterL
                    : which === "before"
                    ? p.beforeR
                    : p.afterR;

                if (!Number.isFinite(v)) continue;
                const x = xFor(side, i, count);
                const y = bandY(L, v);
                d += d ? ` L ${x} ${y}` : `M ${x} ${y}`;
              }
              return d;
            };

            const cuts = groupCuts(L);

            return (
              <g key={`plot-${L}`}>
                {/* group boundary lines (both sides) */}
                {cuts.map((c, idx) => {
                  const b = bands[L];
                  const xL = xFor("L", c.idx, count);
                  const xR = xFor("R", c.idx, count);
                  return (
                    <g key={`cut-${L}-${idx}`}>
                      <line x1={xL} y1={b.y0 + 2} x2={xL} y2={b.y1 - 2} stroke="rgba(255,220,80,0.85)" />
                      <line x1={xR} y1={b.y0 + 2} x2={xR} y2={b.y1 - 2} stroke="rgba(255,220,80,0.85)" />
                    </g>
                  );
                })}


                {/* Group labels (midpoint of each grouping) */}
                {showGroupCuts &&
                  (() => {
                    const b = bands[L];
                    var bandsL = groupBands(L, "L");
                    var bandsR = groupBands(L, "R");
                    const yText = ((b.y0 + b.y1) / 2) + 38;
                    const renderBand = (side, band, i) => {
                      if (!band || !band.name || band.name === "—") return null;
                      const mid = (band.start + band.end) / 2;
                      const x = xFor(side, mid, count);
                      const key = String((band && band.name) || "");
                      const displayName = key.replace(/([LR])$/, "");
                      const baseLoop = baselineLoopByGroupKey && baselineLoopByGroupKey[key] ? baselineLoopByGroupKey[key] : "";
                      // Cosmetic tooltip: show the grouping key AND the Step 3 baseline loop (reference only).
                      // Use newlines so it reads clearly in the native title tooltip.
                      const titleText = key + "\nBaseline: " + (baseLoop ? baseLoop : "—");
                      const cur = (groupLoopChange && groupLoopChange[key]) ? groupLoopChange[key] : "";
                      const w = 68;
                      const h = 22;
                      return (
                        <foreignObject
                          key={`glabel-${L}-${side}-${i}`}
                          x={x - w / 2}
                          y={yText - h + 6}
                          width={w}
                          height={h}
                        >
                          <div
                            xmlns="http://www.w3.org/1999/xhtml"
                            style={{
                              width: "100%",
                              height: "100%",
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "center",
                              borderRadius: 999,
                              border: "1px solid rgba(255,220,80,0.55)",
                              background: "rgba(0,0,0,0.45)",
                              boxSizing: "border-box",
                            }}
                            title={titleText}
                          >
                            <select
                              value={cur}
                              onChange={(e) => {
                                const v = e.target.value || "";
                                if (!setGroupLoopChange) return;
                                setGroupLoopChange(function (prev) {
                                  const next = Object.assign({}, prev || {});
                                  if (v) next[key] = v;
                                  else delete next[key];
                                  return next;
                                });
                              }}
                              style={{
                                width: "100%",
                                height: "100%",
                                border: "none",
                                outline: "none",
                                              userSelect: "auto",
                                background: "transparent",
                                color: (cur ? "rgba(120,200,255,0.95)" : "rgba(255,220,80,0.95)"),
                                fontSize: 11,
                                fontWeight: 950,
                                fontFamily: "ui-monospace, Menlo, Consolas, monospace",
                                textAlignLast: "center",
                                cursor: "pointer",
                                paddingLeft: 6,
                                paddingRight: 6,
                                appearance: "none",
                                WebkitAppearance: "none",
                                MozAppearance: "none",
                              }}
                            >
                              <option value="">{displayName}</option>
                              {((loopTypes && loopTypes.length) ? loopTypes : LOOP_TYPES).map((lt) => (
                                <option key={lt} value={lt}>
                                  {lt}
                                </option>
                              ))}
                            </select>
                          </div>
                        </foreignObject>
                      );
                    };
                    return (
                      <g>
                        {bandsL.map((band, i) => renderBand("L", band, i))}
                        {bandsR.map((band, i) => renderBand("R", band, i))}
                      </g>
                    );
                  })()}
                {/* Before dashed paths */}
                <path d={buildPath("L", "before")} fill="none" stroke="rgba(176,132,255,0.65)" strokeWidth="2" strokeDasharray="6 6" />
                <path d={buildPath("R", "before")} fill="none" stroke="rgba(102,204,255,0.65)" strokeWidth="2" strokeDasharray="6 6" />

                {/* After solid paths */}
                <path d={buildPath("L", "after")} fill="none" stroke="rgba(176,132,255,1)" strokeWidth="3" />
                <path d={buildPath("R", "after")} fill="none" stroke="rgba(102,204,255,1)" strokeWidth="3" />

                {/* Points */}
                {arr.map((p, i) => {
                  const pts = [
                    { side: "L", before: p.beforeL, after: p.afterL },
                    { side: "R", before: p.beforeR, after: p.afterR },
                  ];

                  return pts.map((it) => {
                    const x = xFor(it.side, i, count);

                    // BEFORE points (small hollow circles)
                    const beforeNode =
                      showBeforePoints && Number.isFinite(it.before) ? (
                        <circle
                          key={`${p.lineId}-${it.side}-before`}
                          cx={x}
                          cy={bandY(L, it.before)}
                          r={4}
                          fill="transparent"
                          stroke="rgba(255,255,255,0.30)"
                          strokeWidth="2"
                        />
                      ) : null;

                    // AFTER points (colored)
                    const afterNode = Number.isFinite(it.after) ? (
                      <circle
                        key={`${p.lineId}-${it.side}-after`}
                        cx={x}
                        cy={bandY(L, it.after)}
                        r={5}
                        fill={sevColor(severity(it.after, tol))}
                        stroke="rgba(10,12,16,0.9)"
                        strokeWidth="2"
                        onMouseEnter={() =>
                          setHover({
                            letter: L,
                            lineId: p.lineId,
                            groupName: it.side === "L" ? p.groupNameL : p.groupNameR,
                            side: it.side,
                            before: it.before,
                            after: it.after,
                            sev: severity(it.after, tol),
                            x,
                            y: bandY(L, it.after),
                          })
                        }
                        onMouseLeave={() => setHover(null)}
                      />
                    ) : null;

                    return (
                      <g key={`${p.lineId}-${it.side}`}>
                        {beforeNode}
                        {afterNode}
                      </g>
                    );
                  });
                })}
              </g>
            );
          })}

          {/* Tooltip */}
          {hover ? (
            <g>
              <rect
                x={Math.min(width - 330, Math.max(10, hover.x + 12))}
                y={Math.max(10, hover.y - 80)}
                width={320}
                height={70}
                rx={10}
                ry={10}
                fill="rgba(12,14,22,0.95)"
                stroke="rgba(42,47,63,1)"
              />
              <text
                x={Math.min(width - 312, Math.max(20, hover.x + 22))}
                y={Math.max(28, hover.y - 52)}
                fill="#eef1ff"
                fontSize="12"
                fontFamily="ui-monospace, Menlo, Consolas, monospace"
              >
                {`${hover.lineId} (${hover.side})  group: ${hover.groupName}`}
              </text>
              <text
                x={Math.min(width - 312, Math.max(20, hover.x + 22))}
                y={Math.max(48, hover.y - 32)}
                fill="rgba(170,177,195,0.95)"
                fontSize="12"
                fontFamily="ui-monospace, Menlo, Consolas, monospace"
              >
                {`Before: ${Number.isFinite(hover.before) ? Math.round(hover.before) : "—"}mm   After: ${
                  Number.isFinite(hover.after) ? Math.round(hover.after) : "—"
                }mm   Sev: ${hover.sev}`}
              </text>
            </g>
          ) : null}
        </svg>
      </div>

      <div style={{ color: "#aab1c3", fontSize: 12, marginTop: 8, display: "flex", gap: 14, flexWrap: "wrap" }}>
        <span>Solid = After</span>
        <span>Dashed = Before</span>
        <span>Target (0mm) = dotted line</span>
        {tol > 0 ? (
          <>
            <span>Yellow = within 3mm of tolerance</span>
            <span>Red = outside tolerance</span>
          </>
        ) : null}
      </div>
    </div>
  );
}




