// ── Site version ─────────────────────────────────────────────────────────────
export const SITE_VERSION = "Trim Tuning v1.2";

// ── Loop sizes (mm) – wing-specific, set before baseline loops ───────────────
export const DEFAULT_LOOP_SIZES = {
  SL: 0,
  DL: -7,
  TL: -10,
  AS: -12,
  "AS+": -16,
  "AS++": -20,
  CUSTOM: 0,
};
export const LOOP_TYPES = Object.keys(DEFAULT_LOOP_SIZES);

// ── Theme ─────────────────────────────────────────────────────────────────────
export const theme = {
  bg: "#12151b",
  panel: "rgba(255,255,255,0.08)",
  panel2: "rgba(0,0,0,0.22)",
  bg2: "rgba(0,0,0,0.22)",
  border: "rgba(255,255,255,0.14)",
  text: "rgba(255,255,255,0.92)",
  textSub: "rgba(170,177,195,0.85)",
  green: "rgba(34,197,94,0.95)",
  good: "rgba(34,197,94,0.95)",
  bad: "rgba(239,68,68,0.95)",
  warn: "rgba(245,158,11,0.95)",
  warnBg: "rgba(245,158,11,0.10)",
  warnStroke: "rgba(245,158,11,0.55)",
};

// ── Colour palette per riser group ────────────────────────────────────────────
export const PALETTE = {
  A: { base: "#1e6eff", s2: "#2d7fff", s3: "#5aa6ff", s4: "#86c4ff" },
  B: { base: "#8b5cf6", s2: "#9d73ff", s3: "#b69bff", s4: "#cdbbff" },
  C: { base: "#ff9f43", s2: "#ffb76b", s3: "#ffd1a3", s4: "#ffe2c7" },
  D: { base: "#facc15", s2: "#fde047", s3: "#fef08a", s4: "#fff6bf" },
};

// ── Diagram layout constants ──────────────────────────────────────────────────
export const DIAGRAM_SCALE = 0.9;
export const DIAGRAM_BASE_W = 2400;
export const DIAGRAM_BASE_H = 980;
export const DIAGRAM_W = Math.round(DIAGRAM_BASE_W * DIAGRAM_SCALE);
export const DIAGRAM_H = Math.round(DIAGRAM_BASE_H * DIAGRAM_SCALE);

// ── Bundled example CSV ───────────────────────────────────────────────────────
// Loaded from a real asset file in production; kept here as fallback for dev.
export const ATTACHED_TEST_CSV = `Make ,Model,tolerance ,Korrektur,,,,,,,,,,,,
Ozone,Speedster3,10,-507,,,,,,,,,,,,
A,Soll,Ist L,Ist R,B,Soll,L,R,C,Soll,Ist L,Ist R,D,Soll,L,R
A1,6717,7220,7222,B1,6635,7142,7145,C1,6712,7219,7213,D1,6871,7379,7380
A2,6676,7184,7185,B2,6593,7100,7101,C2,6672,7177,7175,D2,6833,7334,7341
A3,6646,7151,7156,B3,6566,7071,7077,C3,6644,7149,7149,D3,6801,7309,7309
A4,6616,7119,7123,B4,6534,7038,7041,C4,6613,7116,7116,D4,6769,7276,7279
A5,6590,7090,7093,B5,6504,7007,7008,C5,6587,7088,7088,D5,6742,7247,7247
A6,6563,7060,7060,B6,6473,6978,6975,C6,6560,7059,7056,D6,6716,7217,7213
A7,6532,7026,7028,B7,6441,6944,6948,C7,6529,7027,7027,D7,6683,7184,7184
A8,6498,6991,6994,B8,6408,6909,6914,C8,6495,6991,6991,D8,6648,7149,7151
A9,6468,6958,6958,B9,6375,6879,6879,C9,6465,6956,6957,D9,6616,7118,7118
A10,6440,6929,6928,B10,6347,6850,6851,C10,6437,6927,6927,D10,6585,7088,7090
A11,6412,6900,6900,B11,6318,6820,6822,C11,6409,6898,6898,D11,6552,7053,7052
A12,6382,6868,6869,B12,6287,6788,6791,C12,6380,6867,6867,D12,6518,7016,7017
A13,6352,6839,6839,B13,6257,6760,6760,C13,6350,6838,6838,D13,6486,6987,6986
A14,6322,6808,6808,B14,6226,6729,6730,C14,6320,6808,6808,D14,6453,6952,6953
A15,6294,6779,6779,B15,6197,6702,6702,C15,6292,6777,6777,D15,6422,6922,6923
A16,6264,6749,6749,B16,6165,6670,6671,C16,6262,6748,6748,D16,6389,6887,6887`;
