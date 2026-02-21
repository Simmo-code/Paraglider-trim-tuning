import React from "react";
import { theme } from "../../utils/constants.js";

// ── Layout ────────────────────────────────────────────────────────────────────

export function Panel({ title, right, children, tint = false }) {
  return (
    <div style={{ border: `1px solid ${theme.border}`, borderRadius: 18, background: tint ? "linear-gradient(180deg, rgba(59,130,246,0.08), rgba(255,255,255,0.04))" : theme.panel, overflow: "visible" }}>
      <div style={{ padding: "10px 12px", borderBottom: `1px solid ${theme.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", background: "rgba(255,255,255,0.035)" }}>
        <div style={{ fontWeight: 950, letterSpacing: -0.2, fontSize: 16 }}>{title}</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>{right}</div>
      </div>
      <div style={{ padding: 12 }}>{children}</div>
    </div>
  );
}

export function WarningBanner({ title, children }) {
  return (
    <div style={{ border: `0.1px solid ${theme.warnStroke}`, background: theme.warnBg, borderRadius: 16, padding: 8 }}>
      <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
        <div style={{ width: 10, height: 10, borderRadius: 999, background: theme.warnStroke, marginTop: 5 }} />
        <div style={{ display: "grid", gap: 6 }}>
          <div style={{ fontWeight: 950 }}>{title}</div>
          <div style={{ opacity: 0.82, fontWeight: 800, fontSize: 12 }}>{children}</div>
        </div>
      </div>
    </div>
  );
}

// ── Inputs ────────────────────────────────────────────────────────────────────

export function NumInput({ value, onChange, min = -99999, max = 99999, step = 1, width = 120 }) {
  return (
    <input
      type="number"
      value={value}
      min={min}
      max={max}
      step={step}
      onChange={(e) => onChange(Number(e.target.value))}
      style={{ width, padding: "7px 10px", borderRadius: 12, border: `1px solid ${theme.border}`, background: "rgba(0,0,0,0.68)", color: theme.text, outline: "none", userSelect: "auto", fontWeight: 900, fontSize: 14 }}
    />
  );
}

export function Select({ value, onChange, options, width = 140 }) {
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{ width, padding: "7px 10px", borderRadius: 12, border: `1px solid ${theme.border}`, background: "rgba(0,0,0,0.80)", color: theme.text, outline: "none", userSelect: "auto", fontWeight: 950, fontSize: 14 }}
    >
      {options.map((o) => (
        <option key={o.value} value={o.value} style={{ color: "#111" }}>
          {o.label}
        </option>
      ))}
    </select>
  );
}

export function Toggle({ value, onChange, label }) {
  return (
    <label style={{ display: "inline-flex", alignItems: "center", gap: 10, cursor: "pointer", userSelect: "none" }}>
      <span
        style={{ width: 40, height: 24, borderRadius: 999, border: `1px solid ${theme.border}`, background: value ? "rgba(59,130,246,0.35)" : "rgba(255,255,255,0.08)", position: "relative" }}
        onClick={() => onChange(!value)}
      >
        <span style={{ width: 18, height: 18, borderRadius: 999, background: value ? "rgba(255,255,255,0.92)" : "rgba(255,255,255,0.58)", position: "absolute", top: 2.5, left: value ? 19 : 3 }} />
      </span>
      <span style={{ opacity: 0.92, fontWeight: 900, fontSize: 14 }}>{label}</span>
    </label>
  );
}

export function FactorySelect({ value, onChange, disabled, title, minWidth = 160, opacity, children }) {
  return (
    <select
      className="factory-select"
      value={value}
      onChange={onChange}
      disabled={disabled}
      title={title}
      style={{ padding: "8px 10px", borderRadius: 10, border: `1px solid ${theme.border}`, background: theme.panel, color: theme.text, fontWeight: 800, minWidth, opacity: opacity ?? (disabled ? 0.55 : 1) }}
    >
      {children}
    </select>
  );
}

// ── Status & pills ────────────────────────────────────────────────────────────

export function ImportStatusRadio({ loaded, label = "Import status" }) {
  const color = loaded ? theme.good : theme.bad;
  const text  = loaded ? "File loaded" : "No file loaded";
  return (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 10, padding: "8px 12px", borderRadius: 999, border: `1px solid ${theme.border}`, background: "rgba(0,0,0,0.45)" }}>
      <span style={{ width: 14, height: 14, borderRadius: 999, background: color }} />
      <div style={{ display: "grid", lineHeight: 1.05 }}>
        <span style={{ fontSize: 12, opacity: 0.78, fontWeight: 900 }}>{label}</span>
        <span style={{ fontSize: 12, fontWeight: 950, color }}>{text}</span>
      </div>
    </div>
  );
}

export function StatPill({ label, value, n }) {
  const v = Number.isFinite(value) ? Math.round(value) : null;
  return (
    <div style={{ border: `1px solid ${theme.border}`, background: "rgba(0,0,0,0.55)", borderRadius: 999, padding: "8px 10px", display: "flex", gap: 8, alignItems: "baseline" }}>
      <span style={{ fontSize: 12, opacity: 0.8 }}>{label}</span>
      <span style={{ fontWeight: 900 }}>{v == null ? "—" : `${v} mm`}</span>
      <span style={{ fontSize: 12, opacity: 0.6 }}>{typeof n === "number" ? `n=${n}` : ""}</span>
    </div>
  );
}

export function ControlPill({ label, value, onChange, suffix = "mm", width = 90, step = 1, min, max, inputColor = theme.text }) {
  return (
    <div style={{ border: `1px solid ${theme.border}`, background: "rgba(0,0,0,0.55)", borderRadius: 999, padding: "8px 10px", display: "flex", gap: 8, alignItems: "baseline" }}>
      <span style={{ fontSize: 12, opacity: 0.8 }}>{label}</span>
      <input
        type="number"
        value={Number.isFinite(value) ? value : 0}
        step={step}
        min={min}
        max={max}
        onChange={(e) => onChange(Number(e.target.value || 0))}
        style={{ width, background: "transparent", color: inputColor, border: `1px solid ${theme.border}`, borderRadius: 12, padding: "6px 8px", fontWeight: 900, outline: "none", userSelect: "auto" }}
      />
      <span style={{ fontSize: 12, opacity: 0.75 }}>{suffix}</span>
    </div>
  );
}

export function TogglePill({ label, checked, onChange }) {
  return (
    <button
      type="button"
      onClick={() => onChange(!checked)}
      style={{ border: `1px solid ${theme.border}`, background: "rgba(0,0,0,0.55)", borderRadius: 999, padding: "8px 10px", display: "flex", gap: 10, alignItems: "center", cursor: "pointer", color: theme.text }}
      title={label}
    >
      <span style={{ fontSize: 12, opacity: 0.8 }}>{label}</span>
      <span style={{ width: 36, height: 20, borderRadius: 999, border: `1px solid ${theme.border}`, background: checked ? "rgba(34,197,94,0.35)" : "rgba(148,163,184,0.18)", position: "relative" }}>
        <span style={{ position: "absolute", top: 2, left: checked ? 18 : 2, width: 16, height: 16, borderRadius: 999, background: checked ? "rgba(34,197,94,0.95)" : "rgba(148,163,184,0.65)", transition: "left 120ms ease" }} />
      </span>
    </button>
  );
}

export function SegTabs({ value, onChange, tabs }) {
  return (
    <div style={{ display: "inline-flex", border: `1px solid ${theme.border}`, borderRadius: 14, overflow: "hidden", background: "rgba(255,255,255,0.16)", boxShadow: "inset 0 0 0 2px rgba(250,204,21,0.65)", flexWrap: "wrap" }}>
      {tabs.map((t) => {
        const active = value === t.value;
        return (
          <button
            key={t.value}
            onClick={() => onChange(t.value)}
            style={{ padding: "9px 12px", border: "none", background: active ? "rgba(59,130,246,0.25)" : "transparent", color: theme.text, cursor: "pointer", fontWeight: 950, fontSize: 14, whiteSpace: "nowrap" }}
          >
            {t.label}
          </button>
        );
      })}
    </div>
  );
}
