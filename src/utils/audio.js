/**
 * Play a short 880 Hz sine-wave beep via the Web Audio API.
 * Respects window.__ttBeepMuted — set to true to silence all beeps.
 */
export const playBeep = () => {
  try {
    if (window.__ttBeepMuted) return;
    if (!window.__ttBeepCtx) {
      window.__ttBeepCtx = new (window.AudioContext || window.webkitAudioContext)();
    }
    const ctx = window.__ttBeepCtx;
    if (ctx.state === "suspended") ctx.resume();

    const o = ctx.createOscillator();
    const g = ctx.createGain();
    o.type            = "sine";
    o.frequency.value = 880;
    g.gain.value      = 0.12;
    o.connect(g);
    g.connect(ctx.destination);

    const now = ctx.currentTime;
    o.start(now);
    o.stop(now + 0.12);
  } catch {
    // Audio not available – fail silently.
  }
};
