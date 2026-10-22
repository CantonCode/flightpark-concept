// Talks to the Flight Park Apps Script. Every call resolves to the script's JSON reply.
// POST bodies go as plain text so the browser skips the CORS preflight Apps Script can't answer.
(function () {
  const url = () => (window.FP_CONFIG || {}).scriptUrl;
  async function get(params, timeoutMs) {
    const ctl = timeoutMs ? new AbortController() : null;
    const timer = ctl && setTimeout(() => ctl.abort(), timeoutMs);
    try {
      const res = await fetch(url() + "?" + new URLSearchParams(params), ctl ? { signal: ctl.signal } : undefined);
      return await res.json();
    } finally {
      if (timer) clearTimeout(timer);
    }
  }
  async function post(body) {
    const res = await fetch(url(), { method: "POST", body: JSON.stringify(body) });
    return res.json();
  }
  window.FP_API = {
    ping: () => get({ action: "ping" }),
    members: () => get({ action: "members" }),
    confirm: (sessionId, timeoutMs) => get({ action: "confirm", session_id: sessionId }, timeoutMs),
    checkout: (data) => post({ action: "checkout", ...data }),
    commercial: (data) => post({ action: "commercial", ...data }),
  };
})();
