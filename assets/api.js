// Talks to the Flight Park Apps Script. Every call resolves to the script's JSON reply.
// POST bodies go as plain text so the browser skips the CORS preflight Apps Script can't answer.
(function () {
  const url = () => (window.FP_CONFIG || {}).scriptUrl;
  async function get(params) {
    const res = await fetch(url() + "?" + new URLSearchParams(params));
    return res.json();
  }
  async function post(body) {
    const res = await fetch(url(), { method: "POST", body: JSON.stringify(body) });
    return res.json();
  }
  window.FP_API = {
    ping: () => get({ action: "ping" }),
    members: () => get({ action: "members" }),
    confirm: (sessionId) => get({ action: "confirm", session_id: sessionId }),
    checkout: (data) => post({ action: "checkout", ...data }),
    commercial: (data) => post({ action: "commercial", ...data }),
  };
})();
