/* shell.js: the MHS add-in shell for Confirmations, Kinnect Relay and Call Pool PTO (keep this file byte-identical across them).
   The same pieces as the IVK Tracker (its app.js and style.css):
   - light / dark: html[data-theme], chosen per computer (localStorage "<app>Theme"); full screen follows the pane
   - the toolbar, flush right: refresh, full screen, light / dark, each with a hover label (data-tip)
   - full screen: an Office dialog floating over the Excel page (displayInIframe, never a new window). It holds a copy of what
     the pane read and asks the pane for every read and write (Shell.call), so only the pane touches the workbook. Messages go
     in parts because the dialog API caps their size.
   <html data-app="conf"> names the storage keys; data-theme-default="system" starts from the computer's setting instead of light.
   Shell.init({ bell, render, refresh, calls, snap, receive, toast })
     bell       the Dev Update bell's app key (dev-updates.js); the bell shows only on the full-screen page's .du-home
     render()   re-draws the app (it puts Shell.tools() in its header)
     refresh()  read the workbook again (the Refresh button). Full screen without one asks the pane to, then receive()s a snap
     calls      pane: { name: async (...args) => result } full screen may ask for (Shell.call(name, ...args))
     snap()     pane: what full screen needs to draw; sent when full screen opens and on Shell.update()
     receive(m) full screen: { k: "snap", ... } and anything else the pane sends with Shell.send */
(function () {
  "use strict";
  const Q = new URLSearchParams(location.search), FULL = Q.has("full"), H = document.documentElement;
  const CHUNK = 8000, CALL_MS = 120000, APP = H.dataset.app || "mhs";
  const O = { bell: "", render() {}, refresh: null, calls: {}, snap: () => ({}), receive() {}, toast(m) { console.warn(m); } };
  const ls = { get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }, set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* this session only */ } } };
  const T = { theme: "light", swap: false, reading: false, got: false };
  const svg = (d, extra) => `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"${extra || ' stroke-linejoin="round"'}>${d}</svg>`;
  const ICON = {
    refresh: svg('<path d="M20 12a8 8 0 1 1-2.34-5.66"/><path d="M20 4v5h-5"/>'),
    full: svg('<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>'),
    moon: svg('<path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z"/>'),
    sun: svg('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>', " ")
  };
  const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  /* ---------- light / dark ---------- */
  const themeKey = APP + "Theme";
  function applyTheme(t) { T.theme = t === "dark" ? "dark" : "light"; H.dataset.theme = T.theme; }
  (function first() {
    let t = (FULL && Q.get("theme")) || ls.get(themeKey);
    if (!t && H.dataset.themeDefault === "system") try { t = matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light"; } catch (e) { /* light */ }
    applyTheme(t);
  })();
  function setTheme(t, fromOther) {
    applyTheme(t); ls.set(themeKey, T.theme);
    if (!fromOther) { send({ k: "theme", theme: T.theme }); T.swap = true; }
    O.render(); T.swap = false;
  }

  /* ---------- toolbar ---------- */
  /* refresh, full screen, light / dark. The theme label is the mode a press switches to; the sun and moon swap places when the
     mode changes (the one on the right is the one to press). */
  function tools(busy) {
    const dark = T.theme === "dark", spin = !!busy || T.reading;
    const fl = FULL ? "Exit full screen" : pop.dlg ? "Full screen is open" : "Full screen", tl = dark ? "Light mode" : "Dark mode";
    const pair = dark ? `<span class="tcur">${ICON.moon}</span><span class="tsl">/</span><span class="tgo">${ICON.sun}</span>` : `<span class="tcur">${ICON.sun}</span><span class="tsl">/</span><span class="tgo">${ICON.moon}</span>`;
    return `<div class="tools">${O.refresh || FULL ? `<button type="button" class="ibtn${spin ? " spin" : ""}" data-shell="reload" data-tip="${spin ? "Reading the workbook…" : "Refresh"}" aria-label="Refresh from the workbook"${spin ? " disabled" : ""}>${ICON.refresh}</button>` : ""}` +
      `<button type="button" class="ibtn" data-shell="full" data-tip="${fl}" aria-label="${fl}">${ICON.full}</button>` +
      `<button type="button" class="ibtn theme${dark ? " is-dark" : ""}${T.swap ? " swap" : ""}" data-shell="theme" data-tip="${tl}" aria-label="Switch to ${tl.toLowerCase()}" aria-pressed="${dark}">${pair}</button></div>`;
  }
  /* full screen's left rail: the add-in's name with the bell's home right of it, the app's own items, then Exit full screen */
  function rail(o) {
    return `<nav class="rail" aria-label="${esc(o.name)}"><div class="brand"><span>${esc(o.name)}${o.sub ? `<small>${esc(o.sub)}</small>` : ""}</span><span class="du-home" aria-hidden="true"></span></div>
      ${o.nav || ""}<div class="spacer"></div>${o.who || ""}
      <button type="button" class="item fs" data-shell="full">${ICON.full}<span>Exit full screen</span></button>
      <div class="conn${o.busy ? " busy" : ""}"><span class="cdot"></span>${esc(o.busy || "Saving through the task pane")}</div></nav>`;
  }
  async function reload() {
    if (T.reading) return;
    T.reading = true; O.render();
    try { if (O.refresh) await O.refresh(); else if (FULL) { const s = await call("__snap"); O.receive(Object.assign({ k: "snap" }, s)); } }
    catch (e) { console.error(e); O.toast("Couldn’t read the workbook: " + ((e && e.message) || e)); }
    finally { T.reading = false; O.render(); }
  }
  document.addEventListener("click", e => {
    const b = e.target.closest && e.target.closest("[data-shell]"); if (!b || b.disabled) return;
    e.preventDefault();
    const a = b.dataset.shell;
    if (a === "reload") reload(); else if (a === "full") open(); else if (a === "theme") setTheme(T.theme === "dark" ? "light" : "dark");
  });

  /* ---------- pane <-> full screen (as in the IVK Tracker) ---------- */
  const pop = { dlg: null }, rpc = { id: 0, wait: new Map() }, asm = {};
  const parts = s => { const n = Math.ceil(s.length / CHUNK) || 1, seq = Date.now() + Math.random(); return Array.from({ length: n }, (_, i) => JSON.stringify({ k: "part", seq, i, n, part: s.slice(i * CHUNK, (i + 1) * CHUNK) })); };
  function reassemble(raw, cb) {
    let m; try { m = JSON.parse(raw); } catch (e) { return; }
    if (m.k !== "part") return cb(m);
    const b = asm[m.seq] || (asm[m.seq] = { n: m.n, p: [] }); b.p[m.i] = m.part;
    if (b.p.filter(x => x != null).length === m.n) { delete asm[m.seq]; let o; try { o = JSON.parse(b.p.join("")); } catch (e) { return; } cb(o); }
  }
  function toChild(o) { if (!pop.dlg) return; for (const p of parts(JSON.stringify(o))) try { pop.dlg.messageChild(p); } catch (e) { /* window closed */ } }
  function toPane(o) { for (const p of parts(JSON.stringify(o))) Office.context.ui.messageParent(p); }
  function send(o) { if (FULL) { try { toPane(o); } catch (e) { /* pane gone */ } } else toChild(o); }
  /* full screen: ask the pane to do something; the pane pings "wait" every 15 s while it works, so only silence times out.
     The pane's error comes back with its own fields (for example which sheets need a password) and its type in .remote. */
  function call(fn, ...args) {
    if (!FULL) { const f = fn === "__snap" ? paneSnap : O.calls[fn]; return f ? Promise.resolve().then(() => f(...args)) : Promise.reject(new Error("no " + fn)); }
    return new Promise((res, rej) => {
      const id = ++rpc.id, lost = () => { if (rpc.wait.delete(id)) rej(new Error("Full screen lost touch with the task pane. Keep the pane open, then check the workbook before trying again.")); };
      const w = { t: setTimeout(lost, CALL_MS), alive() { clearTimeout(w.t); w.t = setTimeout(lost, CALL_MS); }, res: v => { clearTimeout(w.t); res(v); }, rej: e => { clearTimeout(w.t); rej(e); } };
      rpc.wait.set(id, w);
      try { toPane({ k: "call", id, fn, args }); } catch (e) { clearTimeout(w.t); rpc.wait.delete(id); rej(e); }
    });
  }
  async function paneSnap() { if (O.refresh) await O.refresh(); return O.snap(); }
  function update() { if (!pop.dlg) return; Promise.resolve().then(() => O.snap()).then(s => toChild(Object.assign({ k: "snap" }, s, { theme: T.theme }))).catch(e => console.error(e)); }
  function onChild(arg) {
    reassemble(arg.message, async m => {
      if (window.DevUpdates && DevUpdates.receive(m)) return;   // what full screen has read in Dev updates
      if (m.k === "theme") return setTheme(m.theme, true);
      if (m.k === "ready") return update();
      if (m.k === "close") return close();
      if (m.k !== "call") return;
      const f = m.fn === "__snap" ? paneSnap : O.calls[m.fn];
      let ok = true, err = "", ename = "", edata = null, result = null;
      const ping = setInterval(() => toChild({ k: "wait", id: m.id }), 15000);
      try { if (!f) throw new Error("the task pane can’t do " + m.fn); result = await f(...(m.args || [])); }
      catch (e) { ok = false; err = (e && e.message) || String(e); ename = (e && e.constructor && e.constructor.name) || ""; try { edata = e && typeof e === "object" ? JSON.parse(JSON.stringify(Object.assign({}, e))) : null; } catch (x) { edata = null; } }
      finally { clearInterval(ping); }
      toChild({ k: "ack", id: m.id, ok, err, ename, edata, result: result === undefined ? null : result });
    });
  }
  function onPane(m) {
    if (window.DevUpdates && DevUpdates.receive(m)) return;
    if (m.k === "theme") return setTheme(m.theme, true);
    if (m.k === "wait") { const w = rpc.wait.get(m.id); if (w) w.alive(); return; }
    if (m.k === "ack") {
      const w = rpc.wait.get(m.id); if (!w) return; rpc.wait.delete(m.id);
      if (m.ok) w.res(m.result); else w.rej(Object.assign(new Error(m.err || "the task pane couldn’t do that"), m.edata || {}, { remote: m.ename || "" }));
      return;
    }
    if (m.k === "snap") { T.got = true; if (m.theme && m.theme !== T.theme) { applyTheme(m.theme); ls.set(themeKey, T.theme); } }
    O.receive(m);
  }
  function open() {
    if (FULL) { try { toPane({ k: "close" }); } catch (e) { /* pane gone */ } return; }
    if (pop.dlg) return O.toast("Full screen is already open.");
    const ui = window.Office && Office.context && Office.context.ui;
    if (!ui || !ui.displayDialogAsync) return O.toast("Full screen works inside Excel.");
    if (!Office.context.requirements.isSetSupported("DialogApi", "1.2")) return O.toast("This version of Excel can’t open full screen.");
    const url = new URL(location.href); url.search = "?full=1&theme=" + T.theme; url.hash = "";
    /* what Dev updates this computer has read go along, so the red dot matches */
    ui.displayDialogAsync(window.DevUpdates ? DevUpdates.dialogUrl(url.href) : url.href, { height: 95, width: 95, displayInIframe: true }, res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return O.toast(res.error && res.error.code === 12007 ? "Close the other add-in window first." : "Full screen couldn’t open. Try again.");
      pop.dlg = res.value; if (window.DevUpdates) DevUpdates.setSend(m => toChild(m));
      pop.dlg.addEventHandler(Office.EventType.DialogMessageReceived, onChild);
      pop.dlg.addEventHandler(Office.EventType.DialogEventReceived, () => { pop.dlg = null; if (window.DevUpdates) DevUpdates.setSend(null); O.render(); });
      O.render();
    });
  }
  function close() {
    if (FULL) { try { toPane({ k: "close" }); } catch (e) { /* pane gone */ } return; }
    if (pop.dlg) { try { pop.dlg.close(); } catch (e) { /* closed */ } pop.dlg = null; } if (window.DevUpdates) DevUpdates.setSend(null); O.render(); }

  function init(o) {
    Object.assign(O, o || {});
    if (FULL) {
      document.body.classList.add("fs");
      const hello = () => { if (T.got) return; try { toPane({ k: "ready" }); } catch (e) { /* pane not listening yet */ } setTimeout(hello, 2000); };
      try { Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, a => reassemble(a.message, onPane), hello); } catch (e) { console.error(e); }
    }
    if (O.bell && window.DevUpdates) DevUpdates.init({ app: O.bell, send: FULL ? m => { try { toPane(m); } catch (e) { /* pane gone */ } } : null });
  }

  window.Shell = { FULL, ICON, init, tools, rail, call, send, update, open, close, reload, setTheme, isOpen: () => !!pop.dlg, get theme() { return T.theme; }, get reading() { return T.reading; } };
})();
