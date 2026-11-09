/* MHS Confirmations task pane. Reads the List sheet, shows who needs a confirmation per clinic,
   and builds the email body. Patient names never leave the sheet: only LASFIR codes are shown or emailed. */
(function () {
  'use strict';
  var C = window.MHSCore;
  var LIST_SHEET = 'List', SETTINGS_SHEET = 'Confirm Settings';
  var state = { res: null, optIn: {}, holidays: [], opts: {} };
  var $ = function (id) { return document.getElementById(id); };

  function el(tag, attrs, text) {
    var e = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) { e.setAttribute(k, attrs[k]); });
    if (text != null) e.textContent = text;
    return e;
  }

  // Full screen is an Office dialog floating over the Excel page (shell.js). It never touches the workbook: it draws from a copy
  // of what the pane read (snap) and asks the pane for every write (Shell.call). The dialog has no Office.context.document.
  var FULL = !!(window.Shell && Shell.FULL);
  function init() {
    ['optToday', 'optCancelled', 'optNotes'].forEach(function (id) { $(id).addEventListener('change', render); });
    $('baseUrl').addEventListener('change', function () {
      var v = $('baseUrl').value.trim();
      if (FULL) { Shell.call('setBaseUrl', v).catch(function (e) { toast('Couldn’t save the link: ' + (e.message || e)); }); render(); return; }
      setBaseUrl(v);
    });
    $('pvClose').addEventListener('click', function () { $('preview').hidden = true; });
    $('pvCopy').addEventListener('click', copyEmail);
    $('gate').addEventListener('click', onGate);
    $('gate').addEventListener('keydown', function (e) { if (e.key === 'Enter' && e.target.id === 'whoNew') { e.preventDefault(); addUser(); } });
    $('me').addEventListener('click', function (e) { if (e.target.closest('[data-act=switchUser]')) openPicker(); });
    Shell.init({ bell: 'conf', render: drawShell, refresh: FULL ? null : function () { return load(); }, toast: toast,
      calls: PANE, snap: snap, receive: fromPane });
    if (FULL) { BOOT = ['Getting the workbook from the task pane']; $('frame').className = 'fsframe'; drawShell(); renderGate(); return; }
    state.fallbackUrl = Office.context.document.settings.get('baseUrl') || normalizeUrl(Office.context.document.url) || '';
    $('baseUrl').value = state.fallbackUrl;
    who.wb = workbookKey();
    drawShell();
    renderGate();
    load(true);
    // a pane left open overnight asks for a name again on the new day
    setInterval(function () { if (!boot.on && who.name && who.day !== dayIso()) { who.name = ''; renderGate(); } }, 60000);
  }
  function setBaseUrl(v) {
    Office.context.document.settings.set('baseUrl', v);
    Office.context.document.settings.saveAsync();
    $('baseUrl').value = v;
    writeSheetUrl(v);
    render(); Shell.update();
  }

  // ---------- the toolbar, and full screen's rail ----------
  // Refresh (read the workbook again), full screen and light / dark, flush right (shell.js). Full screen adds the rail: the
  // add-in's name with the Dev Update bell's home right of it, then Exit full screen.
  function drawShell() {
    $('tools').innerHTML = Shell.tools(state.loading);
    if (FULL) $('rail').innerHTML = Shell.rail({ name: 'Confirmations', sub: 'Outreach to clinics' });
  }
  var toastT = null;
  function toast(msg) { var t = $('toast'); t.textContent = msg; t.classList.add('on'); clearTimeout(toastT); toastT = setTimeout(function () { t.classList.remove('on'); }, 4000); }
  // what full screen draws from: the List rows, the settings it read and who's at this computer
  function snap() {
    return { rows: state.rows || null, firstRow: state.firstRow, headerBad: state.headerBad || [], optIn: state.optIn, holidays: state.holidays, ccWarnings: state.ccWarnings || [],
      clinics: C.getClinics(), baseUrl: $('baseUrl').value.trim(), err: state.err || '', users: who.users, name: who.name, wb: who.wb };
  }
  function fromPane(m) {
    if (m.k !== 'snap') return;
    state.rows = m.rows; state.firstRow = m.firstRow; state.headerBad = m.headerBad; state.optIn = m.optIn || {}; state.holidays = m.holidays || []; state.ccWarnings = m.ccWarnings;
    C.setClinics(m.clinics); state.err = m.err;
    if (document.activeElement !== $('baseUrl')) $('baseUrl').value = m.baseUrl || '';
    who.wb = m.wb; who.users = m.users || []; if (!who.picking) { who.name = m.name || ''; who.day = dayIso(); }
    boot.on = false;
    if (state.err) showLoadError(state.err); else render();
    renderGate();
  }
  // what full screen may ask the pane to do
  var PANE = {
    setBaseUrl: function (v) { setBaseUrl(String(v || '').trim()); return true; },
    setWho: function (name) { setWho(name); renderGate(); Shell.update(); return true; },
    addUser: function (n) { return saveName(n).then(function (name) { pickUser(name); return { name: name, users: who.users }; }); },
    users: function () { return refreshDoc().then(function () { who.users = readUsers(); return who.users; }); },
    goTo: function (row) { return goTo(row); }
  };

  // ---------- the loading screen, then "Enter your name" (both as in the IVK Tracker) ----------
  // The people who use this workbook are a list in its document settings (confUsers), so no sheet is added. This computer
  // remembers, per workbook, who picked today and who has used it here before (localStorage "confWho:<workbook key>").
  // The pick lasts the day; nobody is picked for you.
  var USERS_KEY = 'confUsers', WBID_KEY = 'confWbId';
  var BOOT = ['Connecting to the workbook', 'Reading Confirm Settings', 'Reading the List sheet', 'Finding who needs a confirmation'];
  var boot = { on: true, i: 0, m: [] };
  var who = { wb: '', users: [], name: '', day: '', picking: false };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  var txt = function (s) { return String(s == null ? '' : s).trim(); };
  var userKey = function (n) { return txt(n).toLowerCase(); };
  var cleanName = function (s) { return txt(s).replace(/\s+/g, ' ').slice(0, 30).trim(); };
  function dayIso() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function lsGet(k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* this session only */ } }
  function wbKeyOf(s) {
    var h = 0x811c9dc5, t = String(s || '').trim().toLowerCase().replace(/[?#].*$/, '');
    for (var i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return t ? 'w' + h.toString(36) : '';
  }
  function workbookKey() {
    var s = Office.context.document.settings, url = '';
    try { url = Office.context.document.url || ''; } catch (e) { /* no document */ }
    if (url) return wbKeyOf(url);
    var id = s.get(WBID_KEY);
    if (!id) { id = 'id' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8); s.set(WBID_KEY, id); s.saveAsync(); }
    return wbKeyOf(id);
  }
  var whoKey = function () { return who.wb ? 'confWho:' + who.wb : ''; };
  // what this computer remembers for one workbook: { day: "2026-11-07", name: "Sam", used: ["Sam", "Ana"] }
  function whoRead() {
    try { var o = JSON.parse(lsGet(whoKey()) || '{}'); return { day: txt(o.day), name: txt(o.name), used: Array.isArray(o.used) ? o.used.map(txt).filter(Boolean) : [] }; }
    catch (e) { return { day: '', name: '', used: [] }; }
  }
  function readUsers() {
    var v = Office.context.document.settings.get(USERS_KEY), seen = {}, out = [];
    (Array.isArray(v) ? v : []).forEach(function (u) {
      var name = cleanName(u && u.name), k = userKey(name);
      if (name && !seen[k]) { seen[k] = true; out.push({ name: name, added: txt(u.added) }); }
    });
    return out;
  }
  function refreshDoc() { return new Promise(function (res) { if (FULL) return res(); try { Office.context.document.settings.refreshAsync(function () { res(); }); } catch (e) { res(); } }); }
  function saveUsers(list) {
    return new Promise(function (res, rej) {
      var s = Office.context.document.settings; s.set(USERS_KEY, list);
      s.saveAsync(function (r) { if (r.status === Office.AsyncResultStatus.Succeeded) res(); else rej(new Error('Couldn’t save your name to the workbook. Try again.')); });
    });
  }
  // the people who have used this computer on this workbook first (most recent first), then everyone else A-Z
  function pickOrder(users, here) {
    var H = (here || []).map(userKey), rank = function (u) { var i = H.indexOf(userKey(u.name)); return i < 0 ? Infinity : i; };
    return users.map(function (u) { return { name: u.name, here: rank(u) !== Infinity }; })
      .sort(function (a, b) { return (rank(a) - rank(b)) || a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }); });
  }
  // today's pick on this computer, if any, and only while they're still on the workbook's list
  function loadWho() {
    var rec = whoRead(), n = rec.day === dayIso() ? rec.name : '';
    var u = n && who.users.filter(function (x) { return userKey(x.name) === userKey(n); })[0];
    who.name = u ? u.name : ''; who.day = who.name ? rec.day : '';
  }
  function setWho(name) {
    var rec = whoRead(), day = dayIso();
    rec = { day: day, name: name, used: [name].concat(rec.used.filter(function (n) { return userKey(n) !== userKey(name); })).slice(0, 20) };
    if (whoKey()) lsSet(whoKey(), JSON.stringify(rec));
    who.name = name; who.day = day; who.picking = false;
  }

  function bootStep(i, msg) { if (!boot.on) return; boot.m[i] = msg; boot.i = Math.max(boot.i, i + 1); renderGate(); }
  function bootDone() {
    if (!boot.on) return;
    return refreshDoc().then(function () { who.users = readUsers(); loadWho(); boot.on = false; renderGate(); });
  }
  function bootHTML() {
    return '<div class="boot" role="status" aria-live="polite"><div class="bootin"><div class="spinner" aria-hidden="true"></div><h1>Confirmations</h1><p class="muted">Reading the workbook.</p>' +
      '<ol class="steps">' + BOOT.map(function (t, i) {
        return '<li class="step ' + (boot.i > i ? 'done' : boot.i === i ? 'now' : 'next') + '"><span class="k">' + (boot.i > i ? '✓' : '') + '</span><span><span class="t">' + t + '</span><span class="m">' + esc(boot.m[i] || '') + '</span></span></li>';
      }).join('') + '</ol></div></div>';
  }
  // "Enter your name": this workbook's list only, the people who've used this computer here first, nobody picked
  function pickHTML() {
    var L = pickOrder(who.users, whoRead().used), cancel = who.picking && who.name;
    return '<div class="boot who"><div class="bootin"><div class="whohero"><h1>Enter your name</h1>' +
      '<p class="muted">' + (L.length ? 'Pick yours below, or add it if you’re new.' : 'Nobody has added their name to this workbook yet. Add yours to start.') + '</p></div>' +
      (L.length ? '<div class="wholist" role="group" aria-label="People on this workbook">' + L.map(function (u) {
        return '<button type="button" class="whob" data-act="pickUser" data-v="' + esc(u.name) + '">' + esc(u.name) + (u.here ? '<span class="tag">On this computer</span>' : '') + '</button>';
      }).join('') + '</div>' : '') +
      '<details class="whonew"' + (L.length ? '' : ' open') + '><summary>' + (L.length ? 'I’m not on the list' : 'Add your name') + '</summary>' +
        '<label class="fld">Your name<input id="whoNew" autocomplete="off" spellcheck="false" maxlength="30"><span class="hint">Usually your first name. Add a last initial if someone here shares it.</span></label>' +
        '<div class="whoact"><button type="button" data-act="addUser">Add me and start</button></div></details>' +
      (cancel ? '<div class="whoact"><button type="button" class="ghost" data-act="pickCancel">Stay as ' + esc(who.name) + '</button></div>' : '') + '</div></div>';
  }
  function renderGate() {
    var g = $('gate'), me = $('me');
    if (boot.on) g.innerHTML = bootHTML();
    else if (!who.name || who.picking) {
      var open = g.querySelector('.whonew') && g.querySelector('.whonew').open, typed = $('whoNew') ? $('whoNew').value : '';
      g.innerHTML = pickHTML();
      if (open) g.querySelector('.whonew').open = true;
      if (typed) $('whoNew').value = typed;
      if ($('whoNew') && (!who.users.length || open)) $('whoNew').focus();
    } else g.innerHTML = '';
    me.hidden = !who.name;
    me.innerHTML = who.name ? '<b>' + esc(who.name) + '</b> · <button type="button" class="linkbtn" data-act="switchUser">Switch user</button>' : '';
  }
  // full screen asks the pane: the pane keeps the list and today's pick (full screen takes the pane's pick)
  function openPicker() {
    var users = FULL ? Shell.call('users') : refreshDoc().then(function () { return readUsers(); });
    users.then(function (u) { who.users = u; who.picking = true; renderGate(); }, function (e) { toast(String((e && e.message) || e)); });
  }
  function pickUser(name) {
    if (FULL) { Shell.call('setWho', name).then(function () { who.name = name; who.day = dayIso(); who.picking = false; renderGate(); }, function (e) { toast(String((e && e.message) || e)); }); return; }
    setWho(name); renderGate(); Shell.update();
  }
  function saveName(n) {
    return refreshDoc().then(function () {
      var list = readUsers(), had = list.filter(function (u) { return userKey(u.name) === userKey(n); })[0];
      if (had) return had.name;
      return saveUsers(list.concat([{ name: n, added: dayIso() }])).then(function () { return n; });
    }).then(function (name) { who.users = readUsers(); return name; });
  }
  function addUser() {
    var n = cleanName($('whoNew') && $('whoNew').value);
    if (!n) { if ($('whoNew')) $('whoNew').focus(); return; }
    var done = function (name) { if (!FULL) return pickUser(name); who.name = name; who.day = dayIso(); who.picking = false; renderGate(); };
    (FULL ? Shell.call('addUser', n).then(function (r) { who.users = r.users; return r.name; }) : saveName(n)).then(done, function (e) {
      var f = $('gate').querySelector('.hint'); if (f) f.textContent = (e && e.message) || 'Couldn’t save your name. Try again.';
    });
  }
  function onGate(e) {
    var b = e.target.closest('[data-act]'); if (!b) return;
    if (b.dataset.act === 'pickUser') pickUser(b.dataset.v);
    else if (b.dataset.act === 'addUser') addUser();
    else if (b.dataset.act === 'pickCancel') { who.picking = false; renderGate(); }
  }

  function normalizeUrl(u) {
    if (!u || !/^https?:/i.test(u)) return '';
    return /[?&]web=1/.test(u) ? u : u + (u.indexOf('?') < 0 ? '?' : '&') + 'web=1';
  }

  // ---------- settings sheet (the "Confirmations Run" Office Script reads clinic names/emails, opt-ins, CC (D), holidays and the G2 link from it) ----------
  // Mirror of the sheet-creation block in office-scripts/ConfirmationsRun.ts. Keep the two in sync.
  function createSettings(sheets) { // queues only; caller syncs
    var s = sheets.add(SETTINGS_SHEET);
    var data = [['ClinicName', 'ClinicEmail', 'Opt In (Yes/No)']].concat(C.CLINICS.map(function (c) { return [c.name, c.email, 'No']; }));
    s.getRange('A1:C' + data.length).values = data;
    s.getRange('D1').values = [['CC (optional, MHS addresses, separate with ;)']];
    s.getRange('E1').values = [['Holidays (YYYY-MM-DD, optional)']];
    s.getRange('G1').values = [['Workbook link for View (paste in G2 if not using the add-in)']];
    s.getRange('C2:C' + data.length).dataValidation.rule = { list: { inCellDropDown: true, source: 'Yes,No' } };
    s.getRange('A1:G1').format.font.bold = true;
    s.getRange('A:G').format.autofitColumns();
    return s;
  }

  // values: Confirm Settings used range, assumed to start at A1 (A1 always holds "ClinicName").
  function readSettings(values, ws) {
    var optIn = {}, holidays = [], clinics = [], ccWarnings = [];
    values.slice(1).forEach(function (row) {
      var cname = String(row[0] || '').trim();
      if (cname) {
        optIn[cname] = /^y/i.test(String(row[2] || ''));
        var email = String(row[1] || '').trim() || C.emailForClinic(cname);
        var cc = C.parseCc(row[3], email);
        if (cc.dropped.length) ccWarnings.push('CC for "' + cname + '" left out ' + cc.dropped.join(', ') + ': only @mymhs.com, @mindfulhealthsolutions.com or @meetradial.com addresses are copied.');
        if (email) clinics.push({ name: cname, email: email, cc: cc.cc });
      }
      var h = row[4];
      if (h !== '' && h != null) { var s = typeof h === 'number' ? h : C.serialOfDate(new Date(String(h) + 'T00:00:00')); if (!isNaN(s)) holidays.push(s); }
    });
    state.optIn = optIn; state.holidays = holidays; state.ccWarnings = ccWarnings;
    if (clinics.length) C.setClinics(clinics); // Confirm Settings A:B drive matching and emails (mirrors the script)
    state.sheetUrl = values.length > 1 && values[1][6] ? String(values[1][6]).trim() : '';
    state.settingsWs = ws;
  }

  // Rebuild the A:O block from row 1 down to the last used row, as getRangeByIndexes(0, 0, rowIndex + rowCount, 15) would,
  // from a used range that may start below row 1 or right of column A.
  function listValues(used) {
    if (used.isNullObject) return [];
    var out = [];
    for (var r = 0; r < used.rowIndex + used.rowCount; r++) {
      var row = [];
      for (var c = 0; c < 15; c++) {
        var ur = r - used.rowIndex, uc = c - used.columnIndex;
        row.push(ur >= 0 && uc >= 0 && uc < used.columnCount ? used.values[ur][uc] : '');
      }
      out.push(row);
    }
    return out;
  }

  function writeSheetUrl(url) {
    state.sheetUrl = url;
    Excel.run(function (context) {
      var ws = context.workbook.worksheets.getItemOrNullObject(SETTINGS_SHEET);
      return context.sync().then(function () {
        if (ws.isNullObject) return;
        ws.getRange('G2').values = [[url]];
        return context.sync();
      });
    }).catch(function () { /* user can paste the link into Confirm Settings!G2 by hand */ });
  }

  // ---------- load ----------
  // first = the read on open: its steps tick on the loading screen, which then hands over to "Enter your name"
  function load(first) {
    $('window').textContent = 'Loading…';
    var step = first ? bootStep : function () {};
    state.loading = true; drawShell();
    return Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      var sws = sheets.getItemOrNullObject(SETTINGS_SHEET);
      var lws = sheets.getItemOrNullObject(LIST_SHEET);
      return context.sync().then(function () {
        step(0, 'Connected');
        if (sws.isNullObject) sws = createSettings(sheets);
        var sused = sws.getUsedRange(true); sused.load('values');
        var lused = null;
        if (!lws.isNullObject) { lused = lws.getUsedRangeOrNullObject(true); lused.load('values,rowIndex,columnIndex,rowCount,columnCount'); }
        return context.sync().then(function () {
          readSettings(sused.values, sws);
          var nc = Object.keys(state.optIn).length;
          step(1, nc + (nc === 1 ? ' clinic' : ' clinics'));
          if (!lused) throw new Error('No sheet named "' + LIST_SHEET + '" in this workbook.');
          return listValues(lused);
        });
      });
    }).then(function (values) {
      var hdr = -1;
      for (var i = 0; i < Math.min(10, values.length); i++) if (String(values[i][0]).trim().toLowerCase() === 'md rc') { hdr = i; break; }
      if (hdr < 0) throw new Error('Could not find the header row (column A "Md Rc") on the List sheet.');
      state.headerBad = C.checkHeaders(values[hdr]);
      // Sheet G2 is the source of truth (the scheduled script reads it). Fill it from the pane if empty.
      if (state.sheetUrl) $('baseUrl').value = state.sheetUrl;
      else if ($('baseUrl').value.trim()) writeSheetUrl($('baseUrl').value.trim());
      state.rows = values.slice(hdr + 1);
      state.firstRow = hdr + 2; // sheet row number of first data row
      state.err = '';
      step(2, state.rows.length + (state.rows.length === 1 ? ' row' : ' rows'));
      render();
      step(3, state.res.stats.flagged + ' to confirm');
      if (first) return bootDone();
    }).catch(function (e) {
      state.err = String(e && e.message || e);
      showLoadError(state.err);
      if (first) return bootDone();   // the warning shows in the pane once a name is picked
    }).then(function () { state.loading = false; drawShell(); Shell.update(); });
  }
  function showLoadError(msg) {
    $('window').textContent = '';
    $('clinics').textContent = '';
    $('warnings').textContent = '';
    $('warnings').appendChild(el('div', { class: 'warn' }, 'Could not load: ' + msg));
  }

  // ---------- render ----------
  function currentOpts() {
    return { todaySerial: C.serialOfDate(new Date()), includeToday: $('optToday').checked, skipCancelled: $('optCancelled').checked, holidays: state.holidays };
  }

  function render() {
    if (!state.rows) return;
    var res = state.res = C.findOpportunities(state.rows, state.firstRow, currentOpts());
    $('window').textContent = C.fmtDate(res.start) + ' to ' + C.fmtDate(res.end) + ' · ' + res.stats.flagged + ' to confirm · ' +
      res.stats.done + ' already confirmed/rescheduled' + (res.stats.cancelled ? ' · ' + res.stats.cancelled + ' cancelled skipped' : '');

    var w = $('warnings'); w.textContent = '';
    if (state.headerBad && state.headerBad.length) w.appendChild(el('div', { class: 'warn' }, 'Column headers do not match the expected layout (Md Rc, Pat Name, Loc Name, Appt Dt, Beg Tm). Check that columns have not moved before sending anything.'));
    Object.keys(res.unmapped).forEach(function (loc) {
      w.appendChild(el('div', { class: 'warn' }, res.unmapped[loc].length + ' appointment(s) at "' + loc + '" have no clinic email mapped and will not be emailed.'));
    });
    if (res.skipped.length) w.appendChild(el('div', { class: 'warn' }, 'Skipped ' + res.skipped.length + ' row(s) where Appt Dt is not a date: sheet row ' + res.skipped.map(function (s) { return s.row; }).join(', ') + '.'));
    if (res.badNames) w.appendChild(el('div', { class: 'warn' }, res.badNames + ' row(s) have a blank or non "Last, First" Pat Name, so no LASFIR code could be made.'));
    (state.ccWarnings || []).forEach(function (t) { w.appendChild(el('div', { class: 'warn' }, t)); });
    if (!$('baseUrl').value.trim()) w.appendChild(el('div', { class: 'warn' }, 'No workbook link set, so emails will have no "View" links.'));

    var host = $('clinics'); host.textContent = '';
    var names = C.CLINICS.map(function (c) { return c.name; }).filter(function (n) { return res.byClinic[n]; });
    if (!names.length) host.appendChild(el('p', { class: 'muted' }, 'Nothing needs confirmation in this window.'));
    names.forEach(function (name) { host.appendChild(clinicCard(name, res.byClinic[name])); });
  }

  function clinicCard(name, items) {
    var card = el('details', { class: 'card' });
    var sum = el('summary');
    sum.appendChild(el('span', { class: 'cname' }, name));
    sum.appendChild(el('span', { class: 'badge' }, String(items.length)));
    sum.appendChild(el('span', { class: state.optIn[name] ? 'tag on' : 'tag' }, state.optIn[name] ? 'Opted in' : 'Not opted in'));
    card.appendChild(sum);

    var t = el('table');
    items.forEach(function (i) {
      var tr = el('tr');
      tr.appendChild(el('td', {}, C.fmtDate(i.dateSerial) + ' ' + C.fmtTime(i.timeFrac)));
      tr.appendChild(el('td', {}, i.lasfir || (i.nameIssue === 'blank' ? '(missing)' : '(check name)')));
      tr.appendChild(el('td', { class: 'ev' }, i.event));
      var td = el('td'); var a = el('a', { href: '#' }, 'View');
      a.addEventListener('click', function (e) { e.preventDefault(); goTo(i.row); });
      td.appendChild(a); tr.appendChild(td); t.appendChild(tr);
    });
    card.appendChild(t);
    var b = el('button', {}, 'Preview email');
    b.addEventListener('click', function () { showPreview(name, items); });
    card.appendChild(b);
    return card;
  }

  // in full screen, View selects the row through the pane and closes full screen, so the row is in sight
  function goTo(row) {
    if (FULL) return Shell.call('goTo', row).then(function () { Shell.close(); }, function (e) { toast('Couldn’t go to the row: ' + (e.message || e)); });
    return Excel.run(function (context) {
      var ws = context.workbook.worksheets.getItem(LIST_SHEET);
      ws.activate(); ws.getRange('A' + row).select();
      return context.sync();
    });
  }

  // ---------- email preview / copy ----------
  var lastEmail = null;
  function showPreview(name, items) {
    lastEmail = C.buildEmail(name, items, state.res, { baseUrl: $('baseUrl').value.trim(), includeNotes: $('optNotes').checked });
    $('pvTitle').textContent = name;
    $('pvMeta').textContent = 'To: ' + lastEmail.to + (lastEmail.cc ? '  ·  Cc: ' + lastEmail.cc.split(';').join('; ') : '') + '  ·  Subject: ' + lastEmail.subject;
    $('pvBody').innerHTML = lastEmail.html; // built by core.js with all cell text HTML-escaped
    $('preview').hidden = false;
    $('preview').scrollIntoView();
  }

  function copyEmail() {
    if (!lastEmail) return;
    var done = function () { $('pvCopy').textContent = 'Copied'; setTimeout(function () { $('pvCopy').textContent = 'Copy email'; }, 1500); };
    if (navigator.clipboard && window.ClipboardItem) {
      navigator.clipboard.write([new ClipboardItem({
        'text/html': new Blob([lastEmail.html], { type: 'text/html' }),
        'text/plain': new Blob([lastEmail.text], { type: 'text/plain' })
      })]).then(done, fallbackCopy);
    } else fallbackCopy();
    function fallbackCopy() {
      var r = document.createRange(); r.selectNodeContents($('pvBody'));
      var s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
      try { document.execCommand('copy'); done(); } catch (e) { /* user can copy manually */ }
    }
  }

  Office.onReady(init);
})();
