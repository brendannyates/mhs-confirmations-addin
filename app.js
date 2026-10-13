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

  Office.onReady(function () {
    state.fallbackUrl = Office.context.document.settings.get('baseUrl') || normalizeUrl(Office.context.document.url) || '';
    $('baseUrl').value = state.fallbackUrl;
    ['optToday', 'optCancelled', 'optNotes'].forEach(function (id) { $(id).addEventListener('change', render); });
    $('baseUrl').addEventListener('change', function () {
      var v = $('baseUrl').value.trim();
      Office.context.document.settings.set('baseUrl', v);
      Office.context.document.settings.saveAsync();
      writeSheetUrl(v);
      render();
    });
    $('refresh').addEventListener('click', load);
    $('pvClose').addEventListener('click', function () { $('preview').hidden = true; });
    $('pvCopy').addEventListener('click', copyEmail);
    load();
  });

  function normalizeUrl(u) {
    if (!u || !/^https?:/i.test(u)) return '';
    return /[?&]web=1/.test(u) ? u : u + (u.indexOf('?') < 0 ? '?' : '&') + 'web=1';
  }

  // ---------- settings sheet (the "Confirmations Run" Office Script reads clinic names/emails, opt-ins, holidays and the G2 link from it) ----------
  // Mirror of the sheet-creation block in office-scripts/ConfirmationsRun.ts. Keep the two in sync.
  function createSettings(sheets) { // queues only; caller syncs
    var s = sheets.add(SETTINGS_SHEET);
    var data = [['ClinicName', 'ClinicEmail', 'Opt In (Yes/No)']].concat(C.CLINICS.map(function (c) { return [c.name, c.email, 'No']; }));
    s.getRange('A1:C' + data.length).values = data;
    s.getRange('E1').values = [['Holidays (YYYY-MM-DD, optional)']];
    s.getRange('G1').values = [['Workbook link for View (paste in G2 if not using the add-in)']];
    s.getRange('C2:C' + data.length).dataValidation.rule = { list: { inCellDropDown: true, source: 'Yes,No' } };
    s.getRange('A1:G1').format.font.bold = true;
    s.getRange('A:G').format.autofitColumns();
    return s;
  }

  // values: Confirm Settings used range, assumed to start at A1 (A1 always holds "ClinicName").
  function readSettings(values, ws) {
    var optIn = {}, holidays = [], clinics = [];
    values.slice(1).forEach(function (row) {
      var cname = String(row[0] || '').trim();
      if (cname) {
        optIn[cname] = /^y/i.test(String(row[2] || ''));
        var email = String(row[1] || '').trim() || C.emailForClinic(cname);
        if (email) clinics.push({ name: cname, email: email });
      }
      var h = row[4];
      if (h !== '' && h != null) { var s = typeof h === 'number' ? h : C.serialOfDate(new Date(String(h) + 'T00:00:00')); if (!isNaN(s)) holidays.push(s); }
    });
    state.optIn = optIn; state.holidays = holidays;
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
  function load() {
    $('window').textContent = 'Loading…';
    Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      var sws = sheets.getItemOrNullObject(SETTINGS_SHEET);
      var lws = sheets.getItemOrNullObject(LIST_SHEET);
      return context.sync().then(function () {
        if (sws.isNullObject) sws = createSettings(sheets);
        var sused = sws.getUsedRange(true); sused.load('values');
        var lused = null;
        if (!lws.isNullObject) { lused = lws.getUsedRangeOrNullObject(true); lused.load('values,rowIndex,columnIndex,rowCount,columnCount'); }
        return context.sync().then(function () {
          readSettings(sused.values, sws);
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
      render();
    }).catch(function (e) {
      $('window').textContent = '';
      $('clinics').textContent = '';
      $('warnings').textContent = '';
      $('warnings').appendChild(el('div', { class: 'warn' }, 'Could not load: ' + (e && e.message || e)));
    });
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

  function goTo(row) {
    Excel.run(function (context) {
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
    $('pvMeta').textContent = 'To: ' + lastEmail.to + '  ·  Subject: ' + lastEmail.subject;
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
})();
