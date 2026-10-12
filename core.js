/* MHS Confirmations add-in: pure logic. No Office.js in here so it can be unit tested in node.
   Works as a browser global (MHSCore) and as a CommonJS module. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.MHSCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Hardcoded per request. Order preserved for the UI.
  var CLINICS = [
    ['Burlingame - CA', 'helpbur@mindfulhealthsolutions.com'],
    ['Castro - CA', 'helpcastro@mindfulhealthsolutions.com'],
    ['Oakland - CA', 'helpoak@mindfulhealthsolutions.com'],
    ['Post - CA', 'helpsfa@mindfulhealthsolutions.com'],
    ['San Rafael - CA', 'helpsrl@mindfulhealthsolutions.com'],
    ['Walnut Creek - CA', 'helpwck@mindfulhealthsolutions.com'],
    ['El Dorado Hills - CA', 'helpedh@mindfulhealthsolutions.com'],
    ['Fairfield - CA', 'helpffd@mindfulhealthsolutions.com'],
    ['Roseville - CA', 'HelpROS@MindfulHealthSolutions.com'],
    ['Sacramento - CA', 'helpsac@mindfulhealthsolutions.com'],
    ['Sierra College - CA', 'helpscb@mindfulhealthsolutions.com'],
    ['Stockton - CA', 'helpsto@mindfulhealthsolutions.com'],
    ['Los Angeles - CA', 'helplaw@mindfulhealthsolutions.com'],
    ['Mission Viejo - CA', 'helpmvo@mindfulhealthsolutions.com'],
    ['Newport Beach - CA', 'helpnbh@mindfulhealthsolutions.com'],
    ['Santa Monica - CA', 'helpsma@mindfulhealthsolutions.com'],
    ['Austin - TX', 'helpaustin@mindfulhealthsolutions.com'],
    ['Houston - TX', 'helphouston@mindfulhealthsolutions.com'],
    ['Seattle - WA', 'helpSEA@mindfulhealthsolutions.com'],
    ['Tacoma - WA', 'helptacoma@mindfulhealthsolutions.com'],
    ['Wenatchee - WA', 'helpwenatchee@mindfulhealthsolutions.com']
  ].map(function (c) { return { name: c[0], email: c[1] }; });

  // "Loc Name" values in the List sheet carry no state suffix, and a few differ from the clinic name.
  // Anything not matched here or by stripping " - XX" is reported as unmapped, never dropped silently.
  var LOCATION_ALIASES = { 'castro street': 'Castro - CA' };

  var LOC_INDEX = (function () {
    var m = {};
    CLINICS.forEach(function (c) { m[c.name.replace(/\s+-\s+[A-Z]{2}$/, '').toLowerCase()] = c.name; });
    Object.keys(LOCATION_ALIASES).forEach(function (k) { m[k] = LOCATION_ALIASES[k]; });
    return m;
  })();

  function clinicForLocation(loc) {
    var k = String(loc || '').replace(/\s+/g, ' ').trim().toLowerCase();
    return LOC_INDEX[k] || null;
  }
  function emailForClinic(name) {
    for (var i = 0; i < CLINICS.length; i++) if (CLINICS[i].name === name) return CLINICS[i].email;
    return '';
  }

  // ---------- LASFIR: last 3 of last name + first 3 of first name. "Smith, John (Tom)" -> SMIJOH ----------
  function letters(s) { return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z]/g, ''); }
  function toLASFIR(raw) {
    var s = String(raw == null ? '' : raw).trim();
    if (!s) return { code: '', ok: false, reason: 'blank' };
    var i = s.indexOf(',');
    if (i < 0) return { code: '', ok: false, reason: 'format' }; // not "Last, First": do not guess
    var last = letters(s.slice(0, i));
    var first = letters(s.slice(i + 1).replace(/\([^)]*\)/g, '').trim().split(/\s+/)[0] || '');
    if (!last || !first) return { code: '', ok: false, reason: 'format' };
    return { code: (last.slice(0, 3) + first.slice(0, 3)).toUpperCase(), ok: true };
  }

  // ---------- dates ----------
  var EPOCH = Date.UTC(1899, 11, 30);
  function serialOfDate(d) { return Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EPOCH) / 86400000); }
  function dateOfSerial(n) { return new Date(EPOCH + Math.floor(n) * 86400000); } // UTC fields
  function isWeekend(serial) { var d = dateOfSerial(serial).getUTCDay(); return d === 0 || d === 6; }
  function addBusinessDays(serial, n, holidays) {
    var h = {}; (holidays || []).forEach(function (x) { h[x] = 1; });
    var s = serial;
    while (n > 0) { s++; if (!isWeekend(s) && !h[s]) n--; }
    return s;
  }
  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function fmtDate(serial) {
    if (typeof serial !== 'number') return '';
    var d = dateOfSerial(serial);
    return pad(d.getUTCMonth() + 1) + '/' + pad(d.getUTCDate()) + '/' + pad(d.getUTCFullYear() % 100);
  }
  function fmtTime(frac) {
    if (typeof frac !== 'number') return String(frac || '');
    var mins = Math.round((frac - Math.floor(frac)) * 1440);
    var h = Math.floor(mins / 60) % 24, m = mins % 60;
    return (h % 12 || 12) + ':' + pad(m) + ' ' + (h < 12 ? 'AM' : 'PM');
  }
  // Text dates: M/D/YY(YY) or YYYY-MM-DD, optional trailing time. Same parser as the Office Script.
  function textDateSerial(s) {
    s = String(s).trim();
    var m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})\b/.exec(s), y, mo, d;
    if (m) { y = +m[3]; if (y < 100) y += 2000; mo = +m[1]; d = +m[2]; }
    else if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})\b/.exec(s))) { y = +m[1]; mo = +m[2]; d = +m[3]; }
    else return null;
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    return Math.round((Date.UTC(y, mo - 1, d) - EPOCH) / 86400000);
  }
  // Text times: "9:30 AM", "14:00", "9:30:00 PM". Returns a day fraction or null.
  function textTimeFrac(s) {
    var m = /^\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?\s*$/.exec(String(s));
    if (!m) return null;
    var h = +m[1], mi = +m[2];
    if (m[3]) { if (h < 1 || h > 12) return null; h = h % 12 + (/p/i.test(m[3]) ? 12 : 0); }
    if (h > 23 || mi > 59) return null;
    return (h * 60 + mi) / 1440;
  }
  function toSerial(v) { // Office.js gives numbers for dates; tolerate text dates
    if (typeof v === 'number') return v;
    if (typeof v === 'string' && v.trim()) return textDateSerial(v);
    return null;
  }
  function toTimeFrac(v) {
    if (typeof v === 'number') return v;
    if (typeof v === 'string' && v.trim()) return textTimeFrac(v);
    return null;
  }
  // Report freshness: Confirm Settings!I2 holds when the List report was pulled. Must be today.
  function reportFreshness(stamp, todaySerial) {
    var s = typeof stamp === 'number' ? stamp : (stamp ? textDateSerial(stamp) : null);
    if (s === null) return { ok: false, day: null };
    return { ok: Math.floor(s) === todaySerial, day: Math.floor(s) };
  }

  // ---------- status rules ----------
  function hasConfirmed(t) {
    var s = String(t).replace(/\bnot\s+confirm/gi, ' ').replace(/\bunconfirm/gi, ' ');
    return /\bconf/i.test(s); // "confirmed", "CONF", "Luma-Confirmed", "Confirmed - GB"
  }
  function hasRescheduled(t) { return /resched/i.test(t); }
  function hasCancelled(t) { return /cancel/i.test(t); }

  // Column positions (0-based) in the List sheet. Verified against headers at load time.
  var COL = { id: 0, name: 1, loc: 2, event: 3, date: 4, time: 5, rendering: 8, created: 9, k: 10, l: 11, m: 12, n: 13, notes: 14 };
  var EXPECTED_HEADERS = { 0: 'md rc', 1: 'pat name', 2: 'loc name', 4: 'appt dt', 5: 'beg tm' };

  function checkHeaders(headerRow) {
    var bad = [];
    Object.keys(EXPECTED_HEADERS).forEach(function (i) {
      if (String(headerRow[i] || '').trim().toLowerCase() !== EXPECTED_HEADERS[i]) bad.push(Number(i));
    });
    return bad;
  }

  /* rows: array of value arrays (A:O), firstRow: sheet row number of rows[0].
     opts: { todaySerial, includeToday=true, skipCancelled=true, holidays=[] } */
  function findOpportunities(rows, firstRow, opts) {
    opts = opts || {};
    var includeToday = opts.includeToday !== false, skipCancelled = opts.skipCancelled !== false;
    var start = includeToday ? opts.todaySerial : addBusinessDays(opts.todaySerial, 1, opts.holidays);
    var end = addBusinessDays(opts.todaySerial, 2, opts.holidays);
    var res = { start: start, end: end, byClinic: {}, unmapped: {}, skipped: [], badNames: 0, stats: { inWindow: 0, done: 0, cancelled: 0, flagged: 0 } };

    rows.forEach(function (r, idx) {
      var sheetRow = firstRow + idx;
      if (r.every(function (c) { return c === '' || c == null; })) return;
      var d = toSerial(r[COL.date]);
      if (d === null) {
        res.skipped.push({ row: sheetRow, why: 'Appt Dt is not a date' }); return;
      }
      if (d < start || d > end) return;
      res.stats.inWindow++;

      var statusText = [r[COL.k], r[COL.l], r[COL.m], r[COL.n]].map(function (x) { return x == null ? '' : String(x); }).join(' | ');
      if (hasConfirmed(statusText) || hasRescheduled(statusText)) { res.stats.done++; return; }
      if (skipCancelled && hasCancelled(statusText)) { res.stats.cancelled++; return; }

      var locRaw = String(r[COL.loc] || '').trim();
      var clinic = clinicForLocation(locRaw);
      var nm = toLASFIR(r[COL.name]);
      if (!nm.ok) res.badNames++;
      var item = {
        row: sheetRow, mdrc: r[COL.id], lasfir: nm.code, nameIssue: nm.ok ? '' : nm.reason,
        event: r[COL.event] || '', dateSerial: d, timeFrac: toTimeFrac(r[COL.time]),
        rendering: r[COL.rendering] || '', created: toSerial(r[COL.created]),
        k: r[COL.k] || '', l: r[COL.l] || '', m: r[COL.m] || '', n: r[COL.n] || '', notes: r[COL.notes] || '',
        loc: locRaw
      };
      if (!clinic) { (res.unmapped[locRaw || '(blank)'] = res.unmapped[locRaw || '(blank)'] || []).push(item); return; }
      (res.byClinic[clinic] = res.byClinic[clinic] || []).push(item);
      res.stats.flagged++;
    });

    Object.keys(res.byClinic).forEach(function (c) {
      res.byClinic[c].sort(function (a, b) { return a.dateSerial - b.dateSerial || (a.timeFrac || 0) - (b.timeFrac || 0) || a.row - b.row; });
    });
    return res;
  }

  // ---------- email ----------
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function rowLink(baseUrl, row) {
    if (!baseUrl) return '';
    return baseUrl + (baseUrl.indexOf('?') < 0 ? '?' : '&') + 'activeCell=' + encodeURIComponent('List!A' + row);
  }

  var INTRO = 'Hello!';
  var BODY = 'Please see the below patients requiring confirmation within the next 2 business days, including any outstanding for today if applicable.';
  var CHECKLIST = 'Please standardly verify email address, date/time and appointment medium (in-person/tele), reminder to account for annual consent forms during check-in, directions to clinic, clinic extension, help email, enroll in portal, etc.';

  function buildEmail(clinicName, items, res, opts) {
    opts = opts || {};
    var cols = [
      ['Md Rc', function (i) { return i.mdrc; }],
      ['Pat Name', function (i) { return i.lasfir || (i.nameIssue === 'blank' ? '(missing)' : '(check name)'); }],
      ['Event', function (i) { return i.event; }],
      ['Appt Dt', function (i) { return fmtDate(i.dateSerial); }],
      ['Beg Tm', function (i) { return fmtTime(i.timeFrac); }],
      ['Rendering', function (i) { return i.rendering; }],
      ['Create Date', function (i) { return fmtDate(i.created); }],
      ['96h Contact Outcome (Patient Access Team)', function (i) { return i.k; }],
      ['Pre-Reg Setup Call (Patient Access Team)', function (i) { return i.l; }],
      ['Status of 48 hour Outreach (Patient Access Team)', function (i) { return i.m; }],
      ['Confirmation Call 48 Hours (Ops)', function (i) { return i.n; }]
    ];
    if (opts.includeNotes) cols.push(['Clinic Notes', function (i) { return i.notes; }]);

    var th = 'style="background:#1f4e79;color:#fff;text-align:left;padding:6px 8px;border:1px solid #ccc;font-size:12px"';
    var td = 'style="padding:6px 8px;border:1px solid #ccc;font-size:12px;vertical-align:top"';
    var head = '<tr>' + cols.map(function (c) { return '<th ' + th + '>' + esc(c[0]) + '</th>'; }).join('') + '<th ' + th + '>Link</th></tr>';
    var body = items.map(function (i) {
      var link = rowLink(opts.baseUrl, i.row);
      return '<tr>' + cols.map(function (c) { return '<td ' + td + '>' + esc(c[1](i)) + '</td>'; }).join('') +
        '<td ' + td + '>' + (link ? '<a href="' + esc(link) + '">View</a>' : '') + '</td></tr>';
    }).join('');

    var html = '<div style="font-family:Calibri,Arial,sans-serif;font-size:14px">' +
      '<p>' + esc(INTRO) + '</p><p>' + esc(BODY) + '</p><p>' + esc(CHECKLIST) + '</p>' +
      '<table style="border-collapse:collapse">' + head + body + '</table></div>';

    var text = [INTRO, '', BODY, '', CHECKLIST, ''].concat(items.map(function (i) {
      var link = rowLink(opts.baseUrl, i.row);
      return cols.map(function (c) { return c[1](i); }).join(' | ') + (link ? ' | View: ' + link : '');
    })).join('\n');

    var range = res ? fmtDate(res.start) + (res.end !== res.start ? '-' + fmtDate(res.end) : '') : '';
    return { to: emailForClinic(clinicName), subject: 'Appointment confirmations needed: ' + clinicName + (range ? ' (' + range + ')' : ''), html: html, text: text, count: items.length };
  }

  return {
    CLINICS: CLINICS, COL: COL, clinicForLocation: clinicForLocation, emailForClinic: emailForClinic,
    toLASFIR: toLASFIR, serialOfDate: serialOfDate, addBusinessDays: addBusinessDays, fmtDate: fmtDate, fmtTime: fmtTime,
    toSerial: toSerial, toTimeFrac: toTimeFrac, reportFreshness: reportFreshness,
    hasConfirmed: hasConfirmed, hasRescheduled: hasRescheduled, hasCancelled: hasCancelled,
    checkHeaders: checkHeaders, findOpportunities: findOpportunities, buildEmail: buildEmail, rowLink: rowLink
  };
});
