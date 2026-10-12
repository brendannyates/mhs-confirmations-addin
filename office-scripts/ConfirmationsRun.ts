// Office Script: "Confirmations Run"
// Called by Power Automate at 8:00 and 16:00 (Pacific), weekdays. Returns JSON text. Sends nothing itself.
// Same rules as the add-in (core.js). Keep the two in sync.
//
// Input  todayIso: "yyyy-MM-dd" in Pacific time, supplied by the flow (Office Scripts run in UTC).
// Output {status: "ok" | "stale" | "error", message, emails: [{clinic,to,subject,html,count}], warnings: [], ownerHtml}
// ownerHtml: one summary for the workbook owner (status, message, every warning). The flow emails it when warnings is not empty.

type Cell = string | number | boolean;
interface Item {
  row: number; mdrc: Cell; lasfir: string; nameIssue: string; event: string; dateSerial: number; timeFrac: number | null;
  rendering: string; created: number | null; k: string; l: string; m: string; n: string;
}

const CLINICS: string[][] = [
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
];
const ALIASES: { [k: string]: string } = { 'castro street': 'Castro - CA' };

const LIST_SHEET = 'List';
const SETTINGS_SHEET = 'Confirm Settings';
const EPOCH = Date.UTC(1899, 11, 30);
// Signature appended to every clinic email (not the owner alerts).
const SIGNATURE_HTML = "<p style=\"margin:16px 0 0\">Best,</p><p style=\"margin:8px 0 0;line-height:1.35\"><b>Brendan Yates</b> (he/him)<br>Senior Practice Manager<br>SF Castro | San Rafael<br>P: (844) 867- 8444 opt. 2 | F: (415) 964-5419<br><a href=\"https://www.mindfulhealthsolutions.com\">www.mindfulhealthsolutions.com</a></p><p style=\"margin:16px 0 0;font-size:11px;color:#666\">This email communication and any attachments may contain confidential information for the use of the designated recipients named above. Distribution, reproduction or any other use of this transmission by any party other than the intended recipient is prohibited and may be subject to protection under the law, including the Health Insurance Portability and Accountability Act (HIPAA). If you receive this e-mail in error, please notify the sender immediately by reply e-mail and permanently delete this e-mail and any attachments.</p>";
const REPORT_HEADER = 'Report pulled (date/time). Must be today or the scheduled run sends nothing.';
type Out = { status: string; message: string; emails: object[]; warnings: string[]; ownerHtml: string };

function main(workbook: ExcelScript.Workbook, todayIso: string): string {
  const out: Out = { status: 'ok', message: '', emails: [], warnings: [], ownerHtml: '' };
  try {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(todayIso || '');
    if (!m) return fail(out, 'todayIso must be yyyy-MM-dd, got "' + todayIso + '"');
    const today = Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - EPOCH) / 86400000);

    // ---- settings: opt-ins, holidays, workbook link ----
    let sws = workbook.getWorksheet(SETTINGS_SHEET);
    if (!sws) {
      // No add-in needed: the script can create the sheet. Everything on it can be filled in by hand.
      sws = workbook.addWorksheet(SETTINGS_SHEET);
      const data: string[][] = [['ClinicName', 'ClinicEmail', 'Opt In (Yes/No)']].concat(CLINICS.map((c) => [c[0], c[1], 'No']));
      sws.getRange('A1:C' + data.length).setValues(data);
      sws.getRange('E1').setValue('Holidays (YYYY-MM-DD, optional)');
      sws.getRange('G1').setValue('Workbook link for View (paste in G2 if not using the add-in)');
      sws.getRange('I1').setValue(REPORT_HEADER);
      sws.getRange('A1:I1').getFormat().getFont().setBold(true);
      sws.getRange('C2:C' + data.length).getDataValidation().setRule({ list: { inCellDropDown: true, source: 'Yes,No' } });
      return fail(out, 'Created "' + SETTINGS_SHEET + '" with every clinic set to No. Set Opt In in column C, paste the workbook link in G2, then run again. Nothing sent.');
    }
    if (!String(sws.getRange('I1').getValue() || '').trim()) sws.getRange('I1').setValue(REPORT_HEADER); // sheets from older versions
    const sused = sws.getUsedRange(true);
    const sv: Cell[][] = sused ? sused.getValues() : [];
    const optIn: { [k: string]: boolean } = {};
    const holidays: number[] = [];
    for (let i = 1; i < sv.length; i++) {
      if (sv[i][0]) optIn[String(sv[i][0]).trim()] = /^y/i.test(String(sv[i][2] || ''));
      const h = sv[i][4];
      if (typeof h === 'number') holidays.push(h);
      else if (h) { const t = Date.parse(String(h) + 'T00:00:00Z'); if (!isNaN(t)) holidays.push(Math.round((t - EPOCH) / 86400000)); }
    }
    const baseUrl = sv.length > 1 && sv[1][6] ? String(sv[1][6]).trim() : ''; // Confirm Settings!G2
    if (!baseUrl) out.warnings.push('Confirm Settings!G2 (workbook link) is blank, so emails have no View links.');

    // ---- freshness: Confirm Settings!I2 must say the report was pulled today ----
    const stampRaw = sws.getRange('I2').getValue();
    const stamp = typeof stampRaw === 'number' ? Math.floor(stampRaw) : textDateSerial(String(stampRaw || ''));
    if (stamp === null || stamp !== today) {
      out.status = 'stale';
      out.message = (stamp === null ? 'No report date in ' + SETTINGS_SHEET + '!I2' : 'Report date in ' + SETTINGS_SHEET + '!I2 is ' + fmtDate(stamp) + ', not today (' + fmtDate(today) + ')') +
        '. Paste a fresh report into List, then set I2 to today (or click "Mark report pulled now" in the add-in). Nothing sent.';
      return finish(out);
    }

    // ---- list data ----
    const ws = workbook.getWorksheet(LIST_SHEET);
    if (!ws) return fail(out, 'Sheet "' + LIST_SHEET + '" not found.');
    const used = ws.getUsedRange(true);
    if (!used) return fail(out, 'List sheet is empty.');
    const values: Cell[][] = ws.getRangeByIndexes(0, 0, used.getRowIndex() + used.getRowCount(), 15).getValues();
    let hdr = -1;
    for (let i = 0; i < Math.min(10, values.length); i++) if (String(values[i][0]).trim().toLowerCase() === 'md rc') { hdr = i; break; }
    if (hdr < 0) return fail(out, 'Header row (column A "Md Rc") not found on List.');
    const expected: { [k: number]: string } = { 0: 'md rc', 1: 'pat name', 2: 'loc name', 4: 'appt dt', 5: 'beg tm' };
    for (const k of Object.keys(expected)) {
      if (String(values[hdr][+k]).trim().toLowerCase() !== expected[+k]) return fail(out, 'Column headers moved on List (expected Md Rc, Pat Name, Loc Name, Appt Dt, Beg Tm). Nothing sent.');
    }
    const rows = values.slice(hdr + 1);
    const firstRow = hdr + 2;

    // ---- window + stale check ----
    const end = addBusinessDays(today, 2, holidays);
    const dates: (number | null)[] = rows.map((r) => toSerial(r[4]));
    let maxDate = 0;
    const badDateRows: number[] = [];
    for (let i = 0; i < rows.length; i++) {
      const blank = rows[i].every((c) => c === '' || c == null);
      if (blank) continue;
      const d = dates[i];
      if (d === null) badDateRows.push(firstRow + i);
      else if (d > maxDate) maxDate = d;
    }
    if (badDateRows.length) out.warnings.push(badDateRows.length + ' row(s) on List have an Appt Dt that is not a date and were skipped: sheet row ' +
      badDateRows.slice(0, 25).join(', ') + (badDateRows.length > 25 ? ' and more' : '') + '.');
    if (maxDate < end) {
      out.status = 'stale';
      out.message = (maxDate ? 'List only runs through ' + fmtDate(maxDate) : 'List has no readable appointment dates') +
        ' but the window ends ' + fmtDate(end) + '. Paste a fresh report. Nothing sent.';
      return finish(out);
    }

    // ---- find opportunities ----
    const byClinic: { [k: string]: Item[] } = {};
    let badNames = 0;
    const unmapped: { [k: string]: number[] } = {};
    for (let idx = 0; idx < rows.length; idx++) {
      const r = rows[idx];
      const d = dates[idx];
      if (d === null || d < today || d > end) continue;
      const status = [r[10], r[11], r[12], r[13]].map((x) => x == null ? '' : String(x)).join(' | ');
      if (hasConfirmed(status) || /resched/i.test(status) || /cancel/i.test(status)) continue;
      const clinic = clinicFor(String(r[2] || ''));
      if (!clinic) { const key = String(r[2] || '(blank)'); (unmapped[key] = unmapped[key] || []).push(firstRow + idx); continue; }
      const nm = toLASFIR(r[1]);
      if (!nm.ok) badNames++;
      const item: Item = {
        row: firstRow + idx, mdrc: r[0], lasfir: nm.code, nameIssue: nm.ok ? '' : nm.reason, event: String(r[3] || ''),
        dateSerial: d, timeFrac: toTimeFrac(r[5]), rendering: String(r[8] || ''),
        created: toSerial(r[9]),
        k: String(r[10] || ''), l: String(r[11] || ''), m: String(r[12] || ''), n: String(r[13] || '')
      };
      if (!byClinic[clinic]) byClinic[clinic] = [];
      byClinic[clinic].push(item);
    }

    for (const loc of Object.keys(unmapped)) out.warnings.push(unmapped[loc].length + ' appointment(s) at "' + loc + '" have no clinic email mapped and were NOT emailed (sheet row ' + unmapped[loc].slice(0, 25).join(', ') + ').');
    if (badNames) out.warnings.push(badNames + ' row(s) in the window have a blank or non "Last, First" Pat Name, so the email shows (missing) or (check name).');

    // ---- build one email per OPTED-IN clinic with something to confirm ----
    for (const c of CLINICS) {
      const items = byClinic[c[0]];
      if (!items || !items.length || !optIn[c[0]]) continue;
      items.sort((a, b) => a.dateSerial - b.dateSerial || (a.timeFrac || 0) - (b.timeFrac || 0) || a.row - b.row);
      out.emails.push(buildEmail(c[0], c[1], items, today, end, baseUrl));
    }
    out.message = out.emails.length + ' email(s) ready for ' + fmtDate(today) + ' to ' + fmtDate(end) + '.';
    return finish(out);
  } catch (e) {
    return fail(out, 'Unexpected error: ' + String(e));
  }
}

function fail(out: Out, msg: string): string {
  out.status = 'error'; out.message = msg; out.emails = [];
  return finish(out);
}

function finish(out: Out): string {
  if (out.status !== 'ok') out.emails = []; // never send clinic emails on a stale or error run
  out.ownerHtml = '<div style="font-family:Calibri,Arial,sans-serif;font-size:14px"><p><b>Confirmations Run: ' + esc(out.status) + '</b></p><p>' + esc(out.message) + '</p>' +
    (out.warnings.length ? '<p>Warnings:</p><ul>' + out.warnings.map((w) => '<li>' + esc(w) + '</li>').join('') + '</ul>' : '') + '</div>';
  return JSON.stringify(out);
}

// Text dates: M/D/YY(YY) or YYYY-MM-DD, optional trailing time. Same parser as core.js.
function textDateSerial(s: string): number | null {
  s = s.trim();
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})\b/.exec(s);
  let y = 0, mo = 0, d = 0;
  if (m) { y = +m[3]; if (y < 100) y += 2000; mo = +m[1]; d = +m[2]; }
  else {
    m = /^(\d{4})-(\d{1,2})-(\d{1,2})\b/.exec(s);
    if (!m) return null;
    y = +m[1]; mo = +m[2]; d = +m[3];
  }
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return Math.round((Date.UTC(y, mo - 1, d) - EPOCH) / 86400000);
}
function textTimeFrac(s: string): number | null {
  const m = /^\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?\s*$/.exec(s);
  if (!m) return null;
  let h = +m[1];
  const mi = +m[2];
  if (m[3]) { if (h < 1 || h > 12) return null; h = h % 12 + (/p/i.test(m[3]) ? 12 : 0); }
  if (h > 23 || mi > 59) return null;
  return (h * 60 + mi) / 1440;
}
function toSerial(v: Cell): number | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'string' && v.trim()) return textDateSerial(v);
  return null;
}
function toTimeFrac(v: Cell): number | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'string' && v.trim()) return textTimeFrac(v);
  return null;
}

function clinicFor(loc: string): string {
  const k = loc.replace(/\s+/g, ' ').trim().toLowerCase();
  if (ALIASES[k]) return ALIASES[k];
  for (const c of CLINICS) if (c[0].replace(/\s+-\s+[A-Z]{2}$/, '').toLowerCase() === k) return c[0];
  return '';
}

function letters(s: string): string { return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z]/g, ''); }
function toLASFIR(raw: Cell): { code: string; ok: boolean; reason: string } {
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return { code: '', ok: false, reason: 'blank' };
  const i = s.indexOf(',');
  if (i < 0) return { code: '', ok: false, reason: 'format' };
  const last = letters(s.slice(0, i));
  const first = letters(s.slice(i + 1).replace(/\([^)]*\)/g, '').trim().split(/\s+/)[0] || '');
  if (!last || !first) return { code: '', ok: false, reason: 'format' };
  return { code: (last.slice(0, 3) + first.slice(0, 3)).toUpperCase(), ok: true, reason: '' };
}

function hasConfirmed(t: string): boolean {
  return /\bconf/i.test(t.replace(/\bnot\s+confirm/gi, ' ').replace(/\bunconfirm/gi, ' '));
}

function isWeekend(serial: number): boolean { const d = new Date(EPOCH + serial * 86400000).getUTCDay(); return d === 0 || d === 6; }
function addBusinessDays(serial: number, n: number, holidays: number[]): number {
  let s = serial;
  while (n > 0) { s++; if (!isWeekend(s) && holidays.indexOf(s) < 0) n--; }
  return s;
}
function pad(n: number): string { return n < 10 ? '0' + n : '' + n; }
function fmtDate(serial: number | null): string {
  if (typeof serial !== 'number') return '';
  const d = new Date(EPOCH + Math.floor(serial) * 86400000);
  return pad(d.getUTCMonth() + 1) + '/' + pad(d.getUTCDate()) + '/' + pad(d.getUTCFullYear() % 100);
}
function fmtTime(frac: number | null): string {
  if (typeof frac !== 'number') return '';
  const mins = Math.round((frac - Math.floor(frac)) * 1440);
  const h = Math.floor(mins / 60) % 24, mm = mins % 60;
  return (h % 12 || 12) + ':' + pad(mm) + ' ' + (h < 12 ? 'AM' : 'PM');
}
function esc(s: Cell | null): string { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

function buildEmail(clinic: string, to: string, items: Item[], start: number, end: number, baseUrl: string) {
  const cols: [string, (i: Item) => Cell][] = [
    ['Md Rc', (i) => i.mdrc],
    ['Pat Name', (i) => i.lasfir || (i.nameIssue === 'blank' ? '(missing)' : '(check name)')],
    ['Event', (i) => i.event],
    ['Appt Dt', (i) => fmtDate(i.dateSerial)],
    ['Beg Tm', (i) => fmtTime(i.timeFrac)],
    ['Rendering', (i) => i.rendering],
    ['Create Date', (i) => fmtDate(i.created)],
    ['96h Contact Outcome (Patient Access Team)', (i) => i.k],
    ['Pre-Reg Setup Call (Patient Access Team)', (i) => i.l],
    ['Status of 48 hour Outreach (Patient Access Team)', (i) => i.m],
    ['Confirmation Call 48 Hours (Ops)', (i) => i.n]
  ];
  const th = 'style="background:#1f4e79;color:#fff;text-align:left;padding:6px 8px;border:1px solid #ccc;font-size:12px"';
  const td = 'style="padding:6px 8px;border:1px solid #ccc;font-size:12px;vertical-align:top"';
  const head = '<tr>' + cols.map((c) => '<th ' + th + '>' + esc(c[0]) + '</th>').join('') + '<th ' + th + '>Link</th></tr>';
  const body = items.map((i) => {
    const link = baseUrl ? baseUrl + (baseUrl.indexOf('?') < 0 ? '?' : '&') + 'activeCell=' + encodeURIComponent('List!A' + i.row) : '';
    return '<tr>' + cols.map((c) => '<td ' + td + '>' + esc(c[1](i)) + '</td>').join('') +
      '<td ' + td + '>' + (link ? '<a href="' + esc(link) + '">View</a>' : '') + '</td></tr>';
  }).join('');
  const html = '<div style="font-family:Calibri,Arial,sans-serif;font-size:14px">' +
    '<p>Hello!</p><p>Please see the below patients requiring confirmation within the next 2 business days, including any outstanding for today if applicable.</p>' +
    '<p>Please standardly verify email address, date/time and appointment medium (in-person/tele), reminder to account for annual consent forms during check-in, directions to clinic, clinic extension, help email, enroll in portal, etc.</p>' +
    '<table style="border-collapse:collapse">' + head + body + '</table>' + SIGNATURE_HTML + '</div>';
  return { clinic: clinic, to: to, subject: 'Appointment confirmations needed: ' + clinic + ' (' + fmtDate(start) + (end !== start ? '-' + fmtDate(end) : '') + ')', html: html, count: items.length };
}
