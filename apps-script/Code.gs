/**
 * OCP Call for Demos – submission receiver (Google Apps Script web app)
 *
 * One deployment serves every Summit. Each form page sends its formId,
 * and FORMS below routes it to that Summit's own Google Sheet.
 *
 * Deploy: Extensions → Apps Script (or script.google.com) → Deploy → New deployment
 *   Type: Web app · Execute as: Me (use a shared OCP events account) · Who has access: Anyone
 * Copy the /exec URL into each form's config as `endpoint`.
 * After editing this file: Deploy → Manage deployments → Edit → New version (URL stays the same).
 */

// ---- Add one line per Summit/form. The spreadsheet ID is the long string in the Sheet's URL.
const FORMS = {
  'emea-2027-iv': {
    sheetId: '1V1YS5CcBHaC5LiAbBfJOgsRcLpqAQRy97fISHUt4oqo', tab: 'Submissions',
    summit: '2027 OCP EMEA Summit', location: 'Hamburg, Germany', closes: '2027-01-12T07:59:00Z',
    page: 'forms/emea-2027-iv/', closesLabel: 'January 11, 2027 at 11:59 PM PT',
    // Dates shown in the confirmation email's "What happens next"
    timeline: [['Jan 11, 2027', 'Call for demos closes, 11:59 PM PT'], ['Week of Feb 1', 'Decision emails go out'], ['Apr 6–8, 2027', 'OCP EMEA Summit, Hamburg, Germany']],
  },
  'global-2027-iv': {
    sheetId: 'PASTE_GLOBAL_2027_SHEET_ID', tab: 'Submissions',
    summit: '2027 OCP Global Summit', location: 'San Jose, California', closes: '2027-06-01T06:59:00Z',
    page: 'forms/global-2027-iv/', closesLabel: '',
    timeline: [],
  },
};

// Logos for the confirmation email, attached inline so they show even when remote images are blocked.
const SITE = 'https://jameskellynet.github.io/ocp-innovation-village/';
const LOGOS = { ocpLogo: SITE + 'assets/ocp-logo.png', ivLogo: SITE + 'assets/iv-logo.png' };

// Column order in every Sheet. Matches the fields the form sends.
const COLUMNS = [
  ['Submission Date', d => new Date()],
  ['Reference',       d => d._ref],
  ['First Name',      d => d.firstName],
  ['Last Name',       d => d.lastName],
  ['Email',           d => d.email],
  ['Company Name',    d => d.company],
  ['Job Title',       d => d.jobTitle],
  ['Station / Demo Title', d => d.title],
  ['Category',        d => d.category],
  ['Abstract',        d => d.abstract],
  ['Why a demo, not a presentation', d => d.whyDemo],
  ['Attendee takeaway', d => d.takeaway],
  ['Equipment',       d => d.equipment],
  ['Proposed Presenters', d => (d.presenters || []).map(p => [p.name, p.title, p.company, p.email].filter(String).join(', ')).join('\n')],
  ['Eligibility confirmed', d => d.eligibilityConfirmed ? 'Yes' : 'No'],
  ['Privacy & T&C accepted', d => d.consent ? 'Yes' : 'No'],
];

// Staff columns after the form columns. The form never writes these; reviewers and Claude do.
const STAFF_COLUMNS = ['Review Status', 'Notes', 'Acceptance Notice Sent', 'Confirmed'];
const STATUS_OPTIONS = ['New', 'Under review', 'Accepted', 'Waitlist', 'Declined', 'Withdrawn'];

const LIMITS = { abstract: 1000, whyDemo: 500, takeaway: 500 };

/**
 * Run once from the Apps Script editor (select setupSheets → Run) after adding a Sheet to FORMS.
 * Safe to re-run: Sheets that already have headers are left alone.
 */
function setupSheets() {
  const done = [];
  for (const [formId, form] of Object.entries(FORMS)) {
    if (form.sheetId.startsWith('PASTE_')) continue;
    const ss = SpreadsheetApp.openById(form.sheetId);
    let sh = ss.getSheetByName(form.tab);
    if (!sh) {
      const first = ss.getSheets()[0];
      if (ss.getSheets().length === 1 && first.getLastRow() === 0) sh = first.setName(form.tab);
      else sh = ss.insertSheet(form.tab, 0);
    }
    if (sh.getLastRow() > 0) { done.push(`${formId}: already set up`); continue; }

    const headers = COLUMNS.map(c => c[0]).concat(STAFF_COLUMNS);
    const n = headers.length;
    const nForm = COLUMNS.length;
    sh.getRange(1, 1, 1, n).setValues([headers])
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#19194D')
      .setWrap(true).setVerticalAlignment('middle');
    sh.getRange(1, nForm + 1, 1, STAFF_COLUMNS.length).setBackground('#343895');   // staff columns
    sh.setRowHeight(1, 42);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(2);

    // Column widths and text wrapping for long answers
    const widths = { 'Submission Date': 150, 'Reference': 100, 'Email': 220, 'Station / Demo Title': 260,
      'Category': 200, 'Abstract': 420, 'Why a demo, not a presentation': 320, 'Attendee takeaway': 320,
      'Equipment': 280, 'Proposed Presenters': 280, 'Notes': 300 };
    headers.forEach((h, i) => sh.setColumnWidth(i + 1, widths[h] || 140));
    const body = sh.getRange(2, 1, sh.getMaxRows() - 1, n);
    body.setVerticalAlignment('top').setWrap(true);
    sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');

    // Staff columns: status dropdown and checkboxes
    const col = (h) => headers.indexOf(h) + 1;
    const rows = sh.getMaxRows() - 1;
    sh.getRange(2, col('Review Status'), rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS, true).setAllowInvalid(false).build());
    // Checkbox validation without values, so empty rows stay empty and new submissions append at the top
    const box = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sh.getRange(2, col('Acceptance Notice Sent'), rows, 1).setDataValidation(box);
    sh.getRange(2, col('Confirmed'), rows, 1).setDataValidation(box);

    // Status colors
    const st = sh.getRange(2, col('Review Status'), rows, 1);
    const rule = (text, bg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setRanges([st]).build();
    sh.setConditionalFormatRules([rule('Accepted', '#DDEFC7'), rule('Waitlist', '#FDEFC9'), rule('Declined', '#F3D9D9'), rule('Under review', '#E6E7F4')]);

    sh.getRange(2, 1, rows, n).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    done.push(`${formId}: set up "${form.tab}" in ${ss.getName()}`);
  }
  Logger.log(done.join('\n'));
}

/**
 * Run once from the editor after pasting a new version (select authorize → Run), so Google asks
 * for any new permission (fetching the email logos) before a live submission needs it.
 */
function authorize() {
  UrlFetchApp.fetch(LOGOS.ocpLogo);
  editSecret();
  Logger.log('Permissions OK. Email quota left today: ' + MailApp.getRemainingDailyQuota());
}

// ---- Private edit links -------------------------------------------------------------
// Each receipt carries ?edit=<reference>&t=<signature>. The signature is an HMAC of the form and
// reference with a random secret kept in this script's properties, so links can't be guessed.
function editSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('EDIT_SECRET');
  if (!secret) { secret = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('EDIT_SECRET', secret); }
  return secret;
}
function editToken(formId, ref) {
  const sig = Utilities.computeHmacSha256Signature(formId + '|' + ref, editSecret());
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '').slice(0, 32);
}
function editLink(formId, ref) {
  return SITE + FORMS[formId].page + '?edit=' + encodeURIComponent(ref) + '&t=' + editToken(formId, ref);
}
function tokenOk(formId, ref, t) {
  const want = editToken(formId, ref), got = String(t || '');
  if (want.length !== got.length) return false;
  let diff = 0;
  for (let i = 0; i < want.length; i++) diff |= want.charCodeAt(i) ^ got.charCodeAt(i);
  return diff === 0;
}
function findRow(sh, ref) {
  const max = sh.getMaxRows();
  if (max < 2) return 0;
  const refs = sh.getRange(2, 2, max - 1, 1).getValues();
  for (let i = 0; i < refs.length; i++) if (refs[i][0] === ref) return i + 2;
  return 0;
}
function parsePresenters(text) {
  return String(text || '').split('\n').filter(String).map(line => {
    const parts = line.split(', ');
    const email = parts.find(x => x.includes('@')) || '';
    const rest = parts.filter(x => x !== email);
    return { name: rest[0] || '', title: rest[1] || '', company: rest.slice(2).join(', '), email };
  });
}
const isClosed = form => form.closes && new Date() > new Date(form.closes);

// The form calls this to load a submission for editing.
function doGet(e) {
  try {
    const q = (e && e.parameter) || {};
    const form = FORMS[q.formId];
    if (q.action !== 'load' || !form || !q.ref || !tokenOk(q.formId, q.ref, q.t))
      return reply({ ok: false, error: 'This edit link is not valid.' });
    if (isClosed(form)) return reply({ ok: false, error: 'The deadline has passed, so this submission can no longer be edited.' });
    const sh = SpreadsheetApp.openById(form.sheetId).getSheetByName(form.tab);
    const r = sh ? findRow(sh, q.ref) : 0;
    if (!r) return reply({ ok: false, error: 'We could not find this submission.' });
    const v = sh.getRange(r, 1, 1, COLUMNS.length).getValues()[0];
    const get = h => v[COLUMNS.findIndex(c => c[0] === h)];
    return reply({ ok: true, ref: q.ref, data: {
      firstName: get('First Name'), lastName: get('Last Name'), email: get('Email'), company: get('Company Name'),
      jobTitle: get('Job Title'), title: get('Station / Demo Title'), category: get('Category'), abstract: get('Abstract'),
      whyDemo: get('Why a demo, not a presentation'), takeaway: get('Attendee takeaway'), equipment: get('Equipment'),
      presenters: parsePresenters(get('Proposed Presenters')) } });
  } catch (err) {
    return reply({ ok: false, error: 'Something went wrong on our side.' });
  }
}

// Run from the editor to list every submission's edit link (for resending one by hand).
function showEditLinks() {
  for (const [formId, form] of Object.entries(FORMS)) {
    if (form.sheetId.startsWith('PASTE_')) continue;
    const sh = SpreadsheetApp.openById(form.sheetId).getSheetByName(form.tab);
    if (!sh || sh.getMaxRows() < 2) continue;
    sh.getRange(2, 2, sh.getMaxRows() - 1, 4).getValues().forEach(r => {
      if (r[0]) Logger.log(`${r[0]}  ${r[3]}  ${editLink(formId, r[0])}`);
    });
  }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);

    if (d.website) return reply({ ok: true, ref: 'IV-OK' });          // honeypot: silently drop bots
    const form = FORMS[d.formId];
    if (!form) return reply({ ok: false, error: 'This form is not recognised.' });
    if (isClosed(form))
      return reply({ ok: false, error: 'The submission deadline has passed.' });
    const editing = !!d.editRef;
    if (editing && !tokenOk(d.formId, d.editRef, d.editToken))
      return reply({ ok: false, error: 'This edit link is not valid.' });

    // Server-side checks (never trust the browser alone)
    for (const k of ['firstName','lastName','email','company','title','category','abstract','whyDemo','takeaway','equipment'])
      if (!String(d[k] || '').trim()) return reply({ ok: false, error: 'A required answer is missing.' });
    for (const [k, max] of Object.entries(LIMITS))
      if (String(d[k]).length > max) return reply({ ok: false, error: 'An answer is over its character limit.' });
    if (!d.consent) return reply({ ok: false, error: 'Consent is required.' });

    d._ref = editing ? d.editRef : 'IV-' + Utilities.getUuid().slice(0, 6).toUpperCase();

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = SpreadsheetApp.openById(form.sheetId);
      const sh = ss.getSheetByName(form.tab) || ss.insertSheet(form.tab);
      if (sh.getLastRow() === 0) sh.appendRow(COLUMNS.map(c => c[0]));
      if (editing) {
        // Overwrite the answer columns (C onward); keep the original date, reference and staff columns.
        const r = findRow(sh, d._ref);
        if (!r) return reply({ ok: false, error: 'We could not find this submission.' });
        const answers = COLUMNS.slice(2).map(c => c[1](d));
        sh.getRange(r, 3, 1, answers.length).setValues([answers]);
        const notes = sh.getRange(r, COLUMNS.length + 2);   // Review Status, then Notes
        const stamp = 'Edited by submitter ' + Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd HH:mm') + ' PT';
        notes.setValue(notes.getValue() ? notes.getValue() + '\n' + stamp : stamp);
      } else {
        const values = COLUMNS.map(c => c[1](d)).concat(['New']);   // Review Status starts as New
        sh.getRange(nextRow(sh), 1, 1, values.length).setValues([values]);
      }
    } finally {
      lock.releaseLock();
    }

    sendReceipt(d, form, editing);

    return reply({ ok: true, ref: d._ref });
  } catch (err) {
    return reply({ ok: false, error: 'Something went wrong on our side.' });
  }
}

// Branded confirmation email: HTML with a plain-text fallback and inline logos.
function sendReceipt(d, form, updated) {
  const link = editLink(d.formId, d._ref);
  const until = form.closesLabel ? ` until ${form.closesLabel}` : '';
  const esc = v => String(v || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const when = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'MMMM d, yyyy');
  const images = {};
  for (const [cid, url] of Object.entries(LOGOS)) {
    try { images[cid] = UrlFetchApp.fetch(url).getBlob().setName(cid + '.png'); } catch (e) { /* send without that logo */ }
  }
  const logo = (cid, w, h, alt) => images[cid] ? `<img src="cid:${cid}" width="${w}" height="${h}" alt="${alt}" style="display:block;border:0;">` : `<span style="font:bold 14px Arial,sans-serif;color:#5F6062;">${alt}</span>`;
  const label = 'font:bold 11px Arial,sans-serif;letter-spacing:2px;text-transform:uppercase;color:#343895;';
  const row = (k, v, last) => `<tr><td style="padding:6px 16px${last ? ' 14px' : ''};color:#5F6062;width:120px;">${k}</td><td style="padding:6px 16px${last ? ' 14px' : ''};color:#19194D;${k === 'Reference' ? 'font-weight:bold;' : ''}">${v}</td></tr>`;
  const steps = (form.timeline || []).map((t, i, a) => {
    const b = i < a.length - 1 ? 'border-bottom:1px solid #EEF0F2;' : '';
    return `<tr><td style="padding:8px 0;${b}width:130px;color:#19194D;font-weight:bold;">${esc(t[0])}</td><td style="padding:8px 0;${b}">${esc(t[1])}</td></tr>`;
  }).join('');

  const html = `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#E9EBEF;"><tr><td align="center">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#FFFFFF;border-collapse:collapse;">
<tr><td style="padding:16px 32px;border-bottom:1px solid #DADCDF;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
<td align="left" valign="middle">${logo('ocpLogo', 100, 40, 'Open Compute Project')}</td>
<td align="right" valign="middle">${logo('ivLogo', 118, 48, 'OCP Innovation Village')}</td></tr></table></td></tr>
<tr><td style="background:#19194D;padding:28px 32px 26px;border-bottom:6px solid #8DC63F;">
<div style="font:bold 11px/1.4 Arial,sans-serif;letter-spacing:3px;text-transform:uppercase;color:#8DC63F;">${esc(form.summit)} · Innovation Village</div>
<div style="font:bold 26px/1.2 'Franklin Gothic Medium','Arial Narrow',Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;color:#FFFFFF;margin-top:10px;">${updated ? 'Submission updated' : 'Submission received'}</div></td></tr>
<tr><td style="padding:28px 32px 8px;font:15px/1.6 Arial,sans-serif;color:#3F4042;">
<p style="margin:0 0 14px;">Hi ${esc(d.firstName)},</p>
<p style="margin:0 0 22px;">${updated ? 'Your changes are saved. This is now the version the Program Committee will review.' : `Thanks for proposing a demo station for the ${esc(form.summit)} Innovation Village. The Program Committee will review every submission after the call closes and email you with a decision.`}</p></td></tr>
<tr><td style="padding:0 32px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #DADCDF;border-top:3px solid #343895;font:14px/1.5 Arial,sans-serif;">
<tr><td style="padding:12px 16px 4px;${label}" colspan="2">Your submission</td></tr>
${row('Reference', esc(d._ref))}${row('Demo title', esc(d.title))}${row('Category', esc(d.category))}${row(updated ? 'Updated' : 'Submitted', when, true)}
</table></td></tr>
${steps ? `<tr><td style="padding:26px 32px 6px;${label}">What happens next</td></tr><tr><td style="padding:0 32px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font:14px/1.5 Arial,sans-serif;color:#3F4042;">${steps}</table></td></tr>` : ''}
<tr><td style="padding:24px 32px 28px;font:14px/1.6 Arial,sans-serif;color:#3F4042;">
<p style="margin:0 0 14px;">Need to change something? You can edit your submission${esc(until)}.</p>
<a href="${link}" style="display:inline-block;background:#8DC63F;color:#19194D;font:bold 13px Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;text-decoration:none;padding:12px 20px;border-radius:4px;">Edit your submission</a>
<p style="margin:14px 0 0;font-size:12px;color:#5F6062;">This link is private to your submission. Please don't forward it.</p></td></tr>
<tr><td style="background:#F3F4F5;padding:18px 32px;font:12px/1.6 Arial,sans-serif;color:#5F6062;border-top:1px solid #DADCDF;">Open Compute Project Foundation · <a href="${SITE}" style="color:#343895;">All Innovation Village calls for demos</a><br>You are receiving this because you submitted a proposal for the OCP Innovation Village.</td></tr>
</table></td></tr></table>`;

  const text = `Hi ${d.firstName},\n\n` + (updated ? 'Your changes are saved. This is now the version the Program Committee will review.\n\n' :
    `Thanks for proposing a demo station for the ${form.summit} Innovation Village. ` +
    `The Program Committee will review every submission after the call closes and email you with a decision.\n\n`) +
    `YOUR SUBMISSION\nReference: ${d._ref}\nDemo title: ${d.title}\nCategory: ${d.category}\n${updated ? 'Updated' : 'Submitted'}: ${when}\n\n` +
    ((form.timeline || []).length ? `WHAT HAPPENS NEXT\n${form.timeline.map(t => `${t[0]}: ${t[1]}`).join('\n')}\n\n` : '') +
    `Need to change something? Edit your submission${until}:\n${link}\n(This link is private to your submission. Please don't forward it.)\n\nOpen Compute Project Foundation\n${SITE}`;

  MailApp.sendEmail({
    to: d.email,
    subject: `${form.summit} Innovation Village – submission ${updated ? 'updated' : 'received'} (${d._ref})`,
    body: text,
    htmlBody: html,
    inlineImages: images,
    name: 'OCP Events',
  });
}

// The row after the last submission, found by the Reference column (B).
// Checkbox cells in the staff columns hold FALSE on every row, so getLastRow() can't be used.
function nextRow(sh) {
  const max = sh.getMaxRows();
  const refs = max > 1 ? sh.getRange(2, 2, max - 1, 1).getValues() : [];
  let last = 1;                                   // header row
  for (let i = refs.length - 1; i >= 0; i--) {
    if (refs[i][0] !== '') { last = i + 2; break; }
  }
  if (last + 1 > max) sh.insertRowsAfter(max, 100);
  return last + 1;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
