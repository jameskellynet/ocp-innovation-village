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
  'emea-2027-iv':   { sheetId: '1V1YS5CcBHaC5LiAbBfJOgsRcLpqAQRy97fISHUt4oqo',   tab: 'Submissions', summit: 'OCP EMEA Summit 2027',   closes: '2027-01-12T07:59:00Z' },
  'global-2027-iv': { sheetId: 'PASTE_GLOBAL_2027_SHEET_ID', tab: 'Submissions', summit: 'OCP Global Summit 2027', closes: '2027-06-01T06:59:00Z' },
};

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

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);

    if (d.website) return reply({ ok: true, ref: 'IV-OK' });          // honeypot: silently drop bots
    const form = FORMS[d.formId];
    if (!form) return reply({ ok: false, error: 'This form is not recognised.' });
    if (form.closes && new Date() > new Date(form.closes))
      return reply({ ok: false, error: 'The submission deadline has passed.' });

    // Server-side checks (never trust the browser alone)
    for (const k of ['firstName','lastName','email','company','title','category','abstract','whyDemo','takeaway','equipment'])
      if (!String(d[k] || '').trim()) return reply({ ok: false, error: 'A required answer is missing.' });
    for (const [k, max] of Object.entries(LIMITS))
      if (String(d[k]).length > max) return reply({ ok: false, error: 'An answer is over its character limit.' });
    if (!d.consent) return reply({ ok: false, error: 'Consent is required.' });

    d._ref = 'IV-' + Utilities.getUuid().slice(0, 6).toUpperCase();

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = SpreadsheetApp.openById(form.sheetId);
      const sh = ss.getSheetByName(form.tab) || ss.insertSheet(form.tab);
      if (sh.getLastRow() === 0) sh.appendRow(COLUMNS.map(c => c[0]));
      const values = COLUMNS.map(c => c[1](d)).concat(['New']);   // Review Status starts as New
      sh.getRange(nextRow(sh), 1, 1, values.length).setValues([values]);
    } finally {
      lock.releaseLock();
    }

    MailApp.sendEmail({
      to: d.email,
      subject: `${form.summit} Innovation Village – submission received (${d._ref})`,
      body: `Hi ${d.firstName},\n\nThanks for proposing "${d.title}" for the ${form.summit} Innovation Village.\n` +
            `Your reference is ${d._ref}. The OCP Program Committee will review all submissions and notify you by email.\n\n` +
            `OCP Events Team`,
      name: 'OCP Events',
    });

    return reply({ ok: true, ref: d._ref });
  } catch (err) {
    return reply({ ok: false, error: 'Something went wrong on our side.' });
  }
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
