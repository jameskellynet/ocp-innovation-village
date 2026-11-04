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

const LIMITS = { abstract: 1000, whyDemo: 500, takeaway: 500 };

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
      sh.appendRow(COLUMNS.map(c => c[1](d)));
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

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
