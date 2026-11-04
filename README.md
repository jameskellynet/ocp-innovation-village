# OCP Innovation Village – Call for Demos

Branded submission forms for the Open Compute Project Innovation Village, hosted on GitHub Pages. Each Summit gets its own form page and its own Google Sheet; one shared engine and one Apps Script serve them all.

## How it works

```
Submitter's browser → form page (GitHub Pages) → Apps Script web app → that Summit's Google Sheet
```

The page validates answers, shows a review step, then posts the submission as JSON to the Apps Script URL. The script checks it again, writes a row to the right Sheet and emails the submitter a confirmation. Claude skills for scoring and notifications read the Sheet as before.

## Layout

```
index.html                 Landing page listing open calls
assets/form.css            Shared styles (OCP brand palette, Franklin Gothic)
assets/form.js             Shared form engine
assets/ocp-logo.png        Official OCP logo
forms/<form-id>/index.html Page shell for one form
forms/<form-id>/config.js  Everything specific to that form
apps-script/Code.gs        Backend receiver for all forms
```

## Add a new Summit form

1. Copy `forms/emea-2027-iv/` to a new folder, e.g. `forms/global-2027-iv/`.
2. Edit `config.js`: `formId`, Summit name, dates, deadline, categories, timeline, eligibility.
3. Update the `<title>` in that folder's `index.html`.
4. Create the Summit's Google Sheet (blank) and add a line to `FORMS` in `apps-script/Code.gs` with its spreadsheet ID and closing time. In the Apps Script editor, run `setupSheets` once to add the headers and formatting, then redeploy the script as a new version (the URL stays the same).
5. Set `endpoint` in `config.js` to the script URL and add the form to `index.html`.

## Deploy the backend (once)

1. Open [script.google.com](https://script.google.com) from a shared OCP events account and create a project.
2. Paste `apps-script/Code.gs`, fill in the Sheet IDs, and run `setupSheets` once.
3. **Deploy → New deployment → Web app**. Execute as: *Me*. Who has access: *Anyone*.
4. Copy the `/exec` URL into each form's `config.js` as `endpoint`.

While `endpoint` is empty, a form runs in preview mode and sends nothing.

## Sheet layout

`setupSheets` creates a **Submissions** tab with one column per form answer, followed by staff columns (Review Status dropdown, Notes, Acceptance Notice Sent and Confirmed checkboxes). The form only ever writes the answer columns, so staff edits are never overwritten.

## Turn on GitHub Pages

Settings → Pages → Build and deployment → Deploy from a branch → `main` / root.

## Caching

GitHub Pages lets browsers cache files for 10 minutes. Each page loads its CSS and scripts with a `?v=` stamp; when you change `form.js`, `form.css` or a `config.js`, bump the stamp in the pages that load them so visitors get a matching set straight away.

## Brand notes

Colors follow *OCP Brand Guidelines v3-1*, p.24: primary green `#8DC63F`, gray `#5F6062`, white, and the secondary indigo trio `#343895`, `#262673`, `#19194D`. Type is Franklin Gothic where installed, falling back to Libre Franklin (open-source revival). To serve licensed Franklin Gothic web fonts, add them under `assets/fonts/` and an `@font-face` rule in `form.css`.
