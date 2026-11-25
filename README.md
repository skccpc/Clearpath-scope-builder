# ClearPath Scope Builder

Web-based tool for assembling Scope of Work documents for ClearPath Construction projects.
Outputs print-ready PDF and editable .docx.

**Live site:** https://skccpc.github.io/Clearpath-scope-builder/

---

## What's in here

| File | Purpose |
|---|---|
| `index.html` | Entry point — load this in a browser or via GitHub Pages |
| `ScopeLibrary.jsx` | Master library of scope sections and line items |
| `ScopeComponents.jsx` | React UI components for the builder interface |
| `ScopeDocument.jsx` | Renders the print/PDF document layout |
| `ScopeDocx.js` | Generates the editable Microsoft Word (.docx) export |
| `*.compiled.js` | Pre-transpiled versions of the .jsx files (not currently used; HTML loads .jsx through Babel at runtime) |
| `assets/` | ClearPath logo files |

## Review highlights (for edits made by Claude)

Each project has a **Track changes** switch in the top bar. While it's on, any
edit to a scope line or exclusion made in Build (by Claude in the browser or by
hand) is marked for review instead of applied silently. It's on for every new
project and, of the projects that existed before, only Obarski.

The marks live on the line itself:

| Field on the line | Meaning | Preview |
|---|---|---|
| `review: 'new'` | line was added or checked | green (Excel `#C6EFCE`) |
| `review: 'changed'` + `reviewFrom: '<old text>'` | wording changed; `text` holds the new wording | yellow (`#FFEB9C`), with "Was: …" |
| `review: 'deleted'` | line was unchecked or deleted; it stays in place until accepted | red (`#FFC7CE`), struck through |

Highlighted text can be edited in place. **Accept** / **Undo** on each line (or
Accept all / Undo all in the bar above the document) clear the mark. Highlights
appear only in the on-screen preview — PDF export, printing and the .docx are
always clean, and a line marked for deletion still prints until it's accepted.

## Running locally

No build step required. Either:

- Open `index.html` directly in a browser, **or**
- From this folder, run `python3 -m http.server 8000` and visit `http://localhost:8000`

The local server route is more reliable because some browsers restrict `file://` script loading.

## Editing the scope library

The single most common edit is adding/removing sections or line items.
Open `ScopeLibrary.jsx` — everything is in one array at the top of the file. Save, refresh the browser, done.

## Deploying updates

This repo is served by GitHub Pages from the `main` branch root.
Any commit pushed to `main` is live within ~30–60 seconds.

```bash
# typical update flow
git pull
# ...edit files...
git add .
git commit -m "Add new line items to demolition section"
git push
```

If updates aren't appearing, hard-refresh the browser (Cmd+Shift+R / Ctrl+Shift+R) — GitHub Pages caches aggressively.

## Versioning

Tag stable releases so you can roll back if something breaks mid-meeting:

```bash
git tag v1.0.0
git push --tags
```

To roll back: `git revert <bad-commit>` and push, or check out a known-good tag and force-push.
