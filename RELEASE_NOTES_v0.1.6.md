# IMPS Web v0.1.6 — Document Engine Refresh and Hospital Phone Fix

- Version-specific Python generator filename and browser-worker filename to prevent old cached code from running.
- Service worker does not cache files under `/generator/`, including versioned URLs.
- Final-pass hospital phone formatting on the finished DOCX table: Arial 6 pt, no-wrap and character fitting.
- Retains v0.1.5 Google address autocomplete cleanup, mobile support and printing notice.

## Deploy
Upload all extracted files and folders to the same GitHub Pages publishing branch and root as before, commit, then close and reopen the site. Verify the heading shows v0.1.6 and generate a new Word document.
