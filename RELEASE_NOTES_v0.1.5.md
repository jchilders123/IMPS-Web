# IMPS Web v0.1.5 — Hospital Phone and Generator Cache Fix

- Strengthens Word hospital phone fitting: Arial 6pt, condensed characters, nonbreaking spacing and hyphens, Word run-level fitText, and cell-level noWrap.
- Bypasses older cached document generators with a versioned URL and network-first/no-store worker fetch.
- Retains Google autocomplete USA removal, mobile layout, and Word/PDF workflows.
- Updated website version and service-worker cache to v0.1.5.

**Validation needed:** Test a newly generated DOCX from a previously saved incident on Word and in preview after deploying. If phone numbers still wrap, capture the new DOCX for direct inspection; live Word layout could not be verified here.
