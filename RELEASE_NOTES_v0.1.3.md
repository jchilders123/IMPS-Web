# IMPS Web v0.1.3 — Address and Phone Formatting

Built from the existing IMPS Web v0.1.2 Printing Notice distribution.

- Google-selected addresses drop a trailing USA, United States, or United States of America. Manually entered addresses are unchanged.
- ICS-206 Word hospital phone numbers use Arial 8 pt and nonbreaking spaces/hyphens, preventing normal telephone digits from splitting across lines. Phone-cell styling is also applied to matching ground and air ambulance fields.
- Existing mobile, PDF, Word, printing notice, and other logic retained.
- Service-worker cache version bumped to v0.1.3 to refresh clients.

## Publish
Upload the **contents** of this ZIP (including index.html, assets, generator, templates and sw.js) to the same hosting root used for v0.1.2. Do not upload the ZIP as a page. Existing browsers may need one reload after service-worker activation.

## Validation
Python syntax and archive integrity can be checked automatically, but Word rendering and end-to-end Google autocomplete should be tested after deployment.
