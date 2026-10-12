# IMPS Web v0.1.8 — Adaptive Phone Formatting

- Restores 10 pt font for shorter phone numbers.
- Uses 9 pt for standard (###) ###-#### numbers; only longer entries reduce to 8 or 7 pt.
- Preserves nonbreaking digits, hospital cell no-wrap, layout, address autocomplete cleanup and existing incidents.
- Uses versioned application JavaScript, worker and generator files to prevent stale code.
- Web only; no changes to Windows or Mac applications.

## GitHub deployment
Upload **the full contents** of this extracted ZIP, preserving assets/ and generator/ directories. Verify on GitHub that `assets/index-v018.js`, `assets/generator-worker-v018.js` and `generator/generate_ics206_v018.py` are all present before testing. Commit and verify successful Pages deployment. Do not clear Local Storage or IndexedDB.

## QA
Short phone numbers should remain 10 pt; standard full numbers 9 pt; long values reduce as needed. Validate Word output visually (Word may interpret no-wrap differently), especially numbers with prefixes and multiple phones.
