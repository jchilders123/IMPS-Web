# IMPS Web v0.1.7 — Reliable Version Updates

- Navigations now use network-first, with cached offline fallback, so reopening the site prefers the newest published app.
- Service worker updates bypass the browser HTTP cache; on controller changes the page refreshes once to pick up the latest files.
- Retains v0.1.6 Word generator and hospital phone formatting, along with existing Google address cleanup.

## Deployment
Upload the complete folder contents to the same GitHub Pages publishing location and commit. Old users may need to clear site data once if an older service worker blocks the transition.
