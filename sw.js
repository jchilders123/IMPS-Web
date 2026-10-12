const CACHE_NAME = 'imps-web-v0.1.9';
const CACHE_PREFIX = 'imps-web-v';
const APP_SHELL = [
  './', './index.html', './manifest.webmanifest',
  './assets/index-v019.js', './assets/generator-worker-v019.js',
  './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k.startsWith(CACHE_PREFIX) && k !== CACHE_NAME).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET') return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  // A new visit should load the latest published application, not an old app shell.
  // Offline support remains available through the cached shell.
  if (request.mode === 'navigate') {
    event.respondWith((async () => {
      try {
        const fresh = await fetch(request, {cache: 'no-store'});
        if (!fresh.ok) throw new Error('Navigation HTTP ' + fresh.status);
        const cache = await caches.open(CACHE_NAME);
        await cache.put('./index.html', fresh.clone());
        return fresh;
      } catch (err) {
        const cache = await caches.open(CACHE_NAME);
        return (await cache.match('./index.html')) || (await cache.match('./')) || Response.error();
      }
    })());
    return;
  }

  // Never cache the Python document generator or service-worker script.
  const generatorPath = new URL('./generator/', self.location.href).pathname;
  if (url.pathname.startsWith(generatorPath) || url.pathname === new URL('./sw.js', self.location.href).pathname) {
    event.respondWith(fetch(request, {cache:'no-store'}));
    return;
  }

  event.respondWith((async () => {
    const cached = await caches.match(request);
    if (cached) return cached;
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      await cache.put(request, response.clone());
    }
    return response;
  })());
});
