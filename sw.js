// RFAI Tracker service worker: makes the site installable as a phone app.
// It does not cache data — every request goes to the server, so numbers are always current.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
