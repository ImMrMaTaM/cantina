// Service worker minimo: rete prima, copia locale solo per i file dell'app quando si è offline.
const C = 'cantina-v1';
self.addEventListener('install', (e) => { self.skipWaiting(); });
self.addEventListener('activate', (e) => { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== self.location.origin) return;
  e.respondWith(fetch(e.request).then((r) => { const c = r.clone(); caches.open(C).then((ca) => ca.put(e.request, c)); return r; }).catch(() => caches.match(e.request)));
});
