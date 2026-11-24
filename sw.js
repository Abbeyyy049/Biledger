// Keeps the app shell available offline. Pages and scripts: network first, cache as fallback.
const CACHE = 'biledger-v1';
const SHELL = ['./', 'index.html', 'styles.css', 'js/app.js', 'js/backend.js', 'js/excel.js', 'manifest.webmanifest', 'icons/icon-192.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  const lib = u.host === 'www.gstatic.com' || u.host === 'cdnjs.cloudflare.com' || u.host.endsWith('fonts.googleapis.com') || u.host.endsWith('fonts.gstatic.com');
  if (u.origin !== location.origin && !lib) return;
  if (lib) { // versioned libraries: cache first
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(res => { const c = res.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return res; })));
    return;
  }
  e.respondWith(fetch(e.request).then(res => { const c = res.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return res; }).catch(() => caches.match(e.request).then(r => r || caches.match('index.html'))));
});
