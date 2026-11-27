// Keeps the app shell available offline. Pages and scripts: network first, cache as fallback.
const CACHE = 'biledger-v16';
const SHELL = ['./', 'index.html', 'styles.css', 'js/app.js', 'js/backend.js', 'js/excel.js', 'js/icons.js', 'manifest.webmanifest', 'icons/icon-192.png'];
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

// Reminder notifications (sent by scripts/notify.mjs through Firebase Cloud Messaging)
self.addEventListener('push', e => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (err) { d = { notification: { title: 'Biledger', body: e.data ? e.data.text() : '' } }; }
  const n = d.notification || {}, data = d.data || {};
  e.waitUntil(self.registration.showNotification(n.title || 'Biledger', {
    body: n.body || '', icon: 'icons/icon-192.png', badge: 'icons/icon-192.png',
    tag: data.lid || 'biledger', data: { link: data.link || (d.fcmOptions && d.fcmOptions.link) || './#home' },
  }));
});
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const url = (e.notification.data && e.notification.data.link) || './#home';
  e.waitUntil(clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
    for (const w of list) if ('focus' in w) { if (w.navigate) w.navigate(url); return w.focus(); }
    return clients.openWindow(url);
  }));
});
