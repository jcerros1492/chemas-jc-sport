const CACHE = 'chemas-jc-sport-v9-20';
const ASSETS = ['./', './index.html', './styles.css', './app.js', './manifest.json', './logo.png', './apple-touch-icon-v2.png', './icon-192-v2.png', './icon-512-v2.png', './icon-maskable-512-v2.png', './html2canvas.min.js', './jspdf.umd.min.js'];
self.addEventListener('install', event => { self.skipWaiting(); event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(ASSETS)).catch(() => {})); });
self.addEventListener('activate', event => event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key !== CACHE).map(key => caches.delete(key)))).then(() => self.clients.claim())));
// Red primero (siempre la versión más nueva); si no hay internet, se usa lo guardado.
self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET' || !req.url.startsWith(self.location.origin)) return;
  event.respondWith(fetch(req.url,{cache:'no-cache'}).then(res => {
    if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return res;
  }).catch(() => caches.match(req).then(hit => hit || caches.match('./index.html'))));
});
