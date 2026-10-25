// Cache-first : l'app fonctionne hors ligne. Les appels Microsoft / OneDrive ne sont jamais mis en cache.
// Incrémenter CACHE à chaque mise à jour publiée.
const CACHE = 'heures-v22';
const LOCAL = [
  './', 'index.html', 'styles.css', 'db.js', 'csv.js', 'xlsx.js', 'onedrive.js', 'billit.js', 'pdf-fournisseurs.js', 'achats.js', 'app.js',
  'manifest.webmanifest', 'icons/icon.svg', 'icons/icon-192.png', 'icons/icon-512.png'
];
const EXCELJS = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    await c.addAll(LOCAL);
    try { await c.add(new Request(EXCELJS, { mode: 'cors' })); } catch (_) { /* récupéré à la première utilisation */ }
    self.skipWaiting();
  })());
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
    self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  const url = e.request.url;
  if (e.request.method !== 'GET' || !(url.startsWith(self.location.origin) || url === EXCELJS)) return;
  e.respondWith((async () => {
    const hit = await caches.match(e.request, { ignoreSearch: true });
    if (hit) return hit;
    const res = await fetch(e.request);
    if (res.ok && url === EXCELJS) (await caches.open(CACHE)).put(e.request, res.clone());
    return res;
  })());
});
