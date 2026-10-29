// Cache-first : l'app fonctionne hors ligne. Les appels Microsoft / OneDrive ne sont jamais mis en cache.
// Incrémenter CACHE à chaque mise à jour publiée (même nom = le navigateur garde l'ancienne copie).
const CACHE = 'heures-v39';
const LOCAL = [
  './', 'index.html', 'styles.css', 'db.js', 'csv.js', 'xlsx.js', 'onedrive.js', 'billit.js', 'copie-billit.js', 'pdf-fournisseurs.js', 'achats.js', 'app.js',
  'manifest.webmanifest', 'icons/icon.svg', 'icons/icon-192.png', 'icons/icon-512.png',
  'lib/pdf.min.mjs', 'lib/pdf.worker.min.mjs', 'lib/exceljs.min.js'
];

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    // cache: 'reload' : fichiers relus sur le site, pas dans le cache HTTP du navigateur (GitHub Pages le garde 10 min).
    await c.addAll(LOCAL.map(u => new Request(u, { cache: 'reload' })));
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
  if (e.request.method !== 'GET' || !url.startsWith(self.location.origin)) return;
  e.respondWith((async () => (await caches.match(e.request, { ignoreSearch: true })) || fetch(e.request))());
});
