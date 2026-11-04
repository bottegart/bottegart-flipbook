/* Service worker del gestionale Bottegart: permette di aprirlo anche senza internet.
   Strategia "prima la rete": se c'è connessione prende sempre la versione aggiornata
   e ne tiene una copia; se non c'è, usa la copia salvata.
   I dati della libreria NON passano di qui: restano nel browser del tablet. */
const CACHE = 'gestionale-bottegart';
const FILE = ['./', './index.html', './manifest.webmanifest', './icona-192.png', './icona-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return; // ricerche online: sempre dalla rete
  e.respondWith(
    fetch(e.request)
      .then(r => { const copia = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copia)); return r; })
      .catch(() => caches.match(e.request, {ignoreSearch: true}).then(r => r || caches.match('./index.html')))
  );
});
