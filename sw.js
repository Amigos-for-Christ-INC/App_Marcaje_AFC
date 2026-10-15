// Service worker: guarda la app en el teléfono para que abra sin internet
// y recibe las notificaciones de recordatorio (Firebase) aunque la app esté cerrada.
importScripts('config.js');
if (self.FIREBASE_CONFIG && self.FIREBASE_CONFIG.apiKey) {
  importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js',
                'https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js');
  firebase.initializeApp(self.FIREBASE_CONFIG);
  firebase.messaging();   // muestra las notificaciones que llegan con la app cerrada
}

// Al cambiar archivos de la app, sube el número de VERSION para que los teléfonos se actualicen.
const VERSION = 'asistencia-v5';
const BASE = ['./', './index.html', './config.js', './manifest.webmanifest',
  './icons/icon-192.png', './icons/icon-512.png', './icons/icon-maskable-512.png', './icons/apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(BASE)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

function conLimite(promesa, ms) {
  return Promise.race([promesa, new Promise((_, no) => setTimeout(() => no(new Error('lento')), ms))]);
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;            // las llamadas al servidor (POST) no se tocan
  const url = new URL(req.url);

  // Página y config: primero internet (para tener la última versión), si no hay, la copia guardada
  if (url.origin === location.origin && (req.mode === 'navigate' || url.pathname.endsWith('/config.js') || url.pathname.endsWith('/index.html'))) {
    e.respondWith(
      conLimite(fetch(req), 5000)
        .then(r => { const copia = r.clone(); caches.open(VERSION).then(c => c.put(req, copia)); return r; })
        .catch(() => caches.match(req).then(r => r || caches.match('./index.html')))
    );
    return;
  }

  // Íconos y demás archivos propios: copia guardada primero
  if (url.origin === location.origin) {
    e.respondWith(caches.match(req).then(r => r || fetch(req)));
    return;
  }

  // Tipografía Raleway de Google Fonts: copia guardada y se actualiza en segundo plano
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com' || url.hostname === 'alcdn.msauth.net') {
    e.respondWith(caches.open(VERSION).then(c => c.match(req).then(guardada => {
      const red = fetch(req).then(r => { c.put(req, r.clone()); return r; }).catch(() => guardada);
      return guardada || red;
    })));
  }
});
