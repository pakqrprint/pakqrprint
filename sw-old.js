/* Pak QR Print — Admin Dashboard PWA service worker v1.
   Sirf app shell cache hota hai (dashboard ke apne files).
   Firebase / Supabase / Google API calls KABHI cache se nahi milti —
   wo hamesha network se jati hain. */
var CACHE = 'pakqrprint-dash-v14'; // v6: naya shell (duplex wala dashboard.js) dobara cache ho
var SHELL = [
  'dashboard.html',
  'styles.css',
  'dashboard.js',
  'qr.js',
  'firebase-config.js',
  'supabase-config.js',
  'manifest.json',
  'icon-192.png',
  'icon-512.png',
  'icon-512-maskable.png',
  'Pak-QR-Print-Logo.png',
  'favicon.ico'
];

function baseName(path) {
  var p = String(path || '').split('?')[0].split('#')[0];
  return p.substring(p.lastIndexOf('/') + 1);
}

/* PWA v2: FCM background push — agent naya order bhejte hi notification dikhayein.
   firebase-config.js page script hai (window.* assignments) — SW mein window nahi
   hota, is liye self ko window bana kar import karein; config deploy wali live
   file se aati hai (kaam ki file kabhi edit nahi hoti). Template config ho to
   FCM band rehta hai — app shell cache phir bhi chalta hai (neeche try/catch).
   Sender (agent) data message bhejta hai: {token, orderKey, name, files}.
   Notification par tap karein to dashboard khulta hai / focus hota hai. */
var fcm = null;
try {
  importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js');
  importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js');
  var window = self; // sirf firebase-config.js ke window.* assignments ke liye
  importScripts('firebase-config.js');
  if (self.PAKQR_FIREBASE_CONFIG && self.PAKQR_FIREBASE_CONFIG.apiKey &&
      String(self.PAKQR_FIREBASE_CONFIG.apiKey).indexOf('YAHAN-') !== 0) {
    firebase.initializeApp(self.PAKQR_FIREBASE_CONFIG);
    fcm = firebase.messaging();
  }
} catch (e) { fcm = null; }

if (fcm) {
  fcm.onBackgroundMessage(function (payload) {
    var d = (payload && payload.data) || {};
    var n = (payload && payload.notification) || {};
    var token = d.token || d.orderToken || '';
    var name = d.name || d.customerName || 'Naam nahi likha';
    var files = d.files || d.fileCount || '1';
    var title = n.title || ('Naya print order' + (token ? ' — ' + token : ''));
    var body = n.body || (name + ' — ' + files + ' file(s)');
    var orderKey = d.orderKey || '';
    return self.registration.showNotification(title, {
      body: body,
      icon: 'icon-192.png',
      badge: 'icon-192.png',
      tag: 'pakqr-order-' + (orderKey || token || Date.now()),
      data: { token: token, orderKey: orderKey, url: 'dashboard.html' }
    });
  });
}

self.addEventListener('notificationclick', function (e) {
  e.notification.close();
  var url = (e.notification.data && e.notification.data.url) || 'dashboard.html';
  e.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (list) {
      for (var i = 0; i < list.length; i++) {
        if (list[i].url.indexOf('dashboard.html') !== -1 && 'focus' in list[i]) return list[i].focus();
      }
      if (clients.openWindow) return clients.openWindow(url);
    })
  );
});

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (c) { return c.addAll(SHELL); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE; })
        .map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  if (e.request.method !== 'GET') return;
  var url = new URL(e.request.url);
  // Doosri origin (Firebase / Supabase / Google) — kabhi intercept nahi
  if (url.origin !== self.location.origin) return;
  // Sirf shell files cache-first; baqi sab network (intercept nahi)
  if (SHELL.indexOf(baseName(url.pathname)) !== -1) {
    e.respondWith(
      caches.match(e.request).then(function (r) { return r || fetch(e.request); })
    );
  }
});
