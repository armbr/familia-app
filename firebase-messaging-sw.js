// ⚡ Fluxo App — Firebase Service Worker para Push Notifications
// Arquivo: firebase-messaging-sw.js (deve ficar na raiz do site)

importScripts('https://www.gstatic.com/firebasejs/10.12.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.12.0/firebase-messaging-compat.js');

// ▼▼▼ SUBSTITUIR COM SEU CONFIG (mesmo do index.html) ▼▼▼
const firebaseConfig = {
  apiKey:            "AIzaSyDWwLFR5Rd1R8YUx7hLk4ZwXR7Fkcct0sY",
  authDomain:        "fluxo-app-46562.firebaseapp.com",
  projectId:         "fluxo-app-46562",
  storageBucket:     "fluxo-app-46562.firebasestorage.app",
  messagingSenderId: "720780183452",
  appId:             "1:720780183452:web:3a471aae91f5b270472a59"
};
// ▲▲▲ SUBSTITUIR COM SEU CONFIG ▲▲▲

firebase.initializeApp(firebaseConfig);
const messaging = firebase.messaging();

// Receber push com app em background/fechado
messaging.onBackgroundMessage(payload => {
  console.log('[SW] Push recebido em background:', payload);

  const title = payload.notification?.title || '⚡ Fluxo App';
  const body  = payload.notification?.body  || '';
  const icon  = payload.notification?.icon  || '/icon.svg';

  self.registration.showNotification(title, {
    body:  body,
    icon:  icon,
    badge: '/icon.svg',
    data:  payload.data || {},
    actions: [
      { action: 'abrir', title: '📱 Abrir app' }
    ]
  });
});

// Clique na notificação — abrir o app
self.addEventListener('notificationclick', event => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
      for (const client of list) {
        if (client.url.includes(self.location.origin) && 'focus' in client) {
          return client.focus();
        }
      }
      return clients.openWindow('/');
    })
  );
});
