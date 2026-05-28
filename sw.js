const CACHE_NAME = 'card-manager-v3';
const urlsToCache = ['./index.html'];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(urlsToCache))
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request).then(response => response || fetch(event.request))
  );
});

// Push notification received
self.addEventListener('push', event => {
  const data = event.data ? event.data.json() : {};
  const title = data.title || 'Card Manager';
  const options = {
    body: data.body || '期限が近い特典があります',
    icon: '/icon-192.png',
    badge: '/icon-192.png',
    tag: data.tag || 'card-reminder',
    requireInteraction: true,
    data: data
  };
  event.waitUntil(self.registration.showNotification(title, options));
});

// Notification click
self.addEventListener('notificationclick', event => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window' }).then(clientList => {
      for (const client of clientList) {
        if (client.url && 'focus' in client) return client.focus();
      }
      if (clients.openWindow) return clients.openWindow('./index.html');
    })
  );
});

// Background sync for daily notification check
self.addEventListener('periodicsync', event => {
  if (event.tag === 'check-benefits') {
    event.waitUntil(checkAndNotify());
  }
});

async function checkAndNotify() {
  // This will be triggered by the app itself via messaging
  const allClients = await clients.matchAll();
  allClients.forEach(client => {
    client.postMessage({ type: 'CHECK_BENEFITS' });
  });
}
