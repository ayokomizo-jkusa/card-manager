importScripts('https://www.gstatic.com/firebasejs/10.12.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.12.0/firebase-messaging-compat.js');

const CACHE_NAME = 'card-manager-v4';
const urlsToCache = ['./index.html'];

// Firebase初期化
firebase.initializeApp({
  apiKey: "AIzaSyDj84BQGXGOC3327yDcThvIRKI2z94o-Vo",
  authDomain: "card-manager-b9a2f.firebaseapp.com",
  projectId: "card-manager-b9a2f",
  storageBucket: "card-manager-b9a2f.firebasestorage.app",
  messagingSenderId: "334509208344",
  appId: "1:334509208344:web:80f3c6343325c6b283246d"
});

const messaging = firebase.messaging();

// バックグラウンドでプッシュを受信したとき
messaging.onBackgroundMessage((payload) => {
  const title = payload.notification?.title || 'CardMgr';
  const body  = payload.notification?.body  || '期限が近い特典があります';
  self.registration.showNotification(title, {
    body,
    icon: './icon-192.png',
    badge: './icon-192.png',
    requireInteraction: true,
    tag: 'cardmgr-reminder'
  });
});

// キャッシュインストール
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(urlsToCache))
  );
  self.skipWaiting();
});

// 古いキャッシュを削除
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// フェッチ（キャッシュ優先）
self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request).then(res => res || fetch(event.request))
  );
});

// 通知タップでアプリを開く
self.addEventListener('notificationclick', event => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window' }).then(list => {
      for (const client of list) {
        if ('focus' in client) return client.focus();
      }
      if (clients.openWindow) return clients.openWindow('./index.html');
    })
  );
});
