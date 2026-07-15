// Rwa7el Bankai — Service Worker v1

self.addEventListener('install', e => {
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(clients.claim());
});

// Handle push notifications
self.addEventListener('push', e => {
  if (!e.data) return;

  let data;
  try { data = e.data.json(); }
  catch { data = { title: 'Rwa7el', body: e.data.text() }; }

  const title = data.title || 'Rwa7el Bankai';
  const options = {
    body:    data.body    || '',
    icon:    data.icon    || '/static/icon.png',
    badge:   '/static/badge.png',
    tag:     data.tag     || 'rwa7el',
    renotify: true,
    requireInteraction: data.urgent || false,
    vibrate: data.vibrate || [200, 100, 200],
    data: { url: data.url || '/join' }
  };

  e.waitUntil(self.registration.showNotification(title, options));
});

// Click opens the join page
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const url = e.notification.data?.url || '/join';
  e.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
      for (const c of list) {
        if (c.url.includes('/join') && 'focus' in c) return c.focus();
      }
      if (clients.openWindow) return clients.openWindow(url);
    })
  );
});
