// ===============================
// BANKAI / RWA7EL Service Worker
// v4 — Auto Build ID Cache Bust
// ===============================

const CACHE_BASE = "rwa7el";
let CACHE_NAME = CACHE_BASE + "-local"; // fallback لو fetch فشل

const APP_SHELL = [
  "/",
  "/join",
  "/manifest.json",
  "/static/icon-192.png",
  "/static/icon-512.png"
];

// --------------------
// Install
// --------------------
self.addEventListener("install", (event) => {
  self.skipWaiting(); // فعّل الـ SW الجديد فوراً من غير ما ينتظر
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(APP_SHELL).catch(() => {});
    })
  );
});

// --------------------
// Activate
// --------------------
self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {

      // جيب الـ BUILD_ID من السيرفر
      try {
        const res = await fetch('/api/build_id', { cache: 'no-store' });
        const data = await res.json();
        CACHE_NAME = CACHE_BASE + "-" + data.id;
      } catch(_) {
        CACHE_NAME = CACHE_BASE + "-" + Date.now();
      }

      // امسح كل كاش قديم
      const keys = await caches.keys();
      await Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) return caches.delete(key);
        })
      );

      await self.clients.claim();

      // بلّغ كل الصفحات المفتوحة عشان تعمل reload
      const allClients = await self.clients.matchAll({
        type: "window",
        includeUncontrolled: true,
      });
      allClients.forEach((client) => {
        client.postMessage({ type: "NEW_VERSION_AVAILABLE" });
      });

    })()
  );
});

// --------------------
// Fetch — Network First (كل حاجة live)
// --------------------
self.addEventListener("fetch", (event) => {

  if (event.request.method !== "GET") return;

  const url = new URL(event.request.url);

  // مسارات السوكيت والـ API والصفحات الرئيسية — مش بنلمسها (عشان المسافة تبقى live دايمًا)
  if (
    url.pathname.startsWith("/api") ||
    url.pathname.startsWith("/socket.io") ||
    url.pathname.startsWith("/register") ||
    url.pathname === "/join" ||
    url.pathname === "/"
  ) {
    return;
  }

  // كل حاجة تانية: network-first — لو الشبكة فشلت بس يرجع للكاش
  event.respondWith(
    fetch(event.request, { cache: "no-store" })
      .then((response) => {
        // حدّث الكاش بالنسخة الجديدة دايمًا
        if (response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => {
        // fallback للكاش بس لو انقطع النت
        return caches.match(event.request);
      })
  );

});

// --------------------
// Push Notifications
// --------------------
self.addEventListener("push", (event) => {

  let data = {};
  try { data = event.data.json(); } catch (_) {}

  const type = data.type || "default";

  const configs = {
    your_turn: {
      title:   "🚨 دورك دلوقتي!",
      body:    data.body || "الأدمن طالبك — استعد فوراً!",
      vibrate: [400,100,400,100,400,100,400,100,400,100,400,100,400,100,400,100,400,100,400,100],
      requireInteraction: true,
      silent:  false,
      actions: [
        { action: "open",    title: "✅ أنا جاهز" },
        { action: "dismiss", title: "شفت" }
      ]
    },
    almost: {
      title:   data.title || "⚡ قريب — استعد!",
      body:    data.body  || "دورك قرب — جهّز نفسك",
      vibrate: [300,80,300,80,300,80,300,80,300,80,300,80,300,80,400,80,400],
      requireInteraction: false,
      silent:  false,
      actions: [
        { action: "open",    title: "👀 شوف دوري" },
        { action: "dismiss", title: "حسناً" }
      ]
    },
    position: {
      title:   data.title || "📋 تحديث الترتيب",
      body:    data.body  || "",
      vibrate: [200, 100, 200],
      requireInteraction: false,
      silent:  false,
      actions: [
        { action: "open", title: "👀 شوف دوري" }
      ]
    },
    default: {
      title:   data.title || "🔔 رواحل",
      body:    data.body  || "إشعار جديد",
      vibrate: [300,80,300,80,300,80,300,80,300,80,300,80,300,80,400,80,400],
      requireInteraction: true,
      silent:  false,
      actions: [
        { action: "open",    title: "افتح" },
        { action: "dismiss", title: "حسناً" }
      ]
    }
  };

  const cfg = configs[type] || configs.default;
  // your_turn دايمًا unique tag عشان كل إشعار يعمل vibration جديدة من الـ OS
  // almost/position بيستخدموا fixed tag عشان ميتراكموش
  const uniqueTag = (type === 'your_turn')
    ? `rwa7el-your_turn-${Date.now()}`
    : `rwa7el-${type}`;

  event.waitUntil(
    (async () => {
      // أبلّغ الـ foreground دايمًا (your_turn + almost + كل حاجة) عشان يشغّل الصوت
      const allClients = await self.clients.matchAll({
        type: "window",
        includeUncontrolled: true
      });
      allClients.forEach((client) => {
        client.postMessage({ type: "PUSH_RECEIVED", notifType: type });
      });

      return self.registration.showNotification(cfg.title, {
        body:               cfg.body,
        icon:               "/static/icon-192.png",
        badge:              "/static/icon-192.png",
        vibrate:            cfg.vibrate,
        requireInteraction: cfg.requireInteraction,
        silent:             cfg.silent,
        tag:                uniqueTag,
        renotify:           true,
        actions:            cfg.actions,
        data: { url: data.url || "/join", type: type }
      });
    })()
  );

});

// --------------------
// Notification Click
// --------------------
self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const targetUrl = event.notification.data?.url || "/join";
  event.waitUntil(
    clients.matchAll({ type: "window", includeUncontrolled: true }).then((clientList) => {
      for (const client of clientList) {
        if (client.url.includes("/join")) return client.focus();
      }
      return clients.openWindow(targetUrl);
    })
  );
});