/* A small app-shell cache; the Google Apps Script endpoint is never cached. */
const CACHE_NAME = "dog-insulin-tracker-v5";
const APP_FILES = ["./", "./index.html", "./manifest.json", "./icon.svg"];

const firebaseConfigText = new URL(self.location.href).searchParams.get("config");
if (firebaseConfigText) {
  importScripts("https://www.gstatic.com/firebasejs/11.7.0/firebase-app-compat.js");
  importScripts("https://www.gstatic.com/firebasejs/11.7.0/firebase-messaging-compat.js");

  firebase.initializeApp(JSON.parse(firebaseConfigText));
  firebase.messaging().onBackgroundMessage((payload) => {
    const data = payload.data || {};
    self.registration.showNotification(data.title || "Инсулин-пёс", {
      body: data.body || "Пора проверить дневник инъекций.",
      tag: data.tag || "dog-insulin-reminder",
      icon: "./icon.svg",
      data: { url: data.url || "./" }
    });
  });

  self.addEventListener("notificationclick", (event) => {
    event.notification.close();
    const target = new URL(event.notification.data.url || "./", self.location.origin);
    if (target.origin !== self.location.origin) return;
    event.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((windows) => {
      for (const client of windows) {
        if (client.url.startsWith(self.location.origin) && "focus" in client) {
          return client.focus();
        }
      }
      return self.clients.openWindow(target.href);
    }));
  });
}

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_FILES)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((names) => Promise.all(
      names.filter((name) => name !== CACHE_NAME).map((name) => caches.delete(name))
    ))
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== "GET" || url.origin !== self.location.origin) return;
  event.respondWith(
    fetch(request).then((response) => {
      if (response.ok) {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
      }
      return response;
    }).catch(() => caches.match(request).then((cached) => {
      if (cached) return cached;
      if (request.mode === "navigate") return caches.match("./index.html");
      return Response.error();
    }))
  );
});
