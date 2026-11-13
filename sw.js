// Offline support: caches the app so it opens with no signal.
// When you change any app file, bump CACHE so phones pick up the new version.
const CACHE = "vehicle-check-v1.0.3";
const FILES = [
  "./", "./index.html", "./style.css", "./app.js", "./data.js", "./config.js",
  "./manifest.webmanifest", "./icons/icon-192.png", "./icons/icon-512.png", "./icons/icon-180.png", "./icons/logo.png", "./icons/lbh-logo.png"
];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return; // never cache the database calls
  // Network first (so updates arrive), but fall back to the cache after 4 s
  // or when offline, so a weak signal never leaves the driver staring at a blank screen.
  const fromCache = () => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match("./index.html"));
  const network = fetch(req).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return res;
  });
  const timeout = new Promise(resolve => setTimeout(resolve, 4000)).then(fromCache);
  e.respondWith(
    Promise.race([network.catch(fromCache), timeout]).then(r => r || network.catch(fromCache))
  );
});
