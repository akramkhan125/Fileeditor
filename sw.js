/* AK TOOLS service worker — cache-first so the app works fully offline
   after the first visit. Bump CACHE when libraries change. */
const CACHE = "aktools-v3";
const CORE = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icon.svg",
  "./icon-192.png",
  "./icon-512.png",
  "./apple-touch-icon.png",
  "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=Poppins:wght@600;700;800&display=swap",
  "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js",
  "https://unpkg.com/docx@8.5.0/build/index.umd.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js",
  "https://cdn.jsdelivr.net/npm/browser-image-compression@2.0.2/dist/browser-image-compression.js",
  "https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/Sortable/1.15.2/Sortable.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js",
  "https://cdn.jsdelivr.net/npm/tesseract.js@5.0.5/dist/tesseract.min.js",
  "https://cdn.jsdelivr.net/npm/tesseract.js@5.0.5/dist/worker.min.js"
];

self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => Promise.allSettled(CORE.map((u) => c.add(u))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const _u = new URL(e.request.url);
  if (e.request.method === "POST" && _u.pathname === "/share") {
    e.respondWith((async () => {
      try {
        const fd = await e.request.formData();
        const file = fd.get("shared_file");
        if (file && file.size) {
          const c = await caches.open("fb-shared");
          await c.put("/shared-file", new Response(file, { headers: { "Content-Type": file.type || "application/octet-stream" } }));
          await c.put("/shared-meta", new Response(JSON.stringify({ name: file.name || "shared-file", type: file.type || "" }), { headers: { "Content-Type": "application/json" } }));
        }
      } catch (err) {}
      return Response.redirect("/?shared=1", 303);
    })());
    return;
  }
  if (e.request.method !== "GET") return;
  // Page loads: network-first so new versions arrive, cache keeps offline working
  if (e.request.mode === "navigate") {
    e.respondWith(
      fetch(e.request)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put("./index.html", copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match("./index.html", { ignoreSearch: true }))
    );
    return;
  }
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then((hit) => {
      if (hit) return hit;
      return fetch(e.request).then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
        return res;
      });
    })
  );
});
