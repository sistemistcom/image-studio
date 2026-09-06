from pathlib import Path

from starlette.responses import FileResponse, HTMLResponse, Response
from starlette.routing import Route
from streamlit.web.server.starlette import App


BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"


async def manifest(_request):
    return FileResponse(
        STATIC_DIR / "manifest.json",
        media_type="application/manifest+json",
        headers={"Cache-Control": "public, max-age=300"},
    )


async def service_worker(_request):
    script = r"""
const CACHE_NAME = "sistemist-image-studio-v860";
const OFFLINE_URL = "/offline.html";

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll([
        OFFLINE_URL,
        "/app/static/icon-192.png",
        "/app/static/icon-512.png"
      ]))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  if (event.request.mode === "navigate") {
    event.respondWith(fetch(event.request).catch(() => caches.match(OFFLINE_URL)));
  }
});
""".strip()
    return Response(
        script,
        media_type="application/javascript",
        headers={
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Service-Worker-Allowed": "/",
        },
    )


async def offline_page(_request):
    return HTMLResponse(
        """<!doctype html>
<html lang="tr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <meta name="theme-color" content="#0b1119">
  <title>Image Studio — Çevrimdışı</title>
  <style>
    body{margin:0;background:#0b1119;color:#f4f7fb;font-family:Arial,sans-serif;display:grid;place-items:center;min-height:100vh;padding:24px;box-sizing:border-box}
    main{max-width:520px;background:#151f2b;border:1px solid #2a394b;border-radius:20px;padding:30px;text-align:center}
    h1{font-size:25px;margin:0 0 12px}p{color:#9aaabd;line-height:1.6;margin:0 0 20px}button{border:0;border-radius:11px;background:#ff6a00;color:#fff;padding:13px 20px;font-weight:700}
  </style>
</head>
<body><main><h1>İnternet bağlantısı bulunamadı</h1><p>Image Studio görselleri işlemek için internet bağlantısına ihtiyaç duyar.</p><button onclick="location.reload()">Tekrar Dene</button></main></body>
</html>""",
        headers={"Cache-Control": "public, max-age=3600"},
    )


app = App(
    "app.py",
    routes=[
        Route("/manifest.webmanifest", manifest),
        Route("/service-worker.js", service_worker),
        Route("/offline.html", offline_page),
    ],
)
