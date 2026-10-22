// NAAD Admin — service worker.
// Caches only the app's own files (page shell, icons, logos) so the app opens
// fast and shows its screen even with a weak connection. Student data and
// every call to the CMG server (Apps Script) always go to the network and are
// never stored here.
var CACHE = 'naad-admin-v4';
var SW_VERSION = 'push-5';
var SHELL = [
  './', 'admin.html', 'manifest.webmanifest',
  'icon-192.png', 'icon-512.png', 'icon-maskable-512.png', 'apple-touch-icon.png',
  'splash.png', 'bg_image.png', 'boot_logo.png',
  'logo_cmg.png', 'logo_fire.png', 'logo_police.png', 'logo_ambulance.png', 'logo_prisons.png',
  'logo_nadmo.png', 'logo_immigration.png', 'logo_customs.png', 'logo_tourism.png',
  'logo_forestry.png', 'logo_gaf.png', 'logo_nacoc.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (c) {
      return Promise.all(SHELL.map(function (u) { return c.add(u).catch(function () {}); }));
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return;       // server calls, fonts: untouched

  var isPage = req.mode === 'navigate' || url.pathname.slice(-10) === 'admin.html' || url.pathname.slice(-11) === 'portal.html';
  if (isPage) {
    // Network first so agencies always get the latest dashboard; cache is the offline fallback.
    e.respondWith(
      fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put(req, copy); });
        return res;
      }).catch(function () {
        return caches.match(req).then(function (m) { return m || caches.match('admin.html'); });
      })
    );
    return;
  }
  // Images / icons / manifest: cache first, refresh in the background.
  e.respondWith(
    caches.match(req).then(function (m) {
      var net = fetch(req).then(function (res) {
        if (res && res.ok) { var copy = res.clone(); caches.open(CACHE).then(function (c) { c.put(req, copy); }); }
        return res;
      }).catch(function () { return m; });
      return m || net;
    })
  );
});

// ════════════════════════════════════════════════════════════════════
// PUSH ALERTS — a system notification with sound/vibration and an icon
// badge count, even when the app is closed or the phone is on standby.
// The server sends an empty "wake-up" push; this worker then asks the
// server what is new for this device and shows the real message text.
// ════════════════════════════════════════════════════════════════════
function idb() {
  return new Promise(function (res, rej) {
    var r = indexedDB.open('naad-push', 1);
    r.onupgradeneeded = function () { r.result.createObjectStore('kv'); };
    r.onsuccess = function () { res(r.result); };
    r.onerror = function () { rej(r.error); };
  });
}
function kvGet(k) {
  return idb().then(function (db) {
    return new Promise(function (res) {
      var q = db.transaction('kv').objectStore('kv').get(k);
      q.onsuccess = function () { res(q.result); }; q.onerror = function () { res(undefined); };
    });
  }).catch(function () { return undefined; });
}
function kvSet(k, v) {
  return idb().then(function (db) {
    return new Promise(function (res) {
      var t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k);
      t.oncomplete = function () { res(); }; t.onerror = function () { res(); };
    });
  }).catch(function () {});
}
function setBadge(n) {
  try {
    if (n > 0 && self.navigator.setAppBadge) return self.navigator.setAppBadge(n).catch(function () {});
    if (n <= 0 && self.navigator.clearAppBadge) return self.navigator.clearAppBadge().catch(function () {});
  } catch (e) {}
  return Promise.resolve();
}

self.addEventListener('message', function (e) {
  var d = e.data || {};
  if (d.type === 'ping') {
    e.waitUntil(kvGet('cfg').then(function (c) {
      var reply = { v: SW_VERSION, hasCfg: !!(c && c.pt), perm: (self.Notification && Notification.permission) || '?' };
      if (e.ports && e.ports[0]) e.ports[0].postMessage(reply); else if (e.source) e.source.postMessage({ type: 'pong', info: reply });
    }));
    return;
  }
  if (d.type === 'cfg') {            // page tells us who this device is
    e.waitUntil(kvGet('cfg').then(function (old) {
      var keepSeen = old && old.pt === d.pt ? old.seen : 0;
      return kvSet('cfg', { url: d.url, pt: d.pt, home: d.home || 'admin.html', seen: Math.max(keepSeen || 0, d.seen || 0) });
    }));
  } else if (d.type === 'seen') {    // user has looked at the app: clear the badge
    e.waitUntil(kvGet('cfg').then(function (c) {
      if (c) { c.seen = Math.max(c.seen || 0, d.ts || 0); c.pending = 0; return kvSet('cfg', c); }
    }).then(function () { return setBadge(0); }));
  } else if (d.type === 'clear') {   // logged out
    e.waitUntil(kvSet('cfg', null).then(function () { return setBadge(0); }));
  }
});

self.addEventListener('push', function (e) { e.waitUntil(handlePush()); });

function handlePush() {
  return kvGet('cfg').then(function (cfg) {
    return self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (cl) {
      var visible = cl.filter(function (c) { return c.visibilityState === 'visible'; });
      var ios = /iPhone|iPad|iPod/.test(self.navigator.userAgent || '');
      if (visible.length && !ios) {   // app is open and in front: it plays its own sound + pop-up
        visible.forEach(function (c) { c.postMessage({ type: 'push-wake' }); });
        return;
      }
      var peek = Promise.resolve(null);
      if (cfg && cfg.url && cfg.pt) {
        peek = fetch(cfg.url + '?action=pushPeek&pt=' + encodeURIComponent(cfg.pt) + '&since=' + (cfg.seen || 0) + '&t=' + Date.now(), { cache: 'no-store' })
          .then(function (r) { return r.json(); }).then(function (d) { return d && d.status === 'ok' ? d : null; }).catch(function () { return null; });
      }
      return peek.then(function (d) {
        var list = d ? d.notifications : [];
        var count = d ? d.count : ((cfg && cfg.pending) || 0) + 1;
        if (cfg && !d) { cfg.pending = count; kvSet('cfg', cfg); }
        if (count < 1) count = 1;
        var top = list[0];
        var title = count > 1 ? '🔔 ' + count + ' new messages — NAAD' : (top && top.title) || '🔔 New message — NAAD';
        var body = top ? (count > 1 ? top.title + ': ' : '') + top.msg : 'Tap to open and read it.';
        return Promise.all([
          setBadge(count),
          self.registration.showNotification(title, {
            body: String(body).slice(0, 180),
            icon: 'icon-192.png', badge: 'icon-192.png',
            tag: 'naad-msg', renotify: true, requireInteraction: true, silent: false,
            vibrate: [250, 120, 250, 120, 500],
            actions: [{ action: 'reply', title: '↩ Reply', type: 'text', placeholder: 'Type your reply…' }, { action: 'open', title: 'Open' }],
            data: { url: (cfg && cfg.home) || 'admin.html', replyTo: top ? top.title : '' }
          })
        ]);
      });
    });
  });
}

self.addEventListener('notificationclick', function (e) {
  var n = e.notification, d = n.data || {};
  n.close();
  var target = (d.url || 'admin.html');
  var openApp = function () {
    return self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (cl) {
      for (var i = 0; i < cl.length; i++) {
        if ('focus' in cl[i]) { cl[i].postMessage({ type: 'push-open' }); return cl[i].focus(); }
      }
      return self.clients.openWindow(target + (target.indexOf('?') < 0 ? '?push=1' : '&push=1'));
    });
  };
  if (e.action === 'reply' && e.reply) {            // typed straight into the notification
    e.waitUntil(kvGet('cfg').then(function (cfg) {
      if (!cfg || !cfg.pt) return openApp();
      return fetch(cfg.url + '?action=pushReply&pt=' + encodeURIComponent(cfg.pt) + '&text=' + encodeURIComponent(e.reply) +
                   '&to=' + encodeURIComponent(d.replyTo || '') + '&t=' + Date.now(), { cache: 'no-store' })
        .then(function (r) { return r.json(); })
        .then(function (res) {
          var ok = res && res.status === 'ok';
          return self.registration.showNotification(ok ? '✅ Reply sent' : '⚠ Reply not sent', {
            body: ok ? String(e.reply).slice(0, 100) : ((res && res.message) || 'Open the app to reply.'),
            tag: 'naad-reply', silent: true, icon: 'icon-192.png', badge: 'icon-192.png', data: { url: target }
          }).then(function () { setTimeout(function () { self.registration.getNotifications({ tag: 'naad-reply' }).then(function (a) { a.forEach(function (x) { x.close(); }); }); }, 3500); });
        })
        .catch(function () { return openApp(); });
    }));
    return;
  }
  e.waitUntil(openApp());
});
