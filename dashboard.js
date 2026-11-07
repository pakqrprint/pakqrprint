// Pak QR Print - Admin Dashboard (GitHub Pages + Firebase)
// Login: Firebase Auth (email/password). Queue: Realtime Database.
// Auto-Open: naya order aate hi A4 preview modal khud khul jata hai.
// Print: dukaan ke PC par print agent karta hai (status se).
(function () {
'use strict';
function $(id) { return document.getElementById(id); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>\"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function showErr(msg) { var e = $('setupErr'); e.textContent = msg; e.style.display = 'block'; }

var cfg = window.PAKQR_FIREBASE_CONFIG || {};
var CONFIG_OK = cfg.apiKey && cfg.apiKey.indexOf('YAHAN-') !== 0 && cfg.databaseURL && cfg.databaseURL.indexOf('YAHAN-') !== 0;
var sbcfg = window.PAKQR_SUPABASE_CONFIG || {};
var SUPABASE_OK = sbcfg.url && sbcfg.url.indexOf('YAHAN-') !== 0 && sbcfg.anonKey && sbcfg.anonKey.indexOf('YAHAN-') !== 0 && sbcfg.bucket;
var SHOP_ID = window.PAKQR_SHOP_ID || 'pakqrprint';
var DEFAULT_SETTINGS = { shopName: 'Pak QR Print', autoPrint: true, autoOpen: true, permissionMode: false, autoDeleteMinutes: 5, printerName: '' };
var settings = DEFAULT_SETTINGS;
var db = null, auth = null;
var orders = {}; // key -> order
var seenKeys = {};
var loadTime = Date.now();
var a4Order = null, a4Zoom = 1, a4Rot = 0, a4ViewUrl = '';

function customerUrl() {
  if (window.PAKQR_CUSTOMER_URL) return window.PAKQR_CUSTOMER_URL.replace(/\/$/, '') + '/index.html?shop=' + SHOP_ID;
  return location.origin + location.pathname.replace(/[^\/]*$/, '') + 'index.html?shop=' + SHOP_ID;
}

// ---------- Supabase Storage (plain REST, koi CDN library nahi) ----------
// Upload: POST /storage/v1/object/<bucket>/<path> (anon key se).
// Public URL: <url>/storage/v1/object/public/<bucket>/<path>.
// NOTE (v3.32): bucket ki policy mein anon ke liye DELETE bhi lagna
// chahiye (Supabase dashboard > Storage > pakqrprint-files > Policies) —
// tab fileSweep saat (7) din se purani files khud delete karta hai.
// Policy na ho to delete 401/403 dega aur sweep khamoshi se ruk jata
// hai; dashboard ko kuch nahi hota.
function supabaseUpload(file, objectPath) {
  var base = String(sbcfg.url).replace(/\/$/, '');
  var encPath = objectPath.split('/').map(encodeURIComponent).join('/');
  var url = base + '/storage/v1/object/' + encodeURIComponent(sbcfg.bucket) + '/' + encPath;
  return fetch(url, {
    method: 'POST',
    headers: {
      'apikey': sbcfg.anonKey,
      'Authorization': 'Bearer ' + sbcfg.anonKey,
      'Content-Type': file.type || 'application/octet-stream',
      'x-upsert': 'true'
    },
    body: file
  }).then(function (res) {
    if (!res.ok) return res.text().then(function (t) { throw new Error('Supabase upload fail (HTTP ' + res.status + '): ' + String(t || '').slice(0, 200)); });
    return base + '/storage/v1/object/public/' + encodeURIComponent(sbcfg.bucket) + '/' + encPath;
  });
}

// ---------- v3.32: 7-day file sweep (Supabase bucket safai) ----------
// Har file saat (7) din ke baad khud-ba-khud delete ho — haftay mein
// jitni bhi files hon. Ye sweep sirf tab chalta hai jab admin dashboard
// khula ho (load par + har ghantay).
// RULES:
//  - 7 din se purani file hi delete hoti hai (umar = bucket ki updated_at).
//  - Zinda (pending) order ki file KABHI delete nahi hoti — chahe kitni
//    purani ho. Skip-set har sweep mein RTDB se taaza li jati hai.
//  - deleteOrder sirf RTDB order delete karta hai — file is sweep se
//    delete hoti hai, order delete hote waqt nahi.
//  - Supabase mein DELETE policy na ho to delete 401/403 dega — sweep
//    khamoshi se ruk jata hai, dashboard ko kuch nahi hota, aur koi RTDB
//    order kabhi delete nahi hota.
var FILE_SWEEP_DAYS = 7;
var FILE_SWEEP_MS = FILE_SWEEP_DAYS * 24 * 60 * 60 * 1000;
var fileSweepRunning = false;
var fileSweepTimer = null;

// Supabase Storage REST delete (wahi plain-fetch style jo supabaseUpload
// use karta hai — koi library nahi). DELETE policy Supabase mein lagni
// chahiye, warna 401/403 aayega.
function supabaseDelete(objectPath) {
  var base = String(sbcfg.url).replace(/\/$/, '');
  var encPath = objectPath.split('/').map(encodeURIComponent).join('/');
  var url = base + '/storage/v1/object/' + encodeURIComponent(sbcfg.bucket) + '/' + encPath;
  return fetch(url, {
    method: 'DELETE',
    headers: { 'apikey': sbcfg.anonKey, 'Authorization': 'Bearer ' + sbcfg.anonKey }
  }).then(function (res) {
    if (!res.ok) throw new Error('Supabase delete fail (HTTP ' + res.status + ')');
    return true;
  });
}

// Supabase Storage REST list (aik level). Folder entries ki pehchan:
// un ki updated_at nahi hoti.
function supabaseList(prefix, offset) {
  var base = String(sbcfg.url).replace(/\/$/, '');
  var url = base + '/storage/v1/object/list/' + encodeURIComponent(sbcfg.bucket);
  return fetch(url, {
    method: 'POST',
    headers: {
      'apikey': sbcfg.anonKey,
      'Authorization': 'Bearer ' + sbcfg.anonKey,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ prefix: prefix, limit: 100, offset: offset || 0 })
  }).then(function (res) {
    if (!res.ok) throw new Error('Supabase list fail (HTTP ' + res.status + ')');
    return res.json();
  });
}

// Aik level ki saari entries (paging ke sath).
function supabaseListAll(prefix) {
  var out = [];
  function page(offset) {
    return supabaseList(prefix, offset).then(function (entries) {
      entries = entries || [];
      out = out.concat(entries);
      if (entries.length >= 100) return page(offset + 100);
      return out;
    });
  }
  return page(0);
}

// prefix ke neeche saari files (folders mein utar kar). File = entry jis
// ki updated_at ho; folder = updated_at nahi (depth 3 tak).
function listFilesDeep(prefix, depth) {
  return supabaseListAll(prefix).then(function (entries) {
    var files = [];
    var folders = [];
    entries.forEach(function (e) {
      if (!e || !e.name) return;
      if (e.updated_at) files.push({ path: prefix + e.name, updatedAt: e.updated_at });
      else if (depth < 3) folders.push(prefix + e.name + '/');
    });
    var chain = Promise.resolve();
    folders.forEach(function (fp) {
      chain = chain.then(function () { return listFilesDeep(fp, depth + 1); }).then(function (sub) { files = files.concat(sub); });
    });
    return chain.then(function () { return files; });
  });
}

// 7-day file sweep — dashboard khula ho to purani files khud delete hon.
function fileSweep() {
  if (!SUPABASE_OK || !db || fileSweepRunning) return Promise.resolve();
  fileSweepRunning = true;
  var cutoff = Date.now() - FILE_SWEEP_MS;
  var livePaths = {};
  var aborted = false;
  function isPolicyError(e) { return /HTTP 40[13]/.test(String((e && e.message) || e)); }
  return ordersRef().once('value').then(function (snap) {
    var all = snap.val() || {};
    Object.keys(all).forEach(function (k) {
      var o = all[k] || {};
      // Zinda = status Done / Printed / Cancelled NAHI hai (New,
      // PrintRequested, Printing, PrintFailed ya koi aur).
      var live = !(o.status === 'Done' || o.status === 'Printed' || o.status === 'Cancelled');
      (o.files || []).forEach(function (f) { if (f && f.storagePath && live) livePaths[f.storagePath] = true; });
    });
    return listFilesDeep('shops/' + SHOP_ID + '/previews/', 0);
  }).then(function (previewFiles) {
    var files = previewFiles;
    return listFilesDeep('shops/' + SHOP_ID + '/orders/', 0).then(function (orderFiles) { return files.concat(orderFiles); });
  }).then(function (files) {
    var chain = Promise.resolve();
    files.forEach(function (f) {
      chain = chain.then(function () {
        if (aborted) return;
        var age = Date.parse(f.updatedAt);
        if (!age || age > cutoff) return; // umar maloom nahi ya 7 din se chhoti — chhor dein
        if (livePaths[f.path]) return; // zinda order ki file — KABHI delete nahi
        return supabaseDelete(f.path).catch(function (e) {
          if (isPolicyError(e)) aborted = true; // DELETE policy nahi — aagay kuch delete nahi karna
        });
      });
    });
    return chain;
  }).catch(function (e) {
    // List ya RTDB fail: khamoshi se ruk jayein (sirf console par likh dein).
    if (typeof console !== 'undefined' && console.log) console.log('File sweep did not run: ' + (e && e.message ? e.message : e));
  }).then(function () { fileSweepRunning = false; });
}
function ordersRef() { return db.ref('shops/' + SHOP_ID + '/orders'); }
function settingsRef() { return db.ref('shops/' + SHOP_ID + '/settings'); }

function badge(status) {
  var map = { New: 'new', Printing: 'printing', Printed: 'printed', Done: 'done', PrintFailed: 'failed', Cancelled: 'failed' };
  var label = { New: 'NEW', Printing: 'PRINTING', Printed: 'PRINTED', Done: 'DONE', PrintFailed: 'FAILED', Cancelled: 'CANCELLED' };
  return '<span class="badge ' + (map[status] || 'new') + '">' + (label[status] || status) + '</span>';
}
function renderStats() {
  var walkin = 0, pending = 0, printed = 0, failed = 0;
  Object.keys(orders).forEach(function (k) {
    var o = orders[k];
    if (o.source === 'walkin' || o.source === 'direct') walkin++;
    if (o.status === 'New' || o.status === 'Printing' || o.status === 'PrintRequested') pending++;
    else if (o.status === 'Printed' || o.status === 'Done') printed++;
    else if (o.status === 'PrintFailed' || o.status === 'Cancelled') failed++;
  });
  $('stWalkin').textContent = walkin;
  $('stPending').textContent = pending;
  $('stPrinted').textContent = printed;
  $('stFailed').textContent = failed;
}
function renderQueue() {
  var box = $('queueList'); box.innerHTML = '';
  var keys = Object.keys(orders).sort(function (a, b) { return (orders[a].createdAt || 0) - (orders[b].createdAt || 0); });
  if (!keys.length) { box.innerHTML = '<div class="qempty"><div class="qfolder">📁</div><b>No Customer Uploads Yet</b><p>Share your customer link or show your counter QR Code to receive instant print files.</p></div>'; renderStats(); return; }
  keys.forEach(function (k) {
    var o = orders[k];
    var d = document.createElement('div'); d.className = 'qitem';
    var files = (o.files || []).map(function (f) {
      return esc(f.name) + ' — ' + (f.pages || 1) + ' pg × ' + (f.copies || 1) + ' • ' + esc(f.colorMode || o.colorMode || 'BW') + (f.nup > 1 ? ' • ' + f.nup + '/sheet' : '') + (f.range ? ' • ' + esc(f.range) : '');
    }).join('<br>');
    d.innerHTML = '<div class="qhead"><span>Token ' + esc(o.token || o.orderNo || '') + ' — ' + esc(o.customerName || 'Walk-in') + '</span><span>' + badge(o.status) + '</span></div>'
      + '<div class="qfile">' + files + (o.source === 'idcard' ? '<br>ID Card (' + esc(o.idType || '') + ')' : '') + '</div>'
      + (o.printLog && o.printLog.length ? '<div class="qfile" style="color:#991b1b">' + esc(o.printLog[o.printLog.length - 1].error || o.printLog[o.printLog.length - 1].method || '') + '</div>' : '')
      + '<div class="qacts">'
      + '<button type="button" class="primary" data-a4="' + k + '">A4 Preview</button>'
      + '<button type="button" data-pr="' + k + '">Direct Print</button>'
      + '<button type="button" data-vw="' + k + '">View File</button>'
      + '<button type="button" data-done="' + k + '">Done</button>'
      + '<button type="button" class="danger" data-del="' + k + '">Delete</button>'
      + '</div>';
    box.appendChild(d);
  });
  box.querySelectorAll('[data-a4]').forEach(function (b) { b.addEventListener('click', function () { openA4Modal(orders[b.getAttribute('data-a4')]); }); });
  box.querySelectorAll('[data-pr]').forEach(function (b) { b.addEventListener('click', function () { directPrint(b.getAttribute('data-pr'), b); }); });
  box.querySelectorAll('[data-vw]').forEach(function (b) { b.addEventListener('click', function () { var o = orders[b.getAttribute('data-vw')]; if (o && o.files && o.files[0] && o.files[0].downloadURL) window.open(o.files[0].downloadURL, '_blank'); }); });
  box.querySelectorAll('[data-done]').forEach(function (b) { b.addEventListener('click', function () { ordersRef().child(b.getAttribute('data-done')).update({ status: 'Done' }); }); });
  box.querySelectorAll('[data-del]').forEach(function (b) { b.addEventListener('click', function () { if (confirm('Delete this order?')) deleteOrder(b.getAttribute('data-del')); }); });
  renderStats();
}
function deleteOrder(key) {
  // v3.32: Sirf order (Realtime Database se) delete hota hai — file
  // yahan delete NAHI hoti. File saat (7) din ke baad fileSweep se khud
  // delete hoti hai, taake 7-din ka rule aik hi jagah se chale.
  ordersRef().child(key).remove();
}
function directPrint(key, btn) {
  if (btn) { btn.disabled = true; btn.textContent = 'Printing...'; }
  ordersRef().child(key).update({ status: 'PrintRequested', printRequestedAt: firebase.database.ServerValue.TIMESTAMP }).then(function () {
    alert('Print has been sent — the print agent on the shop PC will print it. (The PC must be on, with internet and a printer connected.)');
  }).catch(function (e) { alert('Print could not be sent: ' + e.message); if (btn) { btn.disabled = false; btn.textContent = 'Direct Print'; } });
}

// ---------- A4 preview modal ----------
function applyA4Transform() { var el = $('a4Content'); if (el) el.style.transform = 'scale(' + a4Zoom + ') rotate(' + a4Rot + 'deg)'; }
function openA4Modal(o) {
  if (!o) return;
  a4Order = o; a4Zoom = 1; a4Rot = 0;
  $('a4Title').textContent = (o.customerName || 'Walk-in') + ' — Token ' + (o.token || o.orderNo || '');
  $('a4Meta').textContent = 'Copies ' + (o.files && o.files[0] ? (o.files[0].copies || 1) : 1) + ' • Color Mode ' + (o.colorMode || 'BW') + ' • A4 Printable Layout Preview (210×297 mm)';
  var box = $('a4Content'); box.innerHTML = '<p class="hint">Loading preview...</p>';
  $('a4Modal').style.display = 'block';
  var f = (o.files && o.files[0]) || null;
  if (!f || !f.downloadURL) { box.innerHTML = '<p class="hint">No preview is available for this order file.</p>'; $('a4BrowserBtn').style.display = 'none'; return; }
  a4ViewUrl = f.downloadURL;
  // v3.19: Browser Print sirf PDF aur photo ke liye (Office iframe cross-origin hai - browser se print nahi hota)
  var canBrowser = (f.kind === 'image' || f.kind === 'idcard' || f.kind === 'pdf' || /\.(jpe?g|png|pdf)$/i.test(f.name || ''));
  $('a4BrowserBtn').style.display = canBrowser ? '' : 'none';
  if (f.kind === 'image' || f.kind === 'idcard' || /\.(jpe?g|png)$/i.test(f.name || '')) {
    box.innerHTML = '<img src="' + esc(f.downloadURL) + '" alt="preview">';
  } else if (f.kind === 'pdf' || /\.pdf$/i.test(f.name || '')) {
    box.innerHTML = '<p class="hint">Rendering the first page of the PDF...</p>';
    import('./pdfjs/pdf.mjs').then(function (pdfjsLib) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = './pdfjs/pdf.worker.mjs';
      return pdfjsLib.getDocument({ url: f.downloadURL }).promise;
    }).then(function (pdf) { return pdf.getPage(1); }).then(function (page) {
      var cv = document.createElement('canvas');
      var vp = page.getViewport({ scale: 1.5 });
      cv.width = vp.width; cv.height = vp.height;
      box.innerHTML = ''; box.appendChild(cv);
      return page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
    }).catch(function () { box.innerHTML = '<p class="hint">PDF preview did not open — open the file with "View File".</p>'; });
  } else if (/\.(docx?|xlsx?|pptx?)$/i.test(f.name || '')) {
    box.innerHTML = '<iframe src="https://view.officeapps.live.com/op/embed.aspx?src=' + encodeURIComponent(f.downloadURL) + '" style="width:100%;height:420px;border:1.5px solid #dfe2f7;border-radius:10px" frameborder="0"></iframe>';
  } else {
    box.innerHTML = '<p class="hint">' + esc(f.name) + ' (' + esc((f.kind || '').toUpperCase()) + ') — there is no browser preview for this file. Open it with "View File" or print it with "Direct Print".</p>';
  }
  applyA4Transform();
}
// v3.19: Browser Print - PDF (saare pages) aur photo seedha browser se print (agent ki zaroorat nahi).
// Office files (Word/Excel/PowerPoint) ke liye nahi - un ka viewer iframe cross-origin hai.
function browserPrint() {
  var f = (a4Order && a4Order.files && a4Order.files[0]) || null;
  if (!f || !f.downloadURL) { alert('Browser print is available for PDF and photo files only.'); return; }
  var isImg = f.kind === 'image' || f.kind === 'idcard' || /\.(jpe?g|png)$/i.test(f.name || '');
  var isPdf = f.kind === 'pdf' || /\.pdf$/i.test(f.name || '');
  if (!isImg && !isPdf) { alert('Browser print works for PDF and photo files only. Word, Excel and PowerPoint need the print agent (Direct Print).'); return; }
  var area = $('printArea');
  if (!area) { area = document.createElement('div'); area.id = 'printArea'; document.body.appendChild(area); }
  area.innerHTML = '';
  var done = false;
  function goPrint() { if (!done) { done = true; document.body.classList.add('browser-printing'); window.print(); } } // v3.20: class tabhi lagayein jab Browser Print ho - warna poster blank print hota
  if (isImg) {
    var im = document.createElement('img');
    im.alt = f.name || 'print';
    im.onload = goPrint;
    im.onerror = function () { alert('The image could not be loaded for browser print.'); };
    im.src = f.downloadURL;
    area.appendChild(im);
    if (im.complete) goPrint();
  } else {
    import('./pdfjs/pdf.mjs').then(function (pdfjsLib) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = './pdfjs/pdf.worker.mjs';
      return pdfjsLib.getDocument({ url: f.downloadURL }).promise;
    }).then(function (pdf) {
      var chain = Promise.resolve();
      for (var p = 1; p <= pdf.numPages; p++) {
        (function (pn) {
          chain = chain.then(function () { return pdf.getPage(pn); }).then(function (page) {
            var cv = document.createElement('canvas');
            var vp = page.getViewport({ scale: 1.5 });
            cv.width = vp.width; cv.height = vp.height;
            area.appendChild(cv);
            return page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
          });
        })(p);
      }
      return chain;
    }).then(goPrint).catch(function () { alert('The PDF could not be prepared for browser print. Use Direct Print (agent) instead.'); });
  }
}
function closeA4Modal() { $('a4Modal').style.display = 'none'; a4Order = null; }

// ---------- Shop Counter QR modal ----------
function openQrModal() {
  $('qrShopName').textContent = settings.shopName || 'Pak QR Print';
  var ok = false;
  try { ok = drawQRToCanvas($('qrModalCanvas'), customerUrl()); } catch (e) { ok = false; }
  if (ok) { try { drawQRBadge($('qrModalCanvas')); } catch (e) {} } // v3.25: modal QR bhi poster jaisa (badge)
  if (!ok) { $('qrShopName').textContent = 'QR could not be made — check the link.'; }
  $('qrModal').style.display = 'flex';
}
function closeQrModal() { $('qrModal').style.display = 'none'; }

// ---------- settings / toggles ----------
function bindToggle(id, key) {
  var el = $(id); if (!el) return;
  el.addEventListener('change', function () { var u = {}; u[key] = el.checked; settingsRef().update(u); });
}
function applySettingsToUI() {
  document.querySelectorAll('[data-shopname]').forEach(function (el) { el.textContent = settings.shopName || 'Pak QR Print'; });
  $('togAutoOpen').checked = !!settings.autoOpen;
  $('togAutoPrint').checked = !!settings.autoPrint;
  $('togPerm').checked = !!settings.permissionMode;
  $('selAutoDel').value = String(settings.autoDeleteMinutes == null ? 5 : settings.autoDeleteMinutes);
  $('setShopName').value = settings.shopName || 'Pak QR Print';
  $('setPrinter').value = settings.printerName || '';
}
function saveSettings() {
  settingsRef().update({
    shopName: ($('setShopName').value || 'Pak QR Print').trim(),
    printerName: ($('setPrinter').value || '').trim()
  }).then(function () { alert('Settings have been saved.'); }).catch(function (e) { alert('Settings could not be saved: ' + e.message); });
}

// ---------- auto-delete sweep ----------
function sweepOld() {
  var mins = settings.autoDeleteMinutes;
  if (!mins) return;
  var cutoff = Date.now() - mins * 60 * 1000;
  Object.keys(orders).forEach(function (k) {
    var o = orders[k];
    if ((o.status === 'Done' || o.status === 'Cancelled' || o.status === 'Printed') && o.createdAt && o.createdAt < cutoff) deleteOrder(k);
  });
}

// ---------- direct upload (walk-in) ----------
function directUpload(files) {
  var arr = Array.prototype.slice.call(files || []);
  if (!arr.length) return;
  var orderKey = ordersRef().push().key;
  var uploads = arr.map(function (f, idx) {
    var safeName = String(f.name).replace(/[^\w.\-() ]/g, '_');
    var objectPath = 'shops/' + SHOP_ID + '/orders/' + orderKey + '/' + idx + '_' + safeName;
    return supabaseUpload(f, objectPath).then(function (url) {
      return { name: f.name, kind: 'other', size: f.size, pages: 1, copies: 1, colorMode: 'BW', rangeMode: 'all', range: '', nup: 1, spp: 1, paperSize: 'a4', downloadURL: url, storagePath: objectPath };
    });
  });
  var savedFiles = [];
  Promise.all(uploads).then(function (upFiles) {
    savedFiles = upFiles;
    return db.ref('shops/' + SHOP_ID + '/counter').transaction(function (cur) { return (cur || 1000) + 1; });
  }).then(function (res) {
    var n = res.snapshot.val() || 1001;
    return ordersRef().child(orderKey).set({ orderNo: 'QP-' + n, token: 'QP-' + n, customerName: 'Walk-in', createdAt: firebase.database.ServerValue.TIMESTAMP, status: 'New', colorMode: 'BW', copies: 1, source: 'walkin', files: savedFiles, printLog: [] });
  }).then(function () { alert('Walk-in order has been added to the queue.'); }).catch(function (e) { alert('Upload failed: ' + e.message); });
}

// ---------- init ----------
function startDashboard() {
  $('authBox').style.display = 'none';
  $('dashBox').style.display = 'block';
  settingsRef().on('value', function (snap) {
    var s = snap.val();
    if (s) { settings = Object.assign({}, DEFAULT_SETTINGS, s); applySettingsToUI(); }
  });
  ordersRef().on('child_added', function (snap) {
    var isNew = !seenKeys[snap.key];
    orders[snap.key] = snap.val() || {};
    seenKeys[snap.key] = true;
    renderQueue();
    // Auto-Open: page khulne ke BAAD aane wala naya order khud khul jaye
    var o = orders[snap.key];
    if (isNew && settings.autoOpen && o.createdAt && o.createdAt > loadTime) openA4Modal(o);
  });
  ordersRef().on('child_changed', function (snap) { orders[snap.key] = snap.val() || {}; renderQueue(); });
  ordersRef().on('child_removed', function (snap) { delete orders[snap.key]; renderQueue(); });
  setInterval(sweepOld, 60000);
  fileSweep(); // v3.32: load par pehli safai (7 din se purani files)
  fileSweepTimer = setInterval(fileSweep, 60 * 60 * 1000); // v3.32: khula rahe to har ghantay
  renderQueue();
}

document.addEventListener('DOMContentLoaded', function () {
  if (!CONFIG_OK) { showErr('Setup is pending: your Firebase config is not filled in firebase-config.js yet. Paste the config as per Step 2 of the guide file (Pak-QR-Print-Hidayat-v3.txt).'); return; }
  if (!SUPABASE_OK) { showErr('Setup is pending: the Supabase config is not filled in supabase-config.js yet. Paste the config as per Step 2 of the guide file.'); return; }
  try {
    firebase.initializeApp(cfg);
    db = firebase.database();
    auth = firebase.auth();
  } catch (e) { showErr('Firebase did not start: ' + e.message); return; }
  $('loginBtn').addEventListener('click', function () {
    var em = $('loginEmail').value.trim(), pw = $('loginPass').value;
    $('loginMsg').textContent = 'Logging in...';
    auth.signInWithEmailAndPassword(em, pw).then(function () { startDashboard(); }).catch(function (e) { $('loginMsg').textContent = 'Login failed: ' + e.message; });
  });
  // Admin Portal tabs
  document.querySelectorAll('#portalTabs button').forEach(function (b) {
    b.addEventListener('click', function () {
      document.querySelectorAll('#portalTabs button').forEach(function (x) { x.classList.remove('on'); });
      b.classList.add('on');
      var t = b.getAttribute('data-ptab');
      $('portalLogin').style.display = t === 'login' ? 'block' : 'none';
      $('portalCreate').style.display = t === 'create' ? 'block' : 'none';
    });
  });
  // Create Account: Firebase Auth user + shop name settings mein save
  $('createBtn').addEventListener('click', function () {
    var em = $('createEmail').value.trim(), pw = $('createPass').value;
    var shop = $('createShop').value.trim();
    if (!em || pw.length < 6) { $('createMsg').textContent = 'Enter an email and a password of at least 6 characters.'; return; }
    $('createMsg').textContent = 'Creating account...';
    auth.createUserWithEmailAndPassword(em, pw).then(function () {
      if (shop) settingsRef().update({ shopName: shop }).catch(function () { /* shop name save na ho to bhi account ban gaya */ });
      startDashboard();
    }).catch(function (e) { $('createMsg').textContent = 'Account could not be created: ' + e.message; });
  });
  auth.onAuthStateChanged(function (u) { if (u) startDashboard(); });
  $('logoutBtn').addEventListener('click', function () { auth.signOut().then(function () { location.reload(); }); });
  $('copyLinkBtn').addEventListener('click', function () {
    var u = customerUrl();
    if (navigator.clipboard) navigator.clipboard.writeText(u).then(function () { alert('Customer link copied:\n' + u); });
    else prompt('Copy customer link:', u);
  });
  $('qrBtn').addEventListener('click', function () { openQrModal(); });
  $('qrCloseBtn').addEventListener('click', closeQrModal);
  $('qrModal').addEventListener('click', function (e) { if (e.target === $('qrModal')) closeQrModal(); });
  $('qrDownloadBtn').addEventListener('click', function () {
    // v3.25: modal canvas par badge pehle se hai (openQrModal) - seedha download karein.
    var a = document.createElement('a');
    a.href = $('qrModalCanvas').toDataURL('image/png');
    a.download = 'Pak-QR-Print-QR.png';
    a.click();
  });
  $('qrPrintBtn').addEventListener('click', function () { window.open('qr-poster.html?shop=' + SHOP_ID, '_blank'); });
  $('posterBtn').addEventListener('click', function () { window.open('qr-poster.html?shop=' + SHOP_ID, '_blank'); });
  $('homeBtn').addEventListener('click', function () { location.href = 'index.html?shop=' + SHOP_ID; });
  $('duInput').addEventListener('change', function () { directUpload($('duInput').files); $('duInput').value = ''; });
  $('duBtn').addEventListener('click', function () { $('duInput').click(); });
  $('clearBtn').addEventListener('click', function () {
    if (!confirm('Delete all Printed / Done / Cancelled orders?')) return;
    Object.keys(orders).forEach(function (k) { var o = orders[k]; if (o.status === 'Printed' || o.status === 'Done' || o.status === 'Cancelled') deleteOrder(k); });
  });
  $('refreshBtn').addEventListener('click', function () { renderQueue(); });
  bindToggle('togAutoOpen', 'autoOpen');
  bindToggle('togAutoPrint', 'autoPrint');
  bindToggle('togPerm', 'permissionMode');
  $('selAutoDel').addEventListener('change', function () { settingsRef().update({ autoDeleteMinutes: parseInt($('selAutoDel').value, 10) || 0 }); });
  $('saveSetBtn').addEventListener('click', saveSettings);
  $('a4CloseBtn').addEventListener('click', closeA4Modal);
  $('a4PrintBtn').addEventListener('click', function () { if (a4Order) { var k = Object.keys(orders).find(function (key) { return orders[key] === a4Order; }); if (k) directPrint(k, null); closeA4Modal(); } });
  $('a4ZoomIn').addEventListener('click', function () { a4Zoom = Math.min(2, a4Zoom + 0.2); applyA4Transform(); });
  $('a4ZoomOut').addEventListener('click', function () { a4Zoom = Math.max(0.5, a4Zoom - 0.2); applyA4Transform(); });
  $('a4Rotate').addEventListener('click', function () { a4Rot = (a4Rot + 90) % 360; applyA4Transform(); });
  $('a4ViewBtn').addEventListener('click', function () { if (a4ViewUrl) window.open(a4ViewUrl, '_blank'); });
  $('a4BrowserBtn').addEventListener('click', browserPrint);
  window.addEventListener('afterprint', function () { document.body.classList.remove('browser-printing'); var a = $('printArea'); if (a) a.innerHTML = ''; });
  $('a4Modal').addEventListener('click', function (e) { if (e.target === $('a4Modal')) closeA4Modal(); });
});
})();
