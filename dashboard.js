// Pak QR Print - Shopkeeper Dashboard (GitHub Pages + Firebase)
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
var SHOP_ID = window.PAKQR_SHOP_ID || 'pakqrprint';
var DEFAULT_SETTINGS = { shopName: 'Pak QR Print', autoPrint: true, autoOpen: true, permissionMode: false, autoDeleteMinutes: 5, printerName: '' };
var settings = DEFAULT_SETTINGS;
var db = null, storage = null, auth = null;
var orders = {}; // key -> order
var seenKeys = {};
var loadTime = Date.now();
var a4Order = null, a4Zoom = 1, a4Rot = 0, a4ViewUrl = '';

function customerUrl() {
  if (window.PAKQR_CUSTOMER_URL) return window.PAKQR_CUSTOMER_URL.replace(/\/$/, '') + '/index.html?shop=' + SHOP_ID;
  return location.origin + location.pathname.replace(/[^\/]*$/, '') + 'index.html?shop=' + SHOP_ID;
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
  if (!keys.length) { box.innerHTML = '<p class="hint" style="margin-top:10px">No Customer Uploads Yet — Share your customer link or show your counter QR Code to receive instant print files.</p>'; renderStats(); return; }
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
  box.querySelectorAll('[data-del]').forEach(function (b) { b.addEventListener('click', function () { if (confirm('Ye order delete kar dein?')) deleteOrder(b.getAttribute('data-del')); }); });
  renderStats();
}
function deleteOrder(key) {
  var o = orders[key];
  ordersRef().child(key).remove();
  if (o && o.files) o.files.forEach(function (f) { if (f.storagePath) { try { storage.ref(f.storagePath).delete().catch(function () {}); } catch (e) {} } });
}
function directPrint(key, btn) {
  if (btn) { btn.disabled = true; btn.textContent = 'Printing...'; }
  ordersRef().child(key).update({ status: 'PrintRequested', printRequestedAt: firebase.database.ServerValue.TIMESTAMP }).then(function () {
    alert('Print bhej diya gaya hai — dukaan ke PC par print agent is ko print karega. (PC on + internet + printer hona lazmi hai.)');
  }).catch(function (e) { alert('Print nahi bheja ja saka: ' + e.message); if (btn) { btn.disabled = false; btn.textContent = 'Direct Print'; } });
}

// ---------- A4 preview modal ----------
function applyA4Transform() { var el = $('a4Content'); if (el) el.style.transform = 'scale(' + a4Zoom + ') rotate(' + a4Rot + 'deg)'; }
function openA4Modal(o) {
  if (!o) return;
  a4Order = o; a4Zoom = 1; a4Rot = 0;
  $('a4Title').textContent = (o.customerName || 'Walk-in') + ' — Token ' + (o.token || o.orderNo || '');
  $('a4Meta').textContent = 'Copies ' + (o.files && o.files[0] ? (o.files[0].copies || 1) : 1) + ' • Color Mode ' + (o.colorMode || 'BW') + ' • A4 Printable Layout Preview (210×297 mm)';
  var box = $('a4Content'); box.innerHTML = '<p class="hint">Preview load ho raha hai...</p>';
  $('a4Modal').style.display = 'block';
  var f = (o.files && o.files[0]) || null;
  if (!f || !f.downloadURL) { box.innerHTML = '<p class="hint">Is order ki file ka preview nahi hai.</p>'; return; }
  a4ViewUrl = f.downloadURL;
  if (f.kind === 'image' || f.kind === 'idcard' || /\.(jpe?g|png)$/i.test(f.name || '')) {
    box.innerHTML = '<img src="' + esc(f.downloadURL) + '" alt="preview">';
  } else if (f.kind === 'pdf' || /\.pdf$/i.test(f.name || '')) {
    box.innerHTML = '<p class="hint">PDF ka pehla page render ho raha hai...</p>';
    import('./pdfjs/pdf.mjs').then(function (pdfjsLib) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = './pdfjs/pdf.worker.mjs';
      return pdfjsLib.getDocument({ url: f.downloadURL }).promise;
    }).then(function (pdf) { return pdf.getPage(1); }).then(function (page) {
      var cv = document.createElement('canvas');
      var vp = page.getViewport({ scale: 1.5 });
      cv.width = vp.width; cv.height = vp.height;
      box.innerHTML = ''; box.appendChild(cv);
      return page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
    }).catch(function () { box.innerHTML = '<p class="hint">PDF preview nahi khula — "View File" se file kholein.</p>'; });
  } else {
    box.innerHTML = '<p class="hint">' + esc(f.name) + ' (' + esc((f.kind || '').toUpperCase()) + ') — is file ka browser preview nahi hai. "View File" se file kholein ya "Direct Print" se print karein.</p>';
  }
  applyA4Transform();
}
function closeA4Modal() { $('a4Modal').style.display = 'none'; a4Order = null; }

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
  }).then(function () { alert('Settings save ho gayi hain.'); }).catch(function (e) { alert('Settings save nahi hui: ' + e.message); });
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
    var ref = storage.ref('shops/' + SHOP_ID + '/orders/' + orderKey + '/' + idx + '_' + safeName);
    return ref.put(f).then(function (snap) { return snap.ref.getDownloadURL(); }).then(function (url) {
      return { name: f.name, kind: 'other', size: f.size, pages: 1, copies: 1, colorMode: 'BW', rangeMode: 'all', range: '', nup: 1, spp: 1, downloadURL: url, storagePath: ref.fullPath };
    });
  });
  var savedFiles = [];
  Promise.all(uploads).then(function (upFiles) {
    savedFiles = upFiles;
    return db.ref('shops/' + SHOP_ID + '/counter').transaction(function (cur) { return (cur || 1000) + 1; });
  }).then(function (res) {
    var n = res.snapshot.val() || 1001;
    return ordersRef().child(orderKey).set({ orderNo: 'QP-' + n, token: 'QP-' + n, customerName: 'Walk-in', createdAt: firebase.database.ServerValue.TIMESTAMP, status: 'New', colorMode: 'BW', copies: 1, source: 'walkin', files: savedFiles, printLog: [] });
  }).then(function () { alert('Walk-in order queue mein aa gaya hai.'); }).catch(function (e) { alert('Upload nahi hua: ' + e.message); });
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
  renderQueue();
}

document.addEventListener('DOMContentLoaded', function () {
  if (!CONFIG_OK) { showErr('Setup baaki hai: firebase-config.js mein apni Firebase config abhi bhari nahi gayi. Hidayat file (Pak-QR-Print-GitHub-Hidayat.txt) ke Step 2 ke mutabiq config paste karein.'); return; }
  try {
    firebase.initializeApp(cfg);
    db = firebase.database();
    storage = firebase.storage();
    auth = firebase.auth();
  } catch (e) { showErr('Firebase shuru nahi hua: ' + e.message); return; }
  $('loginBtn').addEventListener('click', function () {
    var em = $('loginEmail').value.trim(), pw = $('loginPass').value;
    $('loginMsg').textContent = 'Login ho raha hai...';
    auth.signInWithEmailAndPassword(em, pw).then(function () { startDashboard(); }).catch(function (e) { $('loginMsg').textContent = 'Login nahi hua: ' + e.message; });
  });
  auth.onAuthStateChanged(function (u) { if (u) startDashboard(); });
  $('logoutBtn').addEventListener('click', function () { auth.signOut().then(function () { location.reload(); }); });
  $('copyLinkBtn').addEventListener('click', function () {
    var u = customerUrl();
    if (navigator.clipboard) navigator.clipboard.writeText(u).then(function () { alert('Customer link copy ho gaya:\n' + u); });
    else prompt('Customer link copy karein:', u);
  });
  $('qrBtn').addEventListener('click', function () { window.open('qr-poster.html?shop=' + SHOP_ID, '_blank'); });
  $('posterBtn').addEventListener('click', function () { window.open('qr-poster.html?shop=' + SHOP_ID, '_blank'); });
  $('homeBtn').addEventListener('click', function () { location.href = 'index.html?shop=' + SHOP_ID; });
  $('duInput').addEventListener('change', function () { directUpload($('duInput').files); $('duInput').value = ''; });
  $('duBtn').addEventListener('click', function () { $('duInput').click(); });
  $('clearBtn').addEventListener('click', function () {
    if (!confirm('Saare Printed / Done / Cancelled orders delete kar dein?')) return;
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
  $('a4Modal').addEventListener('click', function (e) { if (e.target === $('a4Modal')) closeA4Modal(); });
});
})();
