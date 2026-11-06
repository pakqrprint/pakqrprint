// Pak QR Print - Customer page (GitHub Pages + Firebase + Supabase)
// Files seedha Supabase Storage par jati hain, order Realtime
// Database mein likha jata hai. Dukaan ka PC (print agent) wahan
// se uthakar print karta hai. Koi file is PC/server par nahi rehti.
(function () {
'use strict';
function $(id) { return document.getElementById(id); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>\"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function showErr(msg) { var e = $('setupErr'); e.textContent = msg; e.style.display = 'block'; }
function showOk(msg) { var e = $('setupOk'); e.textContent = msg; e.style.display = 'block'; }

var cfg = window.PAKQR_FIREBASE_CONFIG || {};
var CONFIG_OK = cfg.apiKey && cfg.apiKey.indexOf('YAHAN-') !== 0 && cfg.databaseURL && cfg.databaseURL.indexOf('YAHAN-') !== 0;
var sbcfg = window.PAKQR_SUPABASE_CONFIG || {};
var SUPABASE_OK = sbcfg.url && sbcfg.url.indexOf('YAHAN-') !== 0 && sbcfg.anonKey && sbcfg.anonKey.indexOf('YAHAN-') !== 0 && sbcfg.bucket;
var SHOP_ID = 'pakqrprint';
try {
  var m = /[?&]shop=([A-Za-z0-9_-]+)/.exec(location.search);
  if (m) SHOP_ID = m[1];
  else if (window.PAKQR_SHOP_ID) SHOP_ID = window.PAKQR_SHOP_ID;
} catch (e) {}

var DEFAULT_SETTINGS = { shopName: 'Pak QR Print', autoPrint: true, autoOpen: true, permissionMode: false, autoDeleteMinutes: 5 };
var settings = DEFAULT_SETTINGS;
var db = null;
var selFiles = []; // {file, kind, pages, pagesAuto, copies, rangeMode, range, rangePages, nup, spp}
var idLayout = 'stacked';
var idCompositeReady = false, idCompositeBlob = null;
var MAX_MB = 25;

function customerUrl() {
  if (window.PAKQR_CUSTOMER_URL) return window.PAKQR_CUSTOMER_URL.replace(/\/$/, '') + '/index.html?shop=' + SHOP_ID;
  return location.origin + location.pathname.replace(/[^\/]*$/, '') + 'index.html?shop=' + SHOP_ID;
}

// ---------- Supabase Storage (plain REST, koi CDN library nahi) ----------
// Upload: POST /storage/v1/object/<bucket>/<path> (anon key se).
// Public URL: <url>/storage/v1/object/public/<bucket>/<path>.
// Agar upload fail ho to error throw hota hai — order tab tak Realtime
// Database mein likha hi nahi jata, is liye order kabhi gum nahi hota.
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

// ---------- PDF pages (binary scan, v1 jaisa) ----------
function pdfPages(arrayBuffer) {
  try {
    var bytes = new Uint8Array(arrayBuffer), text = '', CH = 65536, i;
    for (i = 0; i < bytes.length; i += CH) text += String.fromCharCode.apply(null, bytes.subarray(i, Math.min(i + CH, bytes.length)));
    var mm = text.match(/\/Type\s*\/Page[^s]/g);
    if (mm && mm.length) return mm.length;
    var counts = text.match(/\/Count\s+(\d+)/g), max = 0;
    if (counts) counts.forEach(function (c) { var n = parseInt(c.replace(/[^\d]/g, ''), 10); if (n > max) max = n; });
    if (max > 0) return max;
  } catch (e) {}
  return 1;
}
function fileKind(f) {
  var n = (f && f.name) || '', t = (f && f.type) || '';
  if (/^image\//.test(t) || /\.(jpe?g|png)$/i.test(n)) return 'image';
  if (t === 'application/pdf' || /\.pdf$/i.test(n)) return 'pdf';
  if (/\.docx?$/i.test(n) || /wordprocessingml/.test(t)) return 'word';
  if (/\.xlsx?$/i.test(n) || /spreadsheetml/.test(t)) return 'excel';
  if (/\.pptx?$/i.test(n) || /presentationml/.test(t)) return 'powerpoint';
  return 'other';
}
// .docx/.pptx ke andar docProps/app.xml mein Pages/Slides count hota hai.
function officeAutoPages(file) {
  return file.arrayBuffer().then(function (buf) {
    var bytes = new Uint8Array(buf), i, found = -1;
    var needle = [100, 111, 99, 80, 114, 111, 112, 115, 47, 97, 112, 112, 46, 120, 109, 108]; // "docProps/app.xml"
    outer: for (i = 0; i < bytes.length - 16; i++) {
      for (var k = 0; k < 16; k++) if (bytes[i + k] !== needle[k]) continue outer;
      found = i; break;
    }
    if (found < 0) return null;
    // local header: nameLen @26, extraLen @28, phir data (stored ya deflate)
    var nameLen = bytes[found - 30 + 26] | (bytes[found - 30 + 27] << 8);
    var extraLen = bytes[found - 30 + 28] | (bytes[found - 30 + 29] << 8);
    var method = bytes[found - 30 + 8] | (bytes[found - 30 + 9] << 8);
    var compSize = (bytes[found - 30 + 18] | (bytes[found - 30 + 19] << 8) | (bytes[found - 30 + 20] << 16) | (bytes[found - 30 + 21] << 24)) >>> 0;
    var dataStart = found + nameLen + extraLen;
    var raw = bytes.subarray(dataStart, dataStart + compSize);
    function parseXmlText(txt) {
      var mm = txt.match(/<(?:Pages|Slides)>(\d+)<\/(?:Pages|Slides)>/);
      return mm ? parseInt(mm[1], 10) : null;
    }
    if (method === 0) return parseXmlText(new TextDecoder().decode(raw));
    if (typeof DecompressionStream !== 'undefined') {
      return new Response(new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text().then(parseXmlText).catch(function () { return null; });
    }
    return null;
  }).catch(function () { return null; });
}
function rangeEligible(it) { return it.kind === 'pdf' || it.kind === 'word' || it.kind === 'excel'; }
function parseRange(str, total) {
  var s = String(str || '').trim();
  if (!s) return null;
  var mm = /^(\d+)\s*-\s*(\d+)$/.exec(s);
  if (mm) { var a = parseInt(mm[1], 10), b = parseInt(mm[2], 10); if (a >= 1 && b >= a && b <= total) return { from: a, to: b, count: b - a + 1 }; return null; }
  if (/^\d+$/.test(s)) { var p = parseInt(s, 10); if (p >= 1 && p <= total) return { from: p, to: p, count: 1 }; }
  return null;
}
function renderFiles() {
  var box = $('fileList'); box.innerHTML = '';
  selFiles.forEach(function (it, idx) {
    var d = document.createElement('div'); d.className = 'fileitem';
    var h = '<div class="fname">' + esc(it.file.name) + '</div><div>' + esc(it.kind.toUpperCase()) + ' • ' + (it.file.size / 1024).toFixed(1) + ' KB • Pages: ';
    h += '<input type="number" min="1" value="' + it.pages + '" data-pg="' + idx + '" style="width:64px"> ';
    h += it.pagesAuto ? '(khud gine gaye)' : '(aap ne likhe)';
    h += '</div>';
    if (rangeEligible(it)) {
      h += '<div class="frow"><label>Pages:</label><select data-rm="' + idx + '">'
        + '<option value="all"' + (it.rangeMode === 'all' ? ' selected' : '') + '>All Pages</option>'
        + '<option value="first"' + (it.rangeMode === 'first' ? ' selected' : '') + '>First Page Only</option>'
        + '<option value="some"' + (it.rangeMode === 'some' ? ' selected' : '') + '>Custom Page Range</option></select>'
        + (it.rangeMode === 'some' ? '<input type="text" placeholder="masalan 2-5" value="' + esc(it.range || '') + '" data-rg="' + idx + '" style="width:90px">' : '')
        + '</div>';
    }
    h += '<div class="frow"><label>Copies:</label><input type="number" min="1" max="99" value="' + (it.copies || 1) + '" data-cp="' + idx + '" style="width:64px">'
      + '<label>Aik sheet par kitne:</label><select data-nup="' + idx + '">'
      + [1, 2, 4, 6].map(function (n) { return '<option value="' + n + '"' + (it.nup === n ? ' selected' : '') + '>' + n + '</option>'; }).join('')
      + '</select>'
      + (it.kind === 'powerpoint' ? '<label>Slides/page:</label><select data-spp="' + idx + '">' + [1, 2, 4, 6, 9].map(function (n) { return '<option value="' + n + '"' + (it.spp === n ? ' selected' : '') + '>' + n + '</option>'; }).join('') + '</select>' : '')
      + '<button type="button" class="x" data-x="' + idx + '">Hatayein</button></div>';
    d.innerHTML = h;
    box.appendChild(d);
  });
  box.querySelectorAll('[data-pg]').forEach(function (el) { el.addEventListener('change', function () { var it = selFiles[+el.getAttribute('data-pg')]; var v = parseInt(el.value, 10); if (v > 0) { it.pages = v; it.pagesAuto = false; } renderFiles(); }); });
  box.querySelectorAll('[data-cp]').forEach(function (el) { el.addEventListener('change', function () { var it = selFiles[+el.getAttribute('data-cp')]; var v = parseInt(el.value, 10); if (v > 0) it.copies = v; }); });
  box.querySelectorAll('[data-nup]').forEach(function (el) { el.addEventListener('change', function () { selFiles[+el.getAttribute('data-nup')].nup = parseInt(el.value, 10) || 1; }); });
  box.querySelectorAll('[data-spp]').forEach(function (el) { el.addEventListener('change', function () { selFiles[+el.getAttribute('data-spp')].spp = parseInt(el.value, 10) || 1; }); });
  box.querySelectorAll('[data-rm]').forEach(function (el) { el.addEventListener('change', function () { var it = selFiles[+el.getAttribute('data-rm')]; it.rangeMode = el.value; it.range = ''; it.rangePages = 0; renderFiles(); }); });
  box.querySelectorAll('[data-rg]').forEach(function (el) { el.addEventListener('change', function () { var it = selFiles[+el.getAttribute('data-rg')]; var r = parseRange(el.value, it.pages); if (r) { it.range = el.value; it.rangePages = r.count; } else { it.range = ''; it.rangePages = 0; el.value = ''; alert('Range theek nahi hai - masalan 2-5 likhein, kul pages ' + it.pages + ' hain.'); } }); });
  box.querySelectorAll('[data-x]').forEach(function (el) { el.addEventListener('click', function () { selFiles.splice(+el.getAttribute('data-x'), 1); renderFiles(); }); });
  $('sendBtn').disabled = selFiles.length === 0;
}

function addFiles(list) {
  var arr = Array.prototype.slice.call(list || []);
  arr.forEach(function (f) {
    if (f.size > MAX_MB * 1024 * 1024) { alert(f.name + ' ' + MAX_MB + 'MB se bara hai - chhoti file bhejein.'); return; }
    if (!/\.(pdf|jpe?g|png|docx?|xlsx?|pptx?)$/i.test(f.name)) { alert(f.name + ' - ye file type support nahi hai. PDF, Word, Excel, PowerPoint, JPG, PNG bhejein.'); return; }
    var kind = fileKind(f);
    var it = { file: f, kind: kind, pages: 1, pagesAuto: false, copies: (parseInt(($('copiesInput') || {}).value, 10) || 1), rangeMode: 'all', range: '', rangePages: 0, nup: 1, spp: 1 };
    selFiles.push(it);
    if (kind === 'pdf') {
      f.arrayBuffer().then(function (buf) { it.pages = pdfPages(buf) || 1; it.pagesAuto = true; renderFiles(); });
    } else if (kind === 'word' || kind === 'powerpoint') {
      officeAutoPages(f).then(function (n) { if (n) { it.pages = n; it.pagesAuto = true; } renderFiles(); });
    }
  });
  renderFiles();
}

// ---------- ID Card 2-in-1 ----------
function loadIdImage(file) {
  return new Promise(function (res, rej) {
    var u = URL.createObjectURL(file); var im = new Image();
    im.onload = function () { URL.revokeObjectURL(u); res(im); };
    im.onerror = function () { URL.revokeObjectURL(u); rej(new Error('image nahi khuli')); };
    im.src = u;
  });
}
function drawIdOne(cx, im, x, y, w, h, label) {
  cx.save();
  cx.strokeStyle = '#1e1b4b'; cx.lineWidth = 2; cx.setLineDash([]);
  cx.drawImage(im, x, y, w, h);
  cx.strokeRect(x, y, w, h);
  cx.setLineDash([6, 4]); cx.strokeStyle = '#dc2626'; cx.lineWidth = 1.5;
  cx.strokeRect(x - 6, y - 6, w + 12, h + 12);
  cx.setLineDash([]);
  cx.fillStyle = '#1e1b4b'; cx.font = 'bold 15px Arial'; cx.textAlign = 'left';
  cx.fillText(label, x, y - 12);
  cx.restore();
}
function drawIdCard(front, back, type, layout) {
  var cv = $('idCardCanvas'); var cx = cv.getContext('2d');
  cx.setLineDash([]);
  cx.fillStyle = '#ffffff'; cx.fillRect(0, 0, cv.width, cv.height);
  cx.fillStyle = '#1e1b4b'; cx.font = 'bold 26px Arial'; cx.textAlign = 'center';
  cx.fillText(type + ' — Front + Back (A4)', cv.width / 2, 46);
  cx.font = '16px Arial'; cx.fillStyle = '#5b5f8a';
  cx.fillText('Kaichi wali lakeer se kaatein — dono taraf aik hi A4 par', cv.width / 2, 72);
  if (layout === 'side') {
    drawIdOne(cx, front, 28, 420, 340, 214, 'Front');
    drawIdOne(cx, back, 426, 420, 340, 214, 'Back');
  } else {
    drawIdOne(cx, front, 77, 110, 640, 403, 'Front');
    drawIdOne(cx, back, 77, 583, 640, 403, 'Back');
  }
}

// ---------- Firebase (queue + login) + Supabase (files) ----------
function initFirebase() {
  if (!CONFIG_OK) {
    showErr('Setup baaki hai: firebase-config.js mein apni Firebase config abhi bhari nahi gayi. Hidayat file (Pak-QR-Print-Hidayat-v2-Supabase.txt) ke Step 2 ke mutabiq config paste karein, phir page dobara kholein.');
    return false;
  }
  if (!SUPABASE_OK) {
    showErr('Setup baaki hai: supabase-config.js mein Supabase ki config abhi bhari nahi gayi. Hidayat file ke Step 2 ke mutabiq config paste karein, phir page dobara kholein.');
    return false;
  }
  try {
    firebase.initializeApp(cfg);
    db = firebase.database();
    return true;
  } catch (e) {
    showErr('Firebase shuru nahi hua: ' + e.message);
    return false;
  }
}
function loadSettings() {
  return db.ref('shops/' + SHOP_ID + '/settings').once('value').then(function (snap) {
    var s = snap.val();
    if (s) settings = Object.assign({}, DEFAULT_SETTINGS, s);
    document.querySelectorAll('[data-shopname]').forEach(function (el) { el.textContent = settings.shopName || 'Pak QR Print'; });
  }).catch(function () { /* defaults */ });
}
function confetti() {
  var c = document.createElement('div'); c.className = 'confetti';
  var colors = ['#4f46e5', '#22c55e', '#f59e0b', '#ec4899', '#06b6d4'];
  for (var i = 0; i < 40; i++) {
    var p = document.createElement('i');
    p.style.left = (Math.random() * 100) + 'vw';
    p.style.background = colors[i % colors.length];
    p.style.animationDelay = (Math.random() * 0.6) + 's';
    c.appendChild(p);
  }
  document.body.appendChild(c);
  setTimeout(function () { c.remove(); }, 3200);
}

function sendOrder() {
  if (!selFiles.length) return;
  var btn = $('sendBtn'); btn.disabled = true; btn.textContent = 'Bhej rahe hain...';
  var name = ($('custName').value || '').trim();
  var colorMode = (document.querySelector('input[name="printType"]:checked') || {}).value || 'BW';
  var paperSize = ($('paperSel') && $('paperSel').value) || 'a4';
  var orderKey = db.ref('shops/' + SHOP_ID + '/orders').push().key;
  var uploads = selFiles.map(function (it, idx) {
    var safeName = String(it.file.name).replace(/[^\w.\-() ]/g, '_');
    var objectPath = 'shops/' + SHOP_ID + '/orders/' + orderKey + '/' + idx + '_' + safeName;
    return supabaseUpload(it.file, objectPath).then(function (url) {
      return { name: it.file.name, kind: it.kind, size: it.file.size, pages: it.pages, copies: it.copies || 1, colorMode: colorMode, rangeMode: it.rangeMode, range: it.range || '', nup: it.nup || 1, spp: it.spp || 1, paperSize: paperSize, downloadURL: url, storagePath: objectPath };
    });
  });
  Promise.all(uploads).then(function (files) {
    var counterRef = db.ref('shops/' + SHOP_ID + '/counter');
    return counterRef.transaction(function (cur) { return (cur || 1000) + 1; }).then(function (res) {
      var n = res.snapshot.val() || 1001;
      var orderNo = 'QP-' + n;
      var order = { orderNo: orderNo, token: orderNo, customerName: name || 'Walk-in', createdAt: firebase.database.ServerValue.TIMESTAMP, status: 'New', colorMode: colorMode, copies: 1, paperSize: paperSize, source: 'customer', files: files, printLog: [] };
      return db.ref('shops/' + SHOP_ID + '/orders/' + orderKey).set(order).then(function () { return orderNo; });
    });
  }).then(function (orderNo) {
    $('formBox').style.display = 'none';
    var s = $('successBox'); s.style.display = 'block';
    $('tokenNo').textContent = orderNo;
    confetti();
  }).catch(function (e) {
    alert('Order nahi bheja ja saka: ' + (e && e.message ? e.message : e) + '\nInternet check karein aur dobara try karein.');
    btn.disabled = false; btn.textContent = 'Send Files to Shopkeeper';
  });
}

function sendIdOrder() {
  if (!idCompositeReady || !idCompositeBlob) { $('idMsg').textContent = 'Pehle "ID Card A4 Composite Banayein" dabayein.'; return; }
  var btn = $('idSendBtn'); btn.disabled = true; btn.textContent = 'Bhej rahe hain...';
  var name = ($('custName').value || '').trim();
  var type = $('idType').value;
  var orderKey = db.ref('shops/' + SHOP_ID + '/orders').push().key;
  var file = new File([idCompositeBlob], 'ID_Card_A4_Composite.jpg', { type: 'image/jpeg' });
  var objectPath = 'shops/' + SHOP_ID + '/orders/' + orderKey + '/ID_Card_A4_Composite.jpg';
  supabaseUpload(file, objectPath).then(function (url) {
    var counterRef = db.ref('shops/' + SHOP_ID + '/counter');
    return counterRef.transaction(function (cur) { return (cur || 1000) + 1; }).then(function (res) {
      var n = res.snapshot.val() || 1001;
      var orderNo = 'QP-' + n;
      var order = { orderNo: orderNo, token: orderNo, customerName: name || 'Walk-in', createdAt: firebase.database.ServerValue.TIMESTAMP, status: 'New', colorMode: 'Color', copies: 1, paperSize: 'a4', source: 'idcard', idType: type, idLayout: idLayout, files: [{ name: 'ID_Card_A4_Composite.jpg', kind: 'idcard', size: file.size, pages: 1, copies: 1, colorMode: 'Color', rangeMode: 'all', range: '', nup: 1, spp: 1, paperSize: 'a4', downloadURL: url, storagePath: objectPath }], printLog: [] };
      return db.ref('shops/' + SHOP_ID + '/orders/' + orderKey).set(order).then(function () { return orderNo; });
    });
  }).then(function (orderNo) {
    $('formBox').style.display = 'none';
    var s = $('successBox'); s.style.display = 'block';
    $('tokenNo').textContent = orderNo;
    confetti();
  }).catch(function (e) {
    alert('ID Card nahi bheja ja saka: ' + (e && e.message ? e.message : e));
    btn.disabled = false; btn.textContent = 'Send ID Card to Shop';
  });
}

function segInit(id, fn) {
  var box = $(id); if (!box) return;
  box.querySelectorAll('button').forEach(function (b) {
    b.addEventListener('click', function () {
      box.querySelectorAll('button').forEach(function (x) { x.classList.remove('on'); });
      b.classList.add('on');
      fn(b.getAttribute('data-v'));
    });
  });
}

document.addEventListener('DOMContentLoaded', function () {
  if (!initFirebase()) return;
  loadSettings();
  // tabs
  document.querySelectorAll('#custTabs button').forEach(function (b) {
    b.addEventListener('click', function () {
      document.querySelectorAll('#custTabs button').forEach(function (x) { x.classList.remove('on'); });
      b.classList.add('on');
      var t = b.getAttribute('data-ctab');
      $('stdBox').style.display = t === 'std' ? 'block' : 'none';
      $('idCardBox').style.display = t === 'idcard' ? 'block' : 'none';
    });
  });
  // dropzone
  var dz = $('dropZone'), fi = $('fileInput');
  dz.addEventListener('click', function () { fi.click(); });
  fi.addEventListener('change', function () { addFiles(fi.files); fi.value = ''; });
  dz.addEventListener('dragover', function (e) { e.preventDefault(); dz.classList.add('over'); });
  dz.addEventListener('dragleave', function () { dz.classList.remove('over'); });
  dz.addEventListener('drop', function (e) { e.preventDefault(); dz.classList.remove('over'); addFiles(e.dataTransfer.files); });
  $('sendBtn').addEventListener('click', sendOrder);
  // ID card
  segInit('segIdLayout', function (v) { idLayout = v; idCompositeReady = false; });
  [['idFront', 'frontBox', 'Choose Front'], ['idBack', 'backBox', 'Choose Back']].forEach(function (pair) {
    var el = $(pair[0]); if (!el) return;
    el.addEventListener('change', function () {
      idCompositeReady = false;
      var box = $(pair[1]);
      if (box) {
        var f = el.files && el.files[0];
        var b = box.querySelector('b');
        if (f) { box.classList.add('has'); if (b) b.textContent = f.name; }
        else { box.classList.remove('has'); if (b) b.textContent = pair[2]; }
      }
    });
  });
  $('idMakeBtn').addEventListener('click', function () {
    var msg = $('idMsg');
    var f = $('idFront').files && $('idFront').files[0];
    var bk = $('idBack').files && $('idBack').files[0];
    if (!f || !bk) { msg.textContent = 'Front aur back dono ki photo choose karein.'; return; }
    msg.textContent = 'Composite ban raha hai...';
    Promise.all([loadIdImage(f), loadIdImage(bk)]).then(function (ims) {
      drawIdCard(ims[0], ims[1], $('idType').value, idLayout);
      $('idCardCanvas').style.display = 'block';
      $('idCardCanvas').toBlob(function (blob) { idCompositeBlob = blob; idCompositeReady = true; msg.textContent = 'A4 Composite tayyar hai — upar preview dekhein, phir "Send ID Card to Shop" dabayein.'; }, 'image/jpeg', 0.92);
    }).catch(function () { msg.textContent = 'Photo nahi khuli — doosri photo try karein.'; });
  });
  $('idSendBtn').addEventListener('click', sendIdOrder);
  $('againBtn').addEventListener('click', function () { location.reload(); });
  renderFiles();
});
})();
