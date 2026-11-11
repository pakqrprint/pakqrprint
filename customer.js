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
function rangeEligible(it) { return it.kind === 'pdf' || it.kind === 'word' || it.kind === 'excel' || it.kind === 'powerpoint'; } // v3.13: PPT range bhi (agent COM PrintOut From/To)
function parseRange(str, total) {
  var s = String(str || '').trim();
  if (!s) return null;
  var mm = /^(\d+)\s*-\s*(\d+)$/.exec(s);
  if (mm) { var a = parseInt(mm[1], 10), b = parseInt(mm[2], 10); if (a >= 1 && b >= a && b <= total) return { from: a, to: b, count: b - a + 1 }; return null; }
  if (/^\d+$/.test(s)) { var p = parseInt(s, 10); if (p >= 1 && p <= total) return { from: p, to: p, count: 1 }; }
  return null;
}
var ICO_COPIES = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';
var ICO_PRINT = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V2h12v7"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>';
var ICO_DUP = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 1l4 4-4 4"/><path d="M3 11V9a4 4 0 0 1 4-4h14"/><path d="M7 23l-4-4 4-4"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/></svg>';
function fileIcon(kind, name) {
  var k = kind;
  if (k !== 'image' && /\.(jpe?g|png|bmp|gif|webp)$/i.test(name || '')) k = 'image';
  if (k === 'image') return '<svg viewBox="0 0 48 48"><rect x="1" y="1" width="46" height="46" rx="10" fill="#dbeafe"/><circle cx="33" cy="14" r="5" fill="#f59e0b"/><path d="M9 38l11-13 8 9 6-7 7 11z" fill="#2563eb"/></svg>';
  var M = { word: ['#dbeafe', '#2b579a', 'W', 15], excel: ['#dcfce7', '#217346', 'X', 15], powerpoint: ['#ffedd5', '#d24726', 'P', 15], pdf: ['#fee2e2', '#f40f02', 'PDF', 9] }[k];
  if (!M) return '<svg viewBox="0 0 48 48"><rect x="1" y="1" width="46" height="46" rx="10" fill="#e0e7ff"/><path d="M14 6h14l8 8v28a2 2 0 0 1-2 2H14a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2z" fill="#fff" stroke="#4f46e5" stroke-width="2.5"/><path d="M28 6v8h8" fill="#fff" opacity="0.4"/></svg>';
  return '<svg viewBox="0 0 48 48"><rect x="1" y="1" width="46" height="46" rx="10" fill="' + M[0] + '"/><path d="M15 7h13l8 8v27a2 2 0 0 1-2 2H15a2 2 0 0 1-2-2V9a2 2 0 0 1 2-2z" fill="#fff" stroke="' + M[1] + '" stroke-width="2"/><path d="M28 7v8h8" fill="none" stroke="' + M[1] + '" stroke-width="2"/><rect x="9" y="24" width="15" height="15" rx="3" fill="' + M[1] + '"/><text x="16.5" y="35.5" text-anchor="middle" font-size="' + M[3] + '" font-weight="800" fill="#fff">' + M[2] + '</text><line x1="28" y1="20" x2="38" y2="20" stroke="' + M[1] + '" stroke-width="2" stroke-linecap="round"/><line x1="28" y1="26" x2="38" y2="26" stroke="' + M[1] + '" stroke-width="2" stroke-linecap="round"/><line x1="28" y1="32" x2="38" y2="32" stroke="' + M[1] + '" stroke-width="2" stroke-linecap="round"/></svg>';
}
function copiesOpts(v) {
  var h = '';
  for (var n = 1; n <= 20; n++) h += '<option value="' + n + '"' + (v === n ? ' selected' : '') + '>' + n + '</option>';
  if (v > 20) h += '<option value="' + v + '" selected>' + v + '</option>';
  return h;
}
var EYE_SVG = '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17a5 5 0 1 1 0-10 5 5 0 0 1 0 10zm0-8a3 3 0 1 0 0 6 3 3 0 0 0 0-6z"/></svg>';
function renderFiles() {
  var box = $('fileList'); box.innerHTML = '';
  selFiles.forEach(function (it, idx) {
    var d = document.createElement('div'); d.className = 'fileitem';
    var sizeKB = (it.file.size / 1024).toFixed(1);
    var h = '<div class="frow1"><div class="fic">' + fileIcon(it.kind, it.file.name) + '</div><div class="fmid">'
      + '<div class="fname2">' + esc(it.file.name) + '</div><div class="fmeta">';
    if (it.kind === 'image') {
      h += '<span>' + sizeKB + ' KB</span><span class="fdot">\u2022</span><span>Pages: 1</span><span class="fdot">\u2022</span><span>Copies: <select data-cp2="' + idx + '">' + copiesOpts(it.copies || 1) + '</select></span>';
    } else {
      h += '<span><span class="slabel">Size:</span> ' + sizeKB + ' KB</span>'
        + '<span><span class="plabel">Pages:</span> <input type="text" inputmode="numeric" value="' + it.pages + '" data-pg="' + idx + '"></span>'
        + '<span>Copies: <select data-cp2="' + idx + '">' + copiesOpts(it.copies || 1) + '</select></span>';
    }
    h += '</div></div><div class="fbtns">'
      + '<button type="button" class="eye" data-pv="' + idx + '" title="Preview">' + EYE_SVG + '</button>'
      + '<button type="button" class="x2" data-x="' + idx + '" title="Remove">\u2715</button>'
      + '</div></div>';
    var dxSel = '<span class="fctl">' + ICO_DUP + '<select class="dxsel" data-dx="' + idx + '">'
      + '<option value="single"' + (it.duplex ? '' : ' selected') + '>Single side</option>'
      + '<option value="long"' + (it.duplex === 'long' ? ' selected' : '') + '>Front/Back Portrait</option>'
      + '<option value="short"' + (it.duplex === 'short' ? ' selected' : '') + '>Front/Back Landscape</option></select></span>';
    h += '<div class="fdiv"></div><div class="frow2">'
      + '<span class="fctl">' + ICO_PRINT + '<label class="pslab">Per Sheet:</label><select data-nup="' + idx + '">'
      + (it.kind === 'powerpoint' ? [1, 2, 4, 6, 9] : [1, 2, 4, 6]).map(function (n) { return '<option value="' + n + '"' + (it.nup === n ? ' selected' : '') + '>Per Sheet ' + n + '</option>'; }).join('')
      + '</select></span>'
      + (rangeEligible(it)
        ? '<span class="fsep"></span><span class="frange-in"><label>Pages:</label><select data-rm="' + idx + '">'
          + '<option value="all"' + (it.rangeMode === 'all' ? ' selected' : '') + '>All Pages</option>'
          + '<option value="first"' + (it.rangeMode === 'first' ? ' selected' : '') + '>First Page Only</option>'
          + '<option value="some"' + (it.rangeMode === 'some' ? ' selected' : '') + '>Custom Range</option></select>'
          + (it.rangeMode === 'some' ? '<input type="text" placeholder="e.g. 2-5" value="' + esc(it.range || '') + '" data-rg="' + idx + '">' : '')
          + '<span class="fsep"></span>' + dxSel + '</span>'
        : '<span class="fsep"></span>' + dxSel)
      + '</div>';
    d.innerHTML = h;
    box.appendChild(d);
  });
  function syncCopies(idx, v) {
    var s = box.querySelector('[data-cp2="' + idx + '"]');
    if (s) { if (!s.querySelector('option[value="' + v + '"]')) { var o = document.createElement('option'); o.value = v; o.textContent = v; s.appendChild(o); } s.value = String(v); }
  }
  box.querySelectorAll('[data-pg]').forEach(function (el) { el.addEventListener('change', function () { var it = selFiles[+el.getAttribute('data-pg')]; var v = parseInt(el.value, 10); if (v > 0) { it.pages = v; it.pagesAuto = false; } renderFiles(); }); });
  box.querySelectorAll('[data-cp2]').forEach(function (el) { el.addEventListener('change', function () { var idx = +el.getAttribute('data-cp2'); var it = selFiles[idx]; var v = parseInt(el.value, 10); if (v > 0) { it.copies = v; syncCopies(idx, v); } }); });
  box.querySelectorAll('[data-nup]').forEach(function (el) { el.addEventListener('change', function () { var it = selFiles[+el.getAttribute('data-nup')]; it.nup = parseInt(el.value, 10) || 1; if (it.kind === 'powerpoint') it.spp = it.nup; }); });
  box.querySelectorAll('[data-dx]').forEach(function (el) { el.addEventListener('change', function () { var it = selFiles[+el.getAttribute('data-dx')]; it.duplex = (el.value === 'single' ? false : el.value); }); }); // v8: har file ki apni Front/Back
  box.querySelectorAll('[data-rm]').forEach(function (el) { el.addEventListener('change', function () { var it = selFiles[+el.getAttribute('data-rm')]; it.rangeMode = el.value; it.range = ''; it.rangePages = 0; renderFiles(); }); });
  box.querySelectorAll('[data-rg]').forEach(function (el) { el.addEventListener('change', function () { var it = selFiles[+el.getAttribute('data-rg')]; var r = parseRange(el.value, it.pages); if (r) { it.range = el.value; it.rangePages = r.count; } else { it.range = ''; it.rangePages = 0; el.value = ''; alert('Range is not correct - for example write 2-5, total pages are ' + it.pages + '.'); } }); });
  box.querySelectorAll('[data-x]').forEach(function (el) { el.addEventListener('click', function () { selFiles.splice(+el.getAttribute('data-x'), 1); renderFiles(); }); });
  box.querySelectorAll('[data-pv]').forEach(function (el) { el.addEventListener('click', function () { var it = selFiles[+el.getAttribute('data-pv')]; if (it && prevItem === it && $('prevModal') && $('prevModal').style.display !== 'none') { closePreview(); } else { openPreview(it); } }); });
  $('sendBtn').disabled = selFiles.length === 0;
}

var prevUrl = null;
var prevItem = null; // v3.31: the file whose preview is currently open in the modal (eye toggle)
function openPreview(it) {
  var m = $('prevModal'), body = $('prevBody'), title = $('prevTitle');
  if (!m || !body || !it) return;
  prevItem = it; // v3.31: remember which file the modal is showing (eye toggle)
  title.textContent = it.file.name;
  // v3.31: if this file's viewer iframe is already in the modal, just show the modal — do not
  // wipe it and recreate it (recreating reloads the viewer, which looked like a "refresh").
  // This check must sit BEFORE the "Loading..." reset below, or the iframe is already gone.
  if ((it.kind === 'word' || it.kind === 'excel' || it.kind === 'powerpoint') && it.previewURL) {
    var ex0 = body.querySelector('iframe');
    if (ex0 && ex0.getAttribute('data-pvsrc') === it.previewURL) { m.style.display = 'flex'; return; }
  }
  // v3.33: PDF reuse — pehle se render hue pages hon to dobara render nahi
  // karna; render chal raha ho to bhi sirf modal dikhayein (chain khud
  // body bhar degi). Ye checks bhi "Loading..." reset se PEHLE honi chahiye.
  if (it.kind === 'pdf') {
    if (it.pdfState === 'ready' && body.querySelector('.pdfpages')) { m.style.display = 'flex'; armPdfLazy(it); return; }
    if (it.pdfState === 'loading') { m.style.display = 'flex'; return; }
  }
  stopPdfLazy(); // v3.33: purani PDF ki lazy rendering rok dein (file badli / dobara khula)
  body.innerHTML = '<p class="hint">Loading preview...</p>';
  m.style.display = 'flex';
  if (prevUrl) { URL.revokeObjectURL(prevUrl); prevUrl = null; }
  if (it.kind === 'image') {
    prevUrl = URL.createObjectURL(it.file);
    body.innerHTML = '<img src="' + prevUrl + '" alt="preview">';
  } else if (it.kind === 'pdf') {
    // v3.33: ALL pages, stacked vertically — popup body scrolls up-down
    // (Word/Excel jaisa). Pehle 2 pages foran render hote hain; baqi pages
    // scroll karne par lazy render hote hain (armPdfLazy), taake lambi PDF
    // page ko freeze na kare. Dobara kholne par dobara render nahi hota.
    if (it.pdfState === 'loading') { return; } // pehle se render ho raha hai — chain khud body bhar degi
    it.pdfState = 'loading';
    var docP = it.pdfDoc ? Promise.resolve(it.pdfDoc) : it.file.arrayBuffer().then(function (buf) {
      return import('./pdfjs/pdf.mjs').then(function (pdfjsLib) {
        pdfjsLib.GlobalWorkerOptions.workerSrc = './pdfjs/pdf.worker.mjs';
        return pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
      });
    });
    docP.then(function (pdf) {
      if (prevItem !== it) return; // modal kisi aur file par chala gaya — is ka preview nahi likhna
      it.pdfDoc = pdf;
      it.pdfTotal = pdf.numPages || 1;
      it.pdfNext = 1;
      it.pdfGen = (it.pdfGen || 0) + 1; // nayi render generation — purani chain kabhi nahi likhegi
      var wrap = document.createElement('div');
      wrap.className = 'pdfpages';
      body.innerHTML = '';
      body.appendChild(wrap);
      it.pdfWrap = wrap;
      return renderPdfPages(it, 2); // pehle 2 pages foran
    }).then(function () {
      if (prevItem !== it) { it.pdfState = 'idle'; return; } // beech mein band ho gaya — dobara kholne par dobara render hoga
      it.pdfState = 'ready';
      armPdfLazy(it); // baqi pages scroll par render honge
    }).catch(function () {
      it.pdfState = 'failed';
      if (prevItem !== it) return;
      body.innerHTML = '<p class="hint">PDF preview did not open.</p>';
    });
  } else if (it.kind === 'word' || it.kind === 'excel' || it.kind === 'powerpoint') {
    // v3.31: Office preview in the SAME popup modal as PDF/image (not the inline pane).
    // First eye tap uploads once to Supabase previews/; Send reuses it (no double upload).
    if (it.previewURL) {
      body.innerHTML = '<iframe data-pvsrc="' + it.previewURL + '" src="https://view.officeapps.live.com/op/embed.aspx?src=' + encodeURIComponent(it.previewURL) + '" frameborder="0"></iframe>';
    } else if (it.pvState === 'converting') {
      body.innerHTML = '<p class="hint">Preview tayyar ho raha hai...</p>';
    } else {
      it.pvState = 'converting';
      body.innerHTML = '<p class="hint">Preview tayyar ho raha hai...</p>';
      var safeName = String(it.file.name).replace(/[^\w.\-() ]/g, '_');
      var previewPath = 'shops/' + SHOP_ID + '/previews/' + Date.now() + '_' + Math.floor(Math.random() * 100000) + '_' + safeName;
      supabaseUpload(it.file, previewPath).then(function (url) {
        it.previewURL = url; it.previewPath = previewPath; it.pvState = 'ready';
        if (prevItem !== it) return; // v3.31: modal moved on to another file — do not overwrite its preview
        body.innerHTML = '<iframe data-pvsrc="' + url + '" src="https://view.officeapps.live.com/op/embed.aspx?src=' + encodeURIComponent(url) + '" frameborder="0"></iframe>';
      }).catch(function () {
        it.pvState = 'failed';
        if (prevItem !== it) return; // v3.31: same guard as above
        body.innerHTML = '<p class="hint">Preview could not be loaded. The file will still be sent as it is.</p>';
      });
    }
  } else {
    body.innerHTML = '<p class="hint">No preview is available for this file type. The file will be sent as it is.</p>';
  }
}
var pdfObserver = null; // v3.33: lazy page rendering ke liye IntersectionObserver
function stopPdfLazy() { // v3.33: lazy rendering rok dein (modal band ya file badli)
  if (pdfObserver) { pdfObserver.disconnect(); pdfObserver = null; }
}
// Aik PDF page render kar ke wrap mein lagayein (sentinel se pehle).
// Page render fail ho to us page ko chhor kar agle par chalein.
function renderPdfPage(it, n) {
  var pdf = it.pdfDoc, wrap = it.pdfWrap, gen = it.pdfGen;
  if (!pdf || !wrap) return Promise.resolve(false);
  return pdf.getPage(n).then(function (page) {
    if (prevItem !== it || gen !== it.pdfGen || !wrap.isConnected) return false;
    var vp = page.getViewport({ scale: 1.5 }); // v3.33: scale waisa hi (1.5) — pages chhote nahi hue
    var cv = document.createElement('canvas');
    cv.width = vp.width; cv.height = vp.height;
    var sen = wrap.querySelector('.pdfmore');
    if (sen) wrap.insertBefore(cv, sen); else wrap.appendChild(cv);
    return page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise.then(function () { return true; });
  }).catch(function () { return false; });
}
// Pehle `count` pages aik ke baad aik render karein. Baqi pages hon to
// end par sentinel lagayein (armPdfLazy us par nazar rakhta hai).
function renderPdfPages(it, count) {
  var gen = it.pdfGen;
  var chain = Promise.resolve();
  var end = Math.min(it.pdfTotal, it.pdfNext + count - 1);
  while (it.pdfNext <= end) {
    (function (n) { chain = chain.then(function () { return renderPdfPage(it, n); }); })(it.pdfNext);
    it.pdfNext++;
  }
  return chain.then(function () {
    if (prevItem !== it || gen !== it.pdfGen || !it.pdfWrap) return;
    var sen = it.pdfWrap.querySelector('.pdfmore');
    if (it.pdfNext <= it.pdfTotal) {
      if (!sen) {
        sen = document.createElement('div');
        sen.className = 'pdfmore';
        sen.textContent = 'Loading more pages...';
        it.pdfWrap.appendChild(sen);
      }
    } else if (sen) { sen.remove(); }
  });
}
// Sentinel dikhte hi agla page render karein, phir dobara arm karein.
// Purana browser (IntersectionObserver nahi) ho to baqi pages aik ke
// baad aik render kar dein.
function armPdfLazy(it) {
  stopPdfLazy();
  if (!it || !it.pdfDoc || it.pdfNext > it.pdfTotal) return;
  var body = $('prevBody');
  var sen = it.pdfWrap && it.pdfWrap.querySelector('.pdfmore');
  if (!body || !sen) return;
  if (typeof IntersectionObserver === 'undefined') {
    renderPdfPages(it, it.pdfTotal - it.pdfNext + 1);
    return;
  }
  var gen = it.pdfGen;
  pdfObserver = new IntersectionObserver(function (entries) {
    if (!entries || !entries[0] || !entries[0].isIntersecting) return;
    stopPdfLazy();
    renderPdfPages(it, 1).then(function () { if (prevItem === it && gen === it.pdfGen) armPdfLazy(it); });
  }, { root: body, rootMargin: '300px' });
  pdfObserver.observe(sen);
}
function closePreview() {
  stopPdfLazy(); // v3.33: band modal mein pages render nahi hone chahiye
  var m = $('prevModal'); if (m) m.style.display = 'none';
  prevItem = null; // v3.31: modal closed — nothing is being previewed now (eye toggle). Only the
  // popup is hidden here — the file stays in the list and the page never reloads.
  if (prevUrl) { URL.revokeObjectURL(prevUrl); prevUrl = null; }
}

function addFiles(list) {
  var arr = Array.prototype.slice.call(list || []);
  arr.forEach(function (f) {
    if (f.size > MAX_MB * 1024 * 1024) { alert(f.name + ' is larger than ' + MAX_MB + 'MB - please send a smaller file.'); return; }
    if (!/\.(pdf|jpe?g|png|docx?|xlsx?|pptx?)$/i.test(f.name)) { alert(f.name + ' - this file type is not supported. Send PDF, Word, Excel, PowerPoint, JPG, PNG.'); return; }
    var kind = fileKind(f);
    var it = { file: f, kind: kind, pages: 1, pagesAuto: false, copies: (parseInt(($('copiesInput') || {}).value, 10) || 1), rangeMode: 'all', range: '', rangePages: 0, nup: 1, spp: 1, duplex: (function () { var v = (($('duplexAllSel') || {}).value || 'single'); return v === 'single' ? false : v; })(), pvState: 'idle', previewURL: '', previewPath: '', pdfState: 'idle', pdfDoc: null, pdfNext: 1, pdfTotal: 0, pdfWrap: null, pdfGen: 0 };
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
    im.onerror = function () { URL.revokeObjectURL(u); rej(new Error('image did not open')); };
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
  cx.fillText('Cut along the dashed line — both sides on one A4', cv.width / 2, 72);
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
    showErr('Setup is pending: your Firebase config is not filled in firebase-config.js yet. Paste the config as per Step 2 of the guide file (Pak-QR-Print-Hidayat-v3.txt), then reopen the page.');
    return false;
  }
  if (!SUPABASE_OK) {
    showErr('Setup is pending: the Supabase config is not filled in supabase-config.js yet. Paste the config as per Step 2 of the guide file, then reopen the page.');
    return false;
  }
  try {
    firebase.initializeApp(cfg);
    db = firebase.database();
    return true;
  } catch (e) {
    showErr('Firebase did not start: ' + e.message);
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
  var btn = $('sendBtn'); btn.disabled = true; btn.textContent = 'Sending...';
  var name = ($('custName').value || '').trim();
  var colorMode = (document.querySelector('input[name="printType"]:checked') || {}).value || 'BW';
  var paperSize = ($('paperSel') && $('paperSel').value) || 'a4';
  var orderKey = db.ref('shops/' + SHOP_ID + '/orders').push().key;
  var uploads = selFiles.map(function (it, idx) {
    var safeName = String(it.file.name).replace(/[^\w.\-() ]/g, '_');
    // v3.29: if this file was already uploaded for preview, reuse it.
    if (it.previewURL && it.previewPath) {
      return Promise.resolve({ name: it.file.name, kind: it.kind, size: it.file.size, pages: it.pages, copies: it.copies || 1, colorMode: colorMode, rangeMode: it.rangeMode, range: it.range || '', nup: it.nup || 1, spp: it.spp || 1, duplex: !!it.duplex, flipEdge: (it.duplex === 'short' ? 'short' : 'long'), paperSize: paperSize, downloadURL: it.previewURL, storagePath: it.previewPath });
    }
    var objectPath = 'shops/' + SHOP_ID + '/orders/' + orderKey + '/' + idx + '_' + safeName;
    return supabaseUpload(it.file, objectPath).then(function (url) {
      return { name: it.file.name, kind: it.kind, size: it.file.size, pages: it.pages, copies: it.copies || 1, colorMode: colorMode, rangeMode: it.rangeMode, range: it.range || '', nup: it.nup || 1, spp: it.spp || 1, duplex: !!it.duplex, flipEdge: (it.duplex === 'short' ? 'short' : 'long'), paperSize: paperSize, downloadURL: url, storagePath: objectPath };
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
    alert('Order could not be sent: ' + (e && e.message ? e.message : e) + '\nCheck the internet and try again.');
    btn.disabled = false; btn.textContent = 'Send Files to Admin';
  });
}

function sendIdOrder() {
  if (!idCompositeReady || !idCompositeBlob) { $('idMsg').textContent = 'First tap "Make ID Card A4 Composite".'; return; }
  var btn = $('idSendBtn'); btn.disabled = true; btn.textContent = 'Sending...';
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
    alert('ID Card could not be sent: ' + (e && e.message ? e.message : e));
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
  // v8: global Front/Back - sab files ki duplex aik saath set ho, rows dobara render ho kar nayi value dikhayein; us ke baad bhi har file ka alag toggle badla ja sakta hai
  var dxAll = $('duplexAllSel');
  if (dxAll) dxAll.addEventListener('change', function () {
    var v = (dxAll.value === 'single' ? false : dxAll.value);
    selFiles.forEach(function (it) { it.duplex = v; });
    renderFiles();
  });
  $('prevClose').addEventListener('click', closePreview);
  $('prevModal').addEventListener('click', function (e) { if (e.target === $('prevModal')) closePreview(); });
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
    if (!f || !bk) { msg.textContent = 'Choose photos of both the front and back.'; return; }
    msg.textContent = 'Making composite...';
    Promise.all([loadIdImage(f), loadIdImage(bk)]).then(function (ims) {
      drawIdCard(ims[0], ims[1], $('idType').value, idLayout);
      $('idCardCanvas').style.display = 'block';
      $('idCardCanvas').toBlob(function (blob) { idCompositeBlob = blob; idCompositeReady = true; msg.textContent = 'A4 Composite is ready — see the preview above, then tap "Send ID Card to Shop".'; }, 'image/jpeg', 0.92);
    }).catch(function () { msg.textContent = 'Photo did not open — try another photo.'; });
  });
  $('idSendBtn').addEventListener('click', sendIdOrder);
  $('againBtn').addEventListener('click', function () { location.reload(); });
  renderFiles();
});
})();
