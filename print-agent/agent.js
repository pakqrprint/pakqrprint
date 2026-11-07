// Pak QR Print - Print Agent (dukaan ke PC ke liye)
// ============================================================
// Ye program dukaan ke Windows PC par chalta hai. Kaam:
//   1. Firebase Realtime Database mein naye orders dekhta hai
//   2. File Supabase Storage se (public URL se) download karta hai
//   3. Print engine (printEngine.js - offline v1 ka azmaya hua
//      engine) se print karta hai: SumatraPDF (PDF), Word COM
//      (Word), Shell verb (Excel / PowerPoint 1-slide), PowerPoint
//      COM (slides-per-page), image sheets + Document N-up.
//   4. Status wapas Firebase par likhta hai (Printed / PrintFailed)
//   5. Auto-Delete: purane Done/Cancelled orders (Realtime Database
//      se) saaf karta hai. NOTE: Supabase ki file agent delete nahi
//      karta — anon key se delete nahi hota (policy: SELECT+INSERT
//      only). File bucket mein rehti hai; Supabase dashboard se
//      waqt par saaf karni hogi (Hidayat dekhein).
// Sirf Node.js built-in modules - koi npm install nahi.
// Windows par hi print karta hai; kisi aur OS par queue dekhta
// hai aur saaf batata hai ke print sirf Windows par hoga.
// ============================================================
'use strict';
const fs = require('fs');
const path = require('path');
const https = require('https');
const { printOneFile, tryDocNupForFile, SUMATRA_FILE, groupImageSheets, printViaImageSheet } = require('./printEngine.js'); // v3.16: photo sheet

const ROOT = __dirname;
const CONFIG_FILE = path.join(ROOT, 'config.json');
const DOWNLOAD_DIR = path.join(ROOT, 'downloads');
const LOG_FILE = path.join(ROOT, 'agent-log.txt');

function log(msg) {
  const line = new Date().toISOString() + '  ' + msg;
  console.log(line);
  try { fs.appendFileSync(LOG_FILE, line + '\n'); } catch (e) {}
}
function loadConfig() {
  if (!fs.existsSync(CONFIG_FILE)) {
    console.error('config.json nahi mili! Pehle config.json bharein (Hidayat Step 3 dekhein).');
    process.exit(1);
  }
  const c = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
  const missing = [];
  ['databaseURL', 'apiKey', 'shopId', 'email', 'password'].forEach(function (k) { if (!c[k] || String(c[k]).indexOf('YAHAN-') === 0) missing.push(k); });
  if (missing.length) {
    console.error('config.json mein ye khaali hain: ' + missing.join(', ') + ' - Hidayat Step 3 ke mutabiq bharein.');
    process.exit(1);
  }
  c.pollSeconds = c.pollSeconds || 5;
  return c;
}
const CFG = loadConfig();
if (!fs.existsSync(DOWNLOAD_DIR)) fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });

function httpsJson(method, urlStr, body, headers) {
  return new Promise(function (resolve, reject) {
    const u = new URL(urlStr);
    const data = body ? JSON.stringify(body) : null;
    const req = https.request({
      method: method, hostname: u.hostname, path: u.pathname + u.search,
      headers: Object.assign({ 'Content-Type': 'application/json' }, headers || {}, data ? { 'Content-Length': Buffer.byteLength(data) } : {})
    }, function (res) {
      const chunks = [];
      res.on('data', function (c) { chunks.push(c); });
      res.on('end', function () {
        const text = Buffer.concat(chunks).toString('utf8');
        let j = null;
        try { j = JSON.parse(text); } catch (e) { j = null; }
        if (res.statusCode >= 200 && res.statusCode < 300) resolve(j);
        else reject(new Error('HTTP ' + res.statusCode + ': ' + text.slice(0, 300)));
      });
    });
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });
}
function httpsDownload(urlStr, destPath) {
  return new Promise(function (resolve, reject) {
    const file = fs.createWriteStream(destPath);
    https.get(urlStr, function (res) {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        https.get(res.headers.location, function (r2) { r2.pipe(file); file.on('finish', function () { file.close(resolve); }); }).on('error', reject);
        return;
      }
      if (res.statusCode !== 200) { reject(new Error('Download HTTP ' + res.statusCode)); return; }
      res.pipe(file);
      file.on('finish', function () { file.close(resolve); });
    }).on('error', reject);
  });
}

// ---------- Firebase Auth (REST) ----------
let idToken = '', refreshToken = '', tokenExp = 0;
async function signIn() {
  const url = 'https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=' + encodeURIComponent(CFG.apiKey);
  const j = await httpsJson('POST', url, { email: CFG.email, password: CFG.password, returnSecureToken: true });
  idToken = j.idToken; refreshToken = j.refreshToken;
  tokenExp = Date.now() + (parseInt(j.expiresIn, 10) - 60) * 1000;
  log('Firebase login ho gaya (shopkeeper: ' + CFG.email + ')');
}
async function refreshIdToken() {
  const j = await httpsJson('POST', 'https://securetoken.googleapis.com/v1/token?key=' + encodeURIComponent(CFG.apiKey), { grant_type: 'refresh_token', refresh_token: refreshToken });
  idToken = j.id_token; refreshToken = j.refresh_token;
  tokenExp = Date.now() + (parseInt(j.expires_in, 10) - 60) * 1000;
}
async function ensureAuth() {
  if (!idToken) await signIn();
  else if (Date.now() > tokenExp) await refreshIdToken();
}
function dbUrl(p) {
  return CFG.databaseURL.replace(/\/$/, '') + p + '.json?auth=' + encodeURIComponent(idToken);
}
async function dbGet(p) { await ensureAuth(); return httpsJson('GET', dbUrl(p), null); }
async function dbPatch(p, obj) { await ensureAuth(); return httpsJson('PATCH', dbUrl(p), obj); }
async function dbPut(p, obj) { await ensureAuth(); return httpsJson('PUT', dbUrl(p), obj); }
async function dbDelete(p) { await ensureAuth(); return httpsJson('DELETE', dbUrl(p), null); }
// ---------- Supabase file delete ----------
// Agent Supabase ki file delete NAHI karta: anon key se delete nahi
// hota (bucket policy: SELECT + INSERT only) aur agent mein koi
// Supabase key nahi hoti. Order Realtime Database se saaf hota hai;
// file bucket mein rehti hai — Supabase dashboard se saaf karein.
async function storageDelete(storagePath) { return; }

// ---------- Print ----------
function safeName(n) { return String(n || 'file').replace(/[\\/:*?"<>|]/g, '_'); }
// Phase 3.16: photo sheet fail ho ya member download na ho to har photo
// alag print karein (poore page par aik) - order kabhi fail nahi.
async function printPhotosOneByOne(grpFiles, orderKey, baseOpts, printLog) {
  let okAll = true;
  for (const gf of grpFiles) {
    const gd = path.join(DOWNLOAD_DIR, orderKey + '_' + safeName(gf.name));
    try {
      if (!gf.downloadURL) throw new Error('File ka download link nahi mila');
      await httpsDownload(gf.downloadURL, gd);
      const gopts = Object.assign({}, baseOpts, { copies: gf.copies || 1 });
      const gr = await printOneFile(gd, gopts);
      printLog.push({ file: gf.name, ok: !!gr.ok, method: gr.method || '', error: gr.ok ? '' : (gr.error || 'Print fail'), at: Date.now() });
      log('  ' + gf.name + ': ' + (gr.ok ? 'OK (' + (gr.method || '') + ')' : 'FAIL - ' + (gr.error || '')));
      if (!gr.ok) okAll = false;
    } catch (e) {
      okAll = false;
      printLog.push({ file: gf.name, ok: false, method: '', error: String((e && e.message) || e), at: Date.now() });
      log('  ' + gf.name + ': FAIL - ' + String((e && e.message) || e));
    }
  }
  return okAll;
}

async function printOrder(orderKey, order, settings) {
  log('Order ' + (order.orderNo || orderKey) + ' print ho raha hai (' + (order.files || []).length + ' file)...');
  await dbPatch('/shops/' + CFG.shopId + '/orders/' + orderKey, { status: 'Printing' });
  const printLog = [];
  let allOk = true;
  // Phase 3.16: photo sheet - order ki photos (pehli photo ki "Per sheet"
  // value se) sheet par lagti hain (4/6 aik page par). ID Card order
  // (source idcard / kind idcard) mein grouping NAHI - composite akela
  // hi print hota hai. groupImageSheets engine mein pehle se tha (3.19)
  // lekin kabhi call hi nahi hota tha. NOTE: saari photos ki Per sheet
  // aik jaisi rakhein - pehli photo wali lagegi.
  const orderFiles = order.files || [];
  const firstPhoto = orderFiles.find((x) => /\.(jpe?g|png|bmp|gif)$/i.test(x.name || '') && x.kind !== 'idcard');
  const isCnicOrder = (order.source === 'idcard') || orderFiles.some((x) => x.kind === 'idcard');
  const photoPpp = firstPhoto ? (parseInt(firstPhoto.nup, 10) || 1) : 1;
  const photoGroups = groupImageSheets(orderFiles.map((x) => x.name), photoPpp, isCnicOrder);
  for (const f of (order.files || [])) {
    const dest = path.join(DOWNLOAD_DIR, orderKey + '_' + safeName(f.name));
    // Phase 3.16: sheet member (head nahi) - pehle hi sheet ke saath print ho chuki, skip.
    if (photoGroups.sheetMember[f.name] && !photoGroups.sheetHead[f.name]) continue;
    try {
      if (!f.downloadURL) throw new Error('File ka download link nahi mila');
      await httpsDownload(f.downloadURL, dest);
      const opts = {
        copies: f.copies || 1,
        colorMode: f.colorMode || order.colorMode || 'BW',
        printType: ((f.colorMode || order.colorMode || 'BW') === 'Color' ? 'Colour' : 'BW'), // v3.16: engine -Mono/color isi se chalta hai (pehle kabhi set hi nahi hota tha - rang wali photo bhi BW print hoti thi; BW wale orders pehle jaisi hi rahenge)
        duplex: false,
        paperSize: (f.paperSize || order.paperSize || 'A4'),
        range: (f.rangeMode === 'some' ? (f.range || '') : (f.rangeMode === 'first' ? '1-1' : '')),
        totalPages: f.pages || 0,
        slidesPerPage: f.spp || 1,
        nup: f.nup || 1,
        printerName: settings.printerName || '',
        orderPrinter: '',
        printerMismatch: false
      };
      // Phase 3.16: photo sheet - pehli photo par poori sheet print hoti hai.
      if (photoGroups.sheetHead[f.name]) {
        const grpNames = photoGroups.sheetHead[f.name];
        const grpFiles = grpNames.map((n) => (order.files || []).find((x) => x.name === n)).filter(Boolean);
        if (grpFiles.length <= 1) {
          // aik photo - sheet nahi, poore page par aik photo.
          const r0 = await printOneFile(dest, opts);
          printLog.push({ file: f.name, ok: !!r0.ok, method: r0.method || '', error: r0.ok ? '' : (r0.error || 'Print fail'), at: Date.now() });
          log('  ' + f.name + ': ' + (r0.ok ? 'OK (' + (r0.method || '') + ')' : 'FAIL - ' + (r0.error || '')));
          if (!r0.ok) allOk = false;
          continue;
        }
        // baqi members download karein (pehli pehle se downloaded hai).
        const dests = [dest];
        let dlFail = null;
        for (let gi = 1; gi < grpFiles.length; gi++) {
          const gf = grpFiles[gi];
          const gd = path.join(DOWNLOAD_DIR, orderKey + '_' + safeName(gf.name));
          try {
            if (!gf.downloadURL) throw new Error('File ka download link nahi mila');
            await httpsDownload(gf.downloadURL, gd);
            dests.push(gd);
          } catch (e) { dlFail = e; break; }
        }
        if (dlFail) {
          log('  Photo sheet (' + grpNames.join(', ') + ') - member download fail, alag alag print: ' + String((dlFail && dlFail.message) || dlFail));
          if (!await printPhotosOneByOne(grpFiles, orderKey, opts, printLog)) allOk = false;
          continue;
        }
        const sheetOpts = Object.assign({}, opts, { picturesPerPage: photoPpp, copies: f.copies || 1 });
        const sr = await printViaImageSheet(dests, sheetOpts);
        if (!sr.ok) {
          log('  Photo sheet fail (' + grpNames.join(', ') + ') - alag alag print: ' + (sr.error || ''));
          if (!await printPhotosOneByOne(grpFiles, orderKey, opts, printLog)) allOk = false;
          continue;
        }
        for (const gf of grpFiles) {
          printLog.push({ file: gf.name, ok: true, method: sr.method || '', error: '', at: Date.now() });
          log('  ' + gf.name + ': OK (' + (sr.method || '') + ')');
        }
        continue;
      }
      let r;
      if ((opts.nup || 1) > 1 && /(\.pdf|\.docx?|\.xlsx?)$/i.test(f.name)) {
        r = await tryDocNupForFile(dest, f.name, opts);
        if (r && r.nupFailed) { log('  N-up fail (' + f.name + ') - purane raste se print kar rahe hain: ' + (r.error || '')); r = await printOneFile(dest, opts); }
      } else {
        r = await printOneFile(dest, opts);
      }
      printLog.push({ file: f.name, ok: !!r.ok, method: r.method || '', error: r.ok ? '' : (r.error || 'Print fail'), at: Date.now() });
      log('  ' + f.name + ': ' + (r.ok ? 'OK (' + (r.method || '') + ')' : 'FAIL - ' + (r.error || '')));
      if (!r.ok) allOk = false;
    } catch (e) {
      allOk = false;
      printLog.push({ file: f.name, ok: false, method: '', error: String((e && e.message) || e), at: Date.now() });
      log('  ' + f.name + ': FAIL - ' + String((e && e.message) || e));
      // Phase 3.16: head fail ho to us ki sheet ke members bhi fail (warna wo bina log ke skip ho jayenge).
      if (photoGroups.sheetHead[f.name]) {
        for (const mn of photoGroups.sheetHead[f.name]) {
          if (mn === f.name) continue;
          printLog.push({ file: mn, ok: false, method: '', error: 'Sheet head fail: ' + String((e && e.message) || e), at: Date.now() });
          log('  ' + mn + ': FAIL - sheet head fail');
        }
      }
    }
  }
  const status = allOk ? 'Printed' : 'PrintFailed';
  await dbPatch('/shops/' + CFG.shopId + '/orders/' + orderKey, { status: status, printLog: printLog, printedAt: Date.now() });
  log('Order ' + (order.orderNo || orderKey) + ' -> ' + status);
}

// ---------- Main loop ----------
let busy = false;
async function pollOnce() {
  if (busy) return;
  busy = true;
  try {
    const settings = (await dbGet('/shops/' + CFG.shopId + '/settings')) || {};
    const orders = (await dbGet('/shops/' + CFG.shopId + '/orders')) || {};
    const autoPrint = settings.autoPrint !== false;
    const permMode = !!settings.permissionMode;
    for (const key of Object.keys(orders)) {
      const o = orders[key];
      if (!o) continue;
      const due = o.status === 'PrintRequested' || (o.status === 'New' && autoPrint && !permMode);
      if (due) {
        if (process.platform !== 'win32') {
          await dbPatch('/shops/' + CFG.shopId + '/orders/' + key, { status: 'PrintFailed', printLog: [{ file: '', ok: false, method: '', error: 'Auto-print sirf Windows PC par chalta hai - ye agent abhi Windows par nahi hai.', at: Date.now() }] });
          log('Order ' + (o.orderNo || key) + ' - Windows nahi hai, is liye print nahi ho saka.');
          continue;
        }
        await printOrder(key, o, settings);
      }
    }
    // Auto-Delete sweep
    const mins = parseInt(settings.autoDeleteMinutes, 10) || 0;
    if (mins > 0) {
      const cutoff = Date.now() - mins * 60 * 1000;
      for (const key of Object.keys(orders)) {
        const o = orders[key];
        if (o && (o.status === 'Done' || o.status === 'Cancelled' || o.status === 'Printed') && o.createdAt && o.createdAt < cutoff) {
          log('Auto-Delete: purana order saaf kar rahe hain ' + (o.orderNo || key));
          for (const f of (o.files || [])) if (f.storagePath) await storageDelete(f.storagePath);
          await dbDelete('/shops/' + CFG.shopId + '/orders/' + key);
        }
      }
    }
  } catch (e) {
    log('Poll error: ' + String((e && e.message) || e));
  }
  busy = false;
}
async function main() {
  console.log('==============================================');
  console.log('Pak QR Print - Print Agent');
  console.log('Dukaan: ' + CFG.shopId);
  console.log('Database: ' + CFG.databaseURL);
  console.log('SumatraPDF: ' + (fs.existsSync(SUMATRA_FILE) ? 'mil gaya' : 'NAHI MILA - PDF print ke liye isi folder mein SumatraPDF.exe rakhein'));
  console.log('Platform: ' + process.platform + ' (print sirf Windows par)');
  console.log('Poll har ' + CFG.pollSeconds + ' second');
  console.log('Band karne ke liye Ctrl+C dabayein.');
  console.log('==============================================');
  await signIn();
  await pollOnce();
  setInterval(pollOnce, (CFG.pollSeconds || 5) * 1000);
}
main().catch(function (e) { console.error('Agent start nahi ho saka: ' + String((e && e.message) || e)); process.exit(1); });
