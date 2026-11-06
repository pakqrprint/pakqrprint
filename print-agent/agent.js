// Pak QR Print - Print Agent (dukaan ke PC ke liye)
// ============================================================
// Ye program dukaan ke Windows PC par chalta hai. Kaam:
//   1. Firebase Realtime Database mein naye orders dekhta hai
//   2. File Firebase Storage se download karta hai
//   3. Print engine (printEngine.js - offline v1 ka azmaya hua
//      engine) se print karta hai: SumatraPDF (PDF), Word COM
//      (Word), Shell verb (Excel / PowerPoint 1-slide), PowerPoint
//      COM (slides-per-page), image sheets + Document N-up.
//   4. Status wapas Firebase par likhta hai (Printed / PrintFailed)
//   5. Auto-Delete: purane Done/Cancelled orders saaf karta hai.
// Sirf Node.js built-in modules - koi npm install nahi.
// Windows par hi print karta hai; kisi aur OS par queue dekhta
// hai aur saaf batata hai ke print sirf Windows par hoga.
// ============================================================
'use strict';
const fs = require('fs');
const path = require('path');
const https = require('https');
const { printOneFile, tryDocNupForFile, SUMATRA_FILE } = require('./printEngine.js');

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
  c.storageBucket = c.storageBucket || '';
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
async function storageDelete(storagePath) {
  if (!CFG.storageBucket) return;
  await ensureAuth();
  const url = 'https://firebasestorage.googleapis.com/v0/b/' + encodeURIComponent(CFG.storageBucket) + '/o/' + encodeURIComponent(storagePath) + '?auth=' + encodeURIComponent(idToken);
  try { await httpsJson('DELETE', url, null); } catch (e) { /* file pehle se na ho to koi baat nahi */ }
}

// ---------- Print ----------
function safeName(n) { return String(n || 'file').replace(/[\\/:*?"<>|]/g, '_'); }
async function printOrder(orderKey, order, settings) {
  log('Order ' + (order.orderNo || orderKey) + ' print ho raha hai (' + (order.files || []).length + ' file)...');
  await dbPatch('/shops/' + CFG.shopId + '/orders/' + orderKey, { status: 'Printing' });
  const printLog = [];
  let allOk = true;
  for (const f of (order.files || [])) {
    const dest = path.join(DOWNLOAD_DIR, orderKey + '_' + safeName(f.name));
    try {
      if (!f.downloadURL) throw new Error('File ka download link nahi mila');
      await httpsDownload(f.downloadURL, dest);
      const opts = {
        copies: f.copies || 1,
        colorMode: f.colorMode || order.colorMode || 'BW',
        duplex: false,
        paperSize: 'A4',
        range: (f.rangeMode === 'some' ? (f.range || '') : (f.rangeMode === 'first' ? '1-1' : '')),
        totalPages: f.pages || 0,
        slidesPerPage: f.spp || 1,
        nup: f.nup || 1,
        printerName: settings.printerName || '',
        orderPrinter: '',
        printerMismatch: false
      };
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
