// Pak QR Print - Print Agent ka print engine
// Offline v1 ke server.js (lines 516-1612) se nikala gaya - wahi azmaya hua
// engine: SumatraPDF (PDF), Word COM (Word), Shell Print verb (Excel /
// PowerPoint 1-slide), PowerPoint COM (slides-per-page), System.Drawing
// image sheets + Document N-up (Word/Excel/PDF). Koi tabdeeli nahi ki gayi,
// sirf HTTP server / orders / auth wala hissa nikala gaya hai.
// Sirf Windows par chalta hai (printOneFile ghair-Windows par saaf mana
// karta hai). SumatraPDF.exe isi folder me hona chahiye (PDF ke liye).
'use strict';
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFile } = require('child_process');
const ROOT = __dirname;
const ORDERS_DIR = path.join(ROOT, 'orders');
const PREVIEW_DIR = path.join(ROOT, 'preview-temp');
const PREVIEW_TTL_MS = 24 * 60 * 60 * 1000;
const SUMATRA_FILE = path.join(ROOT, 'SumatraPDF.exe');
const printQueue = [];
let printBusy = false;
const printQueuedIds = new Set();

function loadMetaById(id) {
  const dir = safeJoin(ORDERS_DIR, id);
  if (!dir) return null;
  const mp = path.join(dir, 'meta.json');
  if (!fs.existsSync(mp)) return null;
  try { return { dir: dir, metaPath: mp, meta: JSON.parse(fs.readFileSync(mp, 'utf8')) }; }
  catch (e) { return null; }
}
function saveMeta(rec) { fs.writeFileSync(rec.metaPath, JSON.stringify(rec.meta, null, 2)); }
function addPrintLog(meta, file, ok, method, error, detail) {
  meta.printLog = meta.printLog || [];
  const entry = { ts: Date.now(), time: new Date().toLocaleString('en-GB'), file: file, ok: !!ok, method: method || '', error: error || '' };
  if (detail) entry.detail = detail;
  meta.printLog.push(entry);
  if (meta.printLog.length > 40) meta.printLog = meta.printLog.slice(-40);
}
function execFileP(file, args, timeoutMs) {
  return new Promise((resolve) => {
    execFile(file, args, { timeout: timeoutMs || 90000, windowsHide: true }, (err, stdout, stderr) => {
      if (err) resolve({ ok: false, error: String((stderr || '') + ' ' + (err.message || err)).trim().slice(0, 300) });
      else resolve({ ok: true });
    });
  });
}
// execFileP jaisa hi, lekin stdout bhi wapas deta hai (Phase 2.7 ki
// Word page-ginti stdout se number parhti hai).
function execFileOut(file, args, timeoutMs) {
  return new Promise((resolve) => {
    execFile(file, args, { timeout: timeoutMs || 90000, windowsHide: true, maxBuffer: 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) resolve({ ok: false, stdout: stdout || '', error: String((stderr || '') + ' ' + (err.message || err)).trim().slice(0, 300) });
      else resolve({ ok: true, stdout: stdout || '' });
    });
  });
}
function psQuote(v) { return "'" + String(v || '').replace(/'/g, "''") + "'"; }

// ================= OFFICE PREVIEW (Phase 3.12) =================
// Shop PC par Office se convert: docx/xlsx/pptx -> ALAG temp PDF.
// - Endpoint POST /api/preview: ek file leta hai, type (docx/xlsx/
//   pptx ONLY) + size (order wali MAX_TOTAL limit) check karta hai,
//   uploaded copy ko preview-temp/ me likhta hai, Office COM se
//   sirf us copy ko PARH kar alag PDF banata hai. Original file
//   (phone par wali, aur yahan likhi hui copy) ki bytes kabhi
//   nahi badalti - SaveAs/ExportAsFixedFormat sirf nayi PDF
//   likhta hai, source ko save NAHI karta (Word/PowerPoint Close(0),
//   Excel Close(false) - bina save band).
// - Ye rasta kabhi order NAHI banata, settings NAHI chherta,
//   orders/ me NAHI likhta. Preview PDF ka naam random id hai,
//   is liye path traversal mumkin nahi (GET bhi sirf 32-hex id).
// - Office na ho / COM na bane / Windows na ho to RAW error
//   kabhi bahar nahi jata - sirf wahi friendly note jo print
//   notes me hai (Word -> Writer, PowerPoint -> Impress,
//   Excel -> Calc).
// - PQ_PREVIEW_MOCK=1 sirf is VM ke test ke liye hai (COM yahan
//   chal hi nahi sakta): asal conversion ki jagah ek chhoti si
//   valid PDF likhi jati hai taake baqi rasta (upload, cache,
//   serve, cleanup) prove ho sake. Shop PC par ye env kabhi
//   set nahi hota, is liye wahan hamesha asal Office COM chalta hai.
function previewKindFor(name) {
  const n = String(name || '').toLowerCase();
  if (/\.docx$/.test(n)) return 'word';
  if (/\.xlsx$/.test(n)) return 'excel';
  if (/\.pptx$/.test(n)) return 'powerpoint';
  return null;
}
function previewFriendlyNote(kind) {
  if (kind === 'word') return 'Is file ka preview dekhne ke liye is PC par MS Word ya LibreOffice Writer chahiye';
  if (kind === 'powerpoint') return 'Is file ka preview dekhne ke liye is PC par MS PowerPoint ya LibreOffice Impress hona chahiye';
  return 'Is file ka preview dekhne ke liye is PC par MS Excel ya LibreOffice Calc chahiye';
}
function sweepPreviewTemp() {
  try {
    if (!fs.existsSync(PREVIEW_DIR)) return;
    const now = Date.now();
    for (const f of fs.readdirSync(PREVIEW_DIR)) {
      try {
        const p = path.join(PREVIEW_DIR, f);
        const st = fs.statSync(p);
        if (!st.isDirectory() && (now - st.mtimeMs) > PREVIEW_TTL_MS) fs.unlinkSync(p);
      } catch (e) {}
    }
  } catch (e) {}
}
function buildOfficePreviewScript(kind, srcPath, pdfPath) {
  const lines = [];
  lines.push("$ErrorActionPreference = 'Stop'");
  lines.push('$fail = $null');
  lines.push('$app = $null');
  lines.push('$doc = $null');
  lines.push('try {');
  if (kind === 'word') {
    lines.push('  $app = New-Object -ComObject Word.Application');
    lines.push('  $app.Visible = $false');
    lines.push('  $app.DisplayAlerts = 0');
    lines.push('  $doc = $app.Documents.Open(' + psQuote(srcPath) + ', $false, $true)');
    lines.push('  $doc.SaveAs(' + psQuote(pdfPath) + ', 17)');
  } else if (kind === 'powerpoint') {
    lines.push('  $app = New-Object -ComObject PowerPoint.Application');
    lines.push('  $doc = $app.Presentations.Open(' + psQuote(srcPath) + ', $true, $false, $false)');
    lines.push('  Start-Sleep -Milliseconds 1000');
    lines.push('  $doc.SaveAs(' + psQuote(pdfPath) + ', 32)');
  } else {
    lines.push('  $app = New-Object -ComObject Excel.Application');
    lines.push('  $app.Visible = $false');
    lines.push('  $app.DisplayAlerts = $false');
    lines.push('  $doc = $app.Workbooks.Open(' + psQuote(srcPath) + ', $false, $true)');
    lines.push('  $doc.ExportAsFixedFormat(0, ' + psQuote(pdfPath) + ')');
  }
  lines.push('} catch {');
  lines.push('  $fail = $_.Exception.Message');
  lines.push('} finally {');
  if (kind === 'excel') lines.push('  if ($doc) { try { $doc.Close($false) } catch {} }');
  else if (kind === 'powerpoint') lines.push('  if ($doc) { try { $doc.Close() } catch {} }');
  else lines.push('  if ($doc) { try { $doc.Close(0) } catch {} }');
  lines.push('  if ($app) { try { $app.Quit() } catch {} }');
  lines.push('  if ($doc) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch {} }');
  lines.push('  if ($app) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {} }');
  lines.push('}');
  lines.push('if ($fail) { exit 1 }');
  lines.push('exit 0');
  return lines.join('\n');
}
const PREVIEW_MOCK_PDF = Buffer.from('%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 200]>>endobj\ntrailer<</Root 1 0 R/Size 4>>\n%%EOF\n', 'utf8');
// Phase 3.14: converted PDF ki tasdeeq — sirf maujood + bara hona
// kaafi nahi; pehle 4 bytes '%PDF' hone chahiye. Warna phone par
// embed khaali safed box dikhata hai (v3.12 ka masla: success URL
// ke peeche kharab file thi). GET endpoint bhi isi se parakhta hai.
function isValidPdfFile(p) {
  try {
    if (!fs.existsSync(p)) return false;
    if (fs.statSync(p).size === 0) return false;
    const fd = fs.openSync(p, 'r');
    try {
      const buf = Buffer.alloc(4);
      const n = fs.readSync(fd, buf, 0, 4, 0);
      return n === 4 && buf.toString('ascii') === '%PDF';
    } finally { fs.closeSync(fd); }
  } catch (e) { return false; }
}
async function convertOfficeToPdf(kind, srcPath, pdfPath) {
  try {
    if (process.env.PQ_PREVIEW_MOCK === '1') { fs.writeFileSync(pdfPath, PREVIEW_MOCK_PDF); return { ok: true }; }
    // PQ_PREVIEW_MOCK_BAD=1 sirf is VM ke test ke liye hai (jaise
    // PQ_PREVIEW_MOCK): kharab conversion ki naqal — tasdeeq ko
    // fail hona chahiye, kabhi success URL nahi milna chahiye.
    if (process.env.PQ_PREVIEW_MOCK_BAD === '1') { fs.writeFileSync(pdfPath, Buffer.from('NOT-A-PDF-GARBAGE', 'utf8')); }
    else {
      if (process.platform !== 'win32') return { ok: false, error: previewFriendlyNote(kind) };
      const cmd = buildOfficePreviewScript(kind, srcPath, pdfPath);
      const r = await execFileOut('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', cmd], 180000);
      if (!r.ok) return { ok: false, error: previewFriendlyNote(kind) };
    }
    if (!isValidPdfFile(pdfPath)) return { ok: false, error: previewFriendlyNote(kind) };
    return { ok: true };
  } catch (e) { return { ok: false, error: previewFriendlyNote(kind) }; }
}


async function printViaSumatra(filePath, opts) {
  const parts = [];
  const copies = Math.max(1, parseInt(opts.copies || '1', 10) || 1);
  if (copies > 1) parts.push(copies + 'x');
  // Phase 3.19: Mazeed options - sirf non-default tokens judte
  // hain; default par parts list v3.18 wali hi rehti hai.
  // Sumatra tokens (official docs): portrait/landscape,
  // paper=a4|letter|legal|a5, duplex/duplexshort/duplexlong,
  // fit/noscale. Collate/quality/margins ka Sumatra me koi
  // token nahi - wo is raste par unchanged (honest limit).
  if (opts.side === 'Double') {
    if (opts.flipEdge === 'short') parts.push('duplexshort');
    else parts.push('duplex');
  } else parts.push('simplex');
  parts.push(opts.printType === 'Colour' ? 'color' : 'monochrome');
  if (opts.orientation === 'portrait') parts.push('portrait');
  else if (opts.orientation === 'landscape') parts.push('landscape');
  if (opts.paperSize && opts.paperSize !== 'a4') parts.push('paper=' + opts.paperSize);
  if (opts.fitMode === 'actual') parts.push('noscale');
  if (opts.range) parts.push(String(opts.range));
  const args = ['-silent'];
  if (opts.printerName) args.push('-print-to', opts.printerName);
  else args.push('-print-to-default');
  args.push('-print-settings', parts.join(','));
  args.push(filePath);
  const r = await execFileP(SUMATRA_FILE, args, 120000);
  return { ok: r.ok, method: 'SumatraPDF (' + parts.join(', ') + ')', error: r.ok ? '' : ('SumatraPDF print fail: ' + r.error) };
}
async function printViaShell(filePath, opts) {
  // Windows Shell "Print" verb - default printer par bhejta hai.
  // Copies ke liye utni dafa repeat. Printer ki apni colour/duplex
  // default setting chalegi - is liye log me wazeh likha jata hai.
  const copies = Math.max(1, Math.min(10, parseInt(opts.copies || '1', 10) || 1));
  for (let i = 0; i < copies; i++) {
    const cmd = 'Start-Process -FilePath ' + psQuote(filePath) + ' -Verb Print -Wait';
    const r = await execFileP('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', cmd], 120000);
    if (!r.ok) return { ok: false, method: 'Windows Shell Print', error: 'Shell print fail: ' + r.error };
  }
  return { ok: true, method: 'Windows Shell Print (default printer, ' + copies + ' copy)', error: '' };
}
// ---- Word direct print (Phase 2.5.1) ----
// MS Word ke zariye direct print: PowerShell se Word.Application
// COM object banta hai, file read-only khulti hai aur PrintOut se
// print hoti hai. PrintBackground OFF rakha hai taake PrintOut
// print job spool hone tak wait kare (warna Word pehle band ho
// sakta hai). SELECTED printer opts me ho to Word ka ActivePrinter
// usi par set karne ki koshish hoti hai (fail ho to Word ka apna
// default printer rehta hai). Har qadam par COM objects saaf
// (Close/Quit/ReleaseComObject) kiye jate hain, null check ke saath.
function buildWordComScript(filePath, opts) {
  const copies = Math.max(1, Math.min(10, parseInt(opts.copies || '1', 10) || 1));
  const lines = [];
  lines.push("$ErrorActionPreference = 'Stop'");
  lines.push('$fail = $null');
  lines.push('$w = $null');
  lines.push('$d = $null');
  lines.push('$origPrinter = $null');
  lines.push('try {');
  lines.push('  $w = New-Object -ComObject Word.Application');
  lines.push('  $w.Visible = $false');
  lines.push('  $w.DisplayAlerts = 0');
  lines.push('  try { $w.Options.PrintBackground = $false } catch {}');
  if (opts.printerName) {
    // Phase 3.5: asal printer pehle note kar lo - PrintOut ke baad
    // wapas wahi laga dena hai (khusoosan jab Word pehle se khula
    // ho aur COM usi instance se jud jaye).
    lines.push('  try { $origPrinter = $w.ActivePrinter } catch {}');
    lines.push('  try { $w.ActivePrinter = ' + psQuote(opts.printerName) + ' } catch {}');
  }
  lines.push('  $d = $w.Documents.Open(' + psQuote(filePath) + ', $false, $true)');
  // Phase 3.19: Mazeed options (best-effort - har set try/catch,
  // na lage to print pehle wale tareeqe se hi nikalta hai; doc
  // read-only khulta aur bina save band hota hai, file kabhi
  // change nahi hoti). Sirf non-default lagte hain.
  if (opts.orientation === 'portrait' || opts.orientation === 'landscape') {
    lines.push('  try { $d.PageSetup.Orientation = ' + (opts.orientation === 'landscape' ? 1 : 0) + ' } catch {}');
  }
  if (opts.paperSize && opts.paperSize !== 'a4') {
    const WORD_PAPER = { letter: 1, legal: 5, a5: 11, a3: 8 };
    const pv = WORD_PAPER[opts.paperSize];
    if (pv) lines.push('  try { $d.PageSetup.PaperSize = ' + pv + ' } catch {}');
    else {
      const WORD_CUSTOM = { a2: [1191, 1684], a1: [1684, 2384], '4x6': [288, 432] };
      const cw = WORD_CUSTOM[opts.paperSize];
      if (cw) lines.push('  try { $d.PageSetup.PaperSize = 41; $d.PageSetup.PageWidth = ' + cw[0] + '; $d.PageSetup.PageHeight = ' + cw[1] + ' } catch {}');
    }
  }
  if (opts.margins === 'narrow') {
    lines.push('  try { $d.PageSetup.TopMargin = 36; $d.PageSetup.BottomMargin = 36; $d.PageSetup.LeftMargin = 36; $d.PageSetup.RightMargin = 36 } catch {}');
  }
  // Collate: pehle wala loop (poori copy ek dafa) khud collated
  // hai. Collate OFF ho to har page ki copies ikatthi (page x
  // copies); page ginti na mile to pehle wala loop fallback.
  if (opts.collate === 'off') {
    lines.push('  $collateDone = $false');
    lines.push('  try {');
    lines.push('    $pgCount = $d.ComputeStatistics(2)');
    lines.push('    if ($pgCount -gt 0) {');
    lines.push('      for ($pg = 1; $pg -le $pgCount; $pg++) { for ($i = 0; $i -lt ' + copies + '; $i++) { $d.PrintOut($false, $false, 3, \'\', $pg, $pg) } }');
    lines.push('      $collateDone = $true');
    lines.push('    }');
    lines.push('  } catch {}');
    lines.push('  if (-not $collateDone) { for ($i = 0; $i -lt ' + copies + '; $i++) { $d.PrintOut() } }');
  } else {
    lines.push('  for ($i = 0; $i -lt ' + copies + '; $i++) { $d.PrintOut() }');
  }
  lines.push('} catch {');
  lines.push('  $fail = $_.Exception.Message');
  lines.push('  [Console]::Error.WriteLine($fail)');
  lines.push('} finally {');
  if (opts.printerName) lines.push('  if ($w -and $origPrinter) { try { $w.ActivePrinter = $origPrinter } catch {} }');
  lines.push('  if ($d) { try { $d.Close(0) } catch {} }');
  lines.push('  if ($w) { try { $w.Quit() } catch {} }');
  lines.push('  if ($d) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($d) } catch {} }');
  lines.push('  if ($w) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch {} }');
  lines.push('}');
  lines.push('if ($fail) { exit 1 }');
  lines.push('exit 0');
  return lines.join('\n');
}
async function printViaWordCOM(filePath, opts) {
  const cmd = buildWordComScript(filePath, opts || {});
  const r = await execFileP('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', cmd], 180000);
  if (!r.ok) return { ok: false, method: 'Word Direct Print', error: 'Word print fail: ' + r.error };
  return { ok: true, method: 'Word Direct Print', error: '' };
}
// ---- PowerPoint direct print (Phase 3.5) ----
// Word wale raste ki nakal: PowerPoint.Application COM se file
// read-only khulti hai aur PrintOptions.OutputType se "ek page
// par kitni slides" (Handouts) set karke PrintOut hota hai.
// PpPrintOutputType (Microsoft Learn - PowerPoint VBA API):
//   Slides = 1, TwoSlideHandouts = 2, ThreeSlideHandouts = 3,
//   SixSlideHandouts = 4, FourSlideHandouts = 8,
//   NineSlideHandouts = 9.
// Customer ke ikhtiyar (slides per page) -> OutputType:
//   1 -> 1, 2 -> 2, 3 -> 3, 4 -> 8, 6 -> 4, 9 -> 9.
// (Source: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype)
// PowerPoint is PC par NA HO (sirf LibreOffice ho) to COM object
// banta hi nahi - script us surat me stdout par PQ_NOPPT likh kar
// exit 3 karti hai, taake caller khamoshi se 1-slide-per-page
// print karne ke bajaye file ko roke aur wajah log kare.
const PPT_OUTPUT_TYPE = { 1: 1, 2: 2, 3: 3, 4: 8, 6: 4, 9: 9 };
function buildPptComScript(filePath, opts) {
  const copies = Math.max(1, Math.min(10, parseInt(opts.copies || '1', 10) || 1));
  const spp = parseInt(opts.slidesPerPage, 10) || 1;
  const outType = PPT_OUTPUT_TYPE[spp] || 1;
  // Phase 3.22: PPT range - opts.range "2-5" ya "1-1" ho to PrintOut(From, To);
  // range na ho to pehle jaisa PrintOut() (poori file). PowerPoint VBA:
  // Presentation.PrintOut(From, To, PrintToFile, Copies, Collate).
  var pptFrom = 0, pptTo = 0;
  (function () {
    var r = String((opts && opts.range) || '').trim();
    var m = /^(\d+)\s*-\s*(\d+)$/.exec(r);
    if (m) { var a = parseInt(m[1], 10), b = parseInt(m[2], 10); if (a >= 1 && b >= a) { pptFrom = a; pptTo = b; } }
    else if (/^\d+$/.test(r)) { var q = parseInt(r, 10); if (q >= 1) { pptFrom = q; pptTo = q; } }
  })();
  const lines = [];
  lines.push("$ErrorActionPreference = 'Stop'");
  lines.push('$fail = $null');
  lines.push('$noppt = $false');
  lines.push('$comErr = $null');
  lines.push('$p = $null');
  lines.push('$pres = $null');
  lines.push('$origPrinter = $null');
  lines.push('try {');
  lines.push('  try { $p = New-Object -ComObject PowerPoint.Application } catch { $comErr = $_.Exception.Message; $noppt = $true; throw }');
  lines.push('  try { $p.DisplayAlerts = 1 } catch {}');
  lines.push('  [Console]::Out.WriteLine("PQ_STEP:COM")');
  if (opts.printerName) {
    lines.push('  try { $origPrinter = $p.ActivePrinter } catch {}');
    lines.push('  try { $p.ActivePrinter = ' + psQuote(opts.printerName) + ' } catch {}');
  }
  lines.push('  $pres = $p.Presentations.Open(' + psQuote(filePath) + ', $true, $false, $false)');
  lines.push('  [Console]::Out.WriteLine("PQ_STEP:OPEN")');
  lines.push('  Start-Sleep -Milliseconds 1500');
  lines.push('  $pres.PrintOptions.OutputType = ' + outType);
  var pptPrintCall = (pptFrom > 0 && pptTo >= pptFrom) ? '$pres.PrintOut(' + pptFrom + ', ' + pptTo + ')' : '$pres.PrintOut()';
  lines.push('  for ($i = 0; $i -lt ' + copies + '; $i++) { ' + pptPrintCall + ' }');
  lines.push('  [Console]::Out.WriteLine("PQ_STEP:PRINT")');
  // Phase 3.17: spool-wait - PrintOut ke baad foran Close/Quit job ko
  // spool hone se pehle cancel kar deta tha (background print race).
  // Printer ki queue khaali hone tak wait (har 2s, cap 60s); query
  // error par 10s sleep fallback; wait kabhi throw nahi karta.
  lines.push('  try {');
  lines.push('    $spoolTries = 0');
  lines.push('    while ($spoolTries -lt 30) {');
  lines.push('      $spoolJobs = @(Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop)');
  if (opts.printerName) lines.push('      $spoolJobs = @($spoolJobs | Where-Object { $_.Name -and $_.Name.StartsWith(' + psQuote(opts.printerName + ',') + ') })');
  lines.push('      if ($spoolJobs.Count -eq 0) { break }');
  lines.push('      Start-Sleep -Seconds 2');
  lines.push('      $spoolTries++');
  lines.push('    }');
  lines.push('  } catch { Start-Sleep -Seconds 10 }');
  lines.push('  [Console]::Out.WriteLine("PQ_STEP:SPOOL")');
  lines.push('} catch {');
  lines.push('  if (-not $noppt) { $fail = $_.Exception.Message; [Console]::Error.WriteLine($fail); [Console]::Out.WriteLine("PQ_FAIL:" + (($fail -replace "`r|`n", " ").Trim())) }');
  lines.push('} finally {');
  if (opts.printerName) lines.push('  if ($p -and $origPrinter) { try { $p.ActivePrinter = $origPrinter } catch {} }');
  lines.push('  if ($pres) { try { $pres.Close() } catch {} }');
  lines.push('  if ($p) { try { $quitOk = $true; try { if ($p.Presentations.Count -ne 0) { $quitOk = $false } } catch { $quitOk = $true }; if ($quitOk) { $p.Quit() } } catch {} }');
  lines.push('  if ($pres) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) } catch {} }');
  lines.push('  if ($p) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($p) } catch {} }');
  lines.push('}');
  lines.push('if (-not $noppt -and -not $fail) { [Console]::Out.WriteLine("PQ_STEP:DONE") }');
  lines.push('if ($noppt) { [Console]::Out.WriteLine("PQ_NOPPT"); if ($comErr) { [Console]::Out.WriteLine("PQ_COMERR:" + (($comErr -replace "`r|`n", " ").Trim())) }; exit 3 }');
  lines.push('if ($fail) { exit 1 }');
  lines.push('exit 0');
  return lines.join('\n');
}
// Phase 3.8 (COM hardening): PowerPoint COM banane ki koshish pehle
// fail ho to khamoshi se haar NAHI maante - kyunke us PC par MS Office
// installed hai, COM aksar pehli koshish par (ya 64-bit shell me)
// nahi banta. Tarteeb: (1) pehli koshish, (2) ~2 second ruk kar
// dobara wahi koshish, (3) 32-bit PowerShell se ek koshish. Teenon
// fail hon tab hi decline hota hai. spp>1 order kabhi chupke
// 1-slide-per-page Shell print par NAHI jata (ye rasta sirf COM ka hai).
// PQ_PPT_PS / PQ_PPT_PS32 sirf dry-run test ke liye hain; aam halat
// me defaults hi chalte hain (powershell.exe / SysWOW64 wala).
// Phase 3.9 (PowerPoint Wajah): COM banne ke BAAD wali ghalti (Open/
// PrintOut) ka $fail text ab PQ_FAIL marker se stdout par aata hai,
// taake order note me asal wajah likhi aaye, poori command nahi.
// PQ_FAIL na mile (jaise timeout kill) to chhota fallback note.
// Open ke baad 1500ms pause (readiness race guard).
// Phase 3.15 (Print Fix): timeout par phansi PowerPoint.
// (1) buildPptComScript har phase ke baad PQ_STEP likhta hai (COM /
// OPEN / PRINT / DONE) taake timeout note me aakhri mukammal phase
// aaye ('waqt khatam — atka: OPEN'; koi marker na ho to COM banate
// waqt atka). (2) runPptComAttempt timeout (ok:false, PQ_FAIL
// nahi) ko timeout:true ke saath wapas karta hai. (3)
// printViaPowerPointCOM timeout par pehle sirf NON-RESPONDING
// POWERPNT.EXE band karta hai (user ki khuli, responding
// PowerPoint mehfooz rehti hai), 2s ruk kar EK taza koshish karta
// hai; PQ_FAIL par dobara nahi. NOPPT wala 3-attempt tareeqa
// waisa hi. Androoni retry printAttempts (koshish) NAHI barhata;
// dono koshishein printLog me aati hain (pehli timeout printLogFn
// se, aakhri queue ke aam entry se). PQ_PPT_TIMEOUT_MS sirf test
// ke liye hai (harness) - aam halat me 180s.
// Phase 3.17 (PowerPoint Print Fix 2): 'atka: PRINT' ki asal wajah.
// PQ_STEP:PRINT, PrintOut loop ke BAAD likha jata hai - is liye
// PrintOut khud mukammal ho chuka tha; hang us ke baad finally
// (pres.Close / p.Quit) me tha. PowerPoint background me print
// karta hai, is liye PrintOut ke foran baad Close/Quit job ko
// spool hone se pehle cancel kar deta tha (khud print me
// PowerPoint khula rehta hai, is liye poora nikalta tha; 26 Oct
// kabhi chala kabhi nahi = ye race). Fix: (1) DisplayAlerts=1
// (ppAlertsNone) taake automation me modal dialog na ruke; (2)
// PrintOut ke baad spool-wait - Win32_PrintJob poll (target
// printer, har 2s, cap 60s; query error par 10s sleep fallback),
// phir PQ_STEP:SPOOL; (3) Quit guard - apni presentation band
// karne ke baad sirf tab Quit jab Presentations.Count 0 ho
// (user ki khuli files par Quit/save-prompt nahi). Timeout
// (180s) + non-responding-only cleanup + ek automatic retry,
// NOPPT 3-attempt tareeqa aur baqi sab kuch v3.15 jaisa hi hai.
const PPT_PS32 = 'C:\\Windows\\SysWOW64\\WindowsPowerShell\\v1.0\\powershell.exe';
async function runPptComAttempt(psExe, filePath, opts) {
  const cmd = buildPptComScript(filePath, opts || {});
  const pptTimeout = parseInt(process.env.PQ_PPT_TIMEOUT_MS, 10) || 180000;
  const r = await execFileOut(psExe, ['-NoProfile', '-NonInteractive', '-Command', cmd], pptTimeout);
  const out = String(r.stdout || '');
  if (out.indexOf('PQ_NOPPT') !== -1) {
    let comError = '';
    const m = /PQ_COMERR:([^\r\n]*)/.exec(out);
    if (m) comError = m[1].trim().slice(0, 200);
    return { noppt: true, comError: comError, error: r.error || '' };
  }
  if (!r.ok) {
    let failError = '';
    const fm = /PQ_FAIL:([^\r\n]*)/.exec(out);
    if (fm) failError = fm[1].trim().slice(0, 200);
    if (failError) return { noppt: false, ok: false, timeout: false, error: 'PowerPoint print fail: ' + failError, fullError: r.error || '' };
    const steps = out.match(/PQ_STEP:(COM|OPEN|PRINT|SPOOL|DONE)/g) || [];
    const step = steps.length ? steps[steps.length - 1].replace('PQ_STEP:', '') : 'COM';
    const note = 'PowerPoint print fail: PowerPoint se jawab nahi aaya (waqt khatam — atka: ' + step + ') - file Download/Open karke khud print karein.';
    return { noppt: false, ok: false, timeout: true, step: step, error: note, fullError: r.error || '' };
  }
  return { noppt: false, ok: true, error: '' };
}
// Phase 3.15: timeout ke baad sirf NON-RESPONDING POWERPNT.EXE band
// karte hain - user ki khuli (responding) PowerPoint ko haath nahi
// lagta. Best-effort: cleanup fail ho to bhi retry hoti hai.
async function killHungPowerPoint(psExe) {
  const cmd = 'Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { -not $_.Responding } | Stop-Process -Force';
  await execFileOut(psExe || 'powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', cmd], 30000);
}
async function printViaPowerPointCOM(filePath, opts) {
  const psExe = process.env.PQ_PPT_PS || 'powershell.exe';
  const ps32Exe = process.env.PQ_PPT_PS32 || PPT_PS32;
  const spp = parseInt((opts || {}).slidesPerPage, 10) || 1;
  const rangeTxt = String((opts || {}).range || '').trim();
  const okMethod = 'PowerPoint Direct Print (' + spp + ' slides per page' + (rangeTxt ? ', slides ' + rangeTxt : '') + ')';
  const logFn = (opts && typeof opts.printLogFn === 'function') ? opts.printLogFn : null;
  let a = await runPptComAttempt(psExe, filePath, opts);
  if (!a.noppt) {
    if (a.ok) return { ok: true, method: okMethod, error: '' };
    if (a.timeout) {
      if (logFn) logFn(false, 'PowerPoint Direct Print', a.error, a.fullError || '');
      await killHungPowerPoint(psExe);
      await new Promise((res) => setTimeout(res, 2000));
      const b = await runPptComAttempt(psExe, filePath, opts);
      if (b.noppt) return { ok: false, method: 'PowerPoint Direct Print', error: 'PowerPoint COM nahi bana' + (b.comError ? ': ' + b.comError : ''), fullError: b.error || '' };
      if (b.ok) return { ok: true, method: okMethod, error: '', fullError: 'Pehli koshish timeout (atka: ' + (a.step || 'COM') + '); phansi PowerPoint band karke dobara koshish kamyab.', detail: 'Pehli koshish timeout (atka: ' + (a.step || 'COM') + '); phansi PowerPoint band karke dobara koshish kamyab.' };
      return { ok: false, method: 'PowerPoint Direct Print', error: b.error + ' | Pehli koshish bhi timeout thi (atka: ' + (a.step || 'COM') + '), phansi PowerPoint band karke dobara koshish hui.', fullError: (a.fullError || '') + ' || ' + (b.fullError || '') };
    }
    return { ok: false, method: 'PowerPoint Direct Print', error: a.error, fullError: a.fullError || '' };
  }
  await new Promise((res) => setTimeout(res, 2000));
  a = await runPptComAttempt(psExe, filePath, opts);
  if (!a.noppt) return { ok: a.ok, method: a.ok ? okMethod : 'PowerPoint Direct Print', error: a.error, fullError: a.fullError || '' };
  a = await runPptComAttempt(ps32Exe, filePath, opts);
  if (!a.noppt) return { ok: a.ok, method: a.ok ? okMethod : 'PowerPoint Direct Print', error: a.error, fullError: a.fullError || '' };
  return { ok: false, noPowerPoint: true, comError: a.comError || '', method: 'PowerPoint Direct Print', error: 'PowerPoint COM nahi bana' + (a.comError ? ': ' + a.comError : '') };
}
// ---- Image print (Phase 2.1) ----
// Embedded PowerShell script: System.Drawing se image load karke
// PrintDocument se print karta hai (page par fit + center, aspect
// ratio mehfooz). -Test switch par koi file nahi - memory me test
// page bana kar wahi pipeline se print karta hai (Test Print button).
// Phase 3.19: Mazeed print options ka shared PowerShell settings
// block - single image aur N-up sheet dono scripts me waisa hi
// lagta hai. Sirf non-default ikhtiyar lagte hain; default par
// kuch nahi chhedte taake pehle wala behaviour hi rahe. Har set
// try/catch me hai - koi setting na lage to print phir bhi
// pehle wale tareeqe se nikalta hai (khamoshi se fail nahi hota).
const IMG_PS_SETTINGS = `
  try { if ($Orient -eq 'landscape') { $pd.DefaultPageSettings.Landscape = $true } elseif ($Orient -eq 'portrait') { $pd.DefaultPageSettings.Landscape = $false } } catch {}
  try {
    if ($Paper -ne 'a4') {
      $dims = @{ 'letter' = @(850, 1100); 'legal' = @(850, 1400); 'a5' = @(583, 827); 'a3' = @(1169, 1654); 'a2' = @(1654, 2339); 'a1' = @(2339, 3311); '4x6' = @(400, 600) }[$Paper]
      if ($dims) { $pd.DefaultPageSettings.PaperSize = New-Object System.Drawing.Printing.PaperSize($Paper, $dims[0], $dims[1]) }
    }
  } catch {}
  try {
    if ($Quality -eq 'draft') { $pd.DefaultPageSettings.PrinterResolution = New-Object System.Drawing.Printing.PrinterResolution(150, 150, [System.Drawing.Printing.PrinterResolutionKind]::Custom) }
    elseif ($Quality -eq 'best') { $pd.DefaultPageSettings.PrinterResolution = New-Object System.Drawing.Printing.PrinterResolution(600, 600, [System.Drawing.Printing.PrinterResolutionKind]::Custom) }
  } catch {}
  try { if ($Collate -eq 1) { $pd.PrinterSettings.Collate = $true } elseif ($Collate -eq 2) { $pd.PrinterSettings.Collate = $false } } catch {}
  try { if ($MarginNarrow) { $pd.DefaultPageSettings.Margins = New-Object System.Drawing.Printing.Margins(25, 25, 25, 25) } } catch {}
  try { if ($DuplexOn -and $Flip -eq 1) { $pd.PrinterSettings.Duplex = [System.Drawing.Printing.Duplex]::Vertical } elseif ($DuplexOn -and $Flip -eq 2) { $pd.PrinterSettings.Duplex = [System.Drawing.Printing.Duplex]::Horizontal } } catch {}
`;
const IMG_PS_SCRIPT = `
param(
  [Parameter(Mandatory=$false)][string]$Path,
  [Parameter(Mandatory=$false)][string]$Printer = '',
  [Parameter(Mandatory=$false)][int]$Copies = 1,
  [switch]$Mono,
  [switch]$Test,
  [Parameter(Mandatory=$false)][string]$Orient = 'auto',
  [Parameter(Mandatory=$false)][string]$Paper = 'a4',
  [Parameter(Mandatory=$false)][string]$Quality = 'normal',
  [Parameter(Mandatory=$false)][int]$Collate = 0,
  [Parameter(Mandatory=$false)][string]$FitMode = 'fit',
  [switch]$MarginNarrow,
  [Parameter(Mandatory=$false)][int]$Flip = 0,
  [switch]$DuplexOn
)
try {
  Add-Type -AssemblyName System.Drawing
  $img = $null
  if ($Test) {
    $bmp = New-Object System.Drawing.Bitmap 800, 1100
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::White)
    $fontTitle = New-Object System.Drawing.Font('Arial', 22, [System.Drawing.FontStyle]::Bold)
    $fontBody = New-Object System.Drawing.Font('Arial', 14)
    $brush = [System.Drawing.Brushes]::Black
    $g.DrawString('Pak QR Print - Pakistan Edition', $fontTitle, $brush, 60, 80)
    $g.DrawString('AUTO-PRINT TEST PAGE (Phase 2.4)', $fontBody, $brush, 60, 140)
    $g.DrawString('Agar ye page printer se nikla hai to Auto-Print theek kaam kar raha hai.', $fontBody, $brush, 60, 180)
    $g.DrawString(('Time: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')), $fontBody, $brush, 60, 220)
    $g.Dispose()
    $img = $bmp
  } else {
    if (-not $Path) { throw 'Image path nahi mila' }
    $img = [System.Drawing.Image]::FromFile($Path)
  }
  $global:pakqrImg = $img
  $pd = New-Object System.Drawing.Printing.PrintDocument
  if ($Printer) { $pd.PrinterSettings.PrinterName = $Printer }
  if ($Copies -lt 1) { $Copies = 1 }
  if ($Copies -gt 10) { $Copies = 10 }
  $pd.PrinterSettings.Copies = $Copies
  if ($Mono) { $pd.DefaultPageSettings.Color = $false }
  $global:pakqrFitMode = $FitMode
${IMG_PS_SETTINGS}
  $handler = {
    param($sender, $e)
    $bounds = $e.MarginBounds
    $iw = $global:pakqrImg.Width
    $ih = $global:pakqrImg.Height
    $useActual = $false
    if ($global:pakqrFitMode -eq 'actual') {
      try {
        $resX = $global:pakqrImg.HorizontalResolution
        $resY = $global:pakqrImg.VerticalResolution
        if ($resX -le 0) { $resX = 96 }
        if ($resY -le 0) { $resY = 96 }
        $w = [int]($iw / $resX * $e.Graphics.DpiX)
        $h = [int]($ih / $resY * $e.Graphics.DpiY)
        $useActual = $true
      } catch { $useActual = $false }
    }
    if (-not $useActual) {
      $scale = [Math]::Min($bounds.Width / $iw, $bounds.Height / $ih)
      $w = [int]($iw * $scale)
      $h = [int]($ih * $scale)
    }
    $x = [int]($bounds.X + (($bounds.Width - $w) / 2))
    $y = [int]($bounds.Y + (($bounds.Height - $h) / 2))
    $e.Graphics.DrawImage($global:pakqrImg, $x, $y, $w, $h)
    $e.HasMorePages = $false
  }
  $pd.add_PrintPage($handler)
  $pd.Print()
  $img.Dispose()
  exit 0
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
`;
let imgPsPath = null;
function ensureImgPsScript() {
  if (imgPsPath && fs.existsSync(imgPsPath)) return imgPsPath;
  const p = path.join(os.tmpdir(), 'pakqr-print-img.ps1');
  try { fs.writeFileSync(p, IMG_PS_SCRIPT, 'utf8'); imgPsPath = p; }
  catch (e) { imgPsPath = null; }
  return imgPsPath;
}
async function runImgPs(extraArgs) {
  const script = ensureImgPsScript();
  if (!script) return { ok: false, error: 'Image print script temp folder me nahi ban saka' };
  const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script].concat(extraArgs);
  return execFileP('powershell.exe', args, 120000);
}
async function printViaPaint(filePath, opts) {
  // MS Paint fallback: /p = default printer, /pt = chuna hua printer.
  const copies = Math.max(1, Math.min(10, parseInt(opts.copies || '1', 10) || 1));
  for (let i = 0; i < copies; i++) {
    const args = opts.printerName ? ['/pt', opts.printerName, filePath] : ['/p', filePath];
    const r = await execFileP('mspaint.exe', args, 120000);
    if (!r.ok) return { ok: false, method: 'MS Paint Print', error: 'Paint print fail: ' + r.error };
  }
  return { ok: true, method: 'MS Paint Print', error: '' };
}
async function printViaImage(filePath, opts) {
  const copies = Math.max(1, Math.min(10, parseInt(opts.copies || '1', 10) || 1));
  const args = ['-Path', filePath, '-Copies', String(copies)];
  if (opts.printerName) args.push('-Printer', opts.printerName);
  if (opts.printType !== 'Colour') args.push('-Mono');
  // Phase 3.19: Mazeed options - sirf non-default bhejte hain,
  // taake default par script ka pehle wala behaviour hi rahe.
  appendImgMazeedArgs(args, opts);
  const r = await runImgPs(args);
  if (r.ok) return { ok: true, method: 'Windows Image Print', error: '' };
  const rp = await printViaPaint(filePath, opts);
  if (rp.ok) return rp;
  return { ok: false, method: 'Windows Image Print', error: ('Image print fail: ' + (r.error || 'unknown')).slice(0, 200) + ' | ' + ('Paint fail: ' + (rp.error || 'unknown')).slice(0, 160) };
}
async function printTestPageImage(opts) {
  // Test Print button: wahi image pipeline, file ke baghair.
  const args = ['-Test'];
  if (opts.printerName) args.push('-Printer', opts.printerName);
  const r = await runImgPs(args);
  if (r.ok) return { ok: true, method: 'Windows Image Print (Test Page)', error: '' };
  return { ok: false, method: 'Windows Image Print (Test Page)', error: 'Test print fail: ' + (r.error || 'unknown') };
}
// Phase 3.19: Mazeed options -> image/sheet script args (shared).
// Sirf non-default ikhtiyar judte hain; default par args wahi
// rehte hain jo v3.18 me the (behaviour byte-identical).
function appendImgMazeedArgs(args, opts) {
  if (opts.orientation && opts.orientation !== 'auto') args.push('-Orient', opts.orientation);
  if (opts.paperSize && opts.paperSize !== 'a4') args.push('-Paper', opts.paperSize);
  if (opts.quality && opts.quality !== 'normal') args.push('-Quality', opts.quality);
  if (opts.collate === 'on') args.push('-Collate', '1');
  else if (opts.collate === 'off') args.push('-Collate', '2');
  if (opts.fitMode && opts.fitMode !== 'fit') args.push('-FitMode', opts.fitMode);
  if (opts.margins === 'narrow') args.push('-MarginNarrow');
  // Duplex flip sirf tab jab Double Side chuna ho; warna side ka
  // pehle wala behaviour (printer default) hi rehta hai.
  if (opts.side === 'Double') {
    args.push('-DuplexOn');
    if (opts.flipEdge === 'long') args.push('-Flip', '1');
    else if (opts.flipEdge === 'short') args.push('-Flip', '2');
  }
}
// Phase 3.19: N-up sheet script - aik page par grid (2/4/6
// pictures). Har picture apne cell me fit + center (aspect
// mehfooz). Sheet ki copies order wali copies hoti hain.
// 'Actual size' sheet par nahi lagta (grid me har picture cell
// me fit hoti hai) - ye Hidayat me likha hai. Paint fallback
// sheet par NAHI (Paint grid nahi bana sakta) - fail ho to
// seedha fail note, chupke alag alag print nahi hote.
const IMG_SHEET_PS_SCRIPT = `
param(
  [Parameter(Mandatory=$false)][string]$ListFile,
  [Parameter(Mandatory=$false)][string]$Printer = '',
  [Parameter(Mandatory=$false)][int]$Copies = 1,
  [switch]$Mono,
  [Parameter(Mandatory=$false)][int]$Cols = 2,
  [Parameter(Mandatory=$false)][int]$Rows = 2,
  [Parameter(Mandatory=$false)][string]$Orient = 'auto',
  [Parameter(Mandatory=$false)][string]$Paper = 'a4',
  [Parameter(Mandatory=$false)][string]$Quality = 'normal',
  [Parameter(Mandatory=$false)][int]$Collate = 0,
  [switch]$MarginNarrow,
  [Parameter(Mandatory=$false)][int]$Flip = 0,
  [switch]$DuplexOn
)
try {
  Add-Type -AssemblyName System.Drawing
  if (-not $ListFile -or -not (Test-Path $ListFile)) { throw 'Sheet file list nahi mili' }
  $paths = Get-Content $ListFile -Encoding UTF8 | Where-Object { $_ -and $_.Trim() -ne '' }
  if ($paths.Count -eq 0) { throw 'Sheet me koi tasveer nahi' }
  $imgs = @()
  foreach ($p in $paths) { $imgs += [System.Drawing.Image]::FromFile($p) }
  $global:pakqrSheetImgs = $imgs
  $global:pakqrSheetCols = $Cols
  $global:pakqrSheetRows = $Rows
  $pd = New-Object System.Drawing.Printing.PrintDocument
  if ($Printer) { $pd.PrinterSettings.PrinterName = $Printer }
  if ($Copies -lt 1) { $Copies = 1 }
  if ($Copies -gt 10) { $Copies = 10 }
  $pd.PrinterSettings.Copies = $Copies
  if ($Mono) { $pd.DefaultPageSettings.Color = $false }
${IMG_PS_SETTINGS}
  $handler = {
    param($sender, $e)
    $bounds = $e.MarginBounds
    $cols = $global:pakqrSheetCols
    $rows = $global:pakqrSheetRows
    $cellW = $bounds.Width / $cols
    $cellH = $bounds.Height / $rows
    $pad = 20
    for ($i = 0; $i -lt $global:pakqrSheetImgs.Count; $i++) {
      $img = $global:pakqrSheetImgs[$i]
      $col = $i % $cols
      $row = [int]($i / $cols)
      $cx = $bounds.X + ($col * $cellW) + $pad
      $cy = $bounds.Y + ($row * $cellH) + $pad
      $cw = $cellW - (2 * $pad)
      $ch = $cellH - (2 * $pad)
      if ($cw -le 0 -or $ch -le 0) { continue }
      $scale = [Math]::Min($cw / $img.Width, $ch / $img.Height)
      $w = [int]($img.Width * $scale)
      $h = [int]($img.Height * $scale)
      $x = [int]($cx + (($cw - $w) / 2))
      $y = [int]($cy + (($ch - $h) / 2))
      $e.Graphics.DrawImage($img, $x, $y, $w, $h)
    }
    $e.HasMorePages = $false
  }
  $pd.add_PrintPage($handler)
  $pd.Print()
  foreach ($im in $imgs) { $im.Dispose() }
  exit 0
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
`;
let imgSheetPsPath = null;
function ensureImgSheetPsScript() {
  if (imgSheetPsPath && fs.existsSync(imgSheetPsPath)) return imgSheetPsPath;
  const p = path.join(os.tmpdir(), 'pakqr-print-img-sheet.ps1');
  try { fs.writeFileSync(p, IMG_SHEET_PS_SCRIPT, 'utf8'); imgSheetPsPath = p; }
  catch (e) { imgSheetPsPath = null; }
  return imgSheetPsPath;
}
async function printViaImageSheet(filePaths, opts) {
  const copies = Math.max(1, Math.min(10, parseInt(opts.copies || '1', 10) || 1));
  const ppp = parseInt(opts.picturesPerPage, 10) || 1;
  const grid = sheetGrid(ppp, opts.orientation === 'landscape');
  const cols = grid.cols, rows = grid.rows;
  const listPath = path.join(os.tmpdir(), 'pakqr-sheet-' + Date.now() + '-' + Math.floor(Math.random() * 100000) + '.txt');
  try { fs.writeFileSync(listPath, filePaths.join('\n'), 'utf8'); }
  catch (e) { return { ok: false, method: 'Windows Image Print (Sheet)', error: 'Sheet list nahi ban saki: ' + String(e.message || e).slice(0, 160) }; }
  const script = ensureImgSheetPsScript();
  if (!script) { try { fs.unlinkSync(listPath); } catch (e) {} return { ok: false, method: 'Windows Image Print (Sheet)', error: 'Sheet print script temp folder me nahi ban saka' }; }
  const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, '-ListFile', listPath, '-Cols', String(cols), '-Rows', String(rows), '-Copies', String(copies)];
  if (opts.printerName) args.push('-Printer', opts.printerName);
  if (opts.printType !== 'Colour') args.push('-Mono');
  appendImgMazeedArgs(args, opts);
  const r = await execFileP('powershell.exe', args, 120000);
  try { fs.unlinkSync(listPath); } catch (e) {}
  if (r.ok) return { ok: true, method: 'Windows Image Print (Sheet ' + filePaths.length + ' pictures)', error: '' };
  return { ok: false, method: 'Windows Image Print (Sheet)', error: 'Sheet print fail: ' + (r.error || 'unknown').slice(0, 200) };
}
// Phase 3.19: N-up grouping (pure - test ke liye alag).
// Sirf pictures (image extension) group hoti hain, print
// tarteeb me; CNIC par ya ppp 1 par koi group nahi.
function groupImageSheets(files, ppp, isCnic) {
  const sheetHead = {};
  const sheetMember = {};
  if (ppp > 1 && !isCnic) {
    const imgs = (files || []).filter((f) => /\.(jpe?g|png|bmp|gif)$/i.test(f));
    for (let gi = 0; gi < imgs.length; gi += ppp) {
      const grp = imgs.slice(gi, gi + ppp);
      sheetHead[grp[0]] = grp;
      grp.forEach((g) => { sheetMember[g] = true; });
    }
  }
  return { sheetHead: sheetHead, sheetMember: sheetMember };
}
// Phase 3.19: N-up grid dims (pure). Portrait sheet lambi,
// landscape sheet chauri hoti hai; 4 hamesha 2x2.
function sheetGrid(ppp, landscape) {
  if (ppp === 2) return landscape ? { cols: 2, rows: 1 } : { cols: 1, rows: 2 };
  if (ppp === 6) return landscape ? { cols: 3, rows: 2 } : { cols: 2, rows: 3 };
  return { cols: 2, rows: 2 };
}
// Phase 3.21: Document N-up engine (Word/Excel/PDF pages-per-sheet).
// Kyun ye engine: SumatraPDF me N-up nahi hai (sirf orientation /
// fit / paper / duplex tokens), Word ke PrintOut me pages-per-sheet
// nahi hai, Excel ka COM print hai hi nahi. Is liye: PDF seedha,
// Word/Excel pehle Office COM se PDF (convertOfficeToPdf), phir
// PDF ke pages Windows WinRT (Windows.Data.Pdf - Windows 10+ me
// pehle se maujood, koi naya binary nahi, koi download nahi) se
// PNG ban kar wahi image-sheet pipeline (printViaImageSheet) se
// grid me print hote hain. Is liye Mazeed options sheets par
// pictures jaisi hi lagti hain (orientation, paper, quality,
// collate, margins, duplex flip; fit=actual sheets par nahi
// lagta - pictures ki tarah). Engine fail ho to file purane
// tareeqe (1-up) par print hoti hai - order kabhi fail nahi.
// Render sirf Windows par chalta hai; is VM par sirf mock +
// grouping/range/fallback test hote hain (asal render + COM +
// sheet print sirf shop PC par prove honge).
function parseDocNupPages(effRange, totalPages) {
  // '' => 'all' (sab pages); 'a-b' => [a..b] (clamped);
  // ghalat range => null (caller purane raste par girta hai).
  const r = String(effRange || '').trim();
  if (!r) return 'all';
  const parts = r.split('-');
  if (parts.length !== 2) return null;
  const a = parseInt(parts[0].trim(), 10), b = parseInt(parts[1].trim(), 10);
  if (isNaN(a) || isNaN(b) || a < 1 || b < a) return null;
  let hi = b;
  if (totalPages > 0) {
    if (a > totalPages) return null;
    if (hi > totalPages) hi = totalPages;
  }
  const pages = [];
  for (let p = a; p <= hi; p++) pages.push(p);
  return pages;
}
function groupDocPages(pages, ppp) {
  const groups = [];
  for (let i = 0; i < pages.length; i += ppp) groups.push(pages.slice(i, i + ppp));
  return groups;
}
function docNupDpi(quality) {
  if (quality === 'draft') return 150;
  if (quality === 'best') return 300;
  return 200;
}
// WinRT PDF render script (PowerShell, Windows 10+). Koi backtick
// ya ${} nahi (JS template literal me mehfooz), koi regex nahi
// (range Split se parhi jati hai). Stdout: PAGES:n phir OK.
const DOC_NUP_PS_SCRIPT = `
param(
  [Parameter(Mandatory=$true)][string]$Pdf,
  [Parameter(Mandatory=$true)][string]$OutDir,
  [Parameter(Mandatory=$false)][int]$Dpi = 200,
  [Parameter(Mandatory=$false)][string]$Pages = 'all'
)
try {
  if (-not (Test-Path $Pdf)) { throw 'PDF file nahi mili' }
  if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  [Windows.Storage.StorageFile, Windows, ContentType=WindowsRuntime] | Out-Null
  [Windows.Storage.StorageFolder, Windows, ContentType=WindowsRuntime] | Out-Null
  [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows, ContentType=WindowsRuntime] | Out-Null
  [Windows.Storage.Streams.IRandomAccessStream, Windows, ContentType=WindowsRuntime] | Out-Null
  [Windows.Data.Pdf.PdfDocument, Windows, ContentType=WindowsRuntime] | Out-Null
  [Windows.Data.Pdf.PdfPageRenderOptions, Windows, ContentType=WindowsRuntime] | Out-Null
  [Windows.Graphics.Imaging.BitmapDecoder, Windows, ContentType=WindowsRuntime] | Out-Null
  [Windows.Graphics.Imaging.BitmapEncoder, Windows, ContentType=WindowsRuntime] | Out-Null
  $ext = [System.WindowsRuntimeSystemExtensions]
  $asTaskGeneric = ($ext.GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.GetGenericArguments().Count -eq 1 }) | Select-Object -First 1
  $asTaskAction = ($ext.GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.FullName -eq 'Windows.Foundation.IAsyncAction' }) | Select-Object -First 1
  if (-not $asTaskGeneric -or -not $asTaskAction) { throw 'WinRT AsTask nahi mila' }
  function Await-Op($op, $type) {
    $m = $asTaskGeneric.MakeGenericMethod($type)
    $t = $m.Invoke($null, @($op))
    return $t.Result
  }
  function Await-Action($op) {
    $t = $asTaskAction.Invoke($null, @($op))
    $t.Wait()
  }
  if ($Dpi -lt 72 -or $Dpi -gt 600) { $Dpi = 200 }
  $storageFile = Await-Op ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Pdf)) ([Windows.Storage.StorageFile])
  $doc = Await-Op ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) ([Windows.Data.Pdf.PdfDocument])
  $count = $doc.PageCount
  $wanted = @()
  if ($Pages -eq 'all') {
    for ($i = 0; $i -lt $count; $i++) { $wanted += $i }
  } else {
    $parts = $Pages.Split('-')
    if ($parts.Count -ne 2) { throw 'Pages range ghalat hai' }
    $a = 0
    $b = 0
    $parseOk = [int]::TryParse($parts[0].Trim(), [ref]$a) -and [int]::TryParse($parts[1].Trim(), [ref]$b)
    if (-not $parseOk -or $a -lt 1 -or $b -lt $a) { throw 'Pages range ghalat hai' }
    for ($p = $a; $p -le $b; $p++) {
      if ($p -le $count) { $wanted += ($p - 1) }
    }
  }
  if ($wanted.Count -eq 0) { throw 'Koi page render nahi hua' }
  $folder = Await-Op ([Windows.Storage.StorageFolder]::GetFolderFromPathAsync($OutDir)) ([Windows.Storage.StorageFolder])
  $scale = $Dpi / 72.0
  $done = 0
  foreach ($idx in $wanted) {
    $page = $doc.GetPage([uint32]$idx)
    $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $opts.DestinationWidth = [uint32]([Math]::Round($page.Dimensions.Width * $scale))
    $opts.DestinationHeight = [uint32]([Math]::Round($page.Dimensions.Height * $scale))
    $mem = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    Await-Action ($page.RenderToStreamAsync($mem, $opts))
    $dec = Await-Op ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($mem)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $num = ($idx + 1).ToString('0000')
    $outFile = Await-Op ($folder.CreateFileAsync(("page-" + $num + ".png"), [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])
    $outStream = Await-Op ($outFile.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)) ([Windows.Storage.Streams.IRandomAccessStream])
    $enc = Await-Op ([Windows.Graphics.Imaging.BitmapEncoder]::CreateForTranscodingAsync($outStream, $dec)) ([Windows.Graphics.Imaging.BitmapEncoder])
    $enc.IsThumbnailGenerated = $false
    Await-Action ($enc.FlushAsync())
    try { $outStream.Dispose() } catch {}
    try { $mem.Dispose() } catch {}
    $done++
  }
  [Console]::Out.WriteLine("PAGES:" + $done)
  [Console]::Out.WriteLine("OK")
  exit 0
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
`;
let docNupPsPath = null;
function ensureDocNupPsScript() {
  if (docNupPsPath && fs.existsSync(docNupPsPath)) return docNupPsPath;
  const p = path.join(os.tmpdir(), 'pakqr-pdf-render.ps1');
  try { fs.writeFileSync(p, DOC_NUP_PS_SCRIPT, 'utf8'); docNupPsPath = p; }
  catch (e) { docNupPsPath = null; }
  return docNupPsPath;
}
async function renderPdfPagesToPng(pdfPath, pagesSpec, dpi, outDir) {
  // pagesSpec: 'all' ya 'a-b'. Windows par WinRT script chalta
  // hai; PQ_NUP_MOCK_RENDER=1 sirf is VM ke test ke liye hai
  // (har page ki aik chhoti valid PNG - grouping/sheet flow
  // test ke liye; asal PC par ye kabhi ON nahi hota).
  if (process.env.PQ_NUP_MOCK_RENDER === '1') {
    try {
      fs.mkdirSync(outDir, { recursive: true });
      let pages = [];
      if (pagesSpec === 'all') pages = [1, 2, 3, 4, 5, 6];
      else {
        const m = /^(\d+)\s*-\s*(\d+)$/.exec(String(pagesSpec));
        if (!m) return { ok: false, error: 'Mock pages range ghalat' };
        for (let p = parseInt(m[1], 10); p <= parseInt(m[2], 10); p++) pages.push(p);
      }
      const png1 = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAwMCAO+ip1sAAAAASUVORK5CYII=', 'base64');
      const pngs = [];
      for (const pg of pages) {
        const f = path.join(outDir, 'page-' + String(pg).padStart(4, '0') + '.png');
        fs.writeFileSync(f, png1);
        pngs.push(f);
      }
      return { ok: true, pngs: pngs };
    } catch (e) { return { ok: false, error: 'Mock render fail: ' + String(e.message || e).slice(0, 160) }; }
  }
  const script = ensureDocNupPsScript();
  if (!script) return { ok: false, error: 'PDF render script temp folder me nahi ban saka' };
  const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, '-Pdf', pdfPath, '-OutDir', outDir, '-Dpi', String(dpi), '-Pages', pagesSpec];
  const r = await execFileOut('powershell.exe', args, 180000);
  if (!r.ok) return { ok: false, error: 'PDF pages render nahi ho sake: ' + String(r.error || '').slice(0, 200) };
  if (String(r.stdout || '').indexOf('OK') === -1) return { ok: false, error: 'PDF render script ne OK nahi diya' };
  try {
    const files = fs.readdirSync(outDir).filter((f) => /^page-\d+\.png$/i.test(f)).sort();
    if (!files.length) return { ok: false, error: 'PDF render se koi page image nahi bani' };
    return { ok: true, pngs: files.map((f) => path.join(outDir, f)) };
  } catch (e) { return { ok: false, error: 'Render output parha nahi ja saka: ' + String(e.message || e).slice(0, 160) }; }
}
async function printViaDocNupFromPdf(pdfPath, opts, effRange, docLabel) {
  const ppp = parseInt(opts.picturesPerPage, 10) || 1;
  if (ppp <= 1) return { ok: false, nupFailed: true, error: 'N-up off hai' };
  if (process.platform !== 'win32' && process.env.PQ_NUP_MOCK_RENDER !== '1') return { ok: false, nupFailed: true, error: 'N-up sirf Windows PC par chalta hai' };
  const totalPages = parseInt(opts.totalPages, 10) || 0;
  const parsed = parseDocNupPages(effRange, totalPages);
  if (parsed === null) return { ok: false, nupFailed: true, error: 'Range N-up ke liye durust nahi' };
  const pagesSpec = (parsed === 'all') ? 'all' : (parsed[0] + '-' + parsed[parsed.length - 1]);
  const tmpDir = path.join(os.tmpdir(), 'pakqr-nup-' + Date.now() + '-' + Math.floor(Math.random() * 100000));
  try {
    fs.mkdirSync(tmpDir, { recursive: true });
    const rr = await renderPdfPagesToPng(pdfPath, pagesSpec, docNupDpi(opts.quality), tmpDir);
    if (!rr.ok) return { ok: false, nupFailed: true, error: rr.error };
    if (!rr.pngs.length) return { ok: false, nupFailed: true, error: 'Render se koi page nahi mila' };
    const groups = groupDocPages(rr.pngs, ppp);
    let done = 0;
    for (const grp of groups) {
      const rs = await printViaImageSheet(grp, opts);
      if (!rs.ok) return { ok: false, nupFailed: true, error: 'Sheet print fail: ' + String(rs.error || '').slice(0, 200) };
      done++;
    }
    return { ok: true, method: docLabel + ' N-up (' + done + ' sheet' + (done > 1 ? 's' : '') + ', ' + ppp + '/sheet)', error: '', detail: '' };
  } catch (e) {
    return { ok: false, nupFailed: true, error: 'N-up error: ' + String(e.message || e).slice(0, 200) };
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch (e) {}
  }
}
async function printViaDocNupFromOffice(filePath, kind, opts, effRange) {
  // kind: 'word' | 'excel' (convertOfficeToPdf wale naam).
  // Converted PDF temp file me, render apne temp dir me; dono
  // finally me saaf. Convert fail / Office na ho => nupFailed
  // (caller purane raste par girta hai).
  const tmpPdf = path.join(os.tmpdir(), 'pakqr-nup-doc-' + Date.now() + '-' + Math.floor(Math.random() * 100000) + '.pdf');
  try {
    const conv = await convertOfficeToPdf(kind, filePath, tmpPdf);
    if (!conv.ok || !isValidPdfFile(tmpPdf)) return { ok: false, nupFailed: true, error: 'Office se PDF nahi ban saka (Office hona zaroori hai)' };
    return await printViaDocNupFromPdf(tmpPdf, opts, effRange, kind === 'word' ? 'Word' : 'Excel');
  } catch (e) {
    return { ok: false, nupFailed: true, error: 'Office N-up error: ' + String(e.message || e).slice(0, 200) };
  } finally {
    try { fs.unlinkSync(tmpPdf); } catch (e) {}
  }
}
function rangePrintError(range) {
  return 'Page range wala print is tareeqe se auto nahi ho sakta - file khol kar sirf pages ' + range + ' print karein (PDF ki range auto SumatraPDF se lagti hai - ye software ke andar hi maujood hai; agar SumatraPDF.exe folder me na mile to dobara install karein)';
}
// Phase 3.5: customer ne order par printer chuna ho aur wo Admin
// wale default se ALAG ho to jin raston se print us printer ko
// target NAHI kar sakta (Windows Shell verb), un par file auto
// print NAHI hoti - warna print ghalat printer se nikal jayega.
// Order card par printer ka badge laga hota hai; shop wahi file
// khol kar chune hue printer par khud print kar deta hai.
function printerMismatchError(opts) {
  const chosen = (opts && opts.orderPrinter) ? String(opts.orderPrinter) : '';
  return 'Customer ne printer "' + chosen + '" chuna hai - ye file is tareeqe se us printer par auto print nahi ho sakti, is liye roki gayi hai. File khol kar "' + chosen + '" par khud print karein (order par printer ka badge laga hai)';
}
// Phase 3.7: print fail ho aur wajah ye ho ke file ko sambhalne
// wala program PC par nahi hai, to asal error ke saath seedha
// us program ka naam bhi likha jata hai - shop ko khaali
// PowerShell error ('No application is associated') se wajah
// nahi milti thi (QP-1002 wala masla). Asal error text
// hamesha saath rehta hai; ye note sirf us ke aage judta hai.
// Phase 3.14 rule (aik jagah): poori range (1-kul) koi range
// hi nahi hai - printOneFile aur N-up dono isi se chalte hain.
function collapseFullRange(range, totalPages) {
  const r = String(range || '').trim();
  if (!r || !(totalPages > 0)) return r;
  const rm = /^(\d+)\s*-\s*(\d+)$/.exec(r);
  if (rm && parseInt(rm[1], 10) === 1 && parseInt(rm[2], 10) === totalPages) return '';
  return r;
}
// Phase 3.21: Document N-up ki koshish (order loop se). PDF
// seedha; Word/Excel pehle Office COM se PDF (convertOfficeToPdf
// - preview wala rasta, original kabhi change nahi hota), phir
// wahi N-up. Fail ho to { nupFailed: true } - caller purane
// raste par girta hai, order kabhi N-up ki wajah se fail nahi
// hota. PowerPoint is raste par NAHI aata (us ka spp alag hai).
async function tryDocNupForFile(fp, fname, o) {
  const effRange = collapseFullRange(o.range, o.totalPages);
  if (/\.pdf$/i.test(fname)) return printViaDocNupFromPdf(fp, o, effRange, 'PDF');
  if (/\.docx?$/i.test(fname)) return printViaDocNupFromOffice(fp, 'word', o, effRange);
  if (/\.xlsx?$/i.test(fname)) return printViaDocNupFromOffice(fp, 'excel', o, effRange);
  return { ok: false, nupFailed: true, error: 'Is file par N-up nahi lagta' };
}
async function printOneFile(filePath, opts) {
  if (process.platform !== 'win32') {
    return { ok: false, method: '', error: 'Auto-print sirf Windows PC par chalta hai - ye server abhi Windows par nahi hai.' };
  }
  const range = (opts && opts.range) ? String(opts.range).trim() : '';
  // Phase 3.14: poori range (1-kul) koi range hi nahi hai — poori
  // file print hoti hai (QP-1001: range 1-6, kul 6 pages, phir bhi
  // decline ho gaya tha). Partial range par purana decline waisa
  // hi hai (Office ki partial range auto nahi lag sakti). Kul
  // maloom na ho (0) to safe side: decline waisa hi.
  const totalPages = parseInt(opts && opts.totalPages, 10) || 0;
  const effRange = collapseFullRange(range, totalPages);
  const mismatch = !!(opts && opts.printerMismatch);
  const isImage = /\.(jpe?g|png|bmp|gif)$/i.test(filePath);
  if (isImage) return printViaImage(filePath, opts);
  const isPdf = /\.pdf$/i.test(filePath);
  if (isPdf) {
    if (range) {
      // Range sirf SumatraPDF se auto lag sakti hai - Shell par saare
      // pages nikal jayenge, is liye range wali file roki jati hai.
      if (fs.existsSync(SUMATRA_FILE)) return printViaSumatra(filePath, opts);
      return { ok: false, method: '', error: rangePrintError(range) };
    }
    if (fs.existsSync(SUMATRA_FILE)) {
      const r = await printViaSumatra(filePath, opts);
      if (r.ok) return r;
      // Sumatra fail - Shell sirf tab, jab customer ka chuna hua
      // printer default hi ho (Shell naam se target nahi karta).
      if (mismatch) return { ok: false, method: r.method, error: r.error + ' | ' + printerMismatchError(opts) };
      const r2 = await printViaShell(filePath, opts);
      return r2.ok ? r2 : { ok: false, method: r.method, error: r.error + ' | ' + r2.error };
    }
    if (mismatch) return { ok: false, method: '', error: printerMismatchError(opts) };
    return printViaShell(filePath, opts);
  }
  const isOffice = /\.(docx?|xlsx?|pptx?)$/i.test(filePath);
  if (isOffice) {
    const isPpt = /\.pptx?$/i.test(filePath);
    // Phase 3.22: PPT ki range COM se lagti hai (PrintOut From/To),
    // is liye PPT ko range-decline se chhoot hai; Word/Excel pehle jaisa.
    if (effRange && !isPpt) return { ok: false, method: '', error: rangePrintError(effRange) };
    const isWord = /\.docx?$/i.test(filePath);
    if (isWord) {
      // Phase 2.5.1: Word files pehle Word program ke zariye direct
      // (COM) print hoti hain; wo fail hon to purana Shell rasta.
      // COM wala rasta printer ko naam se target kar sakta hai;
      // Shell nahi kar sakta - customer ka printer default se alag
      // ho to COM fail par Shell par NAHI girte (ghalat printer se
      // print nikal jayega).
      const rc = await printViaWordCOM(filePath, opts);
      if (rc.ok) return rc;
      if (mismatch) return { ok: false, method: rc.method, error: rc.error + ' | ' + printerMismatchError(opts) };
      const rs = await printViaShell(filePath, opts);
      if (rs.ok) return rs;
      return { ok: false, method: rc.method, error: rc.error + ' | ' + rs.error + ' | Lagta hai is PC par MS Office ya LibreOffice nahi hai, ya us ne print ka option nahi diya - file Download/Open karke khud print karein. | MS Word ya LibreOffice Writer chahiye' };
    }
    if (isPpt) {
      // Phase 3.5: slides-per-page 1 se ziada ho to PowerPoint COM
      // wala rasta (Handouts). PowerPoint na mile to file auto
      // print NAHI hoti - khamoshi se 1-slide-per-page nahi
      // nikalta. Option 1 ho to pehle wala Shell rasta hi hai.
      // Phase 3.22: range ho to bhi COM hi (Shell range nahi lagata).
      const spp = parseInt(opts.slidesPerPage, 10) || 1;
      if (spp > 1 || effRange) {
        const rp = await printViaPowerPointCOM(filePath, opts);
        if (rp.ok) return rp;
        if (rp.noPowerPoint) {
          const asal = String(rp.comError || rp.error || '').slice(0, 200);
          var whyTxt, tailTxt;
          if (effRange) {
            whyTxt = 'Slides ki range (' + effRange + ') sirf MS PowerPoint se print ho sakti hai - is PC par PowerPoint COM nahi bana, is liye ye file auto print NAHI hui. | Range wali PowerPoint file ke liye PC par MS PowerPoint hona zaroori hai';
            tailTxt = ' | File ko orders folder se khol kar PowerPoint mein range (' + effRange + ') se khud print karein.';
          } else {
            whyTxt = 'Slides wala option (' + spp + ' slides ek page par) sirf MS PowerPoint se print ho sakta hai - is PC par PowerPoint COM nahi bana, is liye ye file auto print NAHI hui. | ' + spp + ' slides per page ke liye PC par MS PowerPoint hona zaroori hai';
            tailTxt = ' | File ko orders folder se khol kar PowerPoint ke Handouts (' + spp + ' slides per page) se khud print karein.';
          }
          return { ok: false, method: '', error: whyTxt + (asal ? ' | Asal wajah: ' + asal : '') + tailTxt };
        }
        return { ok: rp.ok, method: rp.method, error: rp.error, detail: rp.fullError || '' };
      }
      if (mismatch) return { ok: false, method: '', error: printerMismatchError(opts) };
      const rp2 = await printViaShell(filePath, opts);
      if (rp2.ok) return rp2;
      return { ok: false, method: rp2.method, error: rp2.error + ' | Lagta hai is PC par MS Office ya LibreOffice nahi hai, ya us ne print ka option nahi diya - file Download/Open karke khud print karein.' + (String(rp2.error).indexOf('No application is associated') !== -1 ? ' | Is file ke liye PC par MS PowerPoint ya LibreOffice Impress hona chahiye' : '') };
    }
    // Excel: Shell Print verb MS Office / LibreOffice deta hai -
    // naam se printer target nahi hota.
    if (mismatch) return { ok: false, method: '', error: printerMismatchError(opts) };
    const ro = await printViaShell(filePath, opts);
    if (ro.ok) return ro;
    return { ok: false, method: ro.method, error: ro.error + ' | Lagta hai is PC par MS Office ya LibreOffice nahi hai, ya us ne print ka option nahi diya - file Download/Open karke khud print karein.' + (String(ro.error).indexOf('No application is associated') !== -1 ? ' | MS Excel ya LibreOffice Calc chahiye' : '') };
  }
  if (mismatch) return { ok: false, method: '', error: printerMismatchError(opts) };
  return printViaShell(filePath, opts);
}
// Agent ke liye export: printOneFile(filePath, opts) aur
// tryDocNupForFile(filePath, fileName, opts).
// opts: { copies, colorMode ('BW'|'Color'), duplex, paperSize, range,
//         totalPages, slidesPerPage, nup, printerName, orderPrinter,
//         printerMismatch }
// addPrintLog yahan sirf console par likhta hai (agent khud Firebase
// par status likhta hai).
module.exports = { printOneFile, tryDocNupForFile, printTestPageImage, SUMATRA_FILE };
