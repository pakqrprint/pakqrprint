// Pak QR Print - QR encoder (khud ka, koi library/API nahi)
// Error correction level H (30% recovery) - taake QR ke centre mein chhota badge bhi scan ho.
// Offline v1 (PakQRPrint-v1.html) se liya gaya - wahi code jo OpenCV se
// decode karke verify hua tha (5/5 URLs exact).
// makeQR(text) -> matrix (true/false) ya null (agar text lamba ho). Level H, versions 1-8.
// drawQRToCanvas(canvas, text) -> canvas par QR bana deta hai.
var QR_SPEC = {
  1: { total: 26, data: 9, ec: 17, blocks: [9], align: [], rem: 0 },
  2: { total: 44, data: 16, ec: 28, blocks: [16], align: [6, 18], rem: 7 },
  3: { total: 70, data: 26, ec: 22, blocks: [13, 13], align: [6, 22], rem: 7 },
  4: { total: 100, data: 36, ec: 16, blocks: [9, 9, 9, 9], align: [6, 26], rem: 7 },
  5: { total: 134, data: 46, ec: 22, blocks: [11, 11, 12, 12], align: [6, 30], rem: 7 },
  6: { total: 172, data: 60, ec: 28, blocks: [15, 15, 15, 15], align: [6, 34], rem: 7 },
  7: { total: 196, data: 66, ec: 26, blocks: [13, 13, 13, 13, 14], align: [6, 22, 38], rem: 0 },
  8: { total: 242, data: 86, ec: 26, blocks: [14, 14, 14, 14, 15, 15], align: [6, 24, 42], rem: 0 }
};
var QR_FORMAT_H = ['001011010001001','001001110111110','001110011100111','001100111010000','000011101100010','000001001010101','000110100001100','000100000111011'];
var GF_EXP = [], GF_LOG = [];
(function() {
  var x = 1;
  for (var i = 0; i < 255; i++) { GF_EXP[i] = x; GF_LOG[x] = i; x <<= 1; if (x & 0x100) x ^= 0x11d; }
  for (var j = 255; j < 512; j++) GF_EXP[j] = GF_EXP[j - 255];
})();
function gfMul(a, b) { return (a && b) ? GF_EXP[GF_LOG[a] + GF_LOG[b]] : 0; }
function rsEcBytes(data, ecLen) {
  var gen = rsGenPolyRaw(ecLen);
  var res = new Array(ecLen).fill(0);
  for (var i = 0; i < data.length; i++) {
    var coef = data[i] ^ res.shift(); res.push(0);
    for (var j = 0; j < ecLen; j++) res[j] ^= gfMul(gen[j], coef);
  }
  return res;
}
function rsGenPolyRaw(deg) {
  // returns [coef of x^(deg-1), ..., coef of x^0] in GF(256)
  var poly = [1]; // coefficients highest-first
  for (var i = 0; i < deg; i++) {
    var next = new Array(poly.length + 1).fill(0);
    for (var j = 0; j < poly.length; j++) {
      next[j] ^= poly[j];                     // multiply by x
      next[j + 1] ^= gfMul(poly[j], GF_EXP[i]); // multiply by alpha^i
    }
    poly = next;
  }
  return poly.slice(1); // drop leading 1
}
function makeQR(text) {
  var bytes = [];
  var utf8 = unescape(encodeURIComponent(text));
  for (var i = 0; i < utf8.length; i++) bytes.push(utf8.charCodeAt(i));
  var version = 0, spec = null;
  for (var v = 1; v <= 8; v++) {
    if (4 + 8 + bytes.length * 8 <= QR_SPEC[v].data * 8) { version = v; spec = QR_SPEC[v]; break; }
  }
  if (!version) return null;
  // bit stream
  var bits = [];
  function pushBits(val, len) { for (var k = len - 1; k >= 0; k--) bits.push((val >>> k) & 1); }
  pushBits(4, 4); pushBits(bytes.length, 8);
  bytes.forEach(function(b){ pushBits(b, 8); });
  var capBits = spec.data * 8;
  var term = Math.min(4, capBits - bits.length);
  for (var t = 0; t < term; t++) bits.push(0);
  while (bits.length % 8) bits.push(0);
  var dataBytes = [];
  for (var bi = 0; bi < bits.length; bi += 8) {
    var byte = 0; for (var bj = 0; bj < 8; bj++) byte = (byte << 1) | bits[bi + bj];
    dataBytes.push(byte);
  }
  var pads = [0xEC, 0x11], pi = 0;
  while (dataBytes.length < spec.data) dataBytes.push(pads[pi++ % 2]);
  // blocks + EC
  var blocksData = [], blocksEc = [], offset = 0;
  spec.blocks.forEach(function(bl) {
    var chunk = dataBytes.slice(offset, offset + bl); offset += bl;
    blocksData.push(chunk);
    blocksEc.push(rsEcBytes(chunk, spec.ec));
  });
  var finalBytes = [];
  var maxBl = Math.max.apply(null, spec.blocks);
  for (var c = 0; c < maxBl; c++) blocksData.forEach(function(b){ if (c < b.length) finalBytes.push(b[c]); });
  for (var e = 0; e < spec.ec; e++) blocksEc.forEach(function(b){ if (e < b.length) finalBytes.push(b[e]); });
  var finalBits = [];
  finalBytes.forEach(function(b){ for (var k = 7; k >= 0; k--) finalBits.push((b >>> k) & 1); });
  for (var r = 0; r < spec.rem; r++) finalBits.push(0);
  // matrix
  var size = version * 4 + 17;
  var best = null, bestPenalty = Infinity;
  for (var mask = 0; mask < 8; mask++) {
    var m = buildMatrix(size, version, spec, finalBits, mask);
    var pen = qrPenalty(m, size);
    if (pen < bestPenalty) { bestPenalty = pen; best = m; }
  }
  return best;
}
function buildMatrix(size, version, spec, finalBits, mask) {
  var mod = [], func = [];
  for (var r = 0; r < size; r++) { mod.push(new Array(size).fill(false)); func.push(new Array(size).fill(false)); }
  function setF(x, y, v) { if (x >= 0 && x < size && y >= 0 && y < size) { mod[y][x] = !!v; func[y][x] = true; } }
  function finder(cx, cy) {
    for (var dy = -4; dy <= 4; dy++) for (var dx = -4; dx <= 4; dx++) {
      var x = cx + dx, y = cy + dy;
      if (x < 0 || x >= size || y < 0 || y >= size) continue;
      var dist = Math.max(Math.abs(dx), Math.abs(dy));
      setF(x, y, dist !== 2 && dist !== 4);
    }
  }
  finder(3, 3); finder(size - 4, 3); finder(3, size - 4);
  // timing
  for (var i = 0; i < size; i++) {
    if (!func[6][i]) setF(i, 6, i % 2 === 0);
    if (!func[i][6]) setF(6, i, i % 2 === 0);
  }
  // alignment
  if (spec.align.length) {
    var cs = spec.align;
    for (var a = 0; a < cs.length; a++) for (var b = 0; b < cs.length; b++) {
      var cx = cs[a], cy = cs[b];
      if ((cx === 6 && cy === 6) || (cx === 6 && cy === size - 7) || (cx === size - 7 && cy === 6)) continue;
      for (var dy2 = -2; dy2 <= 2; dy2++) for (var dx2 = -2; dx2 <= 2; dx2++) {
        setF(cx + dx2, cy + dy2, Math.max(Math.abs(dx2), Math.abs(dy2)) !== 1);
      }
    }
  }
  // reserve format info areas (mark function, values drawn later)
  function reserve(x, y) { if (x >= 0 && x < size && y >= 0 && y < size) func[y][x] = true; }
  for (var f = 0; f < 9; f++) { reserve(8, f); reserve(f, 8); }
  for (var f2 = 0; f2 < 8; f2++) { reserve(size - 1 - f2, 8); reserve(8, size - 1 - f2); }
  setF(8, size - 8, true); // dark module
  // version info (sirf version 7+): 18-bit BCH, generator 0x1F25, do copies
  if (version >= 7) {
    var vrem = version;
    for (var vi = 0; vi < 12; vi++) vrem = (vrem << 1) ^ ((vrem >> 11) * 0x1F25);
    var vbits = (version << 12) | vrem;
    for (var vj = 0; vj < 18; vj++) {
      var vb = ((vbits >> vj) & 1) === 1;
      setF(size - 11 + vj % 3, vj / 3 | 0, vb); // top-right copy
      setF(vj / 3 | 0, size - 11 + vj % 3, vb); // bottom-left copy
    }
  }
  // place data
  var idx = 0;
  for (var right = size - 1; right >= 1; right -= 2) {
    if (right === 6) right = 5;
    for (var vert = 0; vert < size; vert++) {
      for (var j = 0; j < 2; j++) {
        var x = right - j;
        var upward = ((right + 1) & 2) === 0;
        var y = upward ? size - 1 - vert : vert;
        if (!func[y][x] && idx < finalBits.length) { mod[y][x] = finalBits[idx++] === 1; }
      }
    }
  }
  // apply mask to non-function modules
  for (var my = 0; my < size; my++) for (var mx = 0; mx < size; mx++) {
    if (func[my][mx]) continue;
    var flip = false, rr = my, cc = mx;
    switch (mask) {
      case 0: flip = (rr + cc) % 2 === 0; break;
      case 1: flip = rr % 2 === 0; break;
      case 2: flip = cc % 3 === 0; break;
      case 3: flip = (rr + cc) % 3 === 0; break;
      case 4: flip = (Math.floor(rr / 2) + Math.floor(cc / 3)) % 2 === 0; break;
      case 5: flip = (rr * cc) % 2 + (rr * cc) % 3 === 0; break;
      case 6: flip = ((rr * cc) % 2 + (rr * cc) % 3) % 2 === 0; break;
      case 7: flip = ((rr + cc) % 2 + (rr * cc) % 3) % 2 === 0; break;
    }
    if (flip) mod[my][mx] = !mod[my][mx];
  }
  // format info (nayuki-style placement, bit i = LSB i of the 15-bit value)
  var data = parseInt(QR_FORMAT_H[mask], 2);
  function bitAt(bb, i) { return ((bb >>> i) & 1) === 1; }
  function setRaw(x, y, v) { mod[y][x] = !!v; func[y][x] = true; }
  for (var i1 = 0; i1 <= 5; i1++) setRaw(8, i1, bitAt(data, i1));
  setRaw(8, 7, bitAt(data, 6));
  setRaw(8, 8, bitAt(data, 7));
  setRaw(7, 8, bitAt(data, 8));
  for (var i2 = 9; i2 < 15; i2++) setRaw(14 - i2, 8, bitAt(data, i2));
  for (var i3 = 0; i3 < 8; i3++) setRaw(size - 1 - i3, 8, bitAt(data, i3));
  for (var i4 = 8; i4 < 15; i4++) setRaw(8, size - 15 + i4, bitAt(data, i4));
  setRaw(8, size - 8, true);
  return mod;
}
function qrPenalty(m, size) {
  var score = 0, r, c, i;
  // rule 1: runs
  function runPenalty(line) {
    var s = 0, runColor = line[0], runLen = 1;
    for (var k = 1; k < line.length; k++) {
      if (line[k] === runColor) runLen++;
      else { if (runLen >= 5) s += 3 + (runLen - 5); runColor = line[k]; runLen = 1; }
    }
    if (runLen >= 5) s += 3 + (runLen - 5);
    return s;
  }
  for (r = 0; r < size; r++) {
    score += runPenalty(m[r]);
    var col = []; for (c = 0; c < size; c++) col.push(m[c][r]);
    score += runPenalty(col);
  }
  // rule 2: 2x2 blocks
  for (r = 0; r < size - 1; r++) for (c = 0; c < size - 1; c++)
    if (m[r][c] === m[r][c + 1] && m[r][c] === m[r + 1][c] && m[r][c] === m[r + 1][c + 1]) score += 3;
  // rule 3: finder-like patterns
  function patternPenalty(line) {
    var s = 0;
    for (var k = 0; k + 7 <= line.length; k++) {
      if (line[k] && !line[k + 1] && line[k + 2] && line[k + 3] && line[k + 4] && !line[k + 5] && line[k + 6]) {
        var before = k >= 4 && !line[k - 1] && !line[k - 2] && !line[k - 3] && !line[k - 4];
        var after = k + 11 <= line.length && !line[k + 7] && !line[k + 8] && !line[k + 9] && !line[k + 10];
        if (before || after) s += 40;
      }
    }
    return s;
  }
  for (r = 0; r < size; r++) {
    score += patternPenalty(m[r]);
    var col2 = []; for (c = 0; c < size; c++) col2.push(m[c][r]);
    score += patternPenalty(col2);
  }
  // rule 4: dark proportion
  var dark = 0;
  for (r = 0; r < size; r++) for (c = 0; c < size; c++) if (m[r][c]) dark++;
  var total = size * size;
  var k = Math.floor((Math.abs(dark * 20 - total * 10) + total - 1) / total) - 1;
  score += k * 10;
  return score;
}
function drawQRToCanvas(cv, text) {
  var matrix = makeQR(text);
  if (!matrix) return false;
  var ctx = cv.getContext('2d');
  var n = matrix.length, quiet = 4;
  var cell = Math.max(2, Math.floor(240 / (n + quiet * 2)));
  var size = cell * (n + quiet * 2);
  cv.width = size; cv.height = size;
  ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, size, size);
  ctx.fillStyle = '#000000';
  for (var r = 0; r < n; r++) for (var c = 0; c < n; c++)
    if (matrix[r][c]) ctx.fillRect((c + quiet) * cell, (r + quiet) * cell, cell, cell);
  return true;
}
if (typeof module !== 'undefined' && module.exports) { module.exports = { makeQR: makeQR, drawQRToCanvas: drawQRToCanvas }; }
