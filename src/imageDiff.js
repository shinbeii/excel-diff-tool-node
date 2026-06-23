/**
 * src/imageDiff.js — So sánh ảnh nhúng trong Excel (Node.js port của core/image_diff.py).
 *
 * Pipeline (giống bản Python 10/10):
 *   1. Trích xuất ảnh nhúng từ exceljs (workbook.media + worksheet drawings).
 *   2. Multi-hash ensemble: pHash 16x16 (DCT-II) + dHash 16x16 + aHash 8x8.
 *   3. Composite distance = 0.45*pHash + 0.40*dHash + 0.15*aHash, normalized [0,1].
 *   4. Greedy global assignment (sort tất cả cặp theo distance -> first-fit).
 *   5. Phân loại 7 status: identical / near_identical / similar / resized / moved / added / removed.
 *   6. Trả về score 0..100.
 */
'use strict';

const sharp = require('sharp');
const ExcelJS = require('exceljs');

const PHASH_SIZE = 16;
const DHASH_SIZE = 16;
const AHASH_SIZE = 8;

const THRESH_IDENTICAL      = 0.02;
const THRESH_NEAR_IDENTICAL = 0.05;
const THRESH_SIMILAR        = 0.18;
const THRESH_MATCH          = 0.30;

// ------------ Hash helpers ------------

/** Resize ảnh về size x size, grayscale, trả về Uint8Array (raw bytes). */
async function _grayResize(buffer, size) {
  return await sharp(buffer)
    .removeAlpha()
    .grayscale()
    .resize(size, size, { fit: 'fill', kernel: 'lanczos3' })
    .raw()
    .toBuffer();
}

/** aHash: average of grayscale pixels, bit = 1 nếu pixel >= mean. */
async function _aHash(buffer, size = AHASH_SIZE) {
  const px = await _grayResize(buffer, size);
  let sum = 0;
  for (let i = 0; i < px.length; i++) sum += px[i];
  const mean = sum / px.length;
  const bits = new Uint8Array(px.length);
  for (let i = 0; i < px.length; i++) bits[i] = px[i] >= mean ? 1 : 0;
  return bits;
}

/** dHash: resize (size+1) x size, so sánh pixel ngang -> size*size bit. */
async function _dHash(buffer, size = DHASH_SIZE) {
  const px = await sharp(buffer)
    .removeAlpha()
    .grayscale()
    .resize(size + 1, size, { fit: 'fill', kernel: 'lanczos3' })
    .raw()
    .toBuffer();
  const bits = new Uint8Array(size * size);
  for (let r = 0; r < size; r++) {
    for (let c = 0; c < size; c++) {
      const left  = px[r * (size + 1) + c];
      const right = px[r * (size + 1) + c + 1];
      bits[r * size + c] = left > right ? 1 : 0;
    }
  }
  return bits;
}

// Pre-compute DCT cosine matrix cho size x size (cache theo size).
const _dctMatrixCache = new Map();
function _dctMatrix(N) {
  if (_dctMatrixCache.has(N)) return _dctMatrixCache.get(N);
  const m = new Float64Array(N * N);
  for (let k = 0; k < N; k++) {
    for (let n = 0; n < N; n++) {
      m[k * N + n] = Math.cos(((2 * n + 1) * k * Math.PI) / (2 * N));
    }
  }
  _dctMatrixCache.set(N, m);
  return m;
}

/** DCT-II 2D: input NxN -> output NxN (chỉ cần top-left). */
function _dct2d(input, N) {
  const M = _dctMatrix(N);
  const tmp = new Float64Array(N * N);
  // rows: tmp[i,k] = sum_n M[k,n] * input[i,n]
  for (let i = 0; i < N; i++) {
    for (let k = 0; k < N; k++) {
      let s = 0;
      for (let n = 0; n < N; n++) s += M[k * N + n] * input[i * N + n];
      tmp[i * N + k] = s;
    }
  }
  const out = new Float64Array(N * N);
  // cols: out[k1,k2] = sum_i M[k1,i] * tmp[i,k2]
  for (let k1 = 0; k1 < N; k1++) {
    for (let k2 = 0; k2 < N; k2++) {
      let s = 0;
      for (let i = 0; i < N; i++) s += M[k1 * N + i] * tmp[i * N + k2];
      out[k1 * N + k2] = s;
    }
  }
  return out;
}

/** pHash: resize 4N x 4N, DCT-II, lấy top-left N x N (bỏ DC), bit = >= median. */
async function _pHash(buffer, size = PHASH_SIZE) {
  const big = size * 4;
  const px = await _grayResize(buffer, big);
  const input = new Float64Array(big * big);
  for (let i = 0; i < px.length; i++) input[i] = px[i];
  const dct = _dct2d(input, big);
  const low = new Float64Array(size * size);
  for (let r = 0; r < size; r++) {
    for (let c = 0; c < size; c++) low[r * size + c] = dct[r * big + c];
  }
  // median, bỏ DC (low[0])
  const arr = Array.from(low.slice(1)).sort((a, b) => a - b);
  const med = arr[Math.floor(arr.length / 2)];
  const bits = new Uint8Array(size * size);
  for (let i = 0; i < size * size; i++) bits[i] = low[i] >= med ? 1 : 0;
  return bits;
}

/** Hamming distance giữa 2 bit-vector cùng độ dài. */
function _hamming(a, b) {
  let d = 0;
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) if (a[i] !== b[i]) d++;
  return d;
}

function _bitsToHex(bits) {
  let hex = '';
  for (let i = 0; i < bits.length; i += 4) {
    const v = (bits[i] << 3) | ((bits[i + 1] || 0) << 2) | ((bits[i + 2] || 0) << 1) | (bits[i + 3] || 0);
    hex += v.toString(16);
  }
  return hex;
}

// ------------ Composite distance ------------

function _nh(bitsA, bitsB, size) {
  return _hamming(bitsA, bitsB) / (size * size);
}

function _composite(a, b) {
  const dp = _nh(a.phash, b.phash, PHASH_SIZE);
  const dd = _nh(a.dhash, b.dhash, DHASH_SIZE);
  const da = _nh(a.ahash, b.ahash, AHASH_SIZE);
  return 0.45 * dp + 0.40 * dd + 0.15 * da;
}

// ------------ Extract ------------

/** Convert col index (0-based) -> Excel letter (A, B, ..., Z, AA, ...). */
function _colLetter(col0) {
  let n = col0 + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Trích xuất tất cả ảnh nhúng từ 1 workbook.
 * Trả về { sheetName: [ImageInfo, ...] }.
 */
async function extractImages(filePath, fileIdx) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const out = {};
  const tasks = [];

  wb.eachSheet((ws) => {
    if (!out[ws.name]) out[ws.name] = [];
    const images = ws.getImages() || [];
    for (const img of images) {
      const media = wb.getImage(img.imageId);
      if (!media || !media.buffer) continue;
      const buf = Buffer.isBuffer(media.buffer) ? media.buffer : Buffer.from(media.buffer);
      let anchor = '';
      try {
        const tl = img.range && img.range.tl;
        if (tl) {
          // tl.col / tl.row có thể là số float (offset). Lấy floor + 1 cho row index 1-based.
          const c0 = Math.floor(tl.nativeCol != null ? tl.nativeCol : tl.col);
          const r0 = Math.floor(tl.nativeRow != null ? tl.nativeRow : tl.row);
          anchor = `${_colLetter(c0)}${r0 + 1}`;
        }
      } catch (_) { /* ignore */ }
      tasks.push({ sheet: ws.name, anchor, buffer: buf });
    }
  });

  // Parallel hash (sharp tự dùng libvips multi-thread, nhưng ta vẫn map song song).
  const infos = await Promise.all(tasks.map(async (t) => {
    try {
      const meta = await sharp(t.buffer).metadata();
      const [phash, dhash, ahash] = await Promise.all([
        _pHash(t.buffer, PHASH_SIZE),
        _dHash(t.buffer, DHASH_SIZE),
        _aHash(t.buffer, AHASH_SIZE),
      ]);
      const width = meta.width || 0;
      const height = meta.height || 1;
      return {
        fileIdx,
        sheet: t.sheet,
        anchor: t.anchor,
        width,
        height,
        aspect: width / Math.max(1, height),
        phash,
        dhash,
        ahash,
        phashHex: _bitsToHex(phash),
        dhashHex: _bitsToHex(dhash),
        ahashHex: _bitsToHex(ahash),
        buffer: t.buffer,
      };
    } catch (e) {
      return null;
    }
  }));

  for (const info of infos) {
    if (!info) continue;
    if (!out[info.sheet]) out[info.sheet] = [];
    out[info.sheet].push(info);
  }
  return out;
}

// ------------ Greedy global match ------------

function _greedyGlobalMatch(lists, thresh) {
  const n = lists.length;
  const base = lists[0];
  const result = Array.from({ length: n }, () => ({}));
  for (let fi = 1; fi < n; fi++) {
    const pairs = [];
    for (let i = 0; i < base.length; i++) {
      for (let j = 0; j < lists[fi].length; j++) {
        const d = _composite(base[i], lists[fi][j]);
        if (d <= thresh) pairs.push([d, i, j]);
      }
    }
    pairs.sort((a, b) => a[0] - b[0]);
    const usedI = new Set(), usedJ = new Set();
    for (const [d, i, j] of pairs) {
      if (usedI.has(i) || usedJ.has(j)) continue;
      result[fi][i] = { j, d };
      usedI.add(i); usedJ.add(j);
    }
  }
  return result;
}

function _classify(items, maxDist) {
  const present = items.filter(it => it != null);
  const n = items.length;
  if (present.length === 0) return 'added';
  if (present.length === 1) return items[0] != null ? 'removed' : 'added';
  const anchors = new Set(present.map(it => it.anchor));
  const sizes = new Set(present.map(it => `${it.width}x${it.height}`));
  const aspects = present.map(it => it.aspect);
  const aspectVar = Math.max(...aspects) - Math.min(...aspects);

  if (present.length < n) return 'similar';

  if (maxDist <= THRESH_IDENTICAL) {
    if (sizes.size > 1 && aspectVar > 0.02) return 'resized';
    if (anchors.size > 1) return 'moved';
    return 'identical';
  }
  if (maxDist <= THRESH_NEAR_IDENTICAL) return 'near_identical';
  if (aspectVar > 0.05 && maxDist <= THRESH_SIMILAR) return 'resized';
  return 'similar';
}

function _buildEntries(lists) {
  const n = lists.length;
  if (n === 0 || lists.every(l => l.length === 0)) return [];
  const matches = _greedyGlobalMatch(lists, THRESH_MATCH);
  const used = Array.from({ length: n }, () => new Set());
  const entries = [];
  const base = lists[0];

  for (let i = 0; i < base.length; i++) {
    const items = new Array(n).fill(null);
    items[0] = base[i];
    used[0].add(i);
    let maxD = 0;
    for (let fi = 1; fi < n; fi++) {
      const m = matches[fi][i];
      if (m) {
        items[fi] = lists[fi][m.j];
        used[fi].add(m.j);
        if (m.d > maxD) maxD = m.d;
      }
    }
    const status = _classify(items, maxD);
    const score = Math.max(0, Math.min(100, Math.round((1 - maxD) * 100)));
    entries.push({ status, items, distance: +maxD.toFixed(4), score });
  }

  for (let fi = 1; fi < n; fi++) {
    for (let j = 0; j < lists[fi].length; j++) {
      if (used[fi].has(j)) continue;
      const items = new Array(n).fill(null);
      items[fi] = lists[fi][j];
      entries.push({ status: 'added', items, distance: 1.0, score: 0 });
    }
  }
  return entries;
}

/**
 * So sánh ảnh nhúng giữa 2-3 workbook, gom theo sheet.
 * @param {string[]} paths
 * @returns {Promise<Object<string, {sheet: string, entries: Array, summary: Object}>>}
 */
async function diffImages(paths) {
  const perFile = await Promise.all(paths.map((p, i) => extractImages(p, i)));
  const allSheets = new Set();
  for (const d of perFile) for (const k of Object.keys(d)) allSheets.add(k);
  const results = {};
  const sortedSheets = Array.from(allSheets).sort();
  for (const sname of sortedSheets) {
    const lists = perFile.map(d => d[sname] || []);
    const entries = _buildEntries(lists);
    const summary = {
      identical: 0, near_identical: 0, similar: 0,
      resized: 0, moved: 0, added: 0, removed: 0,
    };
    for (const e of entries) summary[e.status] = (summary[e.status] || 0) + 1;
    results[sname] = { sheet: sname, entries, summary };
  }
  return results;
}

module.exports = {
  diffImages,
  extractImages,
  _composite,
  _pHash, _dHash, _aHash,
  THRESH_MATCH, THRESH_IDENTICAL, THRESH_NEAR_IDENTICAL, THRESH_SIMILAR,
};
