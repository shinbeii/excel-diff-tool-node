/**
 * src/shapeDiff.js — So sánh shapes (吹き出し / callout / speech bubble) trong Excel.
 *
 * exceljs KHÔNG hỗ trợ shapes, nên ta parse trực tiếp XML drawing trong xlsx (zip):
 *   - xl/workbook.xml + xl/_rels/workbook.xml.rels  -> sheet name -> sheet path
 *   - xl/worksheets/sheetN.xml + .rels              -> drawing path
 *   - xl/drawings/drawingN.xml                       -> các <xdr:sp> shape
 *
 * Mỗi shape extract:
 *   { name, prstGeom, text, anchorFrom (col,row,colOff,rowOff), anchorTo,
 *     widthEMU, heightEMU, fill (hex), sheet, fileIdx }
 *
 * Matching: composite distance =
 *   0.50 * textDist (Levenshtein normalized) +
 *   0.20 * geomDist (prstGeom khác = 1) +
 *   0.20 * posDist  (|Δrow|+|Δcol|/30, clamp 0..1) +
 *   0.10 * fillDist (khác màu = 1)
 * Greedy global match -> 7 status:
 *   identical / text_changed / style_changed / moved / resized / added / removed.
 */
'use strict';

const fs = require('fs');
const sax = require('sax');
const JSZip = require('jszip');

// Các prstGeom kiểu callout/吹き出し (Excel "Callout" group + speech bubble Smart Art)
const CALLOUT_PRSTS = new Set([
  'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout',
  'cloudCallout',
  'borderCallout1', 'borderCallout2', 'borderCallout3',
  'accentCallout1', 'accentCallout2', 'accentCallout3',
  'accentBorderCallout1', 'accentBorderCallout2', 'accentBorderCallout3',
  'callout1', 'callout2', 'callout3',
  // Speech bubbles
  'rightArrowCallout', 'leftArrowCallout', 'upArrowCallout', 'downArrowCallout',
]);

function _parseRelsXml(xml) {
  // Trả về map { rId: target }.
  const out = {};
  const re = /<Relationship[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"/g;
  let m;
  while ((m = re.exec(xml)) !== null) out[m[1]] = m[2];
  return out;
}

function _parseWorkbookSheets(xml) {
  // <sheet name="X" sheetId="1" r:id="rId1"/>
  const out = []; // [{name, rId}]
  const re = /<sheet[^>]*\bname="([^"]+)"[^>]*\br:id="([^"]+)"/g;
  let m;
  while ((m = re.exec(xml)) !== null) out.push({ name: m[1], rId: m[2] });
  return out;
}

function _findSheetDrawing(sheetXml) {
  // <drawing r:id="rIdX"/>
  const m = sheetXml.match(/<drawing\s+[^>]*r:id="([^"]+)"/);
  return m ? m[1] : null;
}

function _resolvePath(base, relTarget) {
  // base: 'xl/worksheets/sheet1.xml', relTarget: '../drawings/drawing1.xml'
  // -> 'xl/drawings/drawing1.xml'
  const baseDir = base.split('/').slice(0, -1);
  const parts = relTarget.split('/');
  const stack = baseDir.slice();
  for (const p of parts) {
    if (p === '..') stack.pop();
    else if (p === '.' || p === '') continue;
    else stack.push(p);
  }
  return stack.join('/');
}

/** Levenshtein distance giữa 2 string. */
function _levenshtein(a, b) {
  if (a === b) return 0;
  if (!a) return b.length;
  if (!b) return a.length;
  const al = a.length, bl = b.length;
  let prev = new Array(bl + 1);
  let cur  = new Array(bl + 1);
  for (let j = 0; j <= bl; j++) prev[j] = j;
  for (let i = 1; i <= al; i++) {
    cur[0] = i;
    for (let j = 1; j <= bl; j++) {
      const cost = a.charCodeAt(i - 1) === b.charCodeAt(j - 1) ? 0 : 1;
      cur[j] = Math.min(
        cur[j - 1] + 1,
        prev[j] + 1,
        prev[j - 1] + cost,
      );
    }
    [prev, cur] = [cur, prev];
  }
  return prev[bl];
}

function _textDist(a, b) {
  a = (a || '').trim();
  b = (b || '').trim();
  if (!a && !b) return 0;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 0;
  return _levenshtein(a, b) / maxLen;
}

// ------------ Drawing XML parser (sax-based) ------------

/**
 * Parse 1 drawing XML, trả về danh sách shapes.
 * @returns {Array<Shape>}
 */
function _parseDrawingXml(xml) {
  const parser = sax.parser(true, { xmlns: false, lowercase: false });
  const shapes = [];
  // anchor state
  let inAnchor = false;
  let anchor = null;     // {from:{col,row,colOff,rowOff}, to:{...}, type:'two'|'one'|'absolute'}
  let cursor = null;     // 'from' | 'to' | null
  let pendingTag = '';   // current xdr:col/row/colOff/rowOff text
  // shape state (chỉ khi đang trong xdr:sp, không phải xdr:pic)
  let inSp = false;
  let curShape = null;
  // text body state
  let inTxBody = false;
  let inT = false;
  // xfrm/prstGeom/fill
  let inSpPr = false;
  let inXfrm = false;
  let inFillSolid = false;
  let inLn = false; // bỏ qua fill nằm trong <a:ln> (màu border)

  parser.onopentag = (node) => {
    const tag = node.name; // có namespace prefix vì xmlns:false
    const attrs = node.attributes;

    if (tag === 'xdr:twoCellAnchor' || tag === 'xdr:oneCellAnchor' || tag === 'xdr:absoluteAnchor') {
      inAnchor = true;
      anchor = { from: null, to: null, type: tag.replace('xdr:', '').replace('Anchor', '') };
      return;
    }
    if (!inAnchor) return;

    if (tag === 'xdr:from') { cursor = 'from'; anchor.from = { col: 0, row: 0, colOff: 0, rowOff: 0 }; return; }
    if (tag === 'xdr:to')   { cursor = 'to';   anchor.to   = { col: 0, row: 0, colOff: 0, rowOff: 0 }; return; }
    if (cursor && (tag === 'xdr:col' || tag === 'xdr:row' || tag === 'xdr:colOff' || tag === 'xdr:rowOff')) {
      pendingTag = tag.replace('xdr:', '');
      return;
    }

    if (tag === 'xdr:sp') {
      inSp = true;
      curShape = {
        name: '', prstGeom: '', text: '',
        anchorFrom: null, anchorTo: null,
        widthEMU: 0, heightEMU: 0,
        offX: 0, offY: 0,
        fill: null,
      };
      return;
    }
    if (!inSp) return;

    if (tag === 'xdr:cNvPr' && attrs && attrs.name) {
      curShape.name = attrs.name;
      return;
    }
    if (tag === 'xdr:spPr') { inSpPr = true; return; }
    if (inSpPr && tag === 'a:xfrm') { inXfrm = true; return; }
    if (inXfrm && tag === 'a:off' && attrs) {
      curShape.offX = parseInt(attrs.x || '0', 10);
      curShape.offY = parseInt(attrs.y || '0', 10);
      return;
    }
    if (inXfrm && tag === 'a:ext' && attrs) {
      curShape.widthEMU  = parseInt(attrs.cx || '0', 10);
      curShape.heightEMU = parseInt(attrs.cy || '0', 10);
      return;
    }
    if (inSpPr && tag === 'a:prstGeom' && attrs && attrs.prst) {
      curShape.prstGeom = attrs.prst;
      return;
    }
    if (inSpPr && tag === 'a:ln') { inLn = true; return; }
    if (inSpPr && !inLn && tag === 'a:solidFill') { inFillSolid = true; return; }
    if (inFillSolid && !inLn && tag === 'a:srgbClr' && attrs && attrs.val && !curShape.fill) {
      curShape.fill = ('#' + attrs.val).toUpperCase();
      return;
    }
    if (inFillSolid && !inLn && tag === 'a:schemeClr' && attrs && attrs.val && !curShape.fill) {
      curShape.fill = 'scheme:' + attrs.val;
      return;
    }
    if (tag === 'xdr:txBody') { inTxBody = true; return; }
    if (inTxBody && tag === 'a:t') { inT = true; return; }
  };

  parser.ontext = (txt) => {
    if (cursor && pendingTag) {
      const v = parseInt(String(txt).trim(), 10);
      if (!isNaN(v)) anchor[cursor][pendingTag] = v;
      return;
    }
    if (inT) curShape.text += txt;
  };

  parser.onclosetag = (tag) => {
    if (cursor && pendingTag && (tag === 'xdr:col' || tag === 'xdr:row' || tag === 'xdr:colOff' || tag === 'xdr:rowOff')) {
      pendingTag = '';
      return;
    }
    if (tag === 'xdr:from' || tag === 'xdr:to') { cursor = null; return; }
    if (tag === 'a:t') { inT = false; if (inTxBody) curShape.text += '\n'; return; }
    if (tag === 'xdr:txBody') { inTxBody = false; return; }
    if (tag === 'a:solidFill') { inFillSolid = false; return; }
    if (tag === 'a:ln') { inLn = false; return; }
    if (tag === 'a:xfrm') { inXfrm = false; return; }
    if (tag === 'xdr:spPr') { inSpPr = false; return; }
    if (tag === 'xdr:sp') {
      if (inSp && curShape) {
        curShape.text = curShape.text.replace(/\n+$/, '').replace(/\n+/g, '\n').trim();
        curShape.anchorFrom = anchor && anchor.from ? { ...anchor.from } : null;
        curShape.anchorTo   = anchor && anchor.to   ? { ...anchor.to   } : null;
        // chỉ giữ shape có prstGeom (loại bỏ pic + group)
        if (curShape.prstGeom) shapes.push(curShape);
      }
      inSp = false; curShape = null;
      return;
    }
    if (tag === 'xdr:twoCellAnchor' || tag === 'xdr:oneCellAnchor' || tag === 'xdr:absoluteAnchor') {
      inAnchor = false; anchor = null; return;
    }
  };

  parser.onerror = (err) => { /* ignore - drawing có thể có namespace lạ */ parser.error = null; parser.resume(); };
  parser.write(xml).close();
  return shapes;
}

// ------------ Extract ------------

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

function _anchorLabel(anchor) {
  if (!anchor) return '';
  return `${_colLetter(anchor.col || 0)}${(anchor.row || 0) + 1}`;
}

/**
 * Extract tất cả shapes từ 1 file xlsx, gom theo sheet name.
 * @returns {Promise<Object<string, Shape[]>>}
 */
async function extractShapes(filePath, fileIdx = 0) {
  const buf = await fs.promises.readFile(filePath);
  const zip = await JSZip.loadAsync(buf);

  // 1. workbook -> sheets
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const wbRelsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (!wbRelsFile) return {};
  const wbRelsXml = await wbRelsFile.async('string');
  const wbRels = _parseRelsXml(wbRelsXml);
  const sheets = _parseWorkbookSheets(wbXml); // [{name, rId}]

  const out = {};

  for (const s of sheets) {
    const sheetTarget = wbRels[s.rId];
    if (!sheetTarget) continue;
    // sheetTarget thường là 'worksheets/sheet1.xml' (relative tới xl/)
    const sheetPath = sheetTarget.startsWith('/') ? sheetTarget.slice(1) : `xl/${sheetTarget}`;
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) continue;
    const sheetXml = await sheetFile.async('string');
    const drawingRId = _findSheetDrawing(sheetXml);
    if (!drawingRId) { out[s.name] = []; continue; }

    const sheetRelsPath = sheetPath.replace(/([^/]+)$/, '_rels/$1.rels');
    const sheetRelsFile = zip.file(sheetRelsPath);
    if (!sheetRelsFile) { out[s.name] = []; continue; }
    const sheetRels = _parseRelsXml(await sheetRelsFile.async('string'));
    const drawingRel = sheetRels[drawingRId];
    if (!drawingRel) { out[s.name] = []; continue; }
    const drawingPath = _resolvePath(sheetPath, drawingRel);
    const drawingFile = zip.file(drawingPath);
    if (!drawingFile) { out[s.name] = []; continue; }

    const drawingXml = await drawingFile.async('string');
    const shapes = _parseDrawingXml(drawingXml).map((sp) => ({
      ...sp,
      sheet: s.name,
      fileIdx,
      anchor: _anchorLabel(sp.anchorFrom),
      isCallout: CALLOUT_PRSTS.has(sp.prstGeom),
      // EMU -> pixel xấp xỉ (1 EMU = 1/914400 inch, 96 dpi)
      widthPx:  Math.round(sp.widthEMU  / 9525),
      heightPx: Math.round(sp.heightEMU / 9525),
    }));
    out[s.name] = shapes;
  }
  return out;
}

// ------------ Diff logic ------------

function _composite(a, b) {
  const tDist = _textDist(a.text, b.text);
  const gDist = a.prstGeom === b.prstGeom ? 0 : 1;
  const dCol = Math.abs((a.anchorFrom?.col || 0) - (b.anchorFrom?.col || 0));
  const dRow = Math.abs((a.anchorFrom?.row || 0) - (b.anchorFrom?.row || 0));
  const pDist = Math.min(1, (dCol + dRow) / 30);
  const fDist = (a.fill || '') === (b.fill || '') ? 0 : 1;
  return 0.50 * tDist + 0.20 * gDist + 0.20 * pDist + 0.10 * fDist;
}

const MATCH_THRESH = 0.55;

function _greedyMatch(lists) {
  const n = lists.length;
  const base = lists[0];
  const result = Array.from({ length: n }, () => ({}));
  for (let fi = 1; fi < n; fi++) {
    const pairs = [];
    for (let i = 0; i < base.length; i++) {
      for (let j = 0; j < lists[fi].length; j++) {
        const d = _composite(base[i], lists[fi][j]);
        if (d <= MATCH_THRESH) pairs.push([d, i, j]);
      }
    }
    pairs.sort((a, b) => a[0] - b[0]);
    const ui = new Set(), uj = new Set();
    for (const [d, i, j] of pairs) {
      if (ui.has(i) || uj.has(j)) continue;
      result[fi][i] = { j, d };
      ui.add(i); uj.add(j);
    }
  }
  return result;
}

function _classify(items) {
  const present = items.filter(it => it != null);
  const n = items.length;
  if (present.length === 0) return 'added';
  if (present.length === 1) return items[0] != null ? 'removed' : 'added';
  if (present.length < n) return 'modified';

  const a = items[0];
  let textChanged = false, geomChanged = false, posChanged = false, sizeChanged = false, fillChanged = false;
  for (let k = 1; k < items.length; k++) {
    const b = items[k];
    if ((a.text || '') !== (b.text || '')) textChanged = true;
    if (a.prstGeom !== b.prstGeom) geomChanged = true;
    const dCol = Math.abs((a.anchorFrom?.col || 0) - (b.anchorFrom?.col || 0));
    const dRow = Math.abs((a.anchorFrom?.row || 0) - (b.anchorFrom?.row || 0));
    if (dCol + dRow > 0) posChanged = true;
    const dW = Math.abs(a.widthEMU - b.widthEMU);
    const dH = Math.abs(a.heightEMU - b.heightEMU);
    if (dW > a.widthEMU * 0.05 || dH > a.heightEMU * 0.05) sizeChanged = true;
    if ((a.fill || '') !== (b.fill || '')) fillChanged = true;
  }
  if (!textChanged && !geomChanged && !posChanged && !sizeChanged && !fillChanged) return 'identical';
  if (textChanged) return 'text_changed';
  if (geomChanged) return 'shape_changed';
  if (sizeChanged) return 'resized';
  if (posChanged) return 'moved';
  if (fillChanged) return 'style_changed';
  return 'modified';
}

function _buildEntries(lists) {
  const n = lists.length;
  if (n === 0 || lists.every(l => l.length === 0)) return [];
  const matches = _greedyMatch(lists);
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
    const status = _classify(items);
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
 * So sánh shapes giữa 2-3 workbook.
 * @param {string[]} paths
 * @returns {Promise<Object<string, {sheet, entries, summary}>>}
 */
async function diffShapes(paths) {
  const perFile = await Promise.all(paths.map((p, i) => extractShapes(p, i)));
  const allSheets = new Set();
  for (const d of perFile) for (const k of Object.keys(d)) allSheets.add(k);
  const results = {};
  for (const sname of Array.from(allSheets).sort()) {
    const lists = perFile.map(d => d[sname] || []);
    if (lists.every(l => l.length === 0)) continue;
    const entries = _buildEntries(lists);
    const summary = {
      identical: 0, text_changed: 0, shape_changed: 0,
      resized: 0, moved: 0, style_changed: 0, modified: 0,
      added: 0, removed: 0,
    };
    let calloutCount = 0;
    for (const e of entries) {
      summary[e.status] = (summary[e.status] || 0) + 1;
      if (e.items.some(it => it && it.isCallout)) calloutCount++;
    }
    summary.callouts = calloutCount;
    results[sname] = { sheet: sname, entries, summary };
  }
  return results;
}

module.exports = { diffShapes, extractShapes, CALLOUT_PRSTS };
