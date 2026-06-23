/**
 * src/excelDiff.js — So sánh nội dung 2-3 workbook Excel theo sheet và cell.
 * Port từ core/excel_diff.py (Python). Dùng exceljs.
 */
'use strict';

const ExcelJS = require('exceljs');

function _colLetter(col1) {
  let n = col1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function _isEmpty(v) {
  return v === null || v === undefined || v === '';
}

/**
 * exceljs trả về cell.value có thể là:
 *   - primitive (string/number/boolean/Date)
 *   - { formula: '=A1+B1', result: 42 }  (formula cell)
 *   - { richText: [...] }                 (rich text)
 *   - { sharedFormula, result }           (shared formula)
 *   - { error: '#REF!' }
 *   - { hyperlink, text }
 * Trả về [valueOnly, formulaOrNull].
 */
function _splitValueFormula(raw) {
  if (raw == null) return [null, null];
  if (typeof raw === 'object' && !(raw instanceof Date)) {
    if ('formula' in raw || 'sharedFormula' in raw) {
      const formula = raw.formula || raw.sharedFormula || '';
      const result = raw.result != null ? raw.result : '';
      return [result, formula ? '=' + formula : null];
    }
    if ('richText' in raw && Array.isArray(raw.richText)) {
      return [raw.richText.map(t => t.text || '').join(''), null];
    }
    if ('text' in raw) return [raw.text, null];
    if ('error' in raw) return [raw.error, null];
    if ('result' in raw) return [raw.result, null];
  }
  return [raw, null];
}

function _normalize(v) {
  if (v instanceof Date) return v.toISOString();
  return v;
}

function _fmtSignature(cell) {
  try {
    const f = cell.font || {};
    const fill = cell.fill || {};
    const bd = cell.border || {};
    return JSON.stringify([
      f.name || null, f.size || null, !!f.bold, !!f.italic,
      (f.color && (f.color.argb || f.color.rgb)) || null,
      (fill.fgColor && (fill.fgColor.argb || fill.fgColor.rgb)) || null,
      bd.left && bd.left.style || null,
      bd.right && bd.right.style || null,
      bd.top && bd.top.style || null,
      bd.bottom && bd.bottom.style || null,
    ]);
  } catch (_) {
    return '';
  }
}

/**
 * @param {string[]} paths - 2 hoặc 3 file Excel
 * @param {object} opts
 * @param {boolean} [opts.compareFormat]
 * @param {(pct:number, msg:string)=>void} [opts.progressCb]
 */
async function diffWorkbooks(paths, opts = {}) {
  const { compareFormat = false, progressCb } = opts;
  if (!Array.isArray(paths) || paths.length < 2 || paths.length > 3) {
    throw new Error('Chỉ hỗ trợ so sánh 2 hoặc 3 file.');
  }
  const wbs = await Promise.all(paths.map(async (p) => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(p);
    return wb;
  }));

  // Gom danh sách sheet (giữ thứ tự xuất hiện đầu tiên).
  const seen = new Set();
  const allSheets = [];
  for (const wb of wbs) {
    wb.eachSheet((ws) => {
      if (!seen.has(ws.name)) { seen.add(ws.name); allSheets.push(ws.name); }
    });
  }

  const sheetsDiff = [];
  const total = allSheets.length || 1;
  for (let i = 0; i < allSheets.length; i++) {
    const sname = allSheets[i];
    if (progressCb) progressCb(Math.floor((i * 100) / total), `Sheet: ${sname}`);
    const present = wbs.map(wb => wb.getWorksheet(sname) != null);
    const sd = _diffSheet(sname, wbs, present, compareFormat);
    sheetsDiff.push(sd);
  }

  const totalSummary = { added: 0, removed: 0, modified: 0, unchanged: 0 };
  for (const sd of sheetsDiff) {
    for (const k of Object.keys(totalSummary)) totalSummary[k] += (sd.summary[k] || 0);
  }
  if (progressCb) progressCb(100, 'Hoàn tất diff workbook');
  return { files: paths.slice(), sheets: sheetsDiff, summary: totalSummary };
}

function _diffSheet(name, wbs, present, compareFormat) {
  const sheets = wbs.map((wb, i) => present[i] ? wb.getWorksheet(name) : null);
  let maxRow = 0, maxCol = 0;
  for (const s of sheets) {
    if (!s) continue;
    if (s.rowCount > maxRow) maxRow = s.rowCount;
    if (s.columnCount > maxCol) maxCol = s.columnCount;
    // exceljs có thể báo rowCount/columnCount = 0 với sheet rỗng; dùng actualRowCount nếu có.
    if (s.actualRowCount > maxRow) maxRow = s.actualRowCount;
    if (s.actualColumnCount > maxCol) maxCol = s.actualColumnCount;
  }

  const summary = { added: 0, removed: 0, modified: 0, unchanged: 0 };
  const cells = [];

  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const vals = [];
      const forms = [];
      const sigs = [];
      for (const s of sheets) {
        if (!s) { vals.push(null); forms.push(null); sigs.push(''); continue; }
        const cell = s.getCell(r, c);
        const [v, f] = _splitValueFormula(cell.value);
        vals.push(_normalize(v));
        forms.push(f);
        sigs.push(compareFormat ? _fmtSignature(cell) : '');
      }

      if (vals.every(_isEmpty)) continue;

      const base = vals[0];
      const others = vals.slice(1);
      let status;
      if (_isEmpty(base) && others.some(v => !_isEmpty(v))) status = 'added';
      else if (!_isEmpty(base) && others.every(_isEmpty)) status = 'removed';
      else if (vals.every(v => _eq(v, base)) && forms.every(fo => fo === forms[0])) status = 'unchanged';
      else status = 'modified';

      let fmtChanged = false;
      if (compareFormat) {
        const set = new Set(sigs.filter(s => s));
        if (set.size > 1) fmtChanged = true;
      }
      if (status === 'unchanged' && !fmtChanged) continue;
      if (status === 'unchanged' && fmtChanged) status = 'modified';

      cells.push({
        coord: `${_colLetter(c)}${r}`,
        row: r, col: c, status,
        values: vals, formulas: forms, fmtChanged,
      });
      summary[status] = (summary[status] || 0) + 1;
    }
  }

  return {
    name,
    status: present.every(Boolean) ? 'common' : 'partial',
    cells, summary,
    maxRow, maxCol,
    filePresent: present,
  };
}

function _eq(a, b) {
  if (a === b) return true;
  if (a == null || b == null) return false;
  if (typeof a === 'number' && typeof b === 'number') return a === b;
  return String(a) === String(b);
}

module.exports = { diffWorkbooks };
