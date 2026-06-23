/**
 * src/report.js — Xuất báo cáo Excel + HTML từ kết quả diff.
 * Port từ core/report.py.
 */
'use strict';

const ExcelJS = require('exceljs');
const sharp = require('sharp');
const fs = require('fs');

const FILL = {
  added:          'FFC6EFCE',
  removed:        'FFFFC7CE',
  modified:       'FFFFEB9C',
  unchanged:      'FFFFFFFF',
  identical:      'FFFFFFFF',
  near_identical: 'FFE2EFDA',
  similar:        'FFFFEB9C',
  resized:        'FFD9E1F2',
  moved:          'FFBDD7EE',
  // Shape statuses
  text_changed:   'FFFFEB9C',
  shape_changed:  'FFFCE4D6',
  style_changed:  'FFE7E6E6',
};

function _safeSheetName(name) {
  const bad = /[\[\]:*?/\\]/g;
  return name.replace(bad, '_').slice(0, 28);
}

function _fillFor(status) {
  return {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: FILL[status] || FILL.unchanged },
  };
}

function _escapeHtml(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * @param {object} diff - kết quả từ diffWorkbooks
 * @param {object} imageResults - kết quả từ diffImages
 * @param {string} outPath
 * @param {object} [extra] - { shapeResults }
 */
async function exportExcelReport(diff, imageResults, outPath, extra = {}) {
  const shapeResults = extra.shapeResults || null;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'ExcelDiffTool';
  wb.created = new Date();

  // ----- Summary -----
  const ws = wb.addWorksheet('Summary');
  ws.getCell('A1').value = 'Excel Diff Report';
  ws.getCell('A1').font = { bold: true, size: 14 };
  ws.getCell('A3').value = 'Generated at';
  ws.getCell('B3').value = new Date().toISOString().slice(0, 19);
  ws.getCell('A4').value = 'Files';
  diff.files.forEach((p, i) => { ws.getCell(4, 2 + i).value = p; });

  const headers = ['Sheet', 'Added', 'Removed', 'Modified', 'Unchanged'];
  headers.forEach((h, i) => {
    const c = ws.getCell(6, i + 1); c.value = h; c.font = { bold: true };
  });
  let r = 7;
  for (const sd of diff.sheets) {
    ws.getCell(r, 1).value = sd.name;
    ws.getCell(r, 2).value = sd.summary.added || 0;
    ws.getCell(r, 3).value = sd.summary.removed || 0;
    ws.getCell(r, 4).value = sd.summary.modified || 0;
    ws.getCell(r, 5).value = sd.summary.unchanged || 0;
    r++;
  }
  for (let col = 1; col <= 5; col++) ws.getColumn(col).width = 18;

  // ----- Per-sheet diff -----
  for (const sd of diff.sheets) {
    if (!sd.cells.length) continue;
    const s = wb.addWorksheet(`diff_${_safeSheetName(sd.name)}`);
    const nFiles = diff.files.length;
    const cols = ['Cell', 'Status'];
    for (let i = 0; i < nFiles; i++) cols.push(`File ${i + 1}`);
    cols.push('Formula(s)');
    cols.forEach((h, i) => {
      const c = s.getCell(1, i + 1); c.value = h; c.font = { bold: true };
    });
    sd.cells.forEach((cd, idx) => {
      const ri = idx + 2;
      s.getCell(ri, 1).value = cd.coord;
      s.getCell(ri, 2).value = cd.status;
      cd.values.forEach((v, j) => {
        s.getCell(ri, 3 + j).value = v == null ? '' : String(v);
      });
      s.getCell(ri, 3 + nFiles).value =
        cd.formulas.map(f => f || '').join(' | ');
      const fill = _fillFor(cd.status);
      for (let col = 1; col <= 3 + nFiles; col++) {
        s.getCell(ri, col).fill = fill;
      }
    });
    for (let col = 1; col <= cols.length; col++) s.getColumn(col).width = 22;
    s.views = [{ state: 'frozen', ySplit: 1 }];
  }

  // ----- Images sheet -----
  if (imageResults && Object.keys(imageResults).length > 0) {
    const ims = wb.addWorksheet('Images');
    const head = ['Sheet', 'Status', 'Score', 'Anchors', 'Sizes', 'pHash', 'Thumbnail (file1)'];
    head.forEach((h, i) => {
      const c = ims.getCell(1, i + 1); c.value = h; c.font = { bold: true };
    });
    let row = 2;
    for (const [sname, res] of Object.entries(imageResults)) {
      for (const e of res.entries) {
        ims.getCell(row, 1).value = sname;
        ims.getCell(row, 2).value = e.status;
        ims.getCell(row, 3).value = `${e.score || 0}%`;
        ims.getCell(row, 4).value = e.items.map(it => it ? it.anchor : '-').join(' / ');
        ims.getCell(row, 5).value = e.items.map(it => it ? `${it.width}x${it.height}` : '-').join(' / ');
        ims.getCell(row, 6).value = e.items.map(it => it ? it.phashHex : '-').join(' / ');
        const first = e.items.find(it => it != null);
        if (first && first.buffer) {
          try {
            const thumb = await sharp(first.buffer)
              .resize(140, 140, { fit: 'inside' })
              .png()
              .toBuffer();
            const imgId = wb.addImage({ buffer: thumb, extension: 'png' });
            ims.addImage(imgId, {
              tl: { col: 6, row: row - 1 },
              ext: { width: 140, height: 100 },
              editAs: 'oneCell',
            });
            ims.getRow(row).height = 80;
          } catch (_) { /* ignore */ }
        }
        const fill = _fillFor(e.status);
        for (let col = 1; col <= 6; col++) ims.getCell(row, col).fill = fill;
        row++;
      }
    }
    const widths = [16, 14, 8, 22, 18, 38, 22];
    widths.forEach((w, i) => { ims.getColumn(i + 1).width = w; });
    ims.views = [{ state: 'frozen', ySplit: 1 }];
  }

  // ----- Shapes sheet (吹き出し / callout) -----
  if (shapeResults && Object.keys(shapeResults).length > 0) {
    const shs = wb.addWorksheet('Shapes');
    const head = ['Sheet', 'Status', 'Score', 'Type', 'Anchors', 'Sizes (px)', 'Fill', 'Text (file1)', 'Text (file2)', 'Text (file3)'];
    const nFiles = diff.files.length;
    const headers = head.slice(0, 7 + nFiles);
    headers.forEach((h, i) => {
      const c = shs.getCell(1, i + 1); c.value = h; c.font = { bold: true };
    });
    let row = 2;
    for (const [sname, res] of Object.entries(shapeResults)) {
      for (const e of res.entries) {
        shs.getCell(row, 1).value = sname;
        shs.getCell(row, 2).value = e.status;
        shs.getCell(row, 3).value = `${e.score || 0}%`;
        shs.getCell(row, 4).value = e.items.map(it => it ? it.prstGeom : '-').join(' / ');
        shs.getCell(row, 5).value = e.items.map(it => it ? it.anchor : '-').join(' / ');
        shs.getCell(row, 6).value = e.items.map(it => it ? `${it.widthPx}x${it.heightPx}` : '-').join(' / ');
        shs.getCell(row, 7).value = e.items.map(it => it ? (it.fill || '-') : '-').join(' / ');
        for (let k = 0; k < nFiles; k++) {
          shs.getCell(row, 8 + k).value = e.items[k] ? (e.items[k].text || '') : '';
        }
        const fill = _fillFor(e.status);
        for (let col = 1; col <= 7 + nFiles; col++) shs.getCell(row, col).fill = fill;
        row++;
      }
    }
    const widths = [14, 14, 8, 22, 14, 14, 12, 26, 26, 26];
    widths.slice(0, 7 + nFiles).forEach((w, i) => { shs.getColumn(i + 1).width = w; });
    shs.views = [{ state: 'frozen', ySplit: 1 }];
  }

  await wb.xlsx.writeFile(outPath);
  return outPath;
}

/**
 * Xuất HTML report đơn giản.
 */
async function exportHtmlReport(diff, imageResults, outPath, extra = {}) {
  const shapeResults = extra.shapeResults || null;
  const parts = [
    '<!doctype html><html><head><meta charset="utf-8">',
    '<title>Excel Diff Report</title>',
    '<style>body{font-family:-apple-system,Segoe UI,Arial;margin:24px;color:#222}',
    'table{border-collapse:collapse;margin-bottom:24px}',
    'td,th{border:1px solid #bbb;padding:4px 8px;font-size:13px}',
    '.added{background:#C6EFCE}.removed{background:#FFC7CE}',
    '.modified{background:#FFEB9C}.moved{background:#BDD7EE}',
    '.near_identical{background:#E2EFDA}.similar{background:#FFEB9C}',
    '.resized{background:#D9E1F2}',
    '.text_changed{background:#FFEB9C}.shape_changed{background:#FCE4D6}',
    '.style_changed{background:#E7E6E6}.identical{background:#FFFFFF}',
    'h1{margin-top:0}h2{border-bottom:2px solid #333;padding-bottom:4px}',
    'code{background:#f3f3f3;padding:1px 4px;border-radius:3px}</style>',
    '</head><body>', '<h1>Excel Diff Report</h1>',
    `<p>Generated: ${_escapeHtml(new Date().toISOString().slice(0, 19))}</p>`,
    '<p>Files: ' + diff.files.map(p => `<code>${_escapeHtml(p)}</code>`).join(' &nbsp;vs&nbsp; ') + '</p>',
    '<h2>Summary</h2><table>',
    '<tr><th>Sheet</th><th>Added</th><th>Removed</th><th>Modified</th><th>Unchanged</th></tr>',
  ];
  for (const sd of diff.sheets) {
    parts.push(
      `<tr><td>${_escapeHtml(sd.name)}</td>` +
      `<td>${sd.summary.added || 0}</td>` +
      `<td>${sd.summary.removed || 0}</td>` +
      `<td>${sd.summary.modified || 0}</td>` +
      `<td>${sd.summary.unchanged || 0}</td></tr>`
    );
  }
  parts.push('</table>');

  for (const sd of diff.sheets) {
    if (!sd.cells.length) continue;
    parts.push(`<h2>Sheet: ${_escapeHtml(sd.name)}</h2><table><tr><th>Cell</th><th>Status</th>`);
    for (let i = 0; i < diff.files.length; i++) parts.push(`<th>File ${i + 1}</th>`);
    parts.push('<th>Formula</th></tr>');
    for (const cd of sd.cells) {
      parts.push(`<tr class="${cd.status}"><td>${cd.coord}</td><td>${cd.status}</td>`);
      for (const v of cd.values) parts.push(`<td>${_escapeHtml(v == null ? '' : String(v))}</td>`);
      parts.push(`<td>${_escapeHtml(cd.formulas.map(f => f || '').join(' | '))}</td></tr>`);
    }
    parts.push('</table>');
  }

  if (imageResults && Object.keys(imageResults).length > 0) {
    parts.push('<h2>Images</h2><table>',
      '<tr><th>Sheet</th><th>Status</th><th>Score</th><th>Anchors</th><th>Sizes</th><th>pHash</th></tr>');
    for (const [sname, res] of Object.entries(imageResults)) {
      for (const e of res.entries) {
        parts.push(
          `<tr class="${e.status}">` +
          `<td>${_escapeHtml(sname)}</td>` +
          `<td>${e.status}</td>` +
          `<td>${e.score || 0}%</td>` +
          `<td>${_escapeHtml(e.items.map(it => it ? it.anchor : '-').join(' / '))}</td>` +
          `<td>${_escapeHtml(e.items.map(it => it ? `${it.width}x${it.height}` : '-').join(' / '))}</td>` +
          `<td>${_escapeHtml(e.items.map(it => it ? it.phashHex : '-').join(' / '))}</td>` +
          '</tr>'
        );
      }
    }
    parts.push('</table>');
  }

  if (shapeResults && Object.keys(shapeResults).length > 0) {
    parts.push('<h2>Shapes (吹き出し / callout)</h2><table>',
      '<tr><th>Sheet</th><th>Status</th><th>Score</th><th>Type</th><th>Anchors</th><th>Sizes (px)</th><th>Fill</th><th>Text</th></tr>');
    for (const [sname, res] of Object.entries(shapeResults)) {
      for (const e of res.entries) {
        parts.push(
          `<tr class="${e.status}">` +
          `<td>${_escapeHtml(sname)}</td>` +
          `<td>${e.status}</td>` +
          `<td>${e.score || 0}%</td>` +
          `<td>${_escapeHtml(e.items.map(it => it ? it.prstGeom : '-').join(' / '))}</td>` +
          `<td>${_escapeHtml(e.items.map(it => it ? it.anchor : '-').join(' / '))}</td>` +
          `<td>${_escapeHtml(e.items.map(it => it ? `${it.widthPx}x${it.heightPx}` : '-').join(' / '))}</td>` +
          `<td>${_escapeHtml(e.items.map(it => it ? (it.fill || '-') : '-').join(' / '))}</td>` +
          `<td>${_escapeHtml(e.items.map(it => it ? (it.text || '') : '').join(' → '))}</td>` +
          '</tr>'
        );
      }
    }
    parts.push('</table>');
  }

  parts.push('</body></html>');
  await fs.promises.writeFile(outPath, parts.join(''), 'utf-8');
  return outPath;
}

module.exports = { exportExcelReport, exportHtmlReport };
