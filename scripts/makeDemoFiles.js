/**
 * scripts/makeDemoFiles.js — Sinh 2 file Excel demo có dữ liệu + ảnh nhúng.
 * Port từ scripts/make_demo_files.py.
 */
'use strict';

const path = require('path');
const fs = require('fs');
const sharp = require('sharp');
const ExcelJS = require('exceljs');

const ROOT = path.resolve(__dirname, '..');
const DEMO = path.join(ROOT, 'demo');
fs.mkdirSync(DEMO, { recursive: true });

/** SVG -> PNG buffer (mỗi lần gọi tạo buffer mới, không bị dedupe). */
async function makeImageBuffer(text, color, size = [220, 140]) {
  const [w, h] = size;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
    <rect width="100%" height="100%" fill="${color}"/>
    <rect x="4" y="4" width="${w - 8}" height="${h - 8}" fill="none" stroke="white" stroke-width="3"/>
    <text x="20" y="${h / 2 + 5}" font-family="Arial, sans-serif" font-size="22" fill="white" font-weight="bold">${text}</text>
  </svg>`;
  return await sharp(Buffer.from(svg)).png().toBuffer();
}

async function addImage(wb, ws, buf, anchorCell) {
  const id = wb.addImage({ buffer: buf, extension: 'png' });
  ws.addImage(id, `${anchorCell}:${anchorCell}`);
}

function populateCommon(ws) {
  ws.getCell('A1').value = 'Sản phẩm';
  ws.getCell('B1').value = 'Số lượng';
  ws.getCell('C1').value = 'Đơn giá';
  ws.getCell('D1').value = 'Thành tiền';
  for (const col of ['A', 'B', 'C', 'D']) {
    const c = ws.getCell(`${col}1`);
    c.font = { bold: true };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
  }
  const rows = [
    ['Bàn phím', 5, 250000],
    ['Chuột',    8, 120000],
    ['Màn hình', 2, 4500000],
    ['Tai nghe', 4, 350000],
  ];
  rows.forEach(([n, q, p], i) => {
    const r = i + 2;
    ws.getCell(`A${r}`).value = n;
    ws.getCell(`B${r}`).value = q;
    ws.getCell(`C${r}`).value = p;
    ws.getCell(`D${r}`).value = { formula: `B${r}*C${r}` };
  });
  ws.getCell('A7').value = 'Tổng';
  ws.getCell('A7').font = { bold: true };
  ws.getCell('D7').value = { formula: 'SUM(D2:D5)' };
}

async function buildFileA(filePath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('BangGia');
  populateCommon(ws);
  await addImage(wb, ws, await makeImageBuffer('LOGO A',  '#0a84ff'), 'F2');
  await addImage(wb, ws, await makeImageBuffer('CHART X', '#34c759'), 'F12');
  const ws2 = wb.addWorksheet('GhiChu');
  ws2.getCell('A1').value = 'Phiên bản 1.0';
  await wb.xlsx.writeFile(filePath);
}

async function buildFileB(filePath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('BangGia');
  populateCommon(ws);
  ws.getCell('B3').value = 10;
  ws.getCell('C5').value = 380000;
  ws.getCell('A6').value = 'Webcam';
  ws.getCell('B6').value = 3;
  ws.getCell('C6').value = 600000;
  ws.getCell('D6').value = { formula: 'B6*C6' };
  ws.getCell('D7').value = { formula: 'SUM(D2:D6)' };
  // LOGO cùng nội dung nhưng đổi vị trí F2 -> H2 (moved)
  await addImage(wb, ws, await makeImageBuffer('LOGO A',  '#0a84ff'), 'H2');
  // CHART đổi màu xanh -> cam (similar/near_identical)
  await addImage(wb, ws, await makeImageBuffer('CHART X', '#ff9500'), 'F12');
  // Ảnh hoàn toàn mới (added)
  await addImage(wb, ws, await makeImageBuffer('NEW PIC', '#af52de'), 'F22');
  const ws2 = wb.addWorksheet('GhiChu');
  ws2.getCell('A1').value = 'Phiên bản 1.1';
  const ws3 = wb.addWorksheet('Extra');
  ws3.getCell('A1').value = 'Sheet mới ở file B';
  await wb.xlsx.writeFile(filePath);
}

async function verify(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const counts = {};
  wb.eachSheet(ws => { counts[ws.name] = ws.getImages().length; });
  console.log(`  ${filePath} -> images per sheet: ${JSON.stringify(counts)}`);
  return counts;
}

async function main() {
  const a = path.join(DEMO, 'file_a.xlsx');
  const b = path.join(DEMO, 'file_b.xlsx');
  await buildFileA(a);
  await buildFileB(b);
  console.log(`[OK] ${a}`);
  console.log(`[OK] ${b}`);
  console.log('Verify:');
  const ca = await verify(a);
  const cb = await verify(b);
  const expA = JSON.stringify({ BangGia: 2, GhiChu: 0 });
  const expB = JSON.stringify({ BangGia: 3, GhiChu: 0, Extra: 0 });
  if (JSON.stringify(ca) !== expA) throw new Error(`file_a expected ${expA}, got ${JSON.stringify(ca)}`);
  if (JSON.stringify(cb) !== expB) throw new Error(`file_b expected ${expB}, got ${JSON.stringify(cb)}`);
  console.log('[OK] Số ảnh nhúng đúng như mong đợi.');
  console.log('Chạy thử:');
  console.log(`  node src/cli.js ${a} ${b} --out ${path.join(DEMO, 'report.xlsx')}`);
}

main().catch(err => { console.error(err); process.exit(1); });
