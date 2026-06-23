/**
 * scripts/makeDemoShapes.js — Sinh demo có khung chú thích (callout shapes).
 *
 * exceljs không hỗ trợ shapes, nên ta:
 *   1. Dùng exceljs tạo workbook trắng với sheet 'Notes'.
 *   2. Mở lại bằng jszip, inject drawing1.xml + .rels + [Content_Types] override
 *      + thêm <drawing r:id="rId1"/> vào sheet1.xml.
 *
 * Sinh ra:
 *   demo/shapes_a.xlsx — 3 callout: wedgeRectCallout, wedgeRoundRectCallout, cloudCallout
 *   demo/shapes_b.xlsx — 1 cùng nội dung (identical), 1 đổi text, 1 di chuyển (moved),
 *                         1 ảnh shape mới (added), bỏ cloudCallout (removed).
 */
'use strict';

const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const ROOT = path.resolve(__dirname, '..');
const DEMO = path.join(ROOT, 'demo');
fs.mkdirSync(DEMO, { recursive: true });

function escXml(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildShapeXml(sp, idx) {
  const { prstGeom, text, fromCol, fromRow, toCol, toRow, fill } = sp;
  return `<xdr:twoCellAnchor editAs="oneCell">
  <xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:sp macro="" textlink="">
    <xdr:nvSpPr>
      <xdr:cNvPr id="${idx + 2}" name="Callout ${idx + 1}"/>
      <xdr:cNvSpPr/>
    </xdr:nvSpPr>
    <xdr:spPr>
      <a:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="2000000" cy="1000000"/>
      </a:xfrm>
      <a:prstGeom prst="${prstGeom}"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>
      <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
    </xdr:spPr>
    <xdr:txBody>
      <a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/>
      <a:lstStyle/>
      <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="vi-VN" sz="1400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>${escXml(text)}</a:t></a:r></a:p>
    </xdr:txBody>
  </xdr:sp>
  <xdr:clientData/>
</xdr:twoCellAnchor>`;
}

function buildDrawingXml(shapes) {
  const body = shapes.map((s, i) => buildShapeXml(s, i)).join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
${body}
</xdr:wsDr>`;
}

async function buildShapesFile(filePath, shapes) {
  // 1. base workbook bằng exceljs
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Notes');
  ws.getCell('A1').value = 'Ghi chú với khung chú thích (callout)';
  ws.getCell('A1').font = { bold: true, size: 13 };
  ws.getCell('A2').value = 'Mỗi shape phía dưới là một speech bubble.';
  for (let r = 1; r <= 30; r++) ws.getRow(r).height = 22;
  await wb.xlsx.writeFile(filePath);

  // 2. mở lại bằng jszip + inject
  const buf = await fs.promises.readFile(filePath);
  const zip = await JSZip.loadAsync(buf);

  zip.file('xl/drawings/drawing1.xml', buildDrawingXml(shapes));

  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', relsXml);

  let sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  if (!sheetXml.includes('xmlns:r=')) {
    sheetXml = sheetXml.replace(
      '<worksheet ',
      '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    );
  }
  if (!/<drawing\s/.test(sheetXml)) {
    sheetXml = sheetXml.replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>');
  }
  zip.file('xl/worksheets/sheet1.xml', sheetXml);

  let ct = await zip.file('[Content_Types].xml').async('string');
  if (!ct.includes('drawing+xml')) {
    ct = ct.replace(
      '</Types>',
      '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
    );
  }
  zip.file('[Content_Types].xml', ct);

  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  await fs.promises.writeFile(filePath, out);
}

const shapesA = [
  { prstGeom: 'wedgeRectCallout',      text: 'Cần kiểm tra',
    fromCol: 1, fromRow: 3,  toCol: 4, toRow: 6,  fill: 'FFD60A' },
  { prstGeom: 'wedgeRoundRectCallout', text: 'Lưu ý quan trọng!',
    fromCol: 1, fromRow: 9,  toCol: 4, toRow: 12, fill: '34C759' },
  { prstGeom: 'cloudCallout',          text: 'Ý tưởng mới',
    fromCol: 1, fromRow: 15, toCol: 4, toRow: 18, fill: '5AC8FA' },
];

const shapesB = [
  // 1: cùng prstGeom + cùng vị trí, đổi text -> text_changed
  { prstGeom: 'wedgeRectCallout',      text: 'Cần kiểm tra GẤP!',
    fromCol: 1, fromRow: 3,  toCol: 4, toRow: 6,  fill: 'FFD60A' },
  // 2: cùng nội dung, di chuyển sang phải -> moved
  { prstGeom: 'wedgeRoundRectCallout', text: 'Lưu ý quan trọng!',
    fromCol: 7, fromRow: 9,  toCol: 10, toRow: 12, fill: '34C759' },
  // 3: shape mới hoàn toàn -> added
  { prstGeom: 'wedgeEllipseCallout',   text: 'Bubble mới',
    fromCol: 1, fromRow: 21, toCol: 4, toRow: 24, fill: 'AF52DE' },
  // (cloudCallout của file A không xuất hiện -> removed)
];

async function main() {
  const a = path.join(DEMO, 'shapes_a.xlsx');
  const b = path.join(DEMO, 'shapes_b.xlsx');
  await buildShapesFile(a, shapesA);
  await buildShapesFile(b, shapesB);
  console.log(`[OK] ${a} (${shapesA.length} shapes)`);
  console.log(`[OK] ${b} (${shapesB.length} shapes)`);
  console.log('Chạy thử:');
  console.log(`  node src/cli.js ${a} ${b} --out ${path.join(DEMO, 'shapes_report.xlsx')} --html ${path.join(DEMO, 'shapes_report.html')}`);
}

main().catch(err => { console.error(err); process.exit(1); });
