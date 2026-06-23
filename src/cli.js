#!/usr/bin/env node
/**
 * src/cli.js — CLI entry point.
 *
 *   node src/cli.js file_a.xlsx file_b.xlsx [file_c.xlsx] [--out report.xlsx]
 *                  [--html report.html] [--no-images] [--no-shapes] [--format]
 */
'use strict';

const path = require('path');
const { diffWorkbooks } = require('./excelDiff');
const { diffImages } = require('./imageDiff');
const { diffShapes } = require('./shapeDiff');
const { exportExcelReport, exportHtmlReport } = require('./report');

function parseArgs(argv) {
  const files = [];
  const opts = { out: null, html: null, images: true, shapes: true, compareFormat: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--out') opts.out = argv[++i];
    else if (a === '--html') opts.html = argv[++i];
    else if (a === '--no-images') opts.images = false;
    else if (a === '--no-shapes') opts.shapes = false;
    else if (a === '--format') opts.compareFormat = true;
    else if (a === '-h' || a === '--help') opts.help = true;
    else files.push(a);
  }
  return { files, opts };
}

function usage() {
  console.log(`Usage:
  node src/cli.js <file_a.xlsx> <file_b.xlsx> [file_c.xlsx] \\
        [--out report.xlsx] [--html report.html] [--no-images] [--no-shapes] [--format]

Flags:
  --out PATH      Xuất báo cáo Excel (mặc định: comparison_report_<ts>.xlsx)
  --html PATH     Xuất thêm HTML report
  --no-images     Bỏ qua so sánh ảnh nhúng
  --no-shapes     Bỏ qua so sánh shapes/吹き出し (callout)
  --format        Bật so sánh định dạng (font/màu/border)
`);
}

async function main() {
  const { files, opts } = parseArgs(process.argv.slice(2));
  if (opts.help || files.length < 2 || files.length > 3) {
    usage();
    process.exit(files.length < 2 ? 1 : 0);
  }
  console.log('[diff] So sánh', files.length, 'file:');
  for (const f of files) console.log('   -', f);

  const diff = await diffWorkbooks(files, {
    compareFormat: opts.compareFormat,
    progressCb: (pct, msg) => console.log(`  [${String(pct).padStart(3)}%] ${msg}`),
  });

  let imgRes = {};
  if (opts.images) imgRes = await diffImages(files);
  let shapeRes = {};
  if (opts.shapes) shapeRes = await diffShapes(files);

  const totals = diff.summary;
  console.log(`\n=> Tổng: added=${totals.added} removed=${totals.removed} ` +
              `modified=${totals.modified} unchanged=${totals.unchanged}`);
  for (const sd of diff.sheets) {
    console.log(`   sheet '${sd.name}':`, JSON.stringify(sd.summary));
  }
  for (const [sname, res] of Object.entries(imgRes)) {
    console.log(`   images[${sname}]:`, JSON.stringify(res.summary));
  }
  for (const [sname, res] of Object.entries(shapeRes)) {
    console.log(`   shapes[${sname}]:`, JSON.stringify(res.summary));
  }

  const outPath = opts.out || path.join(
    process.cwd(), `comparison_report_${Date.now()}.xlsx`);
  await exportExcelReport(diff, imgRes, outPath, { shapeResults: shapeRes });
  console.log(`\n[OK] Báo cáo đã ghi: ${outPath}`);
  if (opts.html) {
    await exportHtmlReport(diff, imgRes, opts.html, { shapeResults: shapeRes });
    console.log(`[OK] HTML report: ${opts.html}`);
  }
}

main().catch(err => {
  console.error('[ERROR]', err.message);
  console.error(err.stack);
  process.exit(2);
});
