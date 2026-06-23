/**
 * tests/smoke.test.mjs — Smoke test (ESM cho Vitest).
 */
import { describe, it, expect, beforeAll } from 'vitest';
import path from 'node:path';
import fs from 'node:fs';
import { execSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const require = createRequire(import.meta.url);

const { diffWorkbooks } = require('../src/excelDiff');
const { diffImages }    = require('../src/imageDiff');
const { diffShapes }    = require('../src/shapeDiff');
const { exportExcelReport, exportHtmlReport } = require('../src/report');

const ROOT = path.resolve(__dirname, '..');
const DEMO = path.join(ROOT, 'demo');
const FA = path.join(DEMO, 'file_a.xlsx');
const FB = path.join(DEMO, 'file_b.xlsx');
const SA = path.join(DEMO, 'shapes_a.xlsx');
const SB = path.join(DEMO, 'shapes_b.xlsx');

describe('excel-diff-tool smoke', () => {
  beforeAll(() => {
    if (!fs.existsSync(FA) || !fs.existsSync(FB)) {
      execSync('node scripts/makeDemoFiles.js', { cwd: ROOT, stdio: 'inherit' });
    }
    if (!fs.existsSync(SA) || !fs.existsSync(SB)) {
      execSync('node scripts/makeDemoShapes.js', { cwd: ROOT, stdio: 'inherit' });
    }
  }, 60000);

  it('diff workbooks: phát hiện modified/added trên BangGia', async () => {
    const r = await diffWorkbooks([FA, FB]);
    expect(r.files.length).toBe(2);
    const bg = r.sheets.find(s => s.name === 'BangGia');
    expect(bg).toBeTruthy();
    expect(bg.summary.modified).toBeGreaterThan(0);
    expect(bg.summary.added).toBeGreaterThan(0);
    const extra = r.sheets.find(s => s.name === 'Extra');
    expect(extra).toBeTruthy();
  }, 60000);

  it('diff images: tìm thấy moved + identical/similar + added', async () => {
    const r = await diffImages([FA, FB]);
    const bg = r.BangGia;
    expect(bg).toBeTruthy();
    expect(bg.entries.length).toBeGreaterThanOrEqual(3);
    const statuses = bg.entries.map(e => e.status);
    expect(statuses).toContain('moved');
    expect(statuses).toContain('added');
    const matchSim = statuses.some(s => ['identical', 'near_identical', 'similar'].includes(s));
    expect(matchSim).toBe(true);
    for (const e of bg.entries) {
      expect(e.score).toBeGreaterThanOrEqual(0);
      expect(e.score).toBeLessThanOrEqual(100);
    }
  }, 60000);

  it('diff shapes (吹き出し): tìm thấy text_changed + moved + added + removed', async () => {
    const r = await diffShapes([SA, SB]);
    const notes = r.Notes;
    expect(notes).toBeTruthy();
    expect(notes.entries.length).toBeGreaterThanOrEqual(4);
    const statuses = notes.entries.map(e => e.status);
    expect(statuses).toContain('text_changed');
    expect(statuses).toContain('moved');
    expect(statuses).toContain('added');
    expect(statuses).toContain('removed');
    expect(notes.summary.callouts).toBeGreaterThan(0);
    // Verify shape có đầy đủ metadata
    const sh = notes.entries.find(e => e.items.some(it => it && it.prstGeom)).items.find(it => it && it.prstGeom);
    expect(sh.prstGeom).toBeTruthy();
    expect(sh.fill).toMatch(/^#[0-9A-F]{6}$/);
    expect(sh.text).toBeTruthy();
    expect(sh.isCallout).toBe(true);
  }, 60000);

  it('xuất báo cáo Excel + HTML có sheet Shapes', async () => {
    const diff = await diffWorkbooks([SA, SB]);
    const shapes = await diffShapes([SA, SB]);
    const xlsx = path.join(DEMO, 'shapes_report.xlsx');
    const html = path.join(DEMO, 'shapes_report.html');
    await exportExcelReport(diff, {}, xlsx, { shapeResults: shapes });
    await exportHtmlReport(diff, {}, html, { shapeResults: shapes });
    expect(fs.existsSync(xlsx)).toBe(true);
    expect(fs.existsSync(html)).toBe(true);
    const htmlContent = fs.readFileSync(html, 'utf-8');
    expect(htmlContent).toContain('Khung chú thích');
    expect(htmlContent).toContain('Đổi nội dung');
  }, 60000);

  it('xuất báo cáo Excel + HTML', async () => {
    const diff = await diffWorkbooks([FA, FB]);
    const imgs = await diffImages([FA, FB]);
    const xlsx = path.join(DEMO, 'report.xlsx');
    const html = path.join(DEMO, 'report.html');
    await exportExcelReport(diff, imgs, xlsx);
    await exportHtmlReport(diff, imgs, html);
    expect(fs.existsSync(xlsx)).toBe(true);
    expect(fs.existsSync(html)).toBe(true);
    expect(fs.statSync(xlsx).size).toBeGreaterThan(1000);
  }, 60000);
});
