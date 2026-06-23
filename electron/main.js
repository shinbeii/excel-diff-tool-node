/**
 * electron/main.js — Main process: tạo cửa sổ + IPC bridge tới core diff.
 */
'use strict';

const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');

const { diffWorkbooks } = require('../src/excelDiff');
const { diffImages }   = require('../src/imageDiff');
const { diffShapes }   = require('../src/shapeDiff');
const { exportExcelReport, exportHtmlReport } = require('../src/report');

let mainWindow = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 980,
    minHeight: 640,
    title: 'Excel Diff Tool',
    backgroundColor: '#1e1e1e',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });
  mainWindow.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));
  if (process.env.EDT_DEVTOOLS) mainWindow.webContents.openDevTools({ mode: 'detach' });
  mainWindow.on('closed', () => { mainWindow = null; });
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ---------------- IPC ----------------

ipcMain.handle('dialog:openFile', async () => {
  const r = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Excel', extensions: ['xlsx', 'xlsm'] }],
  });
  return r.canceled ? null : r.filePaths[0];
});

ipcMain.handle('dialog:saveReport', async (_e, defName) => {
  const r = await dialog.showSaveDialog(mainWindow, {
    defaultPath: defName || `comparison_report_${Date.now()}.xlsx`,
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
  });
  return r.canceled ? null : r.filePath;
});

ipcMain.handle('compare:run', async (event, args) => {
  const { files, compareFormat, compareImages, compareShapes } = args;
  if (!files || files.length < 2) throw new Error('Cần ít nhất 2 file.');
  for (const f of files) {
    if (!fs.existsSync(f)) throw new Error(`File không tồn tại: ${f}`);
  }
  const send = (channel, payload) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send(channel, payload);
    }
  };
  send('progress', { pct: 0, msg: 'Bắt đầu so sánh nội dung...' });
  const diff = await diffWorkbooks(files, {
    compareFormat,
    progressCb: (pct, msg) => send('progress', { pct: Math.floor(pct * 0.5), msg }),
  });
  let imageResults = {};
  if (compareImages) {
    send('progress', { pct: 55, msg: 'Trích xuất + hash ảnh nhúng...' });
    imageResults = await diffImages(files);
  }
  let shapeResults = {};
  if (compareShapes) {
    send('progress', { pct: 80, msg: 'Parse drawing XML + diff khung chú thích...' });
    shapeResults = await diffShapes(files);
  }
  send('progress', { pct: 100, msg: 'Hoàn tất.' });

  // Strip buffers trước khi gửi sang renderer (giảm payload).
  const slimImages = {};
  for (const [s, r] of Object.entries(imageResults)) {
    slimImages[s] = {
      sheet: r.sheet,
      summary: r.summary,
      entries: r.entries.map(e => ({
        status: e.status, score: e.score, distance: e.distance,
        items: e.items.map(it => it ? {
          fileIdx: it.fileIdx, sheet: it.sheet, anchor: it.anchor,
          width: it.width, height: it.height, phashHex: it.phashHex,
          thumb: it.buffer ? `data:image/png;base64,${it.buffer.toString('base64')}` : null,
        } : null),
      })),
    };
  }
  // Shapes chỉ là dữ liệu text + meta, không cần strip.
  return { diff, imageResults: slimImages, shapeResults };
});

ipcMain.handle('report:export', async (_e, args) => {
  const { diff, outPath, alsoHtml } = args;
  // imageResults từ renderer chưa có Buffer; load lại từ file để embed thumbnail.
  const fullImages = await diffImages(diff.files);
  const fullShapes = await diffShapes(diff.files);
  await exportExcelReport(diff, fullImages, outPath, { shapeResults: fullShapes });
  let htmlPath = null;
  if (alsoHtml) {
    htmlPath = outPath.replace(/\.xlsx?$/i, '.html');
    await exportHtmlReport(diff, fullImages, htmlPath, { shapeResults: fullShapes });
  }
  return { outPath, htmlPath };
});

ipcMain.handle('shell:openPath', async (_e, p) => shell.openPath(p));
ipcMain.handle('shell:showInFolder', async (_e, p) => shell.showItemInFolder(p));
