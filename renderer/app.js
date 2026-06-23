/**
 * renderer/app.js — Logic UI: drag-drop, tabs, render kết quả.
 */
'use strict';

const MAX_FILES = 3;
const state = {
  files: [null, null, null],
  use3: false,
  diff: null,
  imageResults: null,
  selectedSheet: null,
};

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

function renderDropzones() {
  const root = $('#dropzones');
  root.innerHTML = '';
  for (let i = 0; i < MAX_FILES; i++) {
    const dz = document.createElement('div');
    dz.className = 'dropzone';
    dz.dataset.idx = i;
    if (i === 2 && !state.use3) dz.classList.add('disabled');
    if (state.files[i]) dz.classList.add('has-file');
    const label = i === 0 ? 'File 1 (baseline)' : `File ${i + 1}`;
    dz.innerHTML = `
      <div class="label">${label}</div>
      <div class="filename">${state.files[i] ? state.files[i].split('/').pop() : 'Kéo thả .xlsx vào đây hoặc click để chọn'}</div>
      ${state.files[i] ? '<button class="clear-btn" data-action="clear">Xóa</button>' : ''}
    `;
    dz.addEventListener('click', async (e) => {
      if (e.target.dataset.action === 'clear') {
        state.files[i] = null; renderDropzones(); return;
      }
      const p = await window.edt.pickFile();
      if (p) { state.files[i] = p; renderDropzones(); }
    });
    dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('hover'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('hover'));
    dz.addEventListener('drop', (e) => {
      e.preventDefault(); dz.classList.remove('hover');
      const f = e.dataTransfer.files[0];
      if (f && /\.xls[xm]$/i.test(f.name)) {
        state.files[i] = f.path;
        renderDropzones();
      }
    });
    root.appendChild(dz);
  }
}

$('#opt3files').addEventListener('change', (e) => {
  state.use3 = e.target.checked;
  if (!state.use3) state.files[2] = null;
  renderDropzones();
});

$('#btnReset').addEventListener('click', () => {
  state.files = [null, null, null];
  state.diff = null; state.imageResults = null; state.selectedSheet = null;
  renderDropzones(); renderAll();
  $('#btnExport').disabled = true;
});

$$('.tab').forEach(t => t.addEventListener('click', () => {
  $$('.tab').forEach(x => x.classList.remove('active'));
  $$('.tab-panel').forEach(x => x.classList.remove('active'));
  t.classList.add('active');
  document.querySelector(`[data-panel="${t.dataset.tab}"]`).classList.add('active');
}));

$('#btnCompare').addEventListener('click', async () => {
  const files = state.files.slice(0, state.use3 ? 3 : 2).filter(Boolean);
  if (files.length < 2) { alert('Cần chọn ít nhất 2 file Excel.'); return; }
  $('#btnCompare').disabled = true; $('#btnExport').disabled = true;
  $('#progressBox').hidden = false;
  $('#progressFill').style.width = '0%';
  $('#progressMsg').textContent = 'Đang khởi động...';
  try {
    const result = await window.edt.compare({
      files,
      compareFormat: $('#optFormat').checked,
      compareImages: $('#optImages').checked,
    });
    state.diff = result.diff;
    state.imageResults = result.imageResults;
    state.selectedSheet = state.diff.sheets[0]?.name || null;
    renderAll();
    $('#btnExport').disabled = false;
  } catch (err) {
    alert('Lỗi: ' + err.message);
  } finally {
    $('#btnCompare').disabled = false;
    setTimeout(() => $('#progressBox').hidden = true, 800);
  }
});

window.edt.onProgress(({ pct, msg }) => {
  $('#progressFill').style.width = pct + '%';
  $('#progressMsg').textContent = `${pct}% — ${msg}`;
});

$('#btnExport').addEventListener('click', async () => {
  if (!state.diff) return;
  const out = await window.edt.saveReport(`comparison_report_${Date.now()}.xlsx`);
  if (!out) return;
  try {
    const r = await window.edt.exportReport({
      diff: state.diff, imageResults: null, outPath: out, alsoHtml: true,
    });
    const open = confirm(`✓ Đã xuất:\n  ${r.outPath}\n  ${r.htmlPath || ''}\n\nMở file ngay?`);
    if (open) window.edt.openPath(r.outPath);
  } catch (e) {
    alert('Lỗi xuất báo cáo: ' + e.message);
  }
});

$('#filterData').addEventListener('input', renderDataTable);

function renderAll() {
  renderSidebar();
  renderOverview();
  renderDataTable();
  renderImages();
}

function renderSidebar() {
  const root = $('#sidebar');
  if (!state.diff) {
    root.innerHTML = '<div class="hint">Chọn 2-3 file Excel rồi nhấn <b>So sánh</b>.</div>';
    return;
  }
  root.innerHTML = '';
  for (const sd of state.diff.sheets) {
    const div = document.createElement('div');
    div.className = 'sheet-row';
    if (sd.name === state.selectedSheet) div.classList.add('selected');
    const s = sd.summary;
    div.innerHTML = `
      <div class="name">${sd.name}</div>
      <div class="stats">+${s.added} −${s.removed} ✎${s.modified}</div>
    `;
    div.addEventListener('click', () => {
      state.selectedSheet = sd.name; renderSidebar(); renderDataTable();
    });
    root.appendChild(div);
  }
}

function renderOverview() {
  const root = $('#panelOverview');
  if (!state.diff) {
    root.innerHTML = '<div class="placeholder">Chưa có kết quả so sánh.</div>';
    return;
  }
  const d = state.diff;
  let html = '<div class="summary-grid">';
  html += `<div class="summary-card">
    <h3>📈 Tổng quan</h3>
    <div class="row kv-added"><span>Added</span><span class="v">${d.summary.added}</span></div>
    <div class="row kv-removed"><span>Removed</span><span class="v">${d.summary.removed}</span></div>
    <div class="row kv-modified"><span>Modified</span><span class="v">${d.summary.modified}</span></div>
    <div class="row"><span>Unchanged</span><span class="v">${d.summary.unchanged}</span></div>
  </div>`;
  for (const sd of d.sheets) {
    html += `<div class="summary-card">
      <h3>📄 ${sd.name}</h3>
      <div class="row kv-added"><span>Added</span><span class="v">${sd.summary.added}</span></div>
      <div class="row kv-removed"><span>Removed</span><span class="v">${sd.summary.removed}</span></div>
      <div class="row kv-modified"><span>Modified</span><span class="v">${sd.summary.modified}</span></div>
    </div>`;
  }
  if (state.imageResults) {
    for (const [s, r] of Object.entries(state.imageResults)) {
      html += `<div class="summary-card">
        <h3>🖼 ${s}</h3>
        <div class="row"><span>identical</span><span class="v">${r.summary.identical}</span></div>
        <div class="row kv-similar"><span>near_identical</span><span class="v">${r.summary.near_identical}</span></div>
        <div class="row kv-similar"><span>similar</span><span class="v">${r.summary.similar}</span></div>
        <div class="row kv-resized"><span>resized</span><span class="v">${r.summary.resized}</span></div>
        <div class="row kv-moved"><span>moved</span><span class="v">${r.summary.moved}</span></div>
        <div class="row kv-added"><span>added</span><span class="v">${r.summary.added}</span></div>
        <div class="row kv-removed"><span>removed</span><span class="v">${r.summary.removed}</span></div>
      </div>`;
    }
  }
  html += '</div>';
  root.innerHTML = html;
}

function renderDataTable() {
  const root = $('#dataTableWrap');
  if (!state.diff || !state.selectedSheet) {
    root.innerHTML = '<div class="placeholder">Chưa có dữ liệu.</div>'; return;
  }
  const sd = state.diff.sheets.find(s => s.name === state.selectedSheet);
  if (!sd || !sd.cells.length) {
    root.innerHTML = '<div class="placeholder">Sheet không có khác biệt nào.</div>'; return;
  }
  const filter = ($('#filterData').value || '').toLowerCase();
  const nFiles = state.diff.files.length;
  let html = '<table class="diff"><thead><tr><th>Cell</th><th>Status</th>';
  for (let i = 0; i < nFiles; i++) html += `<th>File ${i + 1}</th>`;
  html += '<th>Formula</th></tr></thead><tbody>';
  for (const cd of sd.cells) {
    const text = (cd.coord + ' ' + cd.status + ' ' + (cd.values || []).join(' ')).toLowerCase();
    if (filter && !text.includes(filter)) continue;
    html += `<tr class="${cd.status}"><td>${cd.coord}</td><td>${cd.status}</td>`;
    for (const v of cd.values) html += `<td>${esc(v)}</td>`;
    html += `<td>${esc(cd.formulas.map(f => f || '').join(' | '))}</td></tr>`;
  }
  html += '</tbody></table>';
  root.innerHTML = html;
}

function renderImages() {
  const root = $('#imagesWrap');
  if (!state.imageResults || Object.keys(state.imageResults).length === 0) {
    root.innerHTML = '<div class="placeholder">Không có ảnh được so sánh.</div>'; return;
  }
  let html = '';
  for (const [s, r] of Object.entries(state.imageResults)) {
    if (!r.entries.length) continue;
    html += `<h2 style="font-size:14px;color:#fff;margin:8px 4px">📄 Sheet: ${s}</h2><div class="image-grid">`;
    for (const e of r.entries) {
      html += `<div class="image-card">
        <div><span class="badge b-${e.status}">${e.status}</span> <span class="score">${e.score}%</span></div>
        <div class="imgs">`;
      for (const it of e.items) {
        if (it && it.thumb) html += `<img src="${it.thumb}" title="${it.anchor} ${it.width}x${it.height}">`;
        else html += `<div style="width:120px;height:90px;border:1px dashed #555;display:flex;align-items:center;justify-content:center;color:#777;font-size:11px">—</div>`;
      }
      html += `</div><div class="meta">distance: ${e.distance} · anchors: ${e.items.map(it => it ? it.anchor : '-').join(' / ')}</div></div>`;
    }
    html += '</div>';
  }
  root.innerHTML = html || '<div class="placeholder">Không có ảnh khác biệt.</div>';
}

function esc(v) {
  if (v == null) return '';
  return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

renderDropzones();
renderAll();
