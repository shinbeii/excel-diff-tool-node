# Excel Diff Tool (Node.js + Electron)

Desktop tool nhẹ, mở nhanh để **so sánh 2-3 file Excel (.xlsx/.xlsm)** có chứa cả nội dung (text/số/công thức/định dạng) lẫn hình ảnh nhúng. Bản port Node.js của [excel-diff-tool (Python)](https://github.com/shinbeii/excel-diff-tool) — chạy không cần Python.

## ✨ Tính năng
- So sánh **giá trị / công thức / định dạng** từng cell, mọi sheet (match theo tên).
- So sánh **ảnh nhúng** với multi-hash ensemble (pHash 16×16 DCT + dHash 16×16 + aHash 8×8) + greedy global matching → 7 trạng thái: `identical / near_identical / similar / resized / moved / added / removed` + score 0-100%.
- GUI Electron drag-drop, 3 tab: **Tổng quan / Dữ liệu / Hình ảnh**.
- Xuất báo cáo **Excel** (tô màu diff, thumbnail ảnh) và **HTML**.
- CLI cho automation/CI.
- Đóng gói thành **.app + .dmg** (macOS x64 + arm64), **.exe NSIS installer + portable** (Windows), **AppImage + .deb** (Linux).

## 🚀 Cài & chạy dev
```bash
npm install
npm run demo      # tạo demo/file_a.xlsx + file_b.xlsx
npm test          # chạy vitest smoke tests
npm run cli demo/file_a.xlsx demo/file_b.xlsx --out demo/report.xlsx --html demo/report.html
npm start         # mở GUI Electron
```

## 🏗 Build
```bash
npm run build:mac         # macOS .dmg + .zip (x64 + arm64)
npm run build:win         # Windows NSIS + portable
npm run build:linux       # Linux AppImage + .deb
npm run build             # build cho platform hiện tại
```
Artifacts xuất ra `dist/`.

## 📁 Cấu trúc
```
excel-diff-tool-node/
├── electron/
│   ├── main.js           # main process + IPC handlers
│   └── preload.js        # contextBridge API
├── renderer/
│   ├── index.html
│   ├── style.css
│   └── app.js            # logic UI (drag-drop, tabs, render)
├── src/
│   ├── excelDiff.js      # diff cell/sheet/formula/format
│   ├── imageDiff.js      # multi-hash + greedy global match
│   ├── report.js         # xuất Excel + HTML
│   └── cli.js            # CLI entry
├── scripts/
│   └── makeDemoFiles.js  # tạo file demo có ảnh nhúng
├── tests/
│   └── smoke.test.mjs    # vitest (ESM)
├── assets/               # icon, etc.
└── package.json
```

## 🧪 CLI
```
Usage:
  node src/cli.js <file_a.xlsx> <file_b.xlsx> [file_c.xlsx] \
        [--out report.xlsx] [--html report.html] [--no-images] [--format]
```

## 🔧 Stack
- **Node.js 18+** (đã test trên 22.18)
- **Electron 32** — GUI desktop cross-platform
- **exceljs 4.4** — đọc/ghi xlsx, ảnh nhúng, công thức, style
- **sharp 0.33** — xử lý ảnh nhanh (libvips bindings)
- **vitest 2** — testing
- **electron-builder 25** — đóng gói

## 📊 Image diff
Thuật toán giống bản Python 10/10:
1. Trích xuất ảnh nhúng từ exceljs (`worksheet.getImages()` + `workbook.getImage()`).
2. Multi-hash ensemble:
   - **pHash 16×16** = 256-bit (DCT-II, robust với compression/recolor)
   - **dHash 16×16** = 256-bit (so sánh pixel ngang, robust với edge/cấu trúc)
   - **aHash 8×8**   = 64-bit  (mean threshold, robust với tone trung bình)
3. Composite distance = `0.45·pHash + 0.40·dHash + 0.15·aHash`, normalized [0,1].
4. Greedy global assignment: sort tất cả cặp theo distance → first-fit (xấp xỉ Hungarian).
5. Phân loại 7 status dựa trên ngưỡng + anchor + aspect ratio.
6. Score 0-100% = `round((1 - distance) * 100)`.

## 📝 License
MIT
