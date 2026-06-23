# Excel Diff Tool (Node.js + Electron)

Desktop tool nhẹ, mở nhanh để **so sánh 2-3 file Excel (.xlsx/.xlsm)** có chứa cả nội dung (text/số/công thức/định dạng) lẫn hình ảnh nhúng. Bản port Node.js của [excel-diff-tool (Python)](https://github.com/shinbeii/excel-diff-tool) — chạy không cần Python.

## ✨ Tính năng
- So sánh **giá trị / công thức / định dạng** từng cell, mọi sheet (match theo tên).
- So sánh **ảnh nhúng** với multi-hash ensemble (pHash 16×16 DCT + dHash 16×16 + aHash 8×8) + greedy global matching → 7 trạng thái: `identical / near_identical / similar / resized / moved / added / removed` + score 0-100%.
- So sánh **shapes / 吹き出し (callout / speech bubble)** — parse trực tiếp `xl/drawings/*.xml` (vì exceljs không hỗ trợ shapes), trích xuất prstGeom + text + anchor + size + fill, match bằng composite distance (text Levenshtein 50% + geom 20% + position 20% + fill 10%) → status: `identical / text_changed / shape_changed / moved / resized / style_changed / added / removed`.
- GUI Electron drag-drop, 4 tab: **Tổng quan / Dữ liệu / Hình ảnh / 吹き出し**.
- Xuất báo cáo **Excel** (tô màu diff, thumbnail ảnh, sheet Shapes) và **HTML**.
- CLI cho automation/CI: `--no-images`, `--no-shapes`, `--format`.
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
│   ├── imageDiff.js      # multi-hash + greedy global match (ảnh nhúng)
│   ├── shapeDiff.js      # parse drawing XML + diff 吹き出し/callout
│   ├── report.js         # xuất Excel + HTML
│   └── cli.js            # CLI entry
├── scripts/
│   ├── makeDemoFiles.js  # tạo file demo có ảnh nhúng
│   └── makeDemoShapes.js # tạo file demo có 吹き出し (callout shapes)
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

## 💬 Shape diff (吹き出し / Callout)
exceljs **không hỗ trợ shapes**, nên `src/shapeDiff.js` parse trực tiếp XML drawing trong xlsx (xlsx = ZIP):
1. Mở zip qua `jszip`, đọc `xl/workbook.xml` + rels → resolve tên sheet → drawing path.
2. Parse `xl/drawings/drawing*.xml` qua `sax` streaming, trích xuất từng `<xdr:sp>`:
   - `prstGeom` (wedgeRectCallout, wedgeRoundRectCallout, cloudCallout, borderCallout1..3, ...)
   - text (gộp tất cả `<a:t>`)
   - anchor `from/to` (col/row) + offset EMU → pixel
   - fill color (`<a:solidFill><a:srgbClr>`) — bỏ qua màu trong `<a:ln>` (border).
3. Match bằng composite distance:
   - 50% text Levenshtein normalized
   - 20% prstGeom (khác = 1)
   - 20% position (|Δrow|+|Δcol| / 30, clamp 0..1)
   - 10% fill
4. Greedy global matching → phân loại 8 status:
   `identical / text_changed / shape_changed / moved / resized / style_changed / added / removed` + score 0-100%.

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
