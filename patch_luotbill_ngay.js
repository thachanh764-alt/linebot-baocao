const fs = require('fs');
const path = 'server.js';
let content = fs.readFileSync(path, 'utf8');
const ok = [];
const fail = [];

function tryReplace(name, anchor, replacement) {
  if (content.includes(anchor)) {
    content = content.replace(anchor, replacement);
    ok.push(name);
  } else {
    fail.push(name);
  }
}

// 1. Thêm hàm trích ngày từ câu lệnh (vd "Lượt Bill 10/09" hoặc "luot bill 10/09/2026")
tryReplace(
  'ham_trich_ngay',
  "const TRIGGER_LUOTBILL = ['lượt bill', 'luot bill', 'lượt hóa đơn', 'luot hoa don'];",
  [
    "function trichNgayTuCauLenh(text) {",
    "  if (!text) return null;",
    "  const m = text.match(/(\\d{1,2})\\/(\\d{1,2})(?:\\/(\\d{4}))?/);",
    "  if (!m) return null;",
    "  const d = m[1].padStart(2, '0');",
    "  const mo = m[2].padStart(2, '0');",
    "  const y = m[3] || new Date().getFullYear().toString();",
    "  return `${y}-${mo}-${d}`;",
    "}",
    "",
    "const TRIGGER_LUOTBILL = ['lượt bill', 'luot bill', 'lượt hóa đơn', 'luot hoa don'];"
  ].join("\n")
);

// 2. Sửa generateLuotBillReport để nhận tham số ngày yêu cầu
tryReplace(
  'sua_ham_bao_cao',
  [
    "async function generateLuotBillReport() {",
    "  const sheets = getSheetsClient();",
    "  const res = await sheets.spreadsheets.values.get({",
    "    spreadsheetId: GOOGLE_SHEET_ID,",
    "    range: `${GOOGLE_SHEET_TAB_LUOTBILL}!A2:B`,",
    "  });",
    "",
    "  const rows = res.data.values || [];",
    "  if (rows.length === 0) {",
    "    return { type: 'text', text: 'Chưa có dữ liệu Lượt Bill nào được nạp.' };",
    "  }",
    "",
    "  let ngayMoiNhat = null;",
    "  for (const [dt] of rows) {",
    "    if (!dt) continue;",
    "    const ngay = dt.toString().slice(0, 10);",
    "    if (ngayMoiNhat === null || ngay > ngayMoiNhat) ngayMoiNhat = ngay;",
    "  }"
  ].join("\n"),
  [
    "async function generateLuotBillReport(ngayYeuCau) {",
    "  const sheets = getSheetsClient();",
    "  const res = await sheets.spreadsheets.values.get({",
    "    spreadsheetId: GOOGLE_SHEET_ID,",
    "    range: `${GOOGLE_SHEET_TAB_LUOTBILL}!A2:B`,",
    "  });",
    "",
    "  const rows = res.data.values || [];",
    "  if (rows.length === 0) {",
    "    return { type: 'text', text: 'Chưa có dữ liệu Lượt Bill nào được nạp.' };",
    "  }",
    "",
    "  let ngayMoiNhat = null;",
    "  const cacNgayCoDuLieu = new Set();",
    "  for (const [dt] of rows) {",
    "    if (!dt) continue;",
    "    const ngay = dt.toString().slice(0, 10);",
    "    cacNgayCoDuLieu.add(ngay);",
    "    if (ngayMoiNhat === null || ngay > ngayMoiNhat) ngayMoiNhat = ngay;",
    "  }",
    "",
    "  const ngayMucTieu = ngayYeuCau || ngayMoiNhat;",
    "  if (ngayYeuCau && !cacNgayCoDuLieu.has(ngayYeuCau)) {",
    "    return {",
    "      type: 'text',",
    "      text: `Không có dữ liệu Lượt Bill cho ngày ${fmtNgayHienThi(ngayYeuCau)}. Các ngày hiện có: ${[...cacNgayCoDuLieu].sort().map(fmtNgayHienThi).join(', ')}`,",
    "    };",
    "  }",
    "  ngayMoiNhat = ngayMucTieu;"
  ].join("\n")
);

fs.writeFileSync(path, content, 'utf8');

console.log('===== KẾT QUẢ =====');
console.log('THÀNH CÔNG:', ok.join(', ') || '(không có)');
console.log('THẤT BẠI (cần báo lại):', fail.join(', ') || '(không có - TẤT CẢ ĐỀU OK)');
