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

// 1. Thêm tên tab CHITIETXUAT
const anchor1 = "const GOOGLE_SHEET_TAB_HUYMMKK = process.env.GOOGLE_SHEET_TAB_HUYMMKK || 'HUYMMKK';";
const new1 = anchor1 + "\nconst GOOGLE_SHEET_TAB_CHITIETXUAT = process.env.GOOGLE_SHEET_TAB_CHITIETXUAT || 'CHITIETXUAT';";
tryReplace('tab_const', anchor1, new1);

// 2. Thêm nhận diện file "Chi Tiết Phiếu Xuất"
const anchor2 = [
  "  if (co('Ngày xuất') && co('Ngành hàng BHX') && co('Doanh thu')) {",
  "    return { loai: 'doanhthu_nganhhang', tenTab: GOOGLE_SHEET_TAB_DOANHTHU_NGANHHANG };",
  "  }",
  "  return null;",
  "}"
].join("\n");
const new2 = [
  "  if (co('Ngày xuất') && co('Ngành hàng BHX') && co('Doanh thu')) {",
  "    return { loai: 'doanhthu_nganhhang', tenTab: GOOGLE_SHEET_TAB_DOANHTHU_NGANHHANG };",
  "  }",
  "  if (co('Mã phiếu xuất') && co('Tên sản phẩm') && co('Số lượng') && co('Giá bán')) {",
  "    return { loai: 'chitietxuat', tenTab: GOOGLE_SHEET_TAB_CHITIETXUAT };",
  "  }",
  "  return null;",
  "}"
].join("\n");
tryReplace('nhan_dien_file', anchor2, new2);

// 3. Thêm xử lý nạp file vào tab CHITIETXUAT (ghi đè toàn bộ)
const anchor3 = [
  "  const sheets = getSheetsClient();",
  "",
  "  if (nhanDang.loai === 'huymmkk') {"
].join("\n");
const new3 = [
  "  const sheets = getSheetsClient();",
  "",
  "  if (nhanDang.loai === 'chitietxuat') {",
  "    const idxTen = header.indexOf('Tên sản phẩm');",
  "    const idxSL = header.indexOf('Số lượng');",
  "    const idxGia = header.indexOf('Giá bán');",
  "",
  "    const rows = dataRows.map((row) => [row[idxTen], row[idxSL], row[idxGia]]);",
  "",
  "    await sheets.spreadsheets.values.clear({",
  "      spreadsheetId: GOOGLE_SHEET_ID,",
  "      range: `${nhanDang.tenTab}!A2:ZZ`,",
  "    });",
  "    await sheets.spreadsheets.values.update({",
  "      spreadsheetId: GOOGLE_SHEET_ID,",
  "      range: `${nhanDang.tenTab}!A2`,",
  "      valueInputOption: 'USER_ENTERED',",
  "      requestBody: { values: rows },",
  "    });",
  "",
  "    return { loai: nhanDang.loai, tenTab: nhanDang.tenTab, soDong: rows.length };",
  "  }",
  "",
  "  if (nhanDang.loai === 'huymmkk') {"
].join("\n");
tryReplace('xu_ly_nap_file', anchor3, new3);

// 4. Thêm hàm nhận diện lệnh text + hàm tạo báo cáo
const anchor4 = [
  "const TRIGGER_HUYMMKK = ['mmkk huỷ', 'mmkk huy', 'huỷ mmkk', 'huy mmkk', 'hủy mmkk'];",
  "function laTriggerHuyMmkk(text) {",
  "  if (!text) return false;",
  "  const t = text.trim().toLowerCase();",
  "  return TRIGGER_HUYMMKK.some((kw) => t === kw || t.includes(kw));",
  "}"
].join("\n");
const new4 = anchor4 + "\n\n" + [
  "const TRIGGER_SP1DONG = ['sp 1 đồng', 'sp 1 dong', 'sản phẩm 1 đồng', 'san pham 1 dong'];",
  "function laTriggerSp1Dong(text) {",
  "  if (!text) return false;",
  "  const t = text.trim().toLowerCase();",
  "  return TRIGGER_SP1DONG.some((kw) => t === kw || t.includes(kw));",
  "}",
  "",
  "async function generateSp1DongReport() {",
  "  const sheets = getSheetsClient();",
  "  const res = await sheets.spreadsheets.values.get({",
  "    spreadsheetId: GOOGLE_SHEET_ID,",
  "    range: `${GOOGLE_SHEET_TAB_CHITIETXUAT}!A2:C`,",
  "  });",
  "",
  "  const rows = res.data.values || [];",
  "  const grouped = {};",
  "",
  "  for (const [tenSP, soLuong, giaBan] of rows) {",
  "    if (!tenSP) continue;",
  "    if (Number(giaBan) !== 1) continue;",
  "    if (tenSP.toUpperCase().includes('NẤM')) continue;",
  "",
  "    const sl = Number(soLuong) || 0;",
  "    grouped[tenSP] = (grouped[tenSP] || 0) + sl;",
  "  }",
  "",
  "  const sorted = Object.entries(grouped)",
  "    .sort((a, b) => b[1] - a[1])",
  "    .slice(0, 15);",
  "",
  "  if (sorted.length === 0) {",
  "    return 'Không có sản phẩm giá bán 1 đồng nào (đã loại nấm) trong dữ liệu hiện tại.';",
  "  }",
  "",
  "  let message = '📋 BÁO CÁO SP GIÁ BÁN 1 ĐỒNG (đã loại nấm)\\n\\n';",
  "  sorted.forEach(([ten, sl], idx) => {",
  "    message += `${idx + 1}. ${ten}: ${sl}\\n`;",
  "  });",
  "",
  "  return message;",
  "}"
].join("\n");
tryReplace('trigger_function', anchor4, new4);

// 5. Gắn lệnh "SP 1 ĐỒNG" vào bộ điều phối lệnh text
const anchor5 = [
  "  if (laTriggerHuyMmkk(text)) return { ten: 'MMKK Huỷ', ket: await generateHuyMmkkReport() };",
  "  return null;",
  "}"
].join("\n");
const new5 = [
  "  if (laTriggerHuyMmkk(text)) return { ten: 'MMKK Huỷ', ket: await generateHuyMmkkReport() };",
  "  if (laTriggerSp1Dong(text)) return { ten: 'SP 1 Đồng', ket: await generateSp1DongReport() };",
  "  return null;",
  "}"
].join("\n");
tryReplace('dispatcher', anchor5, new5);

fs.writeFileSync(path, content, 'utf8');

console.log('===== KẾT QUẢ =====');
console.log('THÀNH CÔNG:', ok.join(', ') || '(không có)');
console.log('THẤT BẠI (cần báo lại):', fail.join(', ') || '(không có - TẤT CẢ ĐỀU OK)');
