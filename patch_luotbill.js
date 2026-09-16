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

// 1. Thêm tên tab LUOTBILL
tryReplace(
  'tab_const',
  "const GOOGLE_SHEET_TAB_CHITIETXUAT = process.env.GOOGLE_SHEET_TAB_CHITIETXUAT || 'CHITIETXUAT';",
  "const GOOGLE_SHEET_TAB_CHITIETXUAT = process.env.GOOGLE_SHEET_TAB_CHITIETXUAT || 'CHITIETXUAT';\nconst GOOGLE_SHEET_TAB_LUOTBILL = process.env.GOOGLE_SHEET_TAB_LUOTBILL || 'LUOTBILL';"
);

// 2. Thêm nhận diện file "BC Lượt Bill Theo Siêu Thị"
tryReplace(
  'nhan_dien_file',
  [
    "  if (co('Mã phiếu xuất') && co('Tên sản phẩm') && co('Số lượng') && co('Giá bán')) {",
    "    return { loai: 'chitietxuat', tenTab: GOOGLE_SHEET_TAB_CHITIETXUAT };",
    "  }",
    "  return null;",
    "}"
  ].join("\n"),
  [
    "  if (co('Mã phiếu xuất') && co('Tên sản phẩm') && co('Số lượng') && co('Giá bán')) {",
    "    return { loai: 'chitietxuat', tenTab: GOOGLE_SHEET_TAB_CHITIETXUAT };",
    "  }",
    "  if (co('Lượt Bill của Siêu Thị') && co('Ngày xuất') && co('Tổng tiền (VAT)')) {",
    "    return { loai: 'luotbill', tenTab: GOOGLE_SHEET_TAB_LUOTBILL };",
    "  }",
    "  return null;",
    "}"
  ].join("\n")
);

// 3. Thêm xử lý nạp file LUOTBILL (ghi đè toàn bộ)
tryReplace(
  'xu_ly_nap_file',
  [
    "    return { loai: nhanDang.loai, tenTab: nhanDang.tenTab, soDong: rows.length };",
    "  }",
    "",
    "  if (nhanDang.loai === 'huymmkk') {"
  ].join("\n"),
  [
    "    return { loai: nhanDang.loai, tenTab: nhanDang.tenTab, soDong: rows.length };",
    "  }",
    "",
    "  if (nhanDang.loai === 'luotbill') {",
    "    const idxNgay = header.indexOf('Ngày xuất');",
    "    const idxLuotBill = header.indexOf('Lượt Bill của Siêu Thị');",
    "",
    "    const rowsLB = dataRows.map((row) => [",
    "      toDateTimeKeyForBucket(row[idxNgay]),",
    "      row[idxLuotBill],",
    "    ]);",
    "",
    "    await sheets.spreadsheets.values.clear({",
    "      spreadsheetId: GOOGLE_SHEET_ID,",
    "      range: `${nhanDang.tenTab}!A2:ZZ`,",
    "    });",
    "    await sheets.spreadsheets.values.update({",
    "      spreadsheetId: GOOGLE_SHEET_ID,",
    "      range: `${nhanDang.tenTab}!A2`,",
    "      valueInputOption: 'RAW',",
    "      requestBody: { values: rowsLB },",
    "    });",
    "",
    "    return { loai: nhanDang.loai, tenTab: nhanDang.tenTab, soDong: rowsLB.length };",
    "  }",
    "",
    "  if (nhanDang.loai === 'huymmkk') {"
  ].join("\n")
);

// 4. Thêm helper parse ngày giờ + trigger + hàm tạo báo cáo Flex
const anchor4 = [
  "const TRIGGER_SP1DONG = ['sp 1 đồng', 'sp 1 dong', 'sản phẩm 1 đồng', 'san pham 1 dong'];",
  "function laTriggerSp1Dong(text) {",
  "  if (!text) return false;",
  "  const t = text.trim().toLowerCase();",
  "  return TRIGGER_SP1DONG.some((kw) => t === kw || t.includes(kw));",
  "}"
].join("\n");

const new4 = anchor4 + "\n\n" + [
  "function toDateTimeKeyForBucket(value) {",
  "  if (typeof value === 'number') {",
  "    const epoch = Date.UTC(1899, 11, 30);",
  "    const d = new Date(epoch + value * 86400000);",
  "    const y = d.getUTCFullYear();",
  "    const mo = String(d.getUTCMonth() + 1).padStart(2, '0');",
  "    const da = String(d.getUTCDate()).padStart(2, '0');",
  "    const hh = String(d.getUTCHours()).padStart(2, '0');",
  "    const mi = String(d.getUTCMinutes()).padStart(2, '0');",
  "    return `${y}-${mo}-${da} ${hh}:${mi}`;",
  "  }",
  "  const s = (value || '').toString().trim();",
  "  const m = s.match(/^(\\d{1,2})\\/(\\d{1,2})\\/(\\d{4})\\s+(\\d{1,2}):(\\d{1,2})/);",
  "  if (m) {",
  "    const [, d, mo, y, hh, mi] = m;",
  "    return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')} ${hh.padStart(2, '0')}:${mi.padStart(2, '0')}`;",
  "  }",
  "  return s;",
  "}",
  "",
  "const TRIGGER_LUOTBILL = ['lượt bill', 'luot bill', 'lượt hóa đơn', 'luot hoa don'];",
  "function laTriggerLuotBill(text) {",
  "  if (!text) return false;",
  "  const t = text.trim().toLowerCase();",
  "  return TRIGGER_LUOTBILL.some((kw) => t === kw || t.includes(kw));",
  "}",
  "",
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
  "  }",
  "",
  "  const BUCKET_ORDER = ['05:30-06:00'];",
  "  for (let h = 6; h <= 20; h++) {",
  "    BUCKET_ORDER.push(`${String(h).padStart(2, '0')}:00-${String(h + 1).padStart(2, '0')}:00`);",
  "  }",
  "",
  "  const tong = {};",
  "  BUCKET_ORDER.forEach((k) => { tong[k] = 0; });",
  "",
  "  function xacDinhKhung(hh, mm) {",
  "    const tongPhut = hh * 60 + mm;",
  "    if (tongPhut >= 5 * 60 + 30 && tongPhut < 6 * 60) return '05:30-06:00';",
  "    if (tongPhut >= 6 * 60 && tongPhut < 21 * 60) {",
  "      const gioBatDau = Math.floor(tongPhut / 60);",
  "      return `${String(gioBatDau).padStart(2, '0')}:00-${String(gioBatDau + 1).padStart(2, '0')}:00`;",
  "    }",
  "    return null;",
  "  }",
  "",
  "  for (const [dt, luotBillRaw] of rows) {",
  "    if (!dt) continue;",
  "    const s = dt.toString();",
  "    const ngay = s.slice(0, 10);",
  "    if (ngay !== ngayMoiNhat) continue;",
  "",
  "    const gioPhutMatch = s.match(/(\\d{1,2}):(\\d{2})$/);",
  "    if (!gioPhutMatch) continue;",
  "    const hh = Number(gioPhutMatch[1]);",
  "    const mm = Number(gioPhutMatch[2]);",
  "",
  "    const khung = xacDinhKhung(hh, mm);",
  "    if (!khung) continue;",
  "",
  "    tong[khung] += Number(luotBillRaw) || 0;",
  "  }",
  "",
  "  const tongCong = Object.values(tong).reduce((a, b) => a + b, 0);",
  "  const khungCoLuot = BUCKET_ORDER.filter((k) => tong[k] > 0);",
  "  const khungMax = khungCoLuot.length ? khungCoLuot.reduce((a, b) => (tong[a] >= tong[b] ? a : b)) : null;",
  "  const khungMin = khungCoLuot.length ? khungCoLuot.reduce((a, b) => (tong[a] <= tong[b] ? a : b)) : null;",
  "",
  "  const danhSachContents = BUCKET_ORDER.map((khung) => {",
  "    const soLuot = tong[khung];",
  "    const laMax = khung === khungMax && soLuot > 0;",
  "    const laMin = khung === khungMin && soLuot > 0;",
  "    let icon = '';",
  "    let mauChu = '#333333';",
  "    let doDam = 'regular';",
  "    if (laMax) { icon = '🔥 '; mauChu = '#C62828'; doDam = 'bold'; }",
  "    if (laMin) { icon = '❄️ '; mauChu = '#1565C0'; doDam = 'bold'; }",
  "",
  "    return {",
  "      type: 'box',",
  "      layout: 'horizontal',",
  "      margin: 'md',",
  "      contents: [",
  "        { type: 'text', text: `${icon}${khung}`, size: 'sm', color: mauChu, weight: doDam, flex: 3 },",
  "        { type: 'text', text: `${soLuot}`, size: 'sm', color: mauChu, weight: doDam, align: 'end', flex: 1 },",
  "      ],",
  "    };",
  "  });",
  "",
  "  return {",
  "    type: 'flex',",
  "    altText: `Lượt Bill ngày ${fmtNgayHienThi(ngayMoiNhat)}: Tổng ${tongCong} lượt${khungMax ? ', cao nhất ' + khungMax : ''}`,",
  "    contents: {",
  "      type: 'bubble',",
  "      size: 'giga',",
  "      header: {",
  "        type: 'box',",
  "        layout: 'vertical',",
  "        backgroundColor: '#1565C0',",
  "        paddingAll: '20px',",
  "        contents: [",
  "          { type: 'text', text: '🧾 BÁO CÁO LƯỢT BILL THEO KHUNG GIỜ', color: '#FFFFFF', weight: 'bold', size: 'lg', wrap: true },",
  "          { type: 'text', text: `Ngày ${fmtNgayHienThi(ngayMoiNhat)}`, color: '#D6E4F5', size: 'sm', margin: 'sm' },",
  "        ],",
  "      },",
  "      body: {",
  "        type: 'box',",
  "        layout: 'vertical',",
  "        paddingAll: '16px',",
  "        contents: [",
  "          {",
  "            type: 'box',",
  "            layout: 'vertical',",
  "            backgroundColor: '#E3F2FD',",
  "            cornerRadius: '8px',",
  "            paddingAll: '12px',",
  "            contents: [",
  "              { type: 'text', text: 'TỔNG LƯỢT BILL TRONG NGÀY', size: 'xs', color: '#666666' },",
  "              { type: 'text', text: `${tongCong}`, size: 'xxl', weight: 'bold', color: '#1565C0' },",
  "              { type: 'text', text: khungMax ? `🔥 Cao nhất: ${khungMax} (${tong[khungMax]} lượt)` : '', size: 'xs', color: '#C62828', margin: 'sm' },",
  "              { type: 'text', text: khungMin ? `❄️ Thấp nhất: ${khungMin} (${tong[khungMin]} lượt)` : '', size: 'xs', color: '#1565C0', margin: 'xs' },",
  "            ],",
  "          },",
  "          { type: 'separator', margin: 'lg' },",
  "          { type: 'box', layout: 'vertical', margin: 'lg', spacing: 'sm', contents: danhSachContents },",
  "        ],",
  "      },",
  "    },",
  "  };",
  "}"
].join("\n");

tryReplace('trigger_and_report', anchor4, new4);

// 5. Gắn lệnh "Lượt Bill" vào bộ điều phối lệnh text
tryReplace(
  'dispatcher',
  [
    "  if (laTriggerSp1Dong(text)) return { ten: 'SP 1 Đồng', ket: await generateSp1DongReport() };",
    "  return null;",
    "}"
  ].join("\n"),
  [
    "  if (laTriggerSp1Dong(text)) return { ten: 'SP 1 Đồng', ket: await generateSp1DongReport() };",
    "  if (laTriggerLuotBill(text)) return { ten: 'Lượt Bill', ket: await generateLuotBillReport() };",
    "  return null;",
    "}"
  ].join("\n")
);

fs.writeFileSync(path, content, 'utf8');

console.log('===== KẾT QUẢ =====');
console.log('THÀNH CÔNG:', ok.join(', ') || '(không có)');
console.log('THẤT BẠI (cần báo lại):', fail.join(', ') || '(không có - TẤT CẢ ĐỀU OK)');
