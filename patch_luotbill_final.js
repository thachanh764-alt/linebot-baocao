const fs = require('fs');
const path = 'server.js';
let content = fs.readFileSync(path, 'utf8');
const ok = [];
const fail = [];

// Tìm và thay toàn bộ 1 hàm (từ chữ ký hàm tới dấu } đóng khớp), bất kể nội dung bên trong đang là gì
function replaceWholeFunction(name, signature, newFullFunctionText) {
  const startIdx = content.indexOf(signature);
  if (startIdx === -1) {
    fail.push(name + ' (không tìm thấy chữ ký hàm)');
    return;
  }
  const braceStart = content.indexOf('{', startIdx);
  if (braceStart === -1) {
    fail.push(name + ' (không tìm thấy dấu { mở đầu)');
    return;
  }
  let depth = 0;
  let endIdx = -1;
  for (let i = braceStart; i < content.length; i++) {
    if (content[i] === '{') depth++;
    if (content[i] === '}') {
      depth--;
      if (depth === 0) { endIdx = i + 1; break; }
    }
  }
  if (endIdx === -1) {
    fail.push(name + ' (không tìm thấy dấu } đóng khớp)');
    return;
  }
  content = content.slice(0, startIdx) + newFullFunctionText + content.slice(endIdx);
  ok.push(name);
}

// Thay hoặc thêm 1 khối "if (nhanDang.loai === 'luotbill') { ... }" trong napFileVaoSheet
function replaceIfBlockLuotBill(newBlockText) {
  const marker = "if (nhanDang.loai === 'luotbill') {";
  const startIdx = content.indexOf(marker);
  if (startIdx === -1) {
    fail.push('nap_file_luotbill (không tìm thấy khối cũ, có thể đã bị đổi tên)');
    return;
  }
  const braceStart = content.indexOf('{', startIdx);
  let depth = 0;
  let endIdx = -1;
  for (let i = braceStart; i < content.length; i++) {
    if (content[i] === '{') depth++;
    if (content[i] === '}') {
      depth--;
      if (depth === 0) { endIdx = i + 1; break; }
    }
  }
  if (endIdx === -1) {
    fail.push('nap_file_luotbill (không tìm thấy dấu } đóng khớp)');
    return;
  }
  content = content.slice(0, startIdx) + newBlockText + content.slice(endIdx);
  ok.push('nap_file_luotbill');
}

// 1. Thêm hàm trích ngày từ câu lệnh (nếu chưa có)
if (!content.includes('function trichNgayTuCauLenh')) {
  const anchor = "const TRIGGER_LUOTBILL = ['lượt bill', 'luot bill', 'lượt hóa đơn', 'luot hoa don'];";
  if (content.includes(anchor)) {
    content = content.replace(anchor, [
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
      anchor
    ].join("\n"));
    ok.push('them_ham_trich_ngay');
  } else {
    fail.push('them_ham_trich_ngay (không tìm thấy anchor TRIGGER_LUOTBILL)');
  }
} else {
  ok.push('ham_trich_ngay_da_co_san');
}

// 2. Ghi đè khối nạp file luotbill: thêm cột Tên siêu thị
replaceIfBlockLuotBill([
  "if (nhanDang.loai === 'luotbill') {",
  "    const idxNgay = header.indexOf('Ngày xuất');",
  "    const idxLuotBill = header.indexOf('Lượt Bill của Siêu Thị');",
  "    const idxTenST = header.indexOf('Tên siêu thị');",
  "",
  "    const rowsLB = dataRows.map((row) => [",
  "      toDateTimeKeyForBucket(row[idxNgay]),",
  "      row[idxLuotBill],",
  "      idxTenST === -1 ? '' : row[idxTenST],",
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
  "  }"
].join("\n"));

// 3. Ghi đè toàn bộ hàm generateLuotBillReport (bất kể bản cũ đang là gì)
const newGenerateFunc = `async function generateLuotBillReport(ngayYeuCau) {
  const sheets = getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: \`\${GOOGLE_SHEET_TAB_LUOTBILL}!A2:C\`,
  });

  const rows = res.data.values || [];
  if (rows.length === 0) {
    return { type: 'text', text: 'Chưa có dữ liệu Lượt Bill nào được nạp.' };
  }

  let ngayMoiNhat = null;
  let tenSieuThi = '';
  const cacNgayCoDuLieu = new Set();
  for (const [dt, , tenST] of rows) {
    if (!dt) continue;
    const ngay = dt.toString().slice(0, 10);
    cacNgayCoDuLieu.add(ngay);
    if (ngayMoiNhat === null || ngay > ngayMoiNhat) ngayMoiNhat = ngay;
    if (tenST) tenSieuThi = tenST;
  }

  const ngayMucTieu = ngayYeuCau || ngayMoiNhat;
  if (ngayYeuCau && !cacNgayCoDuLieu.has(ngayYeuCau)) {
    return {
      type: 'text',
      text: \`Không có dữ liệu Lượt Bill cho ngày \${fmtNgayHienThi(ngayYeuCau)}. Các ngày hiện có: \${[...cacNgayCoDuLieu].sort().map(fmtNgayHienThi).join(', ')}\`,
    };
  }
  ngayMoiNhat = ngayMucTieu;

  const BUCKET_ORDER = ['05:30-06:00'];
  for (let h = 6; h <= 20; h++) {
    BUCKET_ORDER.push(\`\${String(h).padStart(2, '0')}:00-\${String(h + 1).padStart(2, '0')}:00\`);
  }

  const tong = {};
  BUCKET_ORDER.forEach((k) => { tong[k] = 0; });

  function xacDinhKhung(hh, mm) {
    const tongPhut = hh * 60 + mm;
    if (tongPhut >= 5 * 60 + 30 && tongPhut < 6 * 60) return '05:30-06:00';
    if (tongPhut >= 6 * 60 && tongPhut < 21 * 60) {
      const gioBatDau = Math.floor(tongPhut / 60);
      return \`\${String(gioBatDau).padStart(2, '0')}:00-\${String(gioBatDau + 1).padStart(2, '0')}:00\`;
    }
    return null;
  }

  for (const [dt, luotBillRaw] of rows) {
    if (!dt) continue;
    const s = dt.toString();
    const ngay = s.slice(0, 10);
    if (ngay !== ngayMoiNhat) continue;

    const gioPhutMatch = s.match(/(\\d{1,2}):(\\d{2})$/);
    if (!gioPhutMatch) continue;
    const hh = Number(gioPhutMatch[1]);
    const mm = Number(gioPhutMatch[2]);

    const khung = xacDinhKhung(hh, mm);
    if (!khung) continue;

    tong[khung] += Number(luotBillRaw) || 0;
  }

  const tongCong = Object.values(tong).reduce((a, b) => a + b, 0);
  const khungCoLuot = BUCKET_ORDER.filter((k) => tong[k] > 0);
  const khungMax = khungCoLuot.length ? khungCoLuot.reduce((a, b) => (tong[a] >= tong[b] ? a : b)) : null;
  const khungMin = khungCoLuot.length ? khungCoLuot.reduce((a, b) => (tong[a] <= tong[b] ? a : b)) : null;

  const danhSachContents = BUCKET_ORDER.map((khung) => {
    const soLuot = tong[khung];
    const laMax = khung === khungMax && soLuot > 0;
    const laMin = khung === khungMin && soLuot > 0;
    let icon = '';
    let mauChu = '#333333';
    let doDam = 'regular';
    if (laMax) { icon = '🔥 '; mauChu = '#C62828'; doDam = 'bold'; }
    if (laMin) { icon = '❄️ '; mauChu = '#1565C0'; doDam = 'bold'; }

    return {
      type: 'box',
      layout: 'horizontal',
      margin: 'md',
      contents: [
        { type: 'text', text: \`\${icon}\${khung}\`, size: 'sm', color: mauChu, weight: doDam, flex: 3 },
        { type: 'text', text: \`\${soLuot}\`, size: 'sm', color: mauChu, weight: doDam, align: 'end', flex: 1 },
      ],
    };
  });

  return {
    type: 'flex',
    altText: \`Lượt Bill \${tenSieuThi} ngày \${fmtNgayHienThi(ngayMoiNhat)}: Tổng \${tongCong} lượt\${khungMax ? ', cao nhất ' + khungMax : ''}\`,
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: '#1565C0',
        paddingAll: '20px',
        contents: [
          { type: 'text', text: '🧾 BÁO CÁO LƯỢT BILL THEO KHUNG GIỜ', color: '#FFFFFF', weight: 'bold', size: 'lg', wrap: true },
          { type: 'text', text: tenSieuThi ? \`\${tenSieuThi}\` : '', color: '#D6E4F5', size: 'sm', margin: 'sm', wrap: true },
          { type: 'text', text: \`Ngày \${fmtNgayHienThi(ngayMoiNhat)}\`, color: '#D6E4F5', size: 'sm', margin: 'xs' },
        ],
      },
      body: {
        type: 'box',
        layout: 'vertical',
        paddingAll: '16px',
        contents: [
          {
            type: 'box',
            layout: 'vertical',
            backgroundColor: '#E3F2FD',
            cornerRadius: '8px',
            paddingAll: '12px',
            contents: [
              { type: 'text', text: 'TỔNG LƯỢT BILL TRONG NGÀY', size: 'xs', color: '#666666' },
              { type: 'text', text: \`\${tongCong}\`, size: 'xxl', weight: 'bold', color: '#1565C0' },
              { type: 'text', text: khungMax ? \`🔥 Cao nhất: \${khungMax} (\${tong[khungMax]} lượt)\` : '', size: 'xs', color: '#C62828', margin: 'sm' },
              { type: 'text', text: khungMin ? \`❄️ Thấp nhất: \${khungMin} (\${tong[khungMin]} lượt)\` : '', size: 'xs', color: '#1565C0', margin: 'xs' },
            ],
          },
          { type: 'separator', margin: 'lg' },
          { type: 'box', layout: 'vertical', margin: 'lg', spacing: 'sm', contents: danhSachContents },
        ],
      },
    },
  };
}`;

replaceWholeFunction('generateLuotBillReport', 'async function generateLuotBillReport', newGenerateFunc);

// 4. Đảm bảo dispatcher gọi hàm kèm ngày trích được (thay bất kể đang gọi kiểu nào)
const dispatcherRegex = /if \(laTriggerLuotBill\(text\)\) return \{ ten: 'Lượt Bill', ket: await generateLuotBillReport\([^)]*\) \};/;
if (dispatcherRegex.test(content)) {
  content = content.replace(dispatcherRegex, "if (laTriggerLuotBill(text)) return { ten: 'Lượt Bill', ket: await generateLuotBillReport(trichNgayTuCauLenh(text)) };");
  ok.push('dispatcher_luotbill');
} else {
  fail.push('dispatcher_luotbill (không tìm thấy dòng gọi hàm trong dispatcher)');
}

fs.writeFileSync(path, content, 'utf8');

console.log('===== KẾT QUẢ =====');
console.log('THÀNH CÔNG:', ok.join(', ') || '(không có)');
console.log('THẤT BẠI (cần báo lại):', fail.join(', ') || '(không có - TẤT CẢ ĐỀU OK)');
