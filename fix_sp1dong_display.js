const fs = require('fs');
const path = 'server.js';
let content = fs.readFileSync(path, 'utf8');

const oldBlock = `async function generateSp1DongReport() {
  const sheets = getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: \`\${GOOGLE_SHEET_TAB_CHITIETXUAT}!A2:C\`,
  });

  const rows = res.data.values || [];
  const grouped = {};

  for (const [tenSP, soLuong, giaBan] of rows) {
    if (!tenSP) continue;
    if (Number(giaBan) !== 1) continue;
    if (tenSP.toUpperCase().includes('NẤM')) continue;

    const sl = Number(soLuong) || 0;
    grouped[tenSP] = (grouped[tenSP] || 0) + sl;
  }

  const sorted = Object.entries(grouped)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15);

  if (sorted.length === 0) {
    return { type: 'text', text: 'Không có sản phẩm giá bán 1 đồng nào (đã loại nấm) trong dữ liệu hiện tại.' };
  }

  let message = '📋 BÁO CÁO SP GIÁ BÁN 1 ĐỒNG (đã loại nấm)\\n\\n';
  sorted.forEach(([ten, sl], idx) => {
    message += \`\${idx + 1}. \${ten}: \${sl}\\n\`;
  });

  return { type: 'text', text: message };
}`;

const newBlock = `async function generateSp1DongReport() {
  const sheets = getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: \`\${GOOGLE_SHEET_TAB_CHITIETXUAT}!A2:C\`,
  });

  const rows = res.data.values || [];
  const grouped = {};
  let tongSoLuong = 0;

  for (const [tenSP, soLuong, giaBan] of rows) {
    if (!tenSP) continue;
    if (Number(giaBan) !== 1) continue;
    if (tenSP.toUpperCase().includes('NẤM')) continue;

    const sl = Number(soLuong) || 0;
    grouped[tenSP] = (grouped[tenSP] || 0) + sl;
    tongSoLuong += sl;
  }

  const allSorted = Object.entries(grouped).sort((a, b) => b[1] - a[1]);
  const soMatHang = allSorted.length;
  const sorted = allSorted.slice(0, 15);

  if (sorted.length === 0) {
    return { type: 'text', text: 'Không có sản phẩm giá bán 1 đồng nào (đã loại nấm) trong dữ liệu hiện tại.' };
  }

  let message = '📦 BÁO CÁO SP GIÁ BÁN 1 ĐỒNG (đã loại nấm)\\n';
  message += '━━━━━━━━━━━━━━━━━━━\\n';
  message += \`🔢 Tổng SL đã bán: \${tongSoLuong}\\n\`;
  message += \`📋 Số mặt hàng: \${soMatHang}\\n\`;
  message += '━━━━━━━━━━━━━━━━━━━\\n\\n';
  sorted.forEach(([ten, sl], idx) => {
    message += \`\${idx + 1}. \${ten}\\n    ↳ SL: \${sl}\\n\\n\`;
  });

  return { type: 'text', text: message.trim() };
}`;

if (content.includes(oldBlock)) {
  content = content.replace(oldBlock, newBlock);
  fs.writeFileSync(path, content, 'utf8');
  console.log('✅ Đã cập nhật giao diện báo cáo SP 1 ĐỒNG');
} else {
  console.log('❌ KHÔNG khớp - hàm generateSp1DongReport có thể đã bị sửa khác đi. Báo lại Claude.');
}
