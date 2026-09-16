const fs = require('fs');
const path = 'server.js';
let content = fs.readFileSync(path, 'utf8');

const marker = 'async function generateSp1DongReport() {';
const startIdx = content.indexOf(marker);
if (startIdx === -1) {
  console.log('❌ KHÔNG tìm thấy hàm generateSp1DongReport. Báo lại Claude.');
  process.exit(0);
}

// Tìm dấu đóng ngoặc } của hàm (dòng đầu tiên chỉ có "}" sau marker)
const afterMarker = content.slice(startIdx);
const endMatch = afterMarker.match(/\n}\n/);
if (!endMatch) {
  console.log('❌ KHÔNG xác định được điểm kết thúc hàm. Báo lại Claude.');
  process.exit(0);
}
const endIdx = startIdx + endMatch.index + endMatch[0].length;

const newFunction = `async function generateSp1DongReport() {
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

  const danhSachContents = [];
  sorted.forEach(([ten, sl], idx) => {
    danhSachContents.push({
      type: 'box',
      layout: 'horizontal',
      margin: idx === 0 ? 'none' : 'md',
      contents: [
        { type: 'text', text: \`\${idx + 1}. \${ten}\`, size: 'sm', color: '#333333', wrap: true, flex: 4 },
        { type: 'text', text: \`\${sl}\`, size: 'sm', color: '#2E7D32', weight: 'bold', align: 'end', flex: 1 },
      ],
    });
  });

  return {
    type: 'flex',
    altText: \`Báo cáo SP 1 Đồng: Tổng SL \${tongSoLuong}, \${soMatHang} mặt hàng\`,
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: '#2E7D32',
        paddingAll: '20px',
        contents: [
          { type: 'text', text: '📦 BÁO CÁO SP GIÁ BÁN 1 ĐỒNG', color: '#FFFFFF', weight: 'bold', size: 'lg', wrap: true },
          { type: 'text', text: '(đã loại trừ sản phẩm nấm)', color: '#D6F5D6', size: 'sm', margin: 'sm' },
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
            backgroundColor: '#F1F8E9',
            cornerRadius: '8px',
            paddingAll: '12px',
            contents: [
              { type: 'text', text: 'TỔNG SỐ LƯỢNG ĐÃ BÁN', size: 'xs', color: '#666666' },
              { type: 'text', text: \`\${tongSoLuong}\`, size: 'xxl', weight: 'bold', color: '#2E7D32' },
              { type: 'text', text: \`\${soMatHang} mặt hàng khác nhau\`, size: 'xs', color: '#888888', margin: 'sm' },
            ],
          },
          { type: 'separator', margin: 'lg' },
          { type: 'box', layout: 'vertical', margin: 'lg', spacing: 'sm', contents: danhSachContents },
        ],
      },
    },
  };
}
`;

content = content.slice(0, startIdx) + newFunction + content.slice(endIdx);
fs.writeFileSync(path, content, 'utf8');
console.log('✅ Đã chuyển generateSp1DongReport sang dạng thẻ Flex đẹp');
