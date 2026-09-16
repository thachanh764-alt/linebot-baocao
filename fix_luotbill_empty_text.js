const fs = require('fs');
const path = 'server.js';
let content = fs.readFileSync(path, 'utf8');

const oldHeader = `      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: '#1565C0',
        paddingAll: '20px',
        contents: [
          { type: 'text', text: '🧾 BÁO CÁO LƯỢT BILL THEO KHUNG GIỜ', color: '#FFFFFF', weight: 'bold', size: 'lg', wrap: true },
          { type: 'text', text: tenSieuThi ? \`\${tenSieuThi}\` : '', color: '#D6E4F5', size: 'sm', margin: 'sm', wrap: true },
          { type: 'text', text: \`Ngày \${fmtNgayHienThi(ngayMoiNhat)}\`, color: '#D6E4F5', size: 'sm', margin: 'xs' },
        ],
      },`;

const newHeader = `      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: '#1565C0',
        paddingAll: '20px',
        contents: [
          { type: 'text', text: '🧾 BÁO CÁO LƯỢT BILL THEO KHUNG GIỜ', color: '#FFFFFF', weight: 'bold', size: 'lg', wrap: true },
          ...(tenSieuThi ? [{ type: 'text', text: tenSieuThi, color: '#D6E4F5', size: 'sm', margin: 'sm', wrap: true }] : []),
          { type: 'text', text: \`Ngày \${fmtNgayHienThi(ngayMoiNhat)}\`, color: '#D6E4F5', size: 'sm', margin: tenSieuThi ? 'xs' : 'sm' },
        ],
      },`;

const oldBox = `            contents: [
              { type: 'text', text: 'TỔNG LƯỢT BILL TRONG NGÀY', size: 'xs', color: '#666666' },
              { type: 'text', text: \`\${tongCong}\`, size: 'xxl', weight: 'bold', color: '#1565C0' },
              { type: 'text', text: khungMax ? \`🔥 Cao nhất: \${khungMax} (\${tong[khungMax]} lượt)\` : '', size: 'xs', color: '#C62828', margin: 'sm' },
              { type: 'text', text: khungMin ? \`❄️ Thấp nhất: \${khungMin} (\${tong[khungMin]} lượt)\` : '', size: 'xs', color: '#1565C0', margin: 'xs' },
            ],`;

const newBox = `            contents: [
              { type: 'text', text: 'TỔNG LƯỢT BILL TRONG NGÀY', size: 'xs', color: '#666666' },
              { type: 'text', text: \`\${tongCong}\`, size: 'xxl', weight: 'bold', color: '#1565C0' },
              ...(khungMax ? [{ type: 'text', text: \`🔥 Cao nhất: \${khungMax} (\${tong[khungMax]} lượt)\`, size: 'xs', color: '#C62828', margin: 'sm' }] : []),
              ...(khungMin ? [{ type: 'text', text: \`❄️ Thấp nhất: \${khungMin} (\${tong[khungMin]} lượt)\`, size: 'xs', color: '#1565C0', margin: 'xs' }] : []),
            ],`;

let count = 0;
if (content.includes(oldHeader)) { content = content.replace(oldHeader, newHeader); count++; }
if (content.includes(oldBox)) { content = content.replace(oldBox, newBox); count++; }

fs.writeFileSync(path, content, 'utf8');
console.log('Đã sửa:', count, '/ 2 chỗ');
