const fs = require('fs');
const path = 'server.js';
let content = fs.readFileSync(path, 'utf8');
let count = 0;

const DONG_BAN_QUYEN = "{ type: 'text', text: 'Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị | Thạch Phạm Hoàng Anh -  197042', color: '#FFFFFF', size: 'xxs', margin: 'sm', wrap: true },";

// SP 1 Đồng: chèn ngay sau dòng "(đã loại trừ sản phẩm nấm)"
const anchor1 = "{ type: 'text', text: '(đã loại trừ sản phẩm nấm)', color: '#D6F5D6', size: 'sm', margin: 'sm' },";
if (content.includes(anchor1)) {
  content = content.replace(anchor1, anchor1 + "\n          " + DONG_BAN_QUYEN);
  count++;
}

// Lượt Bill: chèn ngay sau dòng ngày hiển thị trong header
const anchor2 = "{ type: 'text', text: `Ngày ${fmtNgayHienThi(ngayMoiNhat)}`, color: '#D6E4F5', size: 'sm', margin: tenSieuThi ? 'xs' : 'sm' },";
if (content.includes(anchor2)) {
  content = content.replace(anchor2, anchor2 + "\n          " + DONG_BAN_QUYEN);
  count++;
}

fs.writeFileSync(path, content, 'utf8');
console.log('Đã thêm dòng bản quyền vào:', count, '/ 2 báo cáo');
