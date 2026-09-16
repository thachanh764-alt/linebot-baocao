const fs = require('fs');
const path = 'server.js';
let content = fs.readFileSync(path, 'utf8');
const ok = [];
const fail = [];

function chenSauDong(ten, anchor, mauChu) {
  if (!content.includes(anchor)) {
    fail.push(ten + ' (không tìm thấy anchor)');
    return;
  }
  const dongBanQuyen = `{ type: 'text', text: 'Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị | Thạch Phạm Hoàng Anh -  197042', color: '${mauChu}', size: 'xxs', margin: 'sm', wrap: true },`;
  content = content.replace(anchor, anchor + "\n          " + dongBanQuyen);
  ok.push(ten);
}

// 1. Báo cáo ngày
chenSauDong(
  'bao_cao_ngay',
  "{ type: 'text', text: `${maSieuThi} · ${ngayHienThi}`, color: '#DCEAE1', size: 'sm', margin: 'sm' },",
  '#DCEAE1'
);

// 2. Fresh
chenSauDong(
  'fresh',
  "{ type: 'text', text: `${tenSieuThi} · ${ngayHienThi}`, color: '#DCEAE1', size: 'sm', margin: 'sm', wrap: true },",
  '#DCEAE1'
);

// 3. Giá Vốn
chenSauDong(
  'gia_von',
  "{ type: 'text', text: `${maSieuThi} · Tháng ${thang}${ngayBanHienThi ? ' · Bán ngày ' + fmtNgayHienThi(ngayBanHienThi) : ''}`, color: '#F3DFC5', size: 'sm', margin: 'sm', wrap: true },",
  '#F3DFC5'
);

// 4. Huỷ MMKK
chenSauDong(
  'huy_mmkk',
  "{ type: 'text', text: maSieuThi + ' · ' + ngayHienThi, color: '#F5D5D0', size: 'sm', margin: 'sm' },",
  '#F5D5D0'
);

// 5. Luỹ Kế DT
chenSauDong(
  'luy_ke_dt',
  "{ type: 'text', text: `${maSieuThi} · Lũy kế 01-${fmtNgayNgan(so.ngayCuoi)}`, color: '#D6E4F0', size: 'sm', margin: 'sm' },",
  '#D6E4F0'
);

fs.writeFileSync(path, content, 'utf8');

console.log('===== KẾT QUẢ =====');
console.log('THÀNH CÔNG:', ok.join(', ') || '(không có)');
console.log('THẤT BẠI (cần báo lại):', fail.join(', ') || '(không có - TẤT CẢ ĐỀU OK)');
