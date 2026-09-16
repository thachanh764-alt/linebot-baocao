const fs = require('fs');
let s = fs.readFileSync('server.js', 'utf8');
const oldStr = "console.error('[webhook] LOI_CHI_TIET: ' + JSON.stringify(err.response && err.response.data ? err.response.data : err.message));";
const newStr = "const chiTietLoi = (err.originalError && err.originalError.response && err.originalError.response.data) || (err.response && err.response.data) || err.message; console.error('[webhook] LOI_CHI_TIET: ' + JSON.stringify(chiTietLoi));";
if (!s.includes(oldStr)) {
  console.log('KHÔNG TÌM THẤY ĐOẠN CẦN SỬA - báo lại Claude, đừng sửa tay');
  process.exit(1);
}
s = s.replace(oldStr, newStr);
fs.writeFileSync('server.js', s);
console.log('ĐÃ SỬA XONG');
