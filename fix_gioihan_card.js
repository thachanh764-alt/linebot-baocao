const fs = require('fs');
let s = fs.readFileSync('server.js', 'utf8');
const oldStr = ".sort((a, b) => b.slMMKK - a.slMMKK)";
const newStr = ".sort((a, b) => b.slMMKK - a.slMMKK)\n    .slice(0, 15)";
if (!s.includes(oldStr)) {
  console.log('KHÔNG TÌM THẤY ĐOẠN CẦN SỬA - báo lại Claude, đừng sửa tay');
  process.exit(1);
}
s = s.replace(oldStr, newStr);
fs.writeFileSync('server.js', s);
console.log('ĐÃ SỬA XONG');
