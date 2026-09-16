const fs = require('fs');
let s = fs.readFileSync('server.js', 'utf8');
const oldStr = "async function chayLenhCu(text) {\n  if (laTriggerNgay(text))";
const newStr = "async function chayLenhCu(text) {\n  text = (text || '').normalize('NFC');\n  if (laTriggerNgay(text))";
if (!s.includes(oldStr)) {
  console.log('KHÔNG TÌM THẤY ĐOẠN CẦN SỬA - báo lại Claude, đừng sửa tay');
  process.exit(1);
}
s = s.replace(oldStr, newStr);
fs.writeFileSync('server.js', s);
console.log('ĐÃ SỬA XONG');
