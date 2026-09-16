const fs = require('fs');
const path = 'server.js';
let content = fs.readFileSync(path, 'utf8');

const old1 = "    return 'Không có sản phẩm giá bán 1 đồng nào (đã loại nấm) trong dữ liệu hiện tại.';";
const new1 = "    return { type: 'text', text: 'Không có sản phẩm giá bán 1 đồng nào (đã loại nấm) trong dữ liệu hiện tại.' };";

const old2 = "  return message;\n}";
const new2 = "  return { type: 'text', text: message };\n}";

let okCount = 0;
if (content.includes(old1)) { content = content.replace(old1, new1); okCount++; }
if (content.includes(old2)) { content = content.replace(old2, new2); okCount++; }

fs.writeFileSync(path, content, 'utf8');
console.log('Đã sửa:', okCount, '/ 2 chỗ');
