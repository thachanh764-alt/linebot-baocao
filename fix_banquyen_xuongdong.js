const fs = require('fs');
const path = 'server.js';
let content = fs.readFileSync(path, 'utf8');

const oldText = "Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị | Thạch Phạm Hoàng Anh -  197042";
const newText = "Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị\\nThạch Phạm Hoàng Anh -  197042";

const soLan = content.split(oldText).length - 1;
content = content.split(oldText).join(newText);

fs.writeFileSync(path, content, 'utf8');
console.log('Đã sửa xuống dòng cho', soLan, 'chỗ');
