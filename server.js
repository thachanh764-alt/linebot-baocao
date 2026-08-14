/**
 * server.js
 * =========
 * LINE Bot "Báo cáo trà"
 * -----------------------
 * Khi có người nhắn "Báo cáo trà" trong LINE OA, bot sẽ:
 *   1. Đọc trực tiếp 2 TAB trong Google Sheet (bằng service account) -
 *      1 tab chứa data TỒN, 1 tab chứa data DOANH THU (anh copy/dán
 *      nguyên nội dung 2 file Excel vào 2 tab này, giữ nguyên hàng
 *      tiêu đề cột giống file gốc).
 *   2. Lọc đúng 6 sản phẩm trà C2, gộp theo siêu thị, quy đổi thùng.
 *   3. Trả lời lại đúng định dạng báo cáo (dạng thẻ Flex Message đẹp).
 *
 * CẦN CHUẨN BỊ TRƯỚC KHI CHẠY (điền vào file .env cùng thư mục, hoặc biến
 * môi trường trên Render):
 * ------------------------------------------------------------
 *   LINE_CHANNEL_ACCESS_TOKEN=...........(lấy trong LINE Developers Console)
 *   LINE_CHANNEL_SECRET=..................(lấy trong LINE Developers Console)
 *   GOOGLE_SERVICE_ACCOUNT_JSON=..........(nội dung file json service account,
 *      dán nguyên cả JSON vào biến này - dùng khi deploy lên Render)
 *   HOẶC GOOGLE_SERVICE_ACCOUNT_KEY_PATH=./service-account.json
 *      (đường dẫn tới file json, chỉ dùng khi chạy local trên máy)
 *   GOOGLE_SHEET_ID=..............(ID của Google Sheet - lấy trong URL:
 *      docs.google.com/spreadsheets/d/<ID_Ở_ĐÂY>/edit)
 *   GOOGLE_SHEET_TAB_TON=TON       (tên tab chứa data Tồn)
 *   GOOGLE_SHEET_TAB_DOANHTHU=DOANHTHU  (tên tab chứa data Doanh thu)
 *   PORT=3000
 *
 * Và phải SHARE Google Sheet đó cho email của service account
 * (vd: linebot-sheets@linebot-baocao.iam.gserviceaccount.com) với quyền
 * ít nhất là "Viewer" (Người xem), nếu không bot sẽ không đọc được data.
 *
 * QUAN TRỌNG: 2 tab đó phải giữ nguyên tên cột giống file gốc ở hàng đầu
 * tiên: tab TỒN cần có cột "Mã Model", "Tên siêu thị", "Tồn kho siêu thị";
 * tab DOANH THU cần có cột "Mã Model", "Tên siêu thị", "Tổng số lượng".
 *
 * Chạy: npm start
 */

require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const { google } = require('googleapis');

// ---------------------------------------------------------------------------
// CẤU HÌNH
// ---------------------------------------------------------------------------
const config = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

const GOOGLE_SERVICE_ACCOUNT_KEY_PATH = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH;
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SHEET_TAB_TON = process.env.GOOGLE_SHEET_TAB_TON || 'TON';
const GOOGLE_SHEET_TAB_DOANHTHU = process.env.GOOGLE_SHEET_TAB_DOANHTHU || 'DOANHTHU';
const PORT = process.env.PORT || 3000;

// Từ khoá kích hoạt bot (không phân biệt hoa/thường, dấu cách thừa)
const TRIGGER_KEYWORDS = ['báo cáo trà', 'bao cao tra'];

// 6 sản phẩm trà cần báo cáo, khoá theo Mã Model
const SAN_PHAM_TRA = {
  '2601001494': 'Nước sâm C2 Cool',
  '2203000875': 'Trà đen dâu anh đào C2',
  '2602001178': 'Trà đen tắc C2',
  '2006000354': 'Trà hồng vải C2',
  '2204000011': 'Trà xanh chanh bạc hà C2',
  '1607002174': 'Nước C2 trà xanh hương chanh 360ml',
};

const QUY_DOI_THUNG = 24; // 1 thùng = 24 chai

// ---------------------------------------------------------------------------
// GOOGLE SHEETS: đọc trực tiếp 2 tab TON / DOANHTHU
// ---------------------------------------------------------------------------
function getSheetsClient() {
  const scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly'];
  let auth;

  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    // Trên Render: credential được lưu thẳng dưới dạng nội dung JSON trong biến môi trường
    const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    auth = new google.auth.GoogleAuth({ credentials, scopes });
  } else if (GOOGLE_SERVICE_ACCOUNT_KEY_PATH) {
    // Chạy local: credential là 1 file .json nằm trên đĩa
    auth = new google.auth.GoogleAuth({ keyFile: GOOGLE_SERVICE_ACCOUNT_KEY_PATH, scopes });
  } else {
    throw new Error('Thiếu credential Google: cần GOOGLE_SERVICE_ACCOUNT_JSON hoặc GOOGLE_SERVICE_ACCOUNT_KEY_PATH');
  }

  return google.sheets({ version: 'v4', auth });
}

async function docTabThanhMangDong(sheets, tenTab) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: tenTab, // lấy toàn bộ tab
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const rows = res.data.values || [];
  if (rows.length === 0) {
    throw new Error(`Tab "${tenTab}" trong Google Sheet đang trống hoặc không tồn tại`);
  }
  return rows;
}

function timCotTheoTen(headerRow, tenCot) {
  const idx = headerRow.findIndex((h) => (h || '').toString().trim() === tenCot);
  if (idx === -1) throw new Error(`Không tìm thấy cột "${tenCot}"`);
  return idx;
}

function docTon(rows) {
  const header = rows[0];
  const colMa = timCotTheoTen(header, 'Mã Model');
  const colTenST = timCotTheoTen(header, 'Tên siêu thị');
  const colTon = timCotTheoTen(header, 'Tồn kho siêu thị');

  const ton = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const ma = (row[colMa] || '').toString().trim();
    if (SAN_PHAM_TRA[ma]) {
      const st = row[colTenST];
      ton[st] = (ton[st] || 0) + (Number(row[colTon]) || 0);
    }
  }
  return ton;
}

function docBan(rows) {
  const header = rows[0];
  const colMa = timCotTheoTen(header, 'Mã Model');
  const colTenST = timCotTheoTen(header, 'Tên siêu thị');
  const colSL = timCotTheoTen(header, 'Tổng số lượng');

  const ban = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const ma = (row[colMa] || '').toString().trim();
    if (SAN_PHAM_TRA[ma]) {
      const st = row[colTenST];
      ban[st] = (ban[st] || 0) + (Number(row[colSL]) || 0);
    }
  }
  return ban;
}

// ---------------------------------------------------------------------------
// TÍNH TOÁN
// ---------------------------------------------------------------------------
function tenNganSieuThi(tenDayDu) {
  if (!tenDayDu) return '';
  const idx = tenDayDu.indexOf(' - ');
  return idx === -1 ? tenDayDu.trim() : tenDayDu.slice(idx + 3).trim();
}

function fmtSo(n) {
  return Math.round(n)
    .toString()
    .replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function fmtPct(p) {
  const r = Math.round(p * 10) / 10;
  return Number.isInteger(r) ? `${r}%` : `${r.toFixed(1)}%`;
}

function tinhDuLieu(ton, ban) {
  const tatCaSieuThi = new Set([...Object.keys(ton), ...Object.keys(ban)]);
  const rows = [];

  for (const st of tatCaSieuThi) {
    const tonUnits = ton[st] || 0;
    const banUnits = ban[st] || 0;
    const tonThung = tonUnits / QUY_DOI_THUNG;
    const banThung = banUnits / QUY_DOI_THUNG;
    const mauSo = tonThung + banThung;
    // TL% = % đã bán so với tổng lượng có trong kỳ (tồn còn lại + đã bán)
    const tl = mauSo > 0 ? (banThung / mauSo) * 100 : 0;
    rows.push({ ten: tenNganSieuThi(st), ton: tonThung, ban: banThung, tl });
  }

  rows.sort((a, b) => b.ban - a.ban || b.ton - a.ton);

  const tongTon = rows.reduce((s, r) => s + r.ton, 0);
  const tongBan = rows.reduce((s, r) => s + r.ban, 0);
  const tongMauSo = tongTon + tongBan;
  const tongTl = tongMauSo > 0 ? (tongBan / tongMauSo) * 100 : 0;

  return { rows, tong: { ton: tongTon, ban: tongBan, tl: tongTl, soSieuThi: rows.length } };
}

// ---------------------------------------------------------------------------
// FLEX MESSAGE (thẻ đẹp giống ảnh mẫu) - dùng để gửi qua LINE
// ---------------------------------------------------------------------------
const MAU_XANH_HEADER = '#2C4A3B';
const MAU_XANH_TOT = '#2ECC71';
const MAU_DO_XAU = '#E74C3C';

function dongBang(label, ton, ban, tl, dam) {
  return {
    type: 'box',
    layout: 'horizontal',
    contents: [
      { type: 'text', text: label, size: 'sm', flex: 5, wrap: true, weight: dam ? 'bold' : 'regular', color: dam ? '#1a1a1a' : '#333333' },
      { type: 'text', text: ton, size: 'sm', flex: 2, align: 'end', weight: dam ? 'bold' : 'regular' },
      { type: 'text', text: ban, size: 'sm', flex: 2, align: 'end', weight: dam ? 'bold' : 'regular' },
      { type: 'text', text: tl, size: 'sm', flex: 2, align: 'end', weight: 'bold', color: dam ? MAU_DO_XAU : undefined },
    ],
  };
}

function taoFlexBaoCao(ton, ban) {
  const { rows, tong } = tinhDuLieu(ton, ban);
  const now = new Date();
  const thoiGian = now.toLocaleString('vi-VN', {
    hour: '2-digit', minute: '2-digit', day: '2-digit', month: '2-digit', year: 'numeric',
  });

  const headerContents = [
    {
      type: 'text',
      text: '🍵 BÁO CÁO TRÀ',
      color: '#FFFFFF',
      weight: 'bold',
      size: 'lg',
    },
    ...Object.values(SAN_PHAM_TRA).map((ten) => ({
      type: 'text',
      text: `• ${ten}`,
      color: '#EAEAEA',
      size: 'sm',
      wrap: true,
    })),
    {
      type: 'text',
      text: `${Object.keys(SAN_PHAM_TRA).length} sản phẩm · cập nhật lúc ${thoiGian} · ${tong.soSieuThi} siêu thị · quy đổi 1 thùng = ${QUY_DOI_THUNG} chai`,
      color: '#CFCFCF',
      size: 'xs',
      wrap: true,
      margin: 'md',
    },
  ];

  const bodyContents = [
    dongBang('Siêu thị', 'Tồn', 'Bán', 'TL%', false),
    { type: 'separator', margin: 'sm' },
    dongBang('TỔNG TẤT CẢ', fmtSo(tong.ton), fmtSo(tong.ban), fmtPct(tong.tl), true),
    { type: 'separator', margin: 'sm' },
  ];

  rows.forEach((r) => {
    bodyContents.push({
      type: 'box',
      layout: 'horizontal',
      margin: 'sm',
      contents: [
        { type: 'text', text: r.ten, size: 'sm', flex: 5, wrap: true, color: '#333333' },
        { type: 'text', text: fmtSo(r.ton), size: 'sm', flex: 2, align: 'end', color: '#333333' },
        { type: 'text', text: fmtSo(r.ban), size: 'sm', flex: 2, align: 'end', color: '#333333' },
        {
          type: 'text',
          text: fmtPct(r.tl),
          size: 'sm',
          flex: 2,
          align: 'end',
          weight: 'bold',
          color: r.tl >= 20 ? MAU_XANH_TOT : MAU_DO_XAU,
        },
      ],
    });
  });

  const altText = `Báo cáo trà: Tồn ${fmtSo(tong.ton)} thùng | Bán ${fmtSo(tong.ban)} thùng | TL ${fmtPct(tong.tl)} (${tong.soSieuThi} siêu thị)`;

  return {
    type: 'flex',
    altText: altText.slice(0, 400),
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: MAU_XANH_HEADER,
        paddingAll: '20px',
        contents: headerContents,
      },
      body: {
        type: 'box',
        layout: 'vertical',
        paddingAll: '16px',
        spacing: 'sm',
        contents: bodyContents,
      },
    },
  };
}

// ---------------------------------------------------------------------------
// TẠO BÁO CÁO ĐẦY ĐỦ: ĐỌC 2 TAB TRONG GOOGLE SHEET -> FLEX MESSAGE
// ---------------------------------------------------------------------------
async function generateTraReport() {
  const sheets = getSheetsClient();

  const [rowsTon, rowsDoanhThu] = await Promise.all([
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_TON),
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_DOANHTHU),
  ]);

  const ton = docTon(rowsTon);
  const ban = docBan(rowsDoanhThu);

  return taoFlexBaoCao(ton, ban);
}

// ---------------------------------------------------------------------------
// LINE BOT
// ---------------------------------------------------------------------------
const app = express();
const client = new line.Client(config);

function laTrigger(text) {
  if (!text) return false;
  const t = text.trim().toLowerCase();
  return TRIGGER_KEYWORDS.some((kw) => t === kw || t.includes(kw));
}

app.post('/webhook', line.middleware(config), async (req, res) => {
  // Trả 200 ngay để LINE không retry, xử lý reply ở dưới (không chặn response)
  res.status(200).end();

  const events = req.body.events || [];
  console.log(`[webhook] nhận ${events.length} event(s)`);

  for (const event of events) {
    console.log(`[webhook] event type=${event.type} messageType=${event.message?.type} text="${event.message?.text}"`);

    if (event.type !== 'message' || event.message.type !== 'text') continue;
    if (!laTrigger(event.message.text)) {
      console.log('[webhook] text không khớp từ khoá trigger, bỏ qua');
      continue;
    }

    console.log('[webhook] khớp trigger, đang tạo báo cáo...');
    try {
      const flexMessage = await generateTraReport();
      console.log('[webhook] tạo báo cáo thành công, đang reply...');
      await client.replyMessage(event.replyToken, flexMessage);
      console.log('[webhook] đã reply xong');
    } catch (err) {
      console.error('[webhook] Lỗi tạo báo cáo trà:', err);
      try {
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: `⚠️ Không tạo được báo cáo trà: ${err.message}`,
        });
      } catch (replyErr) {
        console.error('[webhook] Lỗi luôn cả khi reply lỗi (có thể replyToken hết hạn):', replyErr.message);
      }
    }
  }
});

app.get('/health', (req, res) => res.send('ok'));

app.listen(PORT, () => {
  console.log(`LINE bot "Báo cáo trà" đang chạy ở port ${PORT}`);
});
