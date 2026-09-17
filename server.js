/**
 * server.js
 * =========
 * LINE Bot "Báo cáo ngày" + "Bánh Trung Thu" + "BC Luỹ Kế DT" + "Giá Vốn"
 * + nhận file Excel tự động nạp (trừ khi đang trong luồng @tag AI phân tích)
 * + MỚI: bot chỉ trả lời trong group khi được @tag; khi @tag kèm câu hỏi tự do
 *   thì tự đoán tab Google Sheet liên quan rồi dùng AI (Claude) phân tích trả lời;
 *   khi @tag kèm file Excel thì AI chỉ phân tích nhanh theo câu hỏi, KHÔNG lưu vào Sheet;
 *   khi @tag kèm ảnh thì dùng AI vision đọc và phân tích.
 * -----------------------
 * ĐÃ XOÁ: "Báo cáo trà" (theo yêu cầu) — không còn đọc tab TON/DOANHTHU trà C2 nữa.
 *
 * "Báo cáo ngày": nhắn "Báo cáo ngày" (trong group PHẢI @tag bot trước) -> đọc 3 tab
 *   DOANHTHU_SIEUTHI/DOANHTHU_NGANHHANG/FRESH_NHAPXUAT, ra thẻ theo từng siêu thị.
 * "Bánh Trung Thu": nhắn "Bánh Trung Thu" -> đọc 2 tab BANHTT_TON/BANHTT_DOANHTHU,
 *   tự tính thưởng Cái 1.000đ / Hộp 4.000đ theo từng siêu thị.
 * "BC Luỹ Kế DT": nhắn "BC Luỹ Kế DT" -> đọc tab LUYKE_DT (nhiều ngày/tháng),
 *   tính lũy kế + dự kiến hết tháng + so sánh MoM Doanh thu/Lượt bill/Giá trị bill.
 * "Giá Vốn": nhắn "Giá Vốn" -> đọc tab GIAVON (giá vốn hôm nay + lũy kế) VÀ tab
 *   FRESH_NHAPXUAT (doanh thu bán thực tế hôm nay), ghép chung theo từng ngành hàng,
 *   ra thẻ theo từng siêu thị.
 * Gửi file Excel trực tiếp vào group -> bot tự nhận diện loại file, GHI ĐÈ vào đúng
 *   tab, tự trả báo cáo — TRỪ khi anh vừa @tag bot hỏi gì đó trước (trong 3 phút),
 *   lúc đó file gửi tiếp theo sẽ được AI phân tích nhanh, KHÔNG ghi vào Sheet.
 *
 * CẦN CHUẨN BỊ (biến môi trường trên Render, hoặc file .env khi chạy local):
 * ------------------------------------------------------------
 *   LINE_CHANNEL_ACCESS_TOKEN, LINE_CHANNEL_SECRET
 *   GOOGLE_SERVICE_ACCOUNT_JSON (nội dung json service account, dùng trên Render)
 *   HOẶC GOOGLE_SERVICE_ACCOUNT_KEY_PATH=./service-account.json (chạy local)
 *   GOOGLE_SHEET_ID
 *   GOOGLE_SHEET_TAB_DOANHTHU_SIEUTHI=DOANHTHU_SIEUTHI
 *   GOOGLE_SHEET_TAB_DOANHTHU_NGANHHANG=DOANHTHU_NGANHHANG
 *   GOOGLE_SHEET_TAB_FRESH=FRESH_NHAPXUAT
 *   GOOGLE_SHEET_TAB_BANHTT_TON=BANHTT_TON
 *   GOOGLE_SHEET_TAB_BANHTT_DOANHTHU=BANHTT_DOANHTHU
 *   GOOGLE_SHEET_TAB_LUYKE_DT=LUYKE_DT
 *   GOOGLE_SHEET_TAB_GIAVON=GIAVON
 *   ANTHROPIC_API_KEY=sk-ant-xxxxx   (lấy tại https://console.anthropic.com/settings/keys — cần nạp tối thiểu 5 USD)
 *   PORT=3000
 *
 * Phải SHARE Google Sheet cho email service account, quyền EDITOR.
 * Cần đủ 7 tab: DOANHTHU_SIEUTHI, DOANHTHU_NGANHHANG, FRESH_NHAPXUAT,
 * BANHTT_TON, BANHTT_DOANHTHU, LUYKE_DT, GIAVON.
 *
 * Tab GIAVON cần đúng các cột (dòng 1, đúng tên như file Excel gửi vào group):
 *   Tháng | Mã siêu thị | Tên siêu thị | Mã ngành hàng | Ngành hàng |
 *   SL thực nhập hôm nay | Giá vốn cơ bản hôm nay | DT FRESH tính giá vốn |
 *   Giá vốn cơ bản lũy kế đến ngày hôm qua | Tỉ lệ doanh thu trên giá vốn cơ bản |
 *   LN lũy kế | LN TB 3 tháng trước | Chênh lệch LN so với 3 tháng trước
 *
 * Cài thêm thư viện mới trước khi chạy: npm install @anthropic-ai/sdk xlsx
 * Chạy: npm start
 */

require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const Anthropic = require('@anthropic-ai/sdk');

// ---------------------------------------------------------------------------
// CẤU HÌNH
// ---------------------------------------------------------------------------
const config = {
  channelAccessToken: (process.env.LINE_CHANNEL_ACCESS_TOKEN || '').replace(/\s+/g, ''),
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

const GOOGLE_SERVICE_ACCOUNT_KEY_PATH = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH;
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const PORT = process.env.PORT || 3000;

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
// Cân bằng chất lượng/giá. Nếu group nhắn nhiều muốn tiết kiệm chi phí,
// đổi thành: 'claude-haiku-4-5-20251001'
const AI_MODEL = 'claude-sonnet-5';

// ---------------------------------------------------------------------------
// GOOGLE SHEETS
// ---------------------------------------------------------------------------
function getSheetsClient() {
  const scopes = ['https://www.googleapis.com/auth/spreadsheets'];
  let auth;

  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    auth = new google.auth.GoogleAuth({ credentials, scopes });
  } else if (GOOGLE_SERVICE_ACCOUNT_KEY_PATH) {
    auth = new google.auth.GoogleAuth({ keyFile: GOOGLE_SERVICE_ACCOUNT_KEY_PATH, scopes });
  } else {
    throw new Error('Thiếu credential Google: cần GOOGLE_SERVICE_ACCOUNT_JSON hoặc GOOGLE_SERVICE_ACCOUNT_KEY_PATH');
  }

  return google.sheets({ version: 'v4', auth });
}

async function docTabThanhMangDong(sheets, tenTab) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: tenTab,
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

// ---------------------------------------------------------------------------
// TIỆN ÍCH DÙNG CHUNG
// ---------------------------------------------------------------------------
function tenNganSieuThi(tenDayDu) {
  if (!tenDayDu) return '';
  const idx = tenDayDu.indexOf(' - ');
  return idx === -1 ? tenDayDu.trim() : tenDayDu.slice(idx + 3).trim();
}

function rutGonTen(ten, maxLen) {
  if (!ten) return '';
  return ten.length > maxLen ? ten.slice(0, maxLen - 1).trim() + '…' : ten;
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

function dongThongTinNgay(icon, nhan, giaTri, dam) {
  return {
    type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: `${icon} ${nhan}`, size: 'sm', flex: 3, color: '#555555' },
      { type: 'text', text: giaTri, size: dam ? 'md' : 'sm', flex: 2, align: 'end', weight: dam ? 'bold' : 'regular', color: dam ? '#22A45D' : '#111111' },
    ],
  };
}

function oThongKe(icon, nhan, giaTri, mau) {
  return {
    type: 'box', layout: 'vertical', flex: 1, backgroundColor: '#F7FAF8', cornerRadius: 'md',
    paddingAll: '10px', spacing: 'xs',
    contents: [
      { type: 'text', text: icon, size: 'lg' },
      { type: 'text', text: nhan, size: 'xxs', color: '#888888' },
      { type: 'text', text: giaTri, size: 'sm', weight: 'bold', color: mau || '#1a1a1a', wrap: true },
    ],
  };
}

// ---------------------------------------------------------------------------
// BÁO CÁO NGÀY
// ---------------------------------------------------------------------------
const GOOGLE_SHEET_TAB_DOANHTHU_SIEUTHI = process.env.GOOGLE_SHEET_TAB_DOANHTHU_SIEUTHI || 'DOANHTHU_SIEUTHI';
const GOOGLE_SHEET_TAB_DOANHTHU_NGANHHANG = process.env.GOOGLE_SHEET_TAB_DOANHTHU_NGANHHANG || 'DOANHTHU_NGANHHANG';
const GOOGLE_SHEET_TAB_FRESH = process.env.GOOGLE_SHEET_TAB_FRESH || 'FRESH_NHAPXUAT';

const CARD1_CATEGORY_ORDER = [
  'Bia Các Loại',
  'Thức uống giải khát các loại',
  'Bánh kẹo - Trà - Cà phê - Bột Dinh Dưỡng các loại',
  'Thực phẩm - Gia vị các loại',
  'Sữa - Thức uống bổ dưỡng các loại',
  'Chăm sóc nhà cửa',
  'Chăm sóc cá nhân',
  'Thực phẩm đông lạnh - Hàng mát các loại',
  'Kem các loại',
  'Sản Phẩm Từ Sữa - Bảo Quản Mát',
  'Thịt',
  'Rau Củ Quả CL',
  'Trái cây',
  'Cá (Hải sản)',
  'Trứng',
  'BHX - Hàng khuyến mãi',
  'Khác',
];

const FRESH_CATEGORY_ORDER = [
  'Thịt Địa Phương',
  'Rau Địa Phương',
  'Trái Cây Tập Trung',
  'Thủy Hải Sản Tập Trung',
  'Trứng Các Loại',
  'Thịt Nhập Khẩu',
  'Rau Đà Lạt',
  'Thủy Hải Sản Nhập Khẩu',
  'Trái Cây Nhập Khẩu',
];

function toDateKey(value) {
  if (typeof value === 'number') {
    const epoch = Date.UTC(1899, 11, 30);
    const d = new Date(epoch + value * 86400000);
    return d.toISOString().slice(0, 10);
  }
  const s = (value || '').toString().trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return s;
}

function fmtNgayHienThi(dateKey) {
  if (!dateKey) return 'N/A';
  const [y, m, d] = dateKey.split('-');
  return `${d}/${m}/${y}`;
}

function chuanHoaMaSieuThi(value) {
  const s = (value === undefined || value === null) ? '' : value.toString().trim();
  const idx = s.indexOf(' - ');
  return idx === -1 ? s : s.slice(0, idx).trim();
}

function timNgayMoiNhat(rows, colNgay) {
  let moiNhat = null;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row[colNgay] === undefined || row[colNgay] === '') continue;
    const key = toDateKey(row[colNgay]);
    if (moiNhat === null || key > moiNhat) moiNhat = key;
  }
  return moiNhat;
}

const NGANH_FRESH_TRONG_CARD1 = ['Thịt', 'Rau Củ Quả CL', 'Trái cây', 'Cá (Hải sản)', 'Trứng'];

function dongNganhHangDon(ten, giaTri) {
  const isFresh = NGANH_FRESH_TRONG_CARD1.includes(ten);
  return {
    type: 'box', layout: 'horizontal', margin: 'sm', contents: [
      { type: 'text', text: `${isFresh ? '🥬' : '🛒'} ${ten}`, size: 'sm', flex: 5, wrap: true, color: isFresh ? '#1F7A45' : '#333333' },
      { type: 'text', text: giaTri, size: 'sm', flex: 3, align: 'end', weight: 'bold', color: '#111111' },
    ],
  };
}

function taoCardBaoCaoTheoNgay(maSieuThi, tenSieuThi, tong, ngayHienThi) {
  const bodyContents = [
    { type: 'text', text: `🏢 ${tenSieuThi}`, weight: 'bold', size: 'md', wrap: true, color: '#1a1a1a' },
    {
      type: 'box', layout: 'vertical', backgroundColor: '#F0F7F2', cornerRadius: 'md', paddingAll: '14px', margin: 'md',
      contents: [
        { type: 'text', text: 'TỔNG DOANH THU', size: 'xs', color: '#888888' },
        { type: 'text', text: fmtSo(tong.tongDoanhThu) + ' đ', size: 'xxl', weight: 'bold', color: '#1a1a1a', margin: 'sm' },
      ],
    },
    {
      type: 'box', layout: 'horizontal', spacing: 'sm', margin: 'md',
      contents: [
        oThongKe('🏬', 'DT Offline', fmtSo(tong.dtOffline) + ' đ'),
        oThongKe('🌐', 'DT Online', fmtSo(tong.dtOnline) + ' đ'),
      ],
    },
    {
      type: 'box', layout: 'horizontal', spacing: 'sm', margin: 'sm',
      contents: [
        oThongKe('🧾', 'Số bill', fmtSo(tong.soBill)),
        oThongKe('💳', 'Giá trị TB', fmtSo(tong.giaTriBillTB) + ' đ'),
      ],
    },
    { type: 'separator', margin: 'lg' },
    { type: 'text', text: '📦 CHI TIẾT NGÀNH HÀNG', size: 'sm', weight: 'bold', color: '#333333', margin: 'lg' },
  ];

  CARD1_CATEGORY_ORDER.forEach((ten) => {
    bodyContents.push(dongNganhHangDon(ten, fmtSo(tong.byNganh[ten] || 0) + ' đ'));
  });

  return {
    type: 'flex',
    altText: `Báo cáo ngày ${ngayHienThi} - ${tenSieuThi}: Tổng doanh thu ${fmtSo(tong.tongDoanhThu)} đ`,
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#2C4A3B', paddingAll: '20px',
        contents: [
          { type: 'text', text: '📅 BÁO CÁO NGÀY', color: '#FFFFFF', weight: 'bold', size: 'lg' },
          { type: 'text', text: `${maSieuThi} · ${ngayHienThi}`, color: '#DCEAE1', size: 'sm', margin: 'sm' },
          { type: 'text', text: 'Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị\nThạch Phạm Hoàng Anh -  197042', color: '#DCEAE1', size: 'xxs', margin: 'sm', wrap: true },
        ],
      },
      body: { type: 'box', layout: 'vertical', paddingAll: '16px', contents: bodyContents },
    },
  };
}

function taoCardFreshNgay(maSieuThi, tenSieuThi, freshData, ngayHienThi) {
  const bodyContents = [
    {
      type: 'box', layout: 'vertical', backgroundColor: '#F0F7F2', cornerRadius: 'md', paddingAll: '14px',
      contents: [
        { type: 'text', text: `TỔNG QUAN — DT FRESH (SIÊU THỊ ${maSieuThi})`, size: 'xs', color: '#888888', wrap: true },
        { type: 'text', text: fmtTrieuTron(freshData.tongDT), size: 'xxl', weight: 'bold', color: '#1a1a1a', margin: 'sm' },
        { type: 'text', text: `SL: ${fmtSo(freshData.tongSL)}`, size: 'sm', color: '#555555', margin: 'sm' },
      ],
    },
    { type: 'separator', margin: 'lg' },
    {
      type: 'box', layout: 'horizontal', margin: 'lg', contents: [
        { type: 'text', text: 'Ngành hàng', size: 'xs', color: '#888888', flex: 5 },
        { type: 'text', text: 'DT', size: 'xs', color: '#888888', flex: 2, align: 'end' },
        { type: 'text', text: 'SL', size: 'xs', color: '#888888', flex: 2, align: 'end' },
      ],
    },
    { type: 'separator', margin: 'sm' },
  ];

  FRESH_CATEGORY_ORDER.forEach((ten) => {
    const d = freshData.byNganh[ten] || { dt: 0, sl: 0 };
    bodyContents.push({
      type: 'box', layout: 'horizontal', margin: 'sm', contents: [
        { type: 'text', text: ten, size: 'sm', flex: 5, wrap: true, color: '#333333' },
        { type: 'text', text: fmtTrieuTron(d.dt), size: 'sm', flex: 2, align: 'end', weight: 'bold' },
        { type: 'text', text: fmtSo(d.sl), size: 'sm', flex: 2, align: 'end' },
      ],
    });
  });

  return {
    type: 'flex',
    altText: `Fresh ngày ${ngayHienThi} - ${tenSieuThi}: ${fmtTrieuTron(freshData.tongDT)}`,
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#2C4A3B', paddingAll: '20px',
        contents: [
          { type: 'text', text: '🌱 BÁO CÁO FRESH THEO NGÀY', color: '#FFFFFF', weight: 'bold', size: 'lg' },
          { type: 'text', text: `${tenSieuThi} · ${ngayHienThi}`, color: '#DCEAE1', size: 'sm', margin: 'sm', wrap: true },
          { type: 'text', text: 'Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị\nThạch Phạm Hoàng Anh -  197042', color: '#DCEAE1', size: 'xxs', margin: 'sm', wrap: true },
        ],
      },
      body: { type: 'box', layout: 'vertical', paddingAll: '16px', contents: bodyContents },
    },
  };
}

function fmtTrieuTron(n) {
  if (Math.abs(n) < 500000) return fmtSo(n) + ' đ';
  return `${Math.round(n / 1e6)} tr`;
}

async function generateDailyReport() {
  const sheets = getSheetsClient();

  const [rowsST, rowsNH, rowsFresh] = await Promise.all([
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_DOANHTHU_SIEUTHI),
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_DOANHTHU_NGANHHANG),
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_FRESH),
  ]);

  const headerST = rowsST[0];
  const colNgayST = timCotTheoTen(headerST, 'Ngày');
  const colMaST = timCotTheoTen(headerST, 'Mã siêu thị');
  const colTenST = timCotTheoTen(headerST, 'Tên siêu thị');
  const colDTOffline = timCotTheoTen(headerST, 'Doanh thu offline');
  const colDTOnline = timCotTheoTen(headerST, 'Doanh thu Online');
  const colSoBill = timCotTheoTen(headerST, 'Tổng số bill');

  const ngayMoiNhatST = timNgayMoiNhat(rowsST, colNgayST);
  const theoSieuThiST = {};
  for (let i = 1; i < rowsST.length; i++) {
    const row = rowsST[i];
    if (!row || row[colNgayST] === undefined || row[colNgayST] === '') continue;
    if (toDateKey(row[colNgayST]) !== ngayMoiNhatST) continue;
    const ma = chuanHoaMaSieuThi(row[colMaST]);
    if (!theoSieuThiST[ma]) {
      theoSieuThiST[ma] = { ma, ten: row[colTenST] || ma, dtOffline: 0, dtOnline: 0, soBill: 0 };
    }
    theoSieuThiST[ma].dtOffline += Number(row[colDTOffline]) || 0;
    theoSieuThiST[ma].dtOnline += Number(row[colDTOnline]) || 0;
    theoSieuThiST[ma].soBill += Number(row[colSoBill]) || 0;
  }

  const headerNH = rowsNH[0];
  const colNgayNH = timCotTheoTen(headerNH, 'Ngày xuất');
  const colMaNH = timCotTheoTen(headerNH, 'Mã siêu thị');
  const colNganhNH = timCotTheoTen(headerNH, 'Ngành hàng BHX');
  const colDoanhThuNH = timCotTheoTen(headerNH, 'Doanh thu');

  const ngayMoiNhatNH = timNgayMoiNhat(rowsNH, colNgayNH);
  const theoSieuThiNH = {};
  for (let i = 1; i < rowsNH.length; i++) {
    const row = rowsNH[i];
    if (!row || row[colNgayNH] === undefined || row[colNgayNH] === '') continue;
    if (toDateKey(row[colNgayNH]) !== ngayMoiNhatNH) continue;
    const ma = chuanHoaMaSieuThi(row[colMaNH]);
    const ten = (row[colNganhNH] || '').toString().trim();
    if (!ten) continue;
    if (!theoSieuThiNH[ma]) theoSieuThiNH[ma] = {};
    theoSieuThiNH[ma][ten] = (theoSieuThiNH[ma][ten] || 0) + (Number(row[colDoanhThuNH]) || 0);
  }

  const ngayHienThi = fmtNgayHienThi(ngayMoiNhatST || ngayMoiNhatNH);
  const tatCaMaSieuThi = new Set([...Object.keys(theoSieuThiST), ...Object.keys(theoSieuThiNH)]);
  const cardsDoanhThu = Array.from(tatCaMaSieuThi).map((ma) => {
    const st = theoSieuThiST[ma] || { ma, ten: ma, dtOffline: 0, dtOnline: 0, soBill: 0 };
    const byNganh = theoSieuThiNH[ma] || {};
    const tongDoanhThu = st.dtOffline + st.dtOnline;
    const giaTriBillTB = st.soBill > 0 ? tongDoanhThu / st.soBill : 0;
    return taoCardBaoCaoTheoNgay(st.ma, st.ten, { tongDoanhThu, dtOffline: st.dtOffline, dtOnline: st.dtOnline, soBill: st.soBill, giaTriBillTB, byNganh }, ngayHienThi);
  });

  const headerFresh = rowsFresh[0];
  const colNgayFresh = timCotTheoTen(headerFresh, 'Ngày');
  const colMaSTFresh = timCotTheoTen(headerFresh, 'Mã siêu thị');
  const colTenSTFresh = timCotTheoTen(headerFresh, 'Tên siêu thị');
  const colNganhFresh = timCotTheoTen(headerFresh, 'Ngành hàng - Phân tích');
  const colDoanhThuFresh = timCotTheoTen(headerFresh, 'Thành tiền phải thu khách hàng (chưa VAT)');
  const colSLFresh = timCotTheoTen(headerFresh, 'SL thực xuất');

  const ngayMoiNhatFresh = timNgayMoiNhat(rowsFresh, colNgayFresh);
  const theoSieuThi = {};
  for (let i = 1; i < rowsFresh.length; i++) {
    const row = rowsFresh[i];
    if (!row || row[colNgayFresh] === undefined || row[colNgayFresh] === '') continue;
    if (toDateKey(row[colNgayFresh]) !== ngayMoiNhatFresh) continue;

    const ma = chuanHoaMaSieuThi(row[colMaSTFresh]);
    if (!theoSieuThi[ma]) {
      theoSieuThi[ma] = { ma, ten: row[colTenSTFresh] || ma, tongDT: 0, tongSL: 0, byNganh: {} };
    }
    const st = theoSieuThi[ma];
    const ten = (row[colNganhFresh] || '').toString().trim();
    const dt = Number(row[colDoanhThuFresh]) || 0;
    const sl = Number(row[colSLFresh]) || 0;
    if (!st.byNganh[ten]) st.byNganh[ten] = { dt: 0, sl: 0 };
    st.byNganh[ten].dt += dt;
    st.byNganh[ten].sl += sl;
    st.tongDT += dt;
    st.tongSL += sl;
  }

  const ngayHienThiFresh = fmtNgayHienThi(ngayMoiNhatFresh);
  const danhSachSieuThi = Object.values(theoSieuThi);
  const cardsFresh = danhSachSieuThi.map((st) =>
    taoCardFreshNgay(st.ma, st.ten, st, ngayHienThiFresh)
  );

  return [...cardsDoanhThu, ...cardsFresh].slice(0, 5);
}

// ---------------------------------------------------------------------------
// GIÁ VỐN — ghép giá vốn (tab GIAVON) + doanh thu bán thực tế hôm nay (tab FRESH_NHAPXUAT)
// ---------------------------------------------------------------------------
const GOOGLE_SHEET_TAB_GIAVON = process.env.GOOGLE_SHEET_TAB_GIAVON || 'GIAVON';

// Ngành hàng bên tab GIAVON là gộp nhóm rộng hơn ngành hàng chi tiết bên FRESH_NHAPXUAT,
// nên cần ánh xạ 1-nhiều để cộng đúng DT/SL bán hôm nay vào từng dòng GIAVON.
const GIAVON_TO_FRESH_MAP = {
  'Thịt gia cầm gia súc các loại': ['Thịt Địa Phương', 'Thịt Nhập Khẩu'],
  'Thủy Hải Sản Các Loại': ['Thủy Hải Sản Tập Trung', 'Thủy Hải Sản Nhập Khẩu'],
  'Rau Đà Lạt': ['Rau Đà Lạt'],
  'Rau Địa Phương': ['Rau Địa Phương'],
  'Trái cây ngoại CL': ['Trái Cây Nhập Khẩu'],
  'Trái cây nội CL': ['Trái Cây Tập Trung'],
  'Trứng gia cầm các loại': ['Trứng Các Loại'],
};

function chuyenTiLeThanhPhanTram(raw) {
  if (typeof raw === 'number') return raw > 3 ? raw : raw * 100; // đã là % (>3) hay dạng phân số (0.xx)
  const s = (raw || '').toString().replace('%', '').replace(',', '.').trim();
  return parseFloat(s) || 0;
}

function dongNganhHangGiaVon(nganh) {
  const mauLN = nganh.lnLuyKe >= 0 ? '#27AE60' : '#E74C3C';
  const mauChenhLech = nganh.chenhLech >= 0 ? '#27AE60' : '#E74C3C';
  const muiChenhLech = nganh.chenhLech >= 0 ? '▲' : '▼';
  const tiLeHomNay = nganh.giaVonHomNay > 0 ? (nganh.dtBanHomNay / nganh.giaVonHomNay) * 100 : 0;
  const mauHomNay = nganh.dtBanHomNay > nganh.giaVonHomNay ? '#27AE60' : '#E74C3C';

  return {
    type: 'box', layout: 'vertical', margin: 'md', paddingAll: '10px',
    backgroundColor: '#FAF7F2', cornerRadius: 'md',
    contents: [
      { type: 'text', text: nganh.ten, size: 'sm', weight: 'bold', color: '#1a1a1a', wrap: true },

      { type: 'text', text: 'HÔM NAY', size: 'xxs', color: '#B8860B', margin: 'sm' },
      {
        type: 'box', layout: 'horizontal', margin: 'xs', contents: [
          { type: 'text', text: 'DT bán hôm nay', size: 'xxs', color: '#888888', flex: 3 },
          { type: 'text', text: `${fmtSo(nganh.dtBanHomNay)} đ (SL ${fmtSo(nganh.slBanHomNay)})`, size: 'xs', flex: 4, align: 'end', weight: 'bold', color: mauHomNay },
        ],
      },
      {
        type: 'box', layout: 'horizontal', margin: 'xs', contents: [
          { type: 'text', text: 'Giá vốn hôm nay / Tỉ lệ', size: 'xxs', color: '#888888', flex: 3 },
          { type: 'text', text: `${fmtSo(nganh.giaVonHomNay)} đ · ${fmtPct(tiLeHomNay)}`, size: 'xs', flex: 4, align: 'end', weight: 'bold', color: mauHomNay },
        ],
      },

      { type: 'separator', margin: 'sm' },
      { type: 'text', text: 'LŨY KẾ', size: 'xxs', color: '#1B4F72', margin: 'sm' },
      {
        type: 'box', layout: 'horizontal', margin: 'xs', contents: [
          { type: 'text', text: 'Giá vốn lũy kế / Tỉ lệ DT', size: 'xxs', color: '#888888', flex: 3 },
          { type: 'text', text: `${fmtSo(nganh.giaVonLuyKe)} đ · ${fmtPct(nganh.tiLeDTGiaVon)}`, size: 'xs', flex: 4, align: 'end', weight: 'bold' },
        ],
      },
      {
        type: 'box', layout: 'horizontal', margin: 'sm', contents: [
          { type: 'text', text: 'LN lũy kế', size: 'xs', color: '#555555', flex: 3 },
          { type: 'text', text: fmtSo(nganh.lnLuyKe) + ' đ', size: 'sm', flex: 4, align: 'end', weight: 'bold', color: mauLN },
        ],
      },
      {
        type: 'box', layout: 'horizontal', margin: 'xs', contents: [
          { type: 'text', text: 'So với TB 3 tháng trước', size: 'xxs', color: '#888888', flex: 3 },
          { type: 'text', text: `${muiChenhLech} ${fmtSo(Math.abs(nganh.chenhLech))} đ`, size: 'xs', flex: 4, align: 'end', weight: 'bold', color: mauChenhLech },
        ],
      },
    ],
  };
}

function taoCardGiaVon(maSieuThi, tenSieuThi, thang, dsNganhHang, ngayBanHienThi) {
  const tongGiaVonHomNay = dsNganhHang.reduce((s, n) => s + n.giaVonHomNay, 0);
  const tongDTBanHomNay = dsNganhHang.reduce((s, n) => s + (n.dtBanHomNay || 0), 0);
  const tongLNLuyKe = dsNganhHang.reduce((s, n) => s + n.lnLuyKe, 0);
  const mauTongLN = tongLNLuyKe >= 0 ? '#27AE60' : '#E74C3C';
  const mauTongHomNay = tongDTBanHomNay > tongGiaVonHomNay ? '#27AE60' : '#E74C3C';

  const bodyContents = [
    { type: 'text', text: `🏢 ${tenSieuThi}`, weight: 'bold', size: 'md', wrap: true, color: '#1a1a1a' },
    {
      type: 'box', layout: 'vertical', backgroundColor: '#FBF3E7', cornerRadius: 'md', paddingAll: '14px', margin: 'md',
      contents: [
        { type: 'text', text: 'TỔNG LN LŨY KẾ', size: 'xs', color: '#888888' },
        { type: 'text', text: fmtSo(tongLNLuyKe) + ' đ', size: 'xxl', weight: 'bold', color: mauTongLN, margin: 'sm' },
      ],
    },
    {
      type: 'box', layout: 'horizontal', spacing: 'sm', margin: 'md',
      contents: [
        oThongKe('🛒', 'DT bán hôm nay', fmtSo(tongDTBanHomNay) + ' đ', mauTongHomNay),
        oThongKe('💰', 'Giá vốn hôm nay', fmtSo(tongGiaVonHomNay) + ' đ', mauTongHomNay),
      ],
    },
    { type: 'separator', margin: 'lg' },
    { type: 'text', text: '📦 CHI TIẾT THEO NGÀNH HÀNG', size: 'sm', weight: 'bold', color: '#333333', margin: 'lg' },
  ];

  dsNganhHang
    .slice()
    .sort((a, b) => b.lnLuyKe - a.lnLuyKe)
    .forEach((n) => bodyContents.push(dongNganhHangGiaVon(n)));

  return {
    type: 'flex',
    altText: `Báo cáo giá vốn ${tenSieuThi}: LN lũy kế ${fmtSo(tongLNLuyKe)} đ`,
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#7B3F00', paddingAll: '20px',
        contents: [
          { type: 'text', text: '💰 BÁO CÁO GIÁ VỐN', color: '#FFFFFF', weight: 'bold', size: 'lg' },
          { type: 'text', text: `${maSieuThi} · Tháng ${thang}${ngayBanHienThi ? ' · Bán ngày ' + fmtNgayHienThi(ngayBanHienThi) : ''}`, color: '#F3DFC5', size: 'sm', margin: 'sm', wrap: true },
          { type: 'text', text: 'Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị\nThạch Phạm Hoàng Anh -  197042', color: '#F3DFC5', size: 'xxs', margin: 'sm', wrap: true },
        ],
      },
      body: { type: 'box', layout: 'vertical', paddingAll: '16px', contents: bodyContents },
    },
  };
}

async function generateGiaVonReport() {
  const sheets = getSheetsClient();
  const [rows, rowsFresh] = await Promise.all([
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_GIAVON),
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_FRESH),
  ]);

  // --- đọc tab GIAVON (giá vốn hôm nay + lũy kế) ---
  const header = rows[0];
  const colThang = timCotTheoTen(header, 'Tháng');
  const colMaST = timCotTheoTen(header, 'Mã siêu thị');
  const colTenST = timCotTheoTen(header, 'Tên siêu thị');
  const colNganh = timCotTheoTen(header, 'Ngành hàng');
  const colGiaVonHomNay = timCotTheoTen(header, 'Giá vốn cơ bản hôm nay');
  const colGiaVonLuyKe = timCotTheoTen(header, 'Giá vốn cơ bản lũy kế đến ngày hôm qua');
  const colTiLe = timCotTheoTen(header, 'Tỉ lệ doanh thu trên giá vốn cơ bản');
  const colLNLuyKe = timCotTheoTen(header, 'LN lũy kế');
  const colChenhLech = timCotTheoTen(header, 'Chênh lệch LN so với 3 tháng trước');

  let thangMoiNhat = null;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row[colThang] === undefined || row[colThang] === '') continue;
    const t = row[colThang].toString().trim();
    if (thangMoiNhat === null || t > thangMoiNhat) thangMoiNhat = t;
  }

  const theoSieuThi = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row[colThang] === undefined || row[colThang] === '') continue;
    if (row[colThang].toString().trim() !== thangMoiNhat) continue;

    const ma = chuanHoaMaSieuThi(row[colMaST]);
    if (!theoSieuThi[ma]) {
      theoSieuThi[ma] = { ma, ten: row[colTenST] || ma, dsNganhHang: [] };
    }
    theoSieuThi[ma].dsNganhHang.push({
      ten: (row[colNganh] || '').toString().trim(),
      giaVonHomNay: Number(row[colGiaVonHomNay]) || 0,
      giaVonLuyKe: Number(row[colGiaVonLuyKe]) || 0,
      tiLeDTGiaVon: chuyenTiLeThanhPhanTram(row[colTiLe]),
      lnLuyKe: Number(row[colLNLuyKe]) || 0,
      chenhLech: Number(row[colChenhLech]) || 0,
    });
  }

  // --- đọc tab FRESH_NHAPXUAT để lấy DT/SL BÁN THỰC TẾ hôm nay theo ngành hàng ---
  const headerFresh = rowsFresh[0];
  const colNgayFresh = timCotTheoTen(headerFresh, 'Ngày');
  const colMaSTFresh = timCotTheoTen(headerFresh, 'Mã siêu thị');
  const colNganhFresh = timCotTheoTen(headerFresh, 'Ngành hàng - Phân tích');
  const colDoanhThuFresh = timCotTheoTen(headerFresh, 'Thành tiền phải thu khách hàng (chưa VAT)');
  const colSLFresh = timCotTheoTen(headerFresh, 'SL thực xuất');

  const ngayMoiNhatFresh = timNgayMoiNhat(rowsFresh, colNgayFresh);
  const banHomNayTheoSieuThi = {};
  for (let i = 1; i < rowsFresh.length; i++) {
    const row = rowsFresh[i];
    if (!row || row[colNgayFresh] === undefined || row[colNgayFresh] === '') continue;
    if (toDateKey(row[colNgayFresh]) !== ngayMoiNhatFresh) continue;

    const ma = chuanHoaMaSieuThi(row[colMaSTFresh]);
    const ten = (row[colNganhFresh] || '').toString().trim();
    if (!ten) continue;
    if (!banHomNayTheoSieuThi[ma]) banHomNayTheoSieuThi[ma] = {};
    if (!banHomNayTheoSieuThi[ma][ten]) banHomNayTheoSieuThi[ma][ten] = { dt: 0, sl: 0 };
    banHomNayTheoSieuThi[ma][ten].dt += Number(row[colDoanhThuFresh]) || 0;
    banHomNayTheoSieuThi[ma][ten].sl += Number(row[colSLFresh]) || 0;
  }

  // --- ghép: mỗi ngành hàng bên GIAVON cộng DT/SL bán hôm nay từ các ngành hàng con bên FRESH ---
  const cards = Object.values(theoSieuThi).map((st) => {
    const banHomNay = banHomNayTheoSieuThi[st.ma] || {};
    const dsNganhHangDayDu = st.dsNganhHang.map((n) => {
      const nhomFresh = GIAVON_TO_FRESH_MAP[n.ten] || [];
      let dtBanHomNay = 0;
      let slBanHomNay = 0;
      nhomFresh.forEach((tenFresh) => {
        const d = banHomNay[tenFresh];
        if (d) {
          dtBanHomNay += d.dt;
          slBanHomNay += d.sl;
        }
      });
      return { ...n, dtBanHomNay, slBanHomNay };
    });
    return taoCardGiaVon(st.ma, st.ten, thangMoiNhat, dsNganhHangDayDu, ngayMoiNhatFresh);
  });

  return cards.slice(0, 5);
}

// ---------------------------------------------------------------------------
// BC LŨY KẾ DOANH THU
// ---------------------------------------------------------------------------
const GOOGLE_SHEET_TAB_LUYKE_DT = process.env.GOOGLE_SHEET_TAB_LUYKE_DT || 'LUYKE_DT';

// ---------------------------------------------------------------------------
// MMKK HUỶ — Tổng SL bán / Doanh thu / SL giảm giá / Tiền giảm giá / SL mất mát KK — THEO NGÀNH HÀNG
// ---------------------------------------------------------------------------
const GOOGLE_SHEET_TAB_HUYMMKK = process.env.GOOGLE_SHEET_TAB_HUYMMKK || 'HUYMMKK';
const GOOGLE_SHEET_TAB_CHITIETXUAT = process.env.GOOGLE_SHEET_TAB_CHITIETXUAT || 'CHITIETXUAT';
const GOOGLE_SHEET_TAB_LUOTBILL = process.env.GOOGLE_SHEET_TAB_LUOTBILL || 'LUOTBILL';

function taoCardHuyMmkk(maSieuThi, tenSieuThi, dsNganhHang, ngayHienThi) {
  const tongDoanhThu = dsNganhHang.reduce((s, n) => s + n.doanhThu, 0);
  const tongMMKK = dsNganhHang.reduce((s, n) => s + n.slMMKK, 0);
  const tongSLBan = dsNganhHang.reduce((s, n) => s + n.slBan, 0);
  const tongTienGiamGia = dsNganhHang.reduce((s, n) => s + n.tienGiamGia, 0);

  const bodyContents = [
    { type: 'text', text: '🏢 ' + tenSieuThi, weight: 'bold', size: 'md', wrap: true, color: '#1a1a1a' },
    {
      type: 'box', layout: 'vertical', backgroundColor: '#FBEAEA', cornerRadius: 'md', paddingAll: '14px', margin: 'md',
      contents: [
        { type: 'text', text: 'TỔNG SL MẤT MÁT KIỂM KÊ', size: 'xs', color: '#888888' },
        { type: 'text', text: fmtSo(tongMMKK) + (tongMMKK < 0 ? ' (thừa)' : tongMMKK > 0 ? ' (thiếu)' : ''), size: 'xxl', weight: 'bold', color: '#C0392B', margin: 'sm' },
      ],
    },
    {
      type: 'box', layout: 'horizontal', spacing: 'sm', margin: 'md',
      contents: [
        oThongKe('💰', 'Doanh thu', fmtSo(tongDoanhThu) + ' đ'),
        oThongKe('🛒', 'SL bán', fmtSo(tongSLBan)),
      ],
    },
    {
      type: 'box', layout: 'horizontal', spacing: 'sm', margin: 'sm',
      contents: [
        oThongKe('🏷️', 'Tiền giảm giá', fmtSo(tongTienGiamGia) + ' đ'),
        oThongKe('🗑️', 'SL MMKK', fmtSo(tongMMKK)),
      ],
    },
    { type: 'separator', margin: 'lg' },
    { type: 'text', text: '📦 CHI TIẾT THEO NGÀNH HÀNG', size: 'sm', weight: 'bold', color: '#333333', margin: 'lg' },
  ];

  dsNganhHang
    .slice()
    .sort((a, b) => b.slMMKK - a.slMMKK)
    .slice(0, 15)
    .forEach((n) => {
      bodyContents.push({
        type: 'box', layout: 'vertical', margin: 'md', paddingAll: '10px',
        backgroundColor: '#FAF7F2', cornerRadius: 'md',
        contents: [
          { type: 'text', text: n.ten, size: 'sm', weight: 'bold', color: '#1a1a1a', wrap: true },
          {
            type: 'box', layout: 'horizontal', margin: 'xs', contents: [
              { type: 'text', text: 'SL bán', size: 'xxs', color: '#888888', flex: 3 },
              { type: 'text', text: fmtSo(n.slBan), size: 'xs', flex: 4, align: 'end', weight: 'bold' },
            ],
          },
          {
            type: 'box', layout: 'horizontal', margin: 'xs', contents: [
              { type: 'text', text: 'Doanh thu', size: 'xxs', color: '#888888', flex: 3 },
              { type: 'text', text: fmtSo(n.doanhThu) + ' đ', size: 'xs', flex: 4, align: 'end', weight: 'bold' },
            ],
          },
          {
            type: 'box', layout: 'horizontal', margin: 'xs', contents: [
              { type: 'text', text: 'SL giảm giá', size: 'xxs', color: '#888888', flex: 3 },
              { type: 'text', text: fmtSo(n.slGiamGia), size: 'xs', flex: 4, align: 'end', weight: 'bold' },
            ],
          },
          {
            type: 'box', layout: 'horizontal', margin: 'xs', contents: [
              { type: 'text', text: 'Tiền giảm giá', size: 'xxs', color: '#888888', flex: 3 },
              { type: 'text', text: fmtSo(n.tienGiamGia) + ' đ', size: 'xs', flex: 4, align: 'end', weight: 'bold' },
            ],
          },
          {
            type: 'box', layout: 'horizontal', margin: 'xs', contents: [
              { type: 'text', text: 'SL mất mát kiểm kê', size: 'xxs', color: '#C0392B', flex: 3 },
              { type: 'text', text: fmtSo(n.slMMKK) + (n.slMMKK < 0 ? ' (thừa)' : n.slMMKK > 0 ? ' (thiếu)' : ''), size: 'xs', flex: 4, align: 'end', weight: 'bold', color: '#C0392B' },
            ],
          },
        ],
      });
    });

  return {
    type: 'flex',
    altText: 'MMKK Huỷ ' + tenSieuThi + ' (' + ngayHienThi + '): SL mất mát kiểm kê ' + fmtSo(tongMMKK),
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#922B21', paddingAll: '20px',
        contents: [
          { type: 'text', text: '🗑️ BÁO CÁO MMKK HUỶ', color: '#FFFFFF', weight: 'bold', size: 'lg' },
          { type: 'text', text: maSieuThi + ' · ' + ngayHienThi, color: '#F5D5D0', size: 'sm', margin: 'sm' },
          { type: 'text', text: 'Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị\nThạch Phạm Hoàng Anh -  197042', color: '#F5D5D0', size: 'xxs', margin: 'sm', wrap: true },
        ],
      },
      body: { type: 'box', layout: 'vertical', paddingAll: '16px', contents: bodyContents },
    },
  };
}

async function generateHuyMmkkReport() {
  const sheets = getSheetsClient();
  const rows = await docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_HUYMMKK);

  const header = rows[0];
  const colNgay = timCotTheoTen(header, 'Ngày');
  const colMa = timCotTheoTen(header, 'Mã siêu thị');
  const colTen = timCotTheoTen(header, 'Tên siêu thị');
  const colNganh = timCotTheoTen(header, 'Ngành hàng');
  const colSLBan = timCotTheoTen(header, 'Tổng SL bán');
  const colDoanhThu = timCotTheoTen(header, 'Doanh thu');
  const colSLGiamGia = timCotTheoTen(header, 'SL giảm giá');
  const colTienGiamGia = timCotTheoTen(header, 'Tiền giảm giá');
  const colSLMMKK = timCotTheoTen(header, 'SL mất mát kiểm kê');

  const theoSieuThi = {};
  let ngayDauChung = null;
  let ngayCuoiChung = null;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row[colNgay] === undefined || row[colNgay] === '') continue;
    const dateKey = toDateKey(row[colNgay]);
    if (!dateKey || dateKey.length !== 10) continue;
    if (ngayDauChung === null || dateKey < ngayDauChung) ngayDauChung = dateKey;
    if (ngayCuoiChung === null || dateKey > ngayCuoiChung) ngayCuoiChung = dateKey;

    const ma = chuanHoaMaSieuThi(row[colMa]);
    if (!theoSieuThi[ma]) theoSieuThi[ma] = { ma, ten: row[colTen] || ma, byNganh: {} };
    const ten = (row[colNganh] || '').toString().trim();
    if (!ten) continue;
    if (!theoSieuThi[ma].byNganh[ten]) {
      theoSieuThi[ma].byNganh[ten] = { slBan: 0, doanhThu: 0, slGiamGia: 0, tienGiamGia: 0, slMMKK: 0 };
    }
    const d = theoSieuThi[ma].byNganh[ten];
    d.slBan += Number(row[colSLBan]) || 0;
    d.doanhThu += Number(row[colDoanhThu]) || 0;
    d.slGiamGia += Number(row[colSLGiamGia]) || 0;
    d.tienGiamGia += Number(row[colTienGiamGia]) || 0;
    d.slMMKK += Number(row[colSLMMKK]) || 0;
  }

  if (!ngayDauChung) {
    throw new Error('Tab "' + GOOGLE_SHEET_TAB_HUYMMKK + '" chưa có dữ liệu ngày hợp lệ');
  }

  const ngayHienThi = 'Lũy kế ' + fmtNgayNgan(ngayDauChung) + '-' + fmtNgayNgan(ngayCuoiChung);
  const cards = Object.values(theoSieuThi).map((st) => {
    const dsNganhHang = Object.entries(st.byNganh).map(([ten, d]) => ({ ten, ...d }));
    return taoCardHuyMmkk(st.ma, st.ten, dsNganhHang, ngayHienThi);
  });

  return cards.slice(0, 5);
}

function fmtNgayNgan(dateKey) {
  const [, m, d] = dateKey.split('-');
  return `${d}/${m}`;
}

function fmtPctCoDau(p) {
  const r = Math.round(p * 10) / 10;
  const mui = r >= 0 ? '▲' : '▼';
  const dau = r >= 0 ? '+' : '';
  return `${mui}${dau}${r.toFixed(1)}%`;
}

function mauTangGiamLuyKe(p) {
  return p >= 0 ? '#27AE60' : '#E74C3C';
}

function oThongKeLuyKe(nhan, giaTri, mau) {
  return {
    type: 'box', layout: 'vertical', flex: 1, backgroundColor: '#F7FAF8', cornerRadius: 'md',
    paddingAll: '10px', spacing: 'xs',
    contents: [
      { type: 'text', text: nhan, size: 'xxs', color: '#888888' },
      { type: 'text', text: giaTri, size: 'sm', weight: 'bold', color: mau || '#1a1a1a', wrap: true },
    ],
  };
}

function dongSoSanhXuHuong(nhan, giaTriHienTai, phanTram) {
  const mau = mauTangGiamLuyKe(phanTram);
  return {
    type: 'box', layout: 'horizontal', margin: 'sm', contents: [
      { type: 'text', text: nhan, size: 'sm', flex: 4, color: '#555555' },
      { type: 'text', text: giaTriHienTai, size: 'sm', flex: 4, align: 'end', weight: 'bold', color: '#1a1a1a' },
      { type: 'text', text: fmtPctCoDau(phanTram), size: 'xs', flex: 3, align: 'end', weight: 'bold', color: mau },
    ],
  };
}

function taoCardLuyKeDT(maSieuThi, tenSieuThi, so) {
  const mom = fmtPctCoDau(so.mom);
  const mauMom = mauTangGiamLuyKe(so.mom);

  const bodyContents = [
    { type: 'text', text: `🏢 ${tenSieuThi}`, weight: 'bold', size: 'md', wrap: true, color: '#1a1a1a' },
    {
      type: 'box', layout: 'vertical', backgroundColor: '#F0F7F2', cornerRadius: 'md', paddingAll: '14px', margin: 'md',
      contents: [
        { type: 'text', text: `LŨY KẾ 01-${fmtNgayNgan(so.ngayCuoi)}`, size: 'xs', color: '#888888' },
        { type: 'text', text: fmtSo(so.tongLuyKe) + ' đ', size: 'xxl', weight: 'bold', color: '#1a1a1a', margin: 'sm' },
        { type: 'text', text: `Trung bình ${fmtSo(so.tongLuyKe / (so.soNgayCoData || 1))} đ/ngày`, size: 'xs', color: '#555555', margin: 'sm' },
      ],
    },
    {
      type: 'box', layout: 'horizontal', spacing: 'sm', margin: 'md',
      contents: [
        oThongKeLuyKe('DỰ KIẾN HẾT THÁNG', fmtSo(so.duKienHetThang) + ' đ'),
        oThongKeLuyKe('THỰC TẾ THÁNG TRƯỚC', fmtSo(so.thangTruoc) + ' đ'),
      ],
    },
    {
      type: 'box', layout: 'horizontal', spacing: 'sm', margin: 'sm',
      contents: [
        oThongKeLuyKe('MoM', mom, mauMom),
        oThongKeLuyKe('CHÊNH LỆCH', `${so.mom >= 0 ? '+' : ''}${fmtSo(so.duKienHetThang - so.thangTruoc)} đ`, mauMom),
      ],
    },
    { type: 'separator', margin: 'lg' },
    { type: 'text', text: '📊 XU HƯỚNG SO VỚI THÁNG TRƯỚC', size: 'sm', weight: 'bold', color: '#333333', margin: 'lg' },
    {
      type: 'box', layout: 'horizontal', margin: 'md', contents: [
        { type: 'text', text: ' ', size: 'xs', flex: 4 },
        { type: 'text', text: 'Dự kiến hết tháng', size: 'xs', flex: 4, align: 'end', color: '#888888' },
        { type: 'text', text: 'MoM', size: 'xs', flex: 3, align: 'end', color: '#888888' },
      ],
    },
    { type: 'separator', margin: 'sm' },
    dongSoSanhXuHuong('💰 Doanh thu', fmtSo(so.duKienHetThang) + ' đ', so.mom),
    dongSoSanhXuHuong('🧾 Lượt bill', fmtSo(so.billDuKienHetThang), so.billMom),
    dongSoSanhXuHuong('💳 Giá trị bill', fmtSo(so.giaTriBillTB) + ' đ', so.giaTriBillMom),
    { type: 'separator', margin: 'lg' },
    dongThongTinNgay('🧾', 'Tổng số bill (lũy kế)', fmtSo(so.tongBill), false),
    dongThongTinNgay('💳', 'Giá trị bill TB (lũy kế)', fmtSo(so.giaTriBillTB) + ' đ', false),
    dongThongTinNgay('📅', 'Số ngày có dữ liệu', `${so.soNgayCoData} ngày`, false),
  ];

  return {
    type: 'flex',
    altText: `Lũy kế DT ${tenSieuThi}: ${fmtSo(so.tongLuyKe)} đ, dự kiến hết tháng ${fmtSo(so.duKienHetThang)} đ`,
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#1B4F72', paddingAll: '20px',
        contents: [
          { type: 'text', text: '📈 BC LŨY KẾ DOANH THU', color: '#FFFFFF', weight: 'bold', size: 'lg' },
          { type: 'text', text: `${maSieuThi} · Lũy kế 01-${fmtNgayNgan(so.ngayCuoi)}`, color: '#D6E4F0', size: 'sm', margin: 'sm' },
          { type: 'text', text: 'Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị\nThạch Phạm Hoàng Anh -  197042', color: '#D6E4F0', size: 'xxs', margin: 'sm', wrap: true },
        ],
      },
      body: { type: 'box', layout: 'vertical', paddingAll: '16px', contents: bodyContents },
    },
  };
}

async function generateLuyKeReport() {
  const sheets = getSheetsClient();
  const rows = await docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_LUYKE_DT);

  const header = rows[0];
  const colNgay = timCotTheoTen(header, 'Ngày');
  const colMa = timCotTheoTen(header, 'Mã siêu thị');
  const colTen = timCotTheoTen(header, 'Tên siêu thị');
  const colDTOffline = timCotTheoTen(header, 'Doanh thu offline');
  const colDTOnline = timCotTheoTen(header, 'Doanh thu Online');
  const colSoBill = timCotTheoTen(header, 'Tổng số bill');

  const theoSieuThi = {};
  let ngayMoiNhatChung = null;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row[colNgay] === undefined || row[colNgay] === '') continue;
    const dateKey = toDateKey(row[colNgay]);
    if (!dateKey || dateKey.length !== 10) continue;
    if (ngayMoiNhatChung === null || dateKey > ngayMoiNhatChung) ngayMoiNhatChung = dateKey;

    const ma = chuanHoaMaSieuThi(row[colMa]);
    const ten = row[colTen] || ma;
    const ym = dateKey.slice(0, 7);
    const dt = (Number(row[colDTOffline]) || 0) + (Number(row[colDTOnline]) || 0);
    const bill = Number(row[colSoBill]) || 0;

    if (!theoSieuThi[ma]) theoSieuThi[ma] = { ten, byThang: {} };
    if (theoSieuThi[ma].ten === ma && ten !== ma) theoSieuThi[ma].ten = ten;
    if (!theoSieuThi[ma].byThang[ym]) theoSieuThi[ma].byThang[ym] = { tongDT: 0, tongBill: 0, ngay: new Set() };
    const o = theoSieuThi[ma].byThang[ym];
    o.tongDT += dt;
    o.tongBill += bill;
    o.ngay.add(dateKey);
  }

  if (!ngayMoiNhatChung) {
    throw new Error(`Tab "${GOOGLE_SHEET_TAB_LUYKE_DT}" chưa có dữ liệu ngày hợp lệ`);
  }

  const [namHienTai, thangHienTaiSo] = ngayMoiNhatChung.slice(0, 7).split('-').map(Number);
  const ymHienTai = ngayMoiNhatChung.slice(0, 7);
  const ymThangTruoc = thangHienTaiSo === 1
    ? `${namHienTai - 1}-12`
    : `${namHienTai}-${String(thangHienTaiSo - 1).padStart(2, '0')}`;
  const soNgayTrongThangHienTai = new Date(namHienTai, thangHienTaiSo, 0).getDate();

  const cards = Object.entries(theoSieuThi).map(([ma, data]) => {
    const cur = data.byThang[ymHienTai] || { tongDT: 0, tongBill: 0, ngay: new Set() };
    const prev = data.byThang[ymThangTruoc] || { tongDT: 0, tongBill: 0, ngay: new Set() };
    const soNgayCoData = cur.ngay.size;
    const duKienHetThang = soNgayCoData > 0 ? (cur.tongDT / soNgayCoData) * soNgayTrongThangHienTai : 0;
    const mom = prev.tongDT > 0 ? ((duKienHetThang - prev.tongDT) / prev.tongDT) * 100 : 0;
    const giaTriBillTB = cur.tongBill > 0 ? cur.tongDT / cur.tongBill : 0;

    const billDuKienHetThang = soNgayCoData > 0 ? (cur.tongBill / soNgayCoData) * soNgayTrongThangHienTai : 0;
    const billMom = prev.tongBill > 0 ? ((billDuKienHetThang - prev.tongBill) / prev.tongBill) * 100 : 0;

    const giaTriBillThangTruoc = prev.tongBill > 0 ? prev.tongDT / prev.tongBill : 0;
    const giaTriBillMom = giaTriBillThangTruoc > 0 ? ((giaTriBillTB - giaTriBillThangTruoc) / giaTriBillThangTruoc) * 100 : 0;

    let ngayCuoiThangHienTai = ngayMoiNhatChung;
    const ngayThangHienTaiSorted = Array.from(cur.ngay).sort();
    if (ngayThangHienTaiSorted.length > 0) ngayCuoiThangHienTai = ngayThangHienTaiSorted[ngayThangHienTaiSorted.length - 1];

    return taoCardLuyKeDT(ma, data.ten, {
      tongLuyKe: cur.tongDT,
      tongBill: cur.tongBill,
      soNgayCoData,
      duKienHetThang,
      thangTruoc: prev.tongDT,
      mom,
      giaTriBillTB,
      billDuKienHetThang,
      billMom,
      giaTriBillMom,
      ngayCuoi: ngayCuoiThangHienTai,
    });
  });

  return cards.slice(0, 5);
}

// ---------------------------------------------------------------------------
// BÁO CÁO BÁNH TRUNG THU
// ---------------------------------------------------------------------------
const GOOGLE_SHEET_TAB_BANHTT_TON = process.env.GOOGLE_SHEET_TAB_BANHTT_TON || 'BANHTT_TON';
const GOOGLE_SHEET_TAB_BANHTT_DOANHTHU = process.env.GOOGLE_SHEET_TAB_BANHTT_DOANHTHU || 'BANHTT_DOANHTHU';

// Danh sách 89 mã SKU Bánh Trung Thu ĐÃ ĐƯỢC DUYỆT tính thưởng
// (theo file "DANH SÁCH SKUS TRUNG THU 20082026.xlsx" sếp gửi ngày 20/08/2026).
// Khớp theo "Mã Model" (10 số) — mã KHÔNG có trong danh sách này thì KHÔNG tính thưởng.
const SKU_TRUNGTHU_MAP = {
  '2607001852': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2407001584': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2407001585': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2407001417': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2607001877': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2607001879': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2407001423': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '1708000229': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000230': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000237': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000239': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001872': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2508000902': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2407001413': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2607001865': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2607001866': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2607001853': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2508000908': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2508000905': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2607001873': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001874': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001875': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001876': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2508000790': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2508000904': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2508000792': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2508000789': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2607001855': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001854': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001856': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001857': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000254': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000259': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000264': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000252': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000253': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000258': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000262': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2508000903': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2508000906': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2508000907': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2607001878': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2607001858': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001859': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2507002809': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2507002810': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2507002811': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2508000460': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '1708000276': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '1708000277': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2407001416': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2507002610': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2607001867': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001868': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001869': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2507002609': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2607001864': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001863': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001870': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001871': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2507002607': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2507002606': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2507002605': { loai: 'cai', tienThuong: 1000, nhom: 'banhtuoi' },
  '2607001860': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001862': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607001861': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2405000267': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2405000268': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2508002911': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2512000349': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2607000922': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2607000923': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2407002830': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2407002831': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2508000666': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2508001049': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '1708000225': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2607006125': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607006126': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2607006124': { loai: 'cai', tienThuong: 1000, nhom: 'btt' },
  '2407001422': { loai: 'hop', tienThuong: 4000, nhom: 'banhtuoi' },
  '2608000290': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2608000292': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2608000291': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2608000289': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2608000288': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2608001890': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2608002585': { loai: 'hop', tienThuong: 4000, nhom: 'btt' },
  '2608001064': { loai: 'hop', tienThuong: 4000, nhom: 'tra' },
};

// Nhận diện 1 dòng là "Hộp" — CHỈ dùng làm fallback cho tồn kho khi mã KHÔNG
// nằm trong SKU_TRUNGTHU_MAP ở trên (để vẫn hiển thị tồn kho đầy đủ).
function laHangHop(donVi, tenModel) {
  const dv = (donVi || '').toString().trim();
  if (dv === 'Hộp' || dv === 'Bộ' || dv === 'Giỏ') return true;
  const ten = (tenModel || '').toString().toUpperCase();
  return ten.includes('HỘP') || ten.includes('HOP ') || ten.startsWith('HOP');
}

function docTonBanhTT(rows) {
  const header = rows[0];
  const colTenST = timCotTheoTen(header, 'Tên siêu thị');
  const colTon = timCotTheoTen(header, 'Tồn kho siêu thị');
  const colMaModel = timCotTheoTen(header, 'Mã Model');

  const ton = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const st = row[colTenST];
    if (!st) continue;
    const maModel = (row[colMaModel] || '').toString().trim();
    const skuInfo = SKU_TRUNGTHU_MAP[maModel];
    if (!skuInfo) continue; // chỉ tính đúng 89 mã đã duyệt

    const soLuong = Number(row[colTon]) || 0;
    ton[st] = (ton[st] || 0) + soLuong;
  }
  return ton;
}

function docBanBanhTT(rows) {
  const header = rows[0];
  const colTenST = timCotTheoTen(header, 'Tên siêu thị');
  const colSLOnline = timCotTheoTen(header, 'Số lượng Online');
  const colSLOffline = timCotTheoTen(header, 'Số lượng Offline');
  const colMaModel = timCotTheoTen(header, 'Mã Model');
  const colSLKM = header.indexOf('SL xuất km');

  const ban = {};

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const st = row[colTenST];
    if (!st) continue;
    const maModel = (row[colMaModel] || '').toString().trim();
    const slKM = colSLKM === -1 ? 0 : (Number(row[colSLKM]) || 0);
    const soLuongBanRa = (Number(row[colSLOnline]) || 0) + (Number(row[colSLOffline]) || 0);
    const soLuongTinhThuong = Math.max(0, soLuongBanRa - slKM);

    if (!ban[st]) ban[st] = { bttCai: 0, bttHop: 0, banhtuoi: 0, tra: 0, thuong: 0 };

    const skuInfo = SKU_TRUNGTHU_MAP[maModel];
    if (skuInfo) {
      if (skuInfo.nhom === 'btt') {
        if (skuInfo.loai === 'hop') ban[st].bttHop += soLuongTinhThuong;
        else ban[st].bttCai += soLuongTinhThuong;
      } else if (skuInfo.nhom === 'banhtuoi') {
        ban[st].banhtuoi += soLuongTinhThuong;
      } else if (skuInfo.nhom === 'tra') {
        ban[st].tra += soLuongTinhThuong;
      }
      ban[st].thuong += soLuongTinhThuong * skuInfo.tienThuong;
    }
  }
  return { ban };
}

function dongBangBanhTT(label, tonTong, bttCai, bttHop, banhtuoi, tra, thuong, dam) {
  return {
    type: 'box', layout: 'horizontal', margin: dam ? 'none' : 'sm',
    contents: [
      { type: 'text', text: label, size: 'xxs', flex: 5, wrap: false, weight: dam ? 'bold' : 'regular', color: dam ? '#1a1a1a' : '#333333' },
      { type: 'text', text: tonTong, size: 'xxs', flex: 2, align: 'end', weight: dam ? 'bold' : 'regular' },
      { type: 'text', text: bttCai, size: 'xxs', flex: 2, align: 'end', weight: dam ? 'bold' : 'regular' },
      { type: 'text', text: bttHop, size: 'xxs', flex: 2, align: 'end', weight: dam ? 'bold' : 'regular' },
      { type: 'text', text: banhtuoi, size: 'xxs', flex: 2, align: 'end', weight: dam ? 'bold' : 'regular' },
      { type: 'text', text: tra, size: 'xxs', flex: 2, align: 'end', weight: dam ? 'bold' : 'regular' },
      { type: 'text', text: thuong, size: 'xxs', flex: 4, align: 'end', weight: 'bold', color: dam ? '#B8860B' : '#D97706' },
    ],
  };
}

function taoFlexBanhTrungThu(ton, ban) {
  const tatCaSieuThi = new Set([...Object.keys(ton), ...Object.keys(ban)]);
  const rongBan = () => ({ bttCai: 0, bttHop: 0, banhtuoi: 0, tra: 0, thuong: 0 });
  const rows = [];

  for (const st of tatCaSieuThi) {
    const t = ton[st] || 0;
    const b = ban[st] || rongBan();
    rows.push({ ten: tenNganSieuThi(st), ton: t, ban: b, thuong: b.thuong });
  }
  rows.sort((a, b) => b.thuong - a.thuong);

  const tong = rows.reduce((acc, r) => ({
    ton: acc.ton + r.ton,
    ban: {
      bttCai: acc.ban.bttCai + r.ban.bttCai, bttHop: acc.ban.bttHop + r.ban.bttHop,
      banhtuoi: acc.ban.banhtuoi + r.ban.banhtuoi, tra: acc.ban.tra + r.ban.tra,
    },
    thuong: acc.thuong + r.thuong,
  }), { ton: 0, ban: rongBan(), thuong: 0 });

  const now = new Date();
  const thoiGian = now.toLocaleString('vi-VN', {
    hour: '2-digit', minute: '2-digit', day: '2-digit', month: '2-digit', year: 'numeric',
    timeZone: 'Asia/Ho_Chi_Minh',
  });

  // ĐỔI SANG TEXT THUẦN (bỏ hẳn Flex Message) để loại trừ mọi rủi ro về giới hạn/cấu trúc
  // JSON của LINE — đảm bảo chắc chắn gửi được dù có bao nhiêu siêu thị đi nữa.
  const pad = (s, len) => { s = String(s); return s.length >= len ? s.slice(0, len) : s + ' '.repeat(len - s.length); };
  const padNum = (s, len) => { s = String(s); return s.length >= len ? s.slice(0, len) : ' '.repeat(len - s.length) + s; };
  const dong = (ten, tn, c, h, bt, tr, th) =>
    pad(ten, 13) + padNum(tn, 6) + padNum(c, 4) + padNum(h, 4) + padNum(bt, 5) + padNum(tr, 4) + padNum(th, 9);

  let text = '🥮 BÁO CÁO BÁNH TRUNG THU\n';
  text += 'Thưởng: Cái 1.000đ / Hộp 4.000đ, đã trừ hàng xuất KM\n';
  text += `Cập nhật ${thoiGian} · ${rows.length} siêu thị\n`;
  text += '━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  text += 'SiêuThị       Tồn  C   H  BT  Trà   Thưởng\n';
  text += dong('TỔNG TẤT CẢ', fmtSo(tong.ton), fmtSo(tong.ban.bttCai), fmtSo(tong.ban.bttHop), fmtSo(tong.ban.banhtuoi), fmtSo(tong.ban.tra), fmtSo(tong.thuong) + 'đ') + '\n';
  text += '━━━━━━━━━━━━━━━━━━━━━━━━━━\n';

  rows.forEach((r) => {
    text += dong(rutGonTen(r.ten, 13), fmtSo(r.ton), fmtSo(r.ban.bttCai), fmtSo(r.ban.bttHop), fmtSo(r.ban.banhtuoi), fmtSo(r.ban.tra), fmtSo(r.thuong) + 'đ') + '\n';
  });

  if (text.length > 4900) {
    text = text.slice(0, 4880) + '\n... (còn tiếp, xem chi tiết trong Google Sheet)';
  }

  return { type: 'text', text };
}

async function generateBanhTrungThuReport() {
  const sheets = getSheetsClient();

  const [rowsTon, rowsBan] = await Promise.all([
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_BANHTT_TON),
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_BANHTT_DOANHTHU),
  ]);

  const ton = docTonBanhTT(rowsTon);
  const { ban } = docBanBanhTT(rowsBan);

  console.log('[DEBUG BanhTT] rowsTon.length=' + rowsTon.length + ', rowsBan.length=' + rowsBan.length);
  console.log('[DEBUG BanhTT] header rowsBan[0]=' + JSON.stringify(rowsBan[0]));
  console.log('[DEBUG BanhTT] sample rowsBan[1]=' + JSON.stringify(rowsBan[1]));
  console.log('[DEBUG BanhTT] Object.keys(ton).length=' + Object.keys(ton).length);
  console.log('[DEBUG BanhTT] Object.keys(ban).length=' + Object.keys(ban).length);
  console.log('[DEBUG BanhTT] SKU_TRUNGTHU_MAP size=' + Object.keys(SKU_TRUNGTHU_MAP).length);

  return taoFlexBanhTrungThu(ton, ban);
}

// ---------------------------------------------------------------------------
// NẠP FILE NGƯỜI DÙNG GỬI TRỰC TIẾP VÀO GROUP (.xlsx/.xls) — GHI ĐÈ VÀO SHEET
// ---------------------------------------------------------------------------
async function taiNoiDungFileLine(messageId) {
  const token = (process.env.LINE_CHANNEL_ACCESS_TOKEN || '').replace(/\s+/g, '');
  const res = await fetch(`https://api-data.line.me/v2/bot/message/${messageId}/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Tai file LINE that bai: ${res.status} ${body}`);
  }
  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

function nhomHangPhoBien(header, dataRows) {
  const idx = header.indexOf('Nhóm hàng');
  if (idx === -1) return null;
  const dem = {};
  for (const row of dataRows) {
    const v = (row[idx] || '').toString().trim();
    if (!v) continue;
    dem[v] = (dem[v] || 0) + 1;
  }
  const entries = Object.entries(dem).sort((a, b) => b[1] - a[1]);
  return entries.length ? entries[0][0] : null;
}

function soNgayPhanBiet(header, dataRows) {
  const idx = header.indexOf('Ngày');
  if (idx === -1) return 0;
  const s = new Set();
  for (const row of dataRows) {
    const v = row[idx];
    if (v === undefined || v === null || v === '') continue;
    s.add(toDateKey(v));
  }
  return s.size;
}

function nhanDangLoaiFile(header, dataRows) {
  const co = (ten) => header.includes(ten);

  // Nhận diện qua CỘT ĐẶC TRƯNG riêng của từng file (không phụ thuộc nhóm hàng phổ biến
  // nữa, vì giờ anh có thể xuất file KHÔNG lọc theo nhóm hàng, đủ mọi ngành hàng).
  // File "BC Tồn Theo Model": có cột "Mã sản phẩm cơ sở" chỉ file này mới có.
  if (co('Mã Model') && co('Tồn kho siêu thị') && co('Mã sản phẩm cơ sở')) {
    return { loai: 'banhtt_ton', tenTab: GOOGLE_SHEET_TAB_BANHTT_TON };
  }
  // File "Doanh Thu Theo Model": có cột "SL xuất km" chỉ file này mới có.
  if (co('Mã Model') && co('Tổng số lượng') && !co('Tồn kho siêu thị') && co('SL xuất km')) {
    return { loai: 'banhtt_ban', tenTab: GOOGLE_SHEET_TAB_BANHTT_DOANHTHU };
  }
  // Đã có: file "Giá Vốn" — nhận diện qua 2 cột đặc trưng chỉ file này mới có
  if (co('Giá vốn cơ bản hôm nay') && co('DT FRESH tính giá vốn')) {
    return { loai: 'giavon', tenTab: GOOGLE_SHEET_TAB_GIAVON };
  }
  if (co('SL hủy tồn') && co('SL mất mát kiểm kê') && co('SL hủy hao hụt NCC')) {
    return { loai: 'huymmkk', tenTab: GOOGLE_SHEET_TAB_HUYMMKK };
  }
  if (co('Ngày') && co('Mã siêu thị') && co('Doanh thu offline')) {
    if (soNgayPhanBiet(header, dataRows || []) > 1) {
      return { loai: 'luyke_dt', tenTab: GOOGLE_SHEET_TAB_LUYKE_DT };
    }
    return { loai: 'doanhthu_sieuthi', tenTab: GOOGLE_SHEET_TAB_DOANHTHU_SIEUTHI };
  }
  if (co('Ngày') && co('Mã siêu thị') && co('Ngành hàng - Phân tích') && co('Thành tiền phải thu khách hàng (chưa VAT)')) {
    return { loai: 'fresh', tenTab: GOOGLE_SHEET_TAB_FRESH };
  }
  if (co('Ngày xuất') && co('Ngành hàng BHX') && co('Doanh thu')) {
    return { loai: 'doanhthu_nganhhang', tenTab: GOOGLE_SHEET_TAB_DOANHTHU_NGANHHANG };
  }
  if (co('Mã phiếu xuất') && co('Tên sản phẩm') && co('Số lượng') && co('Giá bán')) {
    return { loai: 'chitietxuat', tenTab: GOOGLE_SHEET_TAB_CHITIETXUAT };
  }
  if (co('Lượt Bill của Siêu Thị') && co('Ngày xuất') && co('Tổng tiền (VAT)')) {
    return { loai: 'luotbill', tenTab: GOOGLE_SHEET_TAB_LUOTBILL };
  }
  return null;
}

async function napFileVaoSheet(fileName, buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (allRows.length < 2) throw new Error('File trống hoặc không đọc được dữ liệu');

  const header = allRows[0].map((h) => (h || '').toString().trim());
  const dataRows = allRows.slice(1).filter((r) => r && r.some((v) => v !== null && v !== ''));
  if (dataRows.length === 0) throw new Error('File không có dòng dữ liệu nào');

  const nhanDang = nhanDangLoaiFile(header, dataRows);
  if (!nhanDang) {
    throw new Error(
      `Không nhận diện được loại báo cáo từ file "${fileName || ''}". Kiểm tra lại tiêu đề cột trong file có đúng mẫu không.`
    );
  }

  const sheets = getSheetsClient();

  if (nhanDang.loai === 'chitietxuat') {
    const idxTen = header.indexOf('Tên sản phẩm');
    const idxSL = header.indexOf('Số lượng');
    const idxGia = header.indexOf('Giá bán');

    const rows = dataRows.map((row) => [row[idxTen], row[idxSL], row[idxGia]]);

    await sheets.spreadsheets.values.clear({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: `${nhanDang.tenTab}!A2:ZZ`,
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: `${nhanDang.tenTab}!A2`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: rows },
    });

    return { loai: nhanDang.loai, tenTab: nhanDang.tenTab, soDong: rows.length };
  }

  if (nhanDang.loai === 'luotbill') {
    const idxNgay = header.indexOf('Ngày xuất');
    const idxLuotBill = header.indexOf('Lượt Bill của Siêu Thị');
    const idxTenST = header.indexOf('Tên siêu thị');

    const rowsLB = dataRows.map((row) => [
      toDateTimeKeyForBucket(row[idxNgay]),
      row[idxLuotBill],
      idxTenST === -1 ? '' : row[idxTenST],
    ]);

    await sheets.spreadsheets.values.clear({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: `${nhanDang.tenTab}!A2:ZZ`,
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: `${nhanDang.tenTab}!A2`,
      valueInputOption: 'RAW',
      requestBody: { values: rowsLB },
    });

    return { loai: nhanDang.loai, tenTab: nhanDang.tenTab, soDong: rowsLB.length };
  }

  if (nhanDang.loai === 'huymmkk') {
    const idxNgay = header.indexOf('Ngày xuất');
    const idxMa = header.indexOf('Mã siêu thị');
    const idxTen = header.indexOf('Tên siêu thị');
    const idxNganh = header.indexOf('Ngành hàng');
    const idxSLBan = header.indexOf('Tổng SL bán');
    const idxDoanhThu = header.indexOf('Doanh thu');
    const idxSLGiamGia = header.indexOf('SL s.thị bán giảm giá');
    const idxTienGiamGia = header.indexOf('Tiền s.thị bán giảm giá(chưa VAT)');
    const idxSLMMKK = header.indexOf('SL mất mát kiểm kê');

    const gop = {};
    for (const row of dataRows) {
      const ngay = row[idxNgay];
      const ma = row[idxMa];
      const ten = row[idxTen];
      const nganh = (row[idxNganh] || '').toString().trim();
      if (ngay === undefined || ngay === null || ngay === '' || !nganh) continue;
      const key = toDateKey(ngay) + '|' + ma + '|' + nganh;
      if (!gop[key]) {
        gop[key] = { ngay, ma, ten, nganh, slBan: 0, doanhThu: 0, slGiamGia: 0, tienGiamGia: 0, slMMKK: 0 };
      }
      gop[key].slBan += Number(row[idxSLBan]) || 0;
      gop[key].doanhThu += Number(row[idxDoanhThu]) || 0;
      gop[key].slGiamGia += Number(row[idxSLGiamGia]) || 0;
      gop[key].tienGiamGia += Number(row[idxTienGiamGia]) || 0;
      gop[key].slMMKK += Number(row[idxSLMMKK]) || 0;
    }

    const rowsGop = Object.values(gop).map((g) => [
      g.ngay, g.ma, g.ten, g.nganh, g.slBan, g.doanhThu, g.slGiamGia, g.tienGiamGia, g.slMMKK,
    ]);

    await sheets.spreadsheets.values.clear({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: nhanDang.tenTab + '!A2:ZZ',
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: nhanDang.tenTab + '!A2',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: rowsGop },
    });

    return { loai: nhanDang.loai, tenTab: nhanDang.tenTab, soDong: rowsGop.length };
  }

  // Nếu là file Bánh Trung Thu (tồn hoặc bán), lọc trước theo đúng 89 mã đã duyệt
  // để tránh nạp hàng chục nghìn dòng không liên quan vào Sheet, gây tràn bộ nhớ (OOM).
  let dataRowsDaLoc = dataRows;
  if (nhanDang.loai === 'banhtt_ton' || nhanDang.loai === 'banhtt_ban') {
    const idxMaModelGoc = header.indexOf('Mã Model');
    if (idxMaModelGoc !== -1) {
      const soDongGoc = dataRowsDaLoc.length;
      dataRowsDaLoc = dataRowsDaLoc.filter((row) => {
        const ma = (row[idxMaModelGoc] || '').toString().trim();
        return !!SKU_TRUNGTHU_MAP[ma];
      });
      console.log(`[napFileVaoSheet] Đã lọc file Bánh Trung Thu: ${soDongGoc} dòng gốc -> ${dataRowsDaLoc.length} dòng (đúng 89 mã đã duyệt)`);
    }
  }

  const destRows = await docTabThanhMangDong(sheets, nhanDang.tenTab);
  const destHeader = destRows[0];

  const rowsToAppend = dataRowsDaLoc.map((row) =>
    destHeader.map((tenCot) => {
      const idx = header.indexOf(tenCot);
      return idx === -1 ? '' : row[idx] ?? '';
    })
  );

  if (nhanDang.loai === 'luyke_dt') {
    const colNgayDest = timCotTheoTen(destHeader, 'Ngày');
    const colMaDest = timCotTheoTen(destHeader, 'Mã siêu thị');

    const keyMoi = new Set(
      rowsToAppend.map((r) => `${toDateKey(r[colNgayDest])}|${chuanHoaMaSieuThi(r[colMaDest])}`)
    );
    const giuLai = destRows.slice(1).filter((row) => {
      if (!row || row[colNgayDest] === undefined || row[colNgayDest] === '') return false;
      const key = `${toDateKey(row[colNgayDest])}|${chuanHoaMaSieuThi(row[colMaDest])}`;
      return !keyMoi.has(key);
    });
    const ketQua = [...giuLai, ...rowsToAppend];

    await sheets.spreadsheets.values.clear({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: `${nhanDang.tenTab}!A2:ZZ`,
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: `${nhanDang.tenTab}!A2`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: ketQua },
    });

    return { loai: nhanDang.loai, tenTab: nhanDang.tenTab, soDong: rowsToAppend.length };
  }

  await sheets.spreadsheets.values.clear({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${nhanDang.tenTab}!A2:ZZ`,
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${nhanDang.tenTab}!A2`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: rowsToAppend },
  });

  return { loai: nhanDang.loai, tenTab: nhanDang.tenTab, soDong: rowsToAppend.length };
}

// ---------------------------------------------------------------------------
// AI (CLAUDE) — đoán tab liên quan + phân tích bảng dữ liệu / ảnh
// ---------------------------------------------------------------------------

// Danh sách tab hiện có để AI chọn khi trả lời câu hỏi tự do.
// (Đã bỏ TON/DOANHTHU trà vì tính năng trà đã tắt.)
const KNOWN_TABS = [
  { name: GOOGLE_SHEET_TAB_DOANHTHU_SIEUTHI, desc: 'Doanh thu theo từng siêu thị theo ngày (offline/online, số bill)' },
  { name: GOOGLE_SHEET_TAB_DOANHTHU_NGANHHANG, desc: 'Doanh thu theo ngành hàng theo ngày' },
  { name: GOOGLE_SHEET_TAB_FRESH, desc: 'Nhập/xuất hàng Fresh (rau củ, thịt cá, trái cây...) theo siêu thị theo ngày' },
  { name: GOOGLE_SHEET_TAB_BANHTT_TON, desc: 'Tồn kho Bánh Trung Thu theo siêu thị' },
  { name: GOOGLE_SHEET_TAB_BANHTT_DOANHTHU, desc: 'Số lượng bán Bánh Trung Thu theo siêu thị' },
  { name: GOOGLE_SHEET_TAB_LUYKE_DT, desc: 'Doanh thu luỹ kế nhiều ngày/tháng để so sánh MoM' },
  { name: GOOGLE_SHEET_TAB_GIAVON, desc: 'Giá vốn, lợi nhuận lũy kế theo ngành hàng, so với TB 3 tháng trước' },
];

async function pickRelevantTab(question) {
  const tabList = KNOWN_TABS.map((t) => `- ${t.name}: ${t.desc}`).join('\n');
  const msg = await anthropic.messages.create({
    model: AI_MODEL,
    max_tokens: 300,
    system:
      'Bạn là trợ lý chọn đúng tab Google Sheet để trả lời câu hỏi của quản lý cửa hàng. ' +
      'CHỈ trả lời bằng JSON hợp lệ, không thêm chữ nào khác, không dùng markdown code block. ' +
      'Định dạng bắt buộc: {"tab": "TEN_TAB_HOAC_null", "reason": "giải thích ngắn gọn"}',
    messages: [
      {
        role: 'user',
        content: `Danh sách tab hiện có:\n${tabList}\n\nCâu hỏi của quản lý: "${question}"\n\nChọn 1 tab phù hợp nhất. Nếu không tab nào phù hợp, trả "tab": null.`,
      },
    ],
  });
  const text = msg.content.find((b) => b.type === 'text')?.text || '{}';
  try {
    return JSON.parse(text.replace(/```json|```/g, '').trim());
  } catch {
    return { tab: null };
  }
}

async function analyzeTableData(question, sourceLabel, rows, maxRows = 500) {
  const header = rows[0] || [];
  const dataRows = rows.slice(1, maxRows + 1);
  const truncatedNote =
    rows.length - 1 > maxRows
      ? `\n(Lưu ý: dữ liệu có ${rows.length - 1} dòng, chỉ gửi ${maxRows} dòng đầu để phân tích)`
      : '';
  const csvText = [header, ...dataRows]
    .map((r) => r.map((c) => (c === undefined || c === null ? '' : c)).join(','))
    .join('\n');

  const msg = await anthropic.messages.create({
    model: AI_MODEL,
    max_tokens: 1024,
    system:
      'Bạn là trợ lý phân tích dữ liệu bán lẻ cho quản lý cửa hàng Bách Hoá Xanh. ' +
      'Trả lời ngắn gọn, đi thẳng vào số liệu và nhận xét thực tế, dùng tiếng Việt, ' +
      'dùng gạch đầu dòng cho dễ đọc trên LINE (không dùng bảng markdown).',
    messages: [
      {
        role: 'user',
        content: `Nguồn dữ liệu: ${sourceLabel}${truncatedNote}\n\nDữ liệu (CSV):\n${csvText}\n\nCâu hỏi: ${question}`,
      },
    ],
  });
  return msg.content.find((b) => b.type === 'text')?.text || 'Không có phản hồi từ AI.';
}

async function analyzeImage(question, imageBuffer, mimeType) {
  const base64Image = imageBuffer.toString('base64');
  const userQuestion =
    question && question.trim().length > 0
      ? question
      : 'Đọc và phân tích nội dung trong ảnh này giúp anh (số liệu, tình trạng hàng hoá, hoặc điểm bất thường nếu có).';

  const msg = await anthropic.messages.create({
    model: AI_MODEL,
    max_tokens: 1024,
    system:
      'Bạn là trợ lý phân tích ảnh cho quản lý cửa hàng Bách Hoá Xanh. ' +
      'Ảnh có thể là: (1) ảnh chụp bảng số liệu/báo cáo — hãy đọc số liệu và phân tích; ' +
      'hoặc (2) ảnh hàng hoá/kệ hàng thực tế — hãy nhận xét tình trạng trưng bày, tồn kho, hoặc vấn đề quan sát được. ' +
      'Trả lời ngắn gọn, thực tế, tiếng Việt.',
    messages: [
      {
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mimeType, data: base64Image } },
          { type: 'text', text: userQuestion },
        ],
      },
    ],
  });
  return msg.content.find((b) => b.type === 'text')?.text || 'Không có phản hồi từ AI.';
}

function readExcelBufferAsRows(fileBuffer, sheetName) {
  const wb = XLSX.read(fileBuffer, { type: 'buffer' });
  const targetSheet = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[targetSheet];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
}

// Bật/tắt toàn bộ tính năng AI (đoán tab, phân tích ảnh/file/câu hỏi tự do).
// Tự động BẬT khi đã có ANTHROPIC_API_KEY trong biến môi trường — trước đó bot sẽ
// im lặng bỏ qua (không trả lời) thay vì trả lời lỗi công khai trong group.
const AI_ENABLED = !!process.env.ANTHROPIC_API_KEY;

// ---------------------------------------------------------------------------
// LINE BOT
// ---------------------------------------------------------------------------
const app = express();
const client = new line.Client(config);

let botUserId = null; // lấy 1 lần lúc khởi động, dùng để nhận biết bot có bị @tag không

const TRIGGER_NGAY = ['báo cáo ngày', 'bao cao ngay'];
function laTriggerNgay(text) {
  if (!text) return false;
  const t = text.trim().toLowerCase();
  return TRIGGER_NGAY.some((kw) => t === kw || t.includes(kw));
}

const TRIGGER_BANHTT = ['bánh trung thu', 'banh trung thu'];
function laTriggerBanhTT(text) {
  if (!text) return false;
  const t = text.trim().toLowerCase();
  return TRIGGER_BANHTT.some((kw) => t === kw || t.includes(kw));
}

// Đã sửa: trước đây mảng này rỗng nên lệnh "BC Luỹ Kế DT" không bao giờ khớp qua tin nhắn text
// (chỉ chạy được khi tự động sau khi nạp file) — nay thêm từ khoá để gõ tay cũng dùng được.
const TRIGGER_LUYKE = ['bc luỹ kế dt', 'bc luy ke dt', 'luỹ kế dt', 'luy ke dt', 'luỹ kế', 'luy ke', 'dt dự kiến', 'dt du kien'];
function laTriggerLuyKe(text) {
  if (!text) return false;
  const t = text.trim().toLowerCase();
  return TRIGGER_LUYKE.some((kw) => t === kw || t.includes(kw));
}

// Lệnh "Giá Vốn"
const TRIGGER_GIAVON = ['giá vốn', 'gia von'];
function laTriggerGiaVon(text) {
  if (!text) return false;
  const t = text.trim().toLowerCase();
  return TRIGGER_GIAVON.some((kw) => t === kw || t.includes(kw));
}

const TRIGGER_HUYMMKK = ['mmkk huỷ', 'mmkk huy', 'huỷ mmkk', 'huy mmkk', 'hủy mmkk'];
function laTriggerHuyMmkk(text) {
  if (!text) return false;
  const t = text.trim().toLowerCase();
  return TRIGGER_HUYMMKK.some((kw) => t === kw || t.includes(kw));
}

const TRIGGER_SP1DONG = ['sp 1 đồng', 'sp 1 dong', 'sản phẩm 1 đồng', 'san pham 1 dong'];
function laTriggerSp1Dong(text) {
  if (!text) return false;
  const t = text.trim().toLowerCase();
  return TRIGGER_SP1DONG.some((kw) => t === kw || t.includes(kw));
}

function toDateTimeKeyForBucket(value) {
  if (typeof value === 'number') {
    const epoch = Date.UTC(1899, 11, 30);
    const d = new Date(epoch + value * 86400000);
    const y = d.getUTCFullYear();
    const mo = String(d.getUTCMonth() + 1).padStart(2, '0');
    const da = String(d.getUTCDate()).padStart(2, '0');
    const hh = String(d.getUTCHours()).padStart(2, '0');
    const mi = String(d.getUTCMinutes()).padStart(2, '0');
    return `${y}-${mo}-${da} ${hh}:${mi}`;
  }
  const s = (value || '').toString().trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2})/);
  if (m) {
    const [, d, mo, y, hh, mi] = m;
    return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')} ${hh.padStart(2, '0')}:${mi.padStart(2, '0')}`;
  }
  return s;
}

function trichNgayTuCauLenh(text) {
  if (!text) return null;
  const m = text.match(/(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?/);
  if (!m) return null;
  const d = m[1].padStart(2, '0');
  const mo = m[2].padStart(2, '0');
  const y = m[3] || new Date().getFullYear().toString();
  return `${y}-${mo}-${d}`;
}

const TRIGGER_LUOTBILL = ['lượt bill', 'luot bill', 'lượt hóa đơn', 'luot hoa don'];
function laTriggerLuotBill(text) {
  if (!text) return false;
  const t = text.trim().toLowerCase();
  return TRIGGER_LUOTBILL.some((kw) => t === kw || t.includes(kw));
}

async function generateLuotBillReport(ngayYeuCau) {
  const sheets = getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${GOOGLE_SHEET_TAB_LUOTBILL}!A2:C`,
  });

  const rows = res.data.values || [];
  if (rows.length === 0) {
    return { type: 'text', text: 'Chưa có dữ liệu Lượt Bill nào được nạp.' };
  }

  let ngayMoiNhat = null;
  let tenSieuThi = '';
  const cacNgayCoDuLieu = new Set();
  for (const [dt, , tenST] of rows) {
    if (!dt) continue;
    const ngay = dt.toString().slice(0, 10);
    cacNgayCoDuLieu.add(ngay);
    if (ngayMoiNhat === null || ngay > ngayMoiNhat) ngayMoiNhat = ngay;
    if (tenST) tenSieuThi = tenST;
  }

  const ngayMucTieu = ngayYeuCau || ngayMoiNhat;
  if (ngayYeuCau && !cacNgayCoDuLieu.has(ngayYeuCau)) {
    return {
      type: 'text',
      text: `Không có dữ liệu Lượt Bill cho ngày ${fmtNgayHienThi(ngayYeuCau)}. Các ngày hiện có: ${[...cacNgayCoDuLieu].sort().map(fmtNgayHienThi).join(', ')}`,
    };
  }
  ngayMoiNhat = ngayMucTieu;

  const BUCKET_ORDER = ['05:30-06:00'];
  for (let h = 6; h <= 20; h++) {
    BUCKET_ORDER.push(`${String(h).padStart(2, '0')}:00-${String(h + 1).padStart(2, '0')}:00`);
  }

  const tong = {};
  BUCKET_ORDER.forEach((k) => { tong[k] = 0; });

  function xacDinhKhung(hh, mm) {
    const tongPhut = hh * 60 + mm;
    if (tongPhut >= 5 * 60 + 30 && tongPhut < 6 * 60) return '05:30-06:00';
    if (tongPhut >= 6 * 60 && tongPhut < 21 * 60) {
      const gioBatDau = Math.floor(tongPhut / 60);
      return `${String(gioBatDau).padStart(2, '0')}:00-${String(gioBatDau + 1).padStart(2, '0')}:00`;
    }
    return null;
  }

  for (const [dt, luotBillRaw] of rows) {
    if (!dt) continue;
    const s = dt.toString();
    const ngay = s.slice(0, 10);
    if (ngay !== ngayMoiNhat) continue;

    const gioPhutMatch = s.match(/(\d{1,2}):(\d{2})$/);
    if (!gioPhutMatch) continue;
    const hh = Number(gioPhutMatch[1]);
    const mm = Number(gioPhutMatch[2]);

    const khung = xacDinhKhung(hh, mm);
    if (!khung) continue;

    tong[khung] += Number(luotBillRaw) || 0;
  }

  const tongCong = Object.values(tong).reduce((a, b) => a + b, 0);
  const khungCoLuot = BUCKET_ORDER.filter((k) => tong[k] > 0);
  const khungMax = khungCoLuot.length ? khungCoLuot.reduce((a, b) => (tong[a] >= tong[b] ? a : b)) : null;
  const khungMin = khungCoLuot.length ? khungCoLuot.reduce((a, b) => (tong[a] <= tong[b] ? a : b)) : null;

  const danhSachContents = BUCKET_ORDER.map((khung) => {
    const soLuot = tong[khung];
    const laMax = khung === khungMax && soLuot > 0;
    const laMin = khung === khungMin && soLuot > 0;
    let icon = '';
    let mauChu = '#333333';
    let doDam = 'regular';
    if (laMax) { icon = '🔥 '; mauChu = '#C62828'; doDam = 'bold'; }
    if (laMin) { icon = '❄️ '; mauChu = '#1565C0'; doDam = 'bold'; }

    return {
      type: 'box',
      layout: 'horizontal',
      margin: 'md',
      contents: [
        { type: 'text', text: `${icon}${khung}`, size: 'sm', color: mauChu, weight: doDam, flex: 3 },
        { type: 'text', text: `${soLuot}`, size: 'sm', color: mauChu, weight: doDam, align: 'end', flex: 1 },
      ],
    };
  });

  return {
    type: 'flex',
    altText: `Lượt Bill ${tenSieuThi} ngày ${fmtNgayHienThi(ngayMoiNhat)}: Tổng ${tongCong} lượt${khungMax ? ', cao nhất ' + khungMax : ''}`,
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: '#1565C0',
        paddingAll: '20px',
        contents: [
          { type: 'text', text: '🧾 BÁO CÁO LƯỢT BILL THEO KHUNG GIỜ', color: '#FFFFFF', weight: 'bold', size: 'lg', wrap: true },
          ...(tenSieuThi ? [{ type: 'text', text: tenSieuThi, color: '#D6E4F5', size: 'sm', margin: 'sm', wrap: true }] : []),
          { type: 'text', text: `Ngày ${fmtNgayHienThi(ngayMoiNhat)}`, color: '#D6E4F5', size: 'sm', margin: tenSieuThi ? 'xs' : 'sm' },
          { type: 'text', text: 'Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị\nThạch Phạm Hoàng Anh -  197042', color: '#FFFFFF', size: 'xxs', margin: 'sm', wrap: true },
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
            backgroundColor: '#E3F2FD',
            cornerRadius: '8px',
            paddingAll: '12px',
            contents: [
              { type: 'text', text: 'TỔNG LƯỢT BILL TRONG NGÀY', size: 'xs', color: '#666666' },
              { type: 'text', text: `${tongCong}`, size: 'xxl', weight: 'bold', color: '#1565C0' },
              ...(khungMax ? [{ type: 'text', text: `🔥 Cao nhất: ${khungMax} (${tong[khungMax]} lượt)`, size: 'xs', color: '#C62828', margin: 'sm' }] : []),
              ...(khungMin ? [{ type: 'text', text: `❄️ Thấp nhất: ${khungMin} (${tong[khungMin]} lượt)`, size: 'xs', color: '#1565C0', margin: 'xs' }] : []),
            ],
          },
          { type: 'separator', margin: 'lg' },
          { type: 'box', layout: 'vertical', margin: 'lg', spacing: 'sm', contents: danhSachContents },
        ],
      },
    },
  };
}

async function generateSp1DongReport() {
  const sheets = getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${GOOGLE_SHEET_TAB_CHITIETXUAT}!A2:C`,
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
        { type: 'text', text: `${idx + 1}. ${ten}`, size: 'sm', color: '#333333', wrap: true, flex: 4 },
        { type: 'text', text: `${sl}`, size: 'sm', color: '#2E7D32', weight: 'bold', align: 'end', flex: 1 },
      ],
    });
  });

  return {
    type: 'flex',
    altText: `Báo cáo SP 1 Đồng: Tổng SL ${tongSoLuong}, ${soMatHang} mặt hàng`,
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
          { type: 'text', text: 'Báo Cáo Thuộc Bản Quyền Quản Lý Siêu Thị\nThạch Phạm Hoàng Anh -  197042', color: '#FFFFFF', size: 'xxs', margin: 'sm', wrap: true },
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
              { type: 'text', text: `${tongSoLuong}`, size: 'xxl', weight: 'bold', color: '#2E7D32' },
              { type: 'text', text: `${soMatHang} mặt hàng khác nhau`, size: 'xs', color: '#888888', margin: 'sm' },
            ],
          },
          { type: 'separator', margin: 'lg' },
          { type: 'box', layout: 'vertical', margin: 'lg', spacing: 'sm', contents: danhSachContents },
        ],
      },
    },
  };
}

// Chạy đúng lệnh báo cáo cũ theo tên khớp được (dùng chung cho cả group lẫn chat riêng)
async function chayLenhCu(text) {
  text = (text || '').normalize('NFC');
  if (laTriggerNgay(text)) return { ten: 'ngày', ket: await generateDailyReport() };
  if (laTriggerBanhTT(text)) return { ten: 'Bánh Trung Thu', ket: await generateBanhTrungThuReport() };
  if (laTriggerLuyKe(text)) return { ten: 'Lũy Kế', ket: await generateLuyKeReport() };
  if (laTriggerGiaVon(text)) return { ten: 'Giá Vốn', ket: await generateGiaVonReport() };
  if (laTriggerHuyMmkk(text)) return { ten: 'MMKK Huỷ', ket: await generateHuyMmkkReport() };
  if (laTriggerSp1Dong(text)) return { ten: 'SP 1 Đồng', ket: await generateSp1DongReport() };
  if (laTriggerLuotBill(text)) return { ten: 'Lượt Bill', ket: await generateLuotBillReport(trichNgayTuCauLenh(text)) };
  return null;
}

// ---- MỚI: nhận diện @tag bot trong tin nhắn text (chỉ có ý nghĩa trong group/room) ----
function extractMentionQuestion(event) {
  const mention = event.message?.mention;
  if (!mention || !Array.isArray(mention.mentionees)) return null;

  const isTaggingBot = mention.mentionees.some(
    (m) => m.isSelf === true || (botUserId && m.userId === botUserId)
  );
  if (!isTaggingBot) return null;

  let question = event.message.text;
  const sorted = [...mention.mentionees].sort((a, b) => b.index - a.index);
  for (const m of sorted) {
    question = question.slice(0, m.index) + question.slice(m.index + m.length);
  }
  return question.trim();
}

// ---- MỚI: nhớ tạm "vừa @tag hỏi gì" theo từng group, để phân biệt file gửi tiếp theo
// là để AI phân tích nhanh (không lưu Sheet) hay vẫn là file nạp dữ liệu như cũ.
// Hết hạn sau 3 phút nếu không có file/ảnh gửi tiếp theo.
const cho_AI_PhanTich = new Map(); // groupId/roomId -> { question, expiresAt }
const THOI_GIAN_CHO_MS = 3 * 60 * 1000;

function datCoDangChoFile(targetId, question) {
  cho_AI_PhanTich.set(targetId, { question, expiresAt: Date.now() + THOI_GIAN_CHO_MS });
}
function layVaXoaCoDangCho(targetId) {
  const info = cho_AI_PhanTich.get(targetId);
  if (!info) return null;
  cho_AI_PhanTich.delete(targetId);
  if (Date.now() > info.expiresAt) return null;
  return info;
}

app.post('/webhook', line.middleware(config), async (req, res) => {
  res.status(200).end();

  const events = req.body.events || [];
  console.log(`[webhook] nhận ${events.length} event(s)`);

  for (const event of events) {
    if (event.type !== 'message') continue;

    const isGroup = event.source?.type === 'group' || event.source?.type === 'room';
    const targetId = event.source?.groupId || event.source?.roomId;

    // ------------------------------------------------------------------
    // ẢNH — luôn dùng AI vision phân tích (không có tính năng cũ nào xử lý ảnh
    // nên không xung đột; LINE cũng không gắn kèm được @tag vào ảnh)
    // ------------------------------------------------------------------
    if (event.message.type === 'image') {
      if (!AI_ENABLED) {
        console.log('[webhook] nhận ảnh nhưng AI đang tắt (chưa có ANTHROPIC_API_KEY) -> bỏ qua, không trả lời');
        continue;
      }
      console.log('[webhook] nhận ảnh, đang phân tích bằng AI...');
      try {
        const buffer = await taiNoiDungFileLine(event.message.id);
        const cauHoiCho = isGroup ? layVaXoaCoDangCho(targetId) : null;
        const answer = await analyzeImage(cauHoiCho?.question || '', buffer, 'image/jpeg');
        await client.replyMessage(event.replyToken, { type: 'text', text: answer });
      } catch (err) {
        console.error('[webhook] Lỗi phân tích ảnh:', err);
        try {
          await client.replyMessage(event.replyToken, { type: 'text', text: `❌ Em xem ảnh này bị lỗi: ${err.message}` });
        } catch (e2) { console.error('[webhook] Lỗi luôn cả khi reply lỗi:', e2.message); }
      }
      continue;
    }

    // ------------------------------------------------------------------
    // FILE EXCEL
    // - Nếu group vừa @tag hỏi gì đó trong 3 phút gần đây -> AI phân tích nhanh,
    //   KHÔNG lưu vào Sheet.
    // - Ngược lại -> giữ nguyên hành vi cũ: tự nhận diện + GHI ĐÈ vào Sheet + trả báo cáo.
    // ------------------------------------------------------------------
    if (event.message.type === 'file') {
      const fileName = event.message.fileName || '';
      if (!/\.(xlsx|xls)$/i.test(fileName)) {
        console.log(`[webhook] file "${fileName}" không phải Excel, bỏ qua`);
        continue;
      }

      const cauHoiCho = isGroup ? layVaXoaCoDangCho(targetId) : null;

      if (cauHoiCho) {
        console.log(`[webhook] file "${fileName}" đang trong luồng @tag -> AI phân tích nhanh, không lưu Sheet`);
        try {
          const buffer = await taiNoiDungFileLine(event.message.id);
          const rows = readExcelBufferAsRows(buffer);
          const answer = await analyzeTableData(
            cauHoiCho.question || 'Tóm tắt và nhận xét nhanh các điểm đáng chú ý trong file này giúp anh.',
            `File Excel đính kèm: ${fileName}`,
            rows
          );
          await client.replyMessage(event.replyToken, { type: 'text', text: answer });
        } catch (err) {
          console.error('[webhook] Lỗi AI phân tích file:', err);
          try {
            await client.replyMessage(event.replyToken, { type: 'text', text: `❌ Em đọc file này bị lỗi: ${err.message}` });
          } catch (e2) { console.error('[webhook] Lỗi luôn cả khi reply lỗi:', e2.message); }
        }
        continue;
      }

      console.log(`[webhook] nhận file "${fileName}", đang tải + nạp vào Sheet...`);
      try {
        const buffer = await taiNoiDungFileLine(event.message.id);
        const ketQua = await napFileVaoSheet(fileName, buffer);
        console.log(`[webhook] đã GHI ĐÈ ${ketQua.soDong} dòng vào tab "${ketQua.tenTab}" (loại: ${ketQua.loai})`);

        try {
          let baoCao;
          if (ketQua.loai === 'banhtt_ton' || ketQua.loai === 'banhtt_ban') {
            baoCao = await generateBanhTrungThuReport();
          } else if (ketQua.loai === 'luyke_dt') {
            baoCao = await generateLuyKeReport();
          } else if (ketQua.loai === 'giavon') {
            baoCao = await generateGiaVonReport();
          } else if (ketQua.loai === 'huymmkk') {
            baoCao = await generateHuyMmkkReport();
          } else {
            baoCao = await generateDailyReport();
          }
          await client.replyMessage(event.replyToken, baoCao);
        } catch (loiBaoCao) {
          console.error('[webhook] nạp file OK nhưng chưa tạo được báo cáo:', loiBaoCao.message);
          await client.replyMessage(event.replyToken, {
            type: 'text',
            text: `✅ Đã nạp ${ketQua.soDong} dòng vào tab "${ketQua.tenTab}".\n⚠️ Chưa tạo được báo cáo ngay: ${loiBaoCao.message}`,
          });
        }
      } catch (err) {
        console.error('[webhook] Lỗi nạp file:', err);
        try {
          await client.replyMessage(event.replyToken, { type: 'text', text: `❌ Lỗi nạp file: ${err.message}` });
        } catch (replyErr) {
          console.error('[webhook] Lỗi luôn cả khi reply lỗi:', replyErr.message);
        }
      }
      continue;
    }

    if (event.message.type !== 'text') continue;
    const text = event.message.text;

    // ------------------------------------------------------------------
    // TRONG GROUP/ROOM: bắt buộc phải @tag bot mới xử lý (theo yêu cầu mới)
    // ------------------------------------------------------------------
    if (isGroup) {
      const question = extractMentionQuestion(event);
      if (question === null) {
        console.log('[webhook] tin nhắn group không @tag bot, bỏ qua');
        continue;
      }

      // Nếu câu hỏi khớp đúng 1 trong các lệnh báo cáo có sẵn -> chạy lệnh đó
      const ketQuaLenhCu = await chayLenhCu(question).catch((err) => {
        console.error('[webhook] Lỗi khi chạy lệnh cũ:', err);
        return { loi: err };
      });

      if (ketQuaLenhCu && !ketQuaLenhCu.loi) {
        try {
          await client.replyMessage(event.replyToken, ketQuaLenhCu.ket);
          console.log(`[webhook] @tag khớp lệnh "${ketQuaLenhCu.ten}", đã trả báo cáo`);
        } catch (err) {
          const chiTietLoi = (err.originalError && err.originalError.response && err.originalError.response.data) || (err.response && err.response.data) || err.message;
          console.error('[webhook] Reply thất bại (' + JSON.stringify(chiTietLoi) + '), thử push thẳng vào group...');
          try {
            await client.pushMessage(targetId, ketQuaLenhCu.ket);
            console.log('[webhook] Đã push thành công báo cáo "' + ketQuaLenhCu.ten + '" vào group (fallback)');
          } catch (pushErr) {
            const chiTietPush = (pushErr.originalError && pushErr.originalError.response && pushErr.originalError.response.data) || (pushErr.response && pushErr.response.data) || pushErr.message;
            console.error('[webhook] Push fallback cũng thất bại. CHI TIẾT LỖI THẬT TỪ LINE: ' + JSON.stringify(chiTietPush));
            console.error('[webhook] Kích thước JSON của báo cáo (ký tự): ' + JSON.stringify(ketQuaLenhCu.ket).length);
          }
        }
        continue;
      }
      if (ketQuaLenhCu && ketQuaLenhCu.loi) {
        try {
          await client.replyMessage(event.replyToken, { type: 'text', text: `⚠️ Không tạo được báo cáo: ${ketQuaLenhCu.loi.message}` });
        } catch (e2) { console.error('[webhook] Lỗi luôn cả khi reply lỗi:', e2.message); }
        continue;
      }

      // Không khớp lệnh có sẵn -> coi là câu hỏi tự do, để AI tự đoán tab + phân tích
      if (!AI_ENABLED) {
        console.log('[webhook] @tag kèm câu hỏi tự do nhưng AI đang tắt -> bỏ qua, không trả lời');
        continue;
      }
      console.log('[webhook] @tag kèm câu hỏi tự do, đang nhờ AI tra cứu + phân tích...');
      try {
        await client.replyMessage(event.replyToken, { type: 'text', text: 'Anh chờ chút, em đang tra cứu và phân tích...' });

        // Đánh dấu "đang chờ" để nếu anh gửi tiếp file/ảnh trong 3 phút, bot biết là để AI phân tích
        datCoDangChoFile(targetId, question);

        const picked = await pickRelevantTab(question);
        if (!picked.tab) {
          await client.pushMessage(targetId, {
            type: 'text',
            text: 'Em chưa xác định được câu hỏi này liên quan tab dữ liệu nào, anh hỏi cụ thể hơn giúp em nhé.',
          });
          continue;
        }
        const sheets = getSheetsClient();
        const rows = await docTabThanhMangDong(sheets, picked.tab);
        const answer = await analyzeTableData(question, picked.tab, rows);
        await client.pushMessage(targetId, { type: 'text', text: answer });
      } catch (err) {
        console.error('[webhook] Lỗi phân tích câu hỏi:', err);
        try {
          await client.pushMessage(targetId, { type: 'text', text: `❌ Em gặp lỗi khi phân tích: ${err.message}` });
        } catch (e2) { console.error('[webhook] Lỗi luôn cả khi push lỗi:', e2.message); }
      }
      continue;
    }

    // ------------------------------------------------------------------
    // CHAT RIÊNG (1-1): giữ hành vi cũ — không cần @tag, gõ đúng lệnh là chạy,
    // gõ câu hỏi khác thì AI cũng tự phân tích luôn cho tiện.
    // ------------------------------------------------------------------
    let ketQuaLenhCuRieng = null;
    try {
      ketQuaLenhCuRieng = await chayLenhCu(text);
    } catch (err) {
      console.error('[webhook] Lỗi tạo báo cáo:', err);
      try {
        await client.replyMessage(event.replyToken, { type: 'text', text: `⚠️ Không tạo được báo cáo: ${err.message}` });
      } catch (replyErr) {
        console.error('[webhook] Reply lỗi thất bại, thử push:', replyErr.message);
        try {
          await client.pushMessage(event.source.userId, { type: 'text', text: `⚠️ Không tạo được báo cáo: ${err.message}` });
        } catch (pushErr) {
          console.error('[webhook] Push fallback lỗi cũng thất bại:', pushErr.message);
        }
      }
      continue;
    }
    if (ketQuaLenhCuRieng) {
      try {
        await client.replyMessage(event.replyToken, ketQuaLenhCuRieng.ket);
        console.log(`[webhook] khớp lệnh "${ketQuaLenhCuRieng.ten}", đã trả báo cáo`);
      } catch (err) {
        console.error('[webhook] Reply thất bại, thử push thẳng...');
        try {
          await client.pushMessage(event.source.userId, ketQuaLenhCuRieng.ket);
          console.log('[webhook] Đã push thành công báo cáo "' + ketQuaLenhCuRieng.ten + '" (fallback)');
        } catch (pushErr) {
          console.error('[webhook] Push fallback cũng thất bại:', pushErr.message);
        }
      }
      continue;
    }

    // Không khớp lệnh nào -> AI tự đoán tab + phân tích (chat riêng không cần @tag)
    if (!AI_ENABLED) {
      console.log('[webhook] (1-1) không khớp lệnh và AI đang tắt -> bỏ qua, không trả lời');
      continue;
    }
    try {
      await client.replyMessage(event.replyToken, { type: 'text', text: 'Anh chờ chút, em đang tra cứu và phân tích...' });
      const picked = await pickRelevantTab(text);
      if (!picked.tab) {
        await client.pushMessage(event.source.userId, {
          type: 'text',
          text: 'Em chưa xác định được câu hỏi này liên quan tab dữ liệu nào, anh hỏi cụ thể hơn giúp em nhé.',
        });
        continue;
      }
      const sheets = getSheetsClient();
      const rows = await docTabThanhMangDong(sheets, picked.tab);
      const answer = await analyzeTableData(text, picked.tab, rows);
      await client.pushMessage(event.source.userId, { type: 'text', text: answer });
    } catch (err) {
      console.error('[webhook] Lỗi phân tích câu hỏi (1-1):', err);
    }
  }
});

app.get('/health', (req, res) => res.send('ok'));

(async () => {
  try {
    const info = await client.getBotInfo();
    botUserId = info.userId;
    console.log('[startup] Bot userId:', botUserId);
  } catch (err) {
    console.error('[startup] Không lấy được botUserId (tính năng @tag có thể không nhận diện đúng):', err.message);
  }

  app.listen(PORT, () => {
    console.log(`LINE bot đang chạy ở port ${PORT}`);
  });
})();

// ================== SP 1 ĐỒNG (Chi Tiết Phiếu Xuất) ==================

// ---- NHẬN DIỆN FILE "Chi Tiết Phiếu Xuất" ----
function isChiTietPhieuXuatFile(headers) {
  return headers.includes('Mã phiếu xuất') &&
         headers.includes('Tên sản phẩm') &&
         headers.includes('Số lượng') &&
         headers.includes('Giá bán');
}

// ---- NẠP FILE VÀO TAB CHITIETXUAT (ghi đè toàn bộ) ----
async function saveChiTietXuatToSheet(sheets, spreadsheetId, data) {
  const rows = data.map(row => [
    row['Tên sản phẩm'],
    row['Số lượng'],
    row['Giá bán']
  ]);

  await sheets.spreadsheets.values.clear({
    spreadsheetId,
    range: 'CHITIETXUAT!A:C',
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: 'CHITIETXUAT!A1',
    valueInputOption: 'RAW',
    requestBody: {
      values: [['Tên sản phẩm', 'Số lượng', 'Giá bán'], ...rows],
    },
  });
}

// ---- XỬ LÝ LỆNH "SP 1 ĐỒNG" ----
async function handleSp1Dong(sheets, spreadsheetId) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: 'CHITIETXUAT!A2:C',
  });

  const rows = res.data.values || [];
  const grouped = {};

  for (const [tenSP, soLuong, giaBan] of rows) {
    if (!tenSP) continue;
    if (Number(giaBan) !== 1) continue;
    if (tenSP.toUpperCase().includes('NẤM')) continue;

    const sl = Number(soLuong) || 0;
    grouped[tenSP] = (grouped[tenSP] || 0) + sl;
  }

  const sorted = Object.entries(grouped)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15);

  if (sorted.length === 0) {
    return 'Không có sản phẩm giá bán 1 đồng nào (đã loại nấm) trong dữ liệu hiện tại.';
  }

  let message = '📋 BÁO CÁO SP GIÁ BÁN 1 ĐỒNG (đã loại nấm)\n\n';
  sorted.forEach(([ten, sl], idx) => {
    message += `${idx + 1}. ${ten}: ${sl}\n`;
  });

  return message;
}
