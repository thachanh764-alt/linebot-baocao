/**
 * server.js
 * =========
 * LINE Bot "Báo cáo trà" + "Báo cáo ngày" + nhận file tự động nạp
 * -----------------------
 * "Báo cáo trà": nhắn "Báo cáo trà" -> đọc 2 tab TON/DOANHTHU, lọc 6 sản phẩm trà C2,
 *   số bán = cột "Số lượng Online" + cột "Số lượng Offline".
 * "Báo cáo ngày": nhắn "Báo cáo ngày" -> đọc 3 tab DOANHTHU_SIEUTHI/DOANHTHU_NGANHHANG/
 *   FRESH_NHAPXUAT, ra 1 thẻ tổng hợp + N thẻ Fresh (1 thẻ/siêu thị).
 * Gửi file Excel trực tiếp vào group -> bot tự nhận diện loại file qua tiêu đề
 *   cột, GHI ĐÈ HOÀN TOÀN (KHÔNG cộng dồn) vào đúng tab, rồi tự trả báo cáo.
 *
 * CẦN CHUẨN BỊ (biến môi trường trên Render, hoặc file .env khi chạy local):
 * ------------------------------------------------------------
 *   LINE_CHANNEL_ACCESS_TOKEN, LINE_CHANNEL_SECRET
 *   GOOGLE_SERVICE_ACCOUNT_JSON (nội dung json service account, dùng trên Render)
 *   HOẶC GOOGLE_SERVICE_ACCOUNT_KEY_PATH=./service-account.json (chạy local)
 *   GOOGLE_SHEET_ID
 *   GOOGLE_SHEET_TAB_TON=TON
 *   GOOGLE_SHEET_TAB_DOANHTHU=DOANHTHU
 *   GOOGLE_SHEET_TAB_DOANHTHU_SIEUTHI=DOANHTHU_SIEUTHI
 *   GOOGLE_SHEET_TAB_DOANHTHU_NGANHHANG=DOANHTHU_NGANHHANG
 *   GOOGLE_SHEET_TAB_FRESH=FRESH_NHAPXUAT
 *   PORT=3000
 *
 * Phải SHARE Google Sheet cho email service account, quyền EDITOR.
 *
 * Chạy: npm start
 */

require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const { google } = require('googleapis');
const XLSX = require('xlsx');

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

const TRIGGER_KEYWORDS = ['báo cáo trà', 'bao cao tra'];

const SAN_PHAM_TRA = {
  '2601001494': 'Nước sâm C2 Cool',
  '2203000875': 'Trà đen dâu anh đào C2',
  '2602001178': 'Trà đen tắc C2',
  '2006000354': 'Trà hồng vải C2',
  '2204000011': 'Trà xanh chanh bạc hà C2',
  '1607002174': 'Nước C2 trà xanh hương chanh 360ml',
};

const QUY_DOI_THUNG = 24;

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
  const colSLOnline = timCotTheoTen(header, 'Số lượng Online');
  const colSLOffline = timCotTheoTen(header, 'Số lượng Offline');

  const ban = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const ma = (row[colMa] || '').toString().trim();
    if (SAN_PHAM_TRA[ma]) {
      const st = row[colTenST];
      const slOnline = Number(row[colSLOnline]) || 0;
      const slOffline = Number(row[colSLOffline]) || 0;
      ban[st] = (ban[st] || 0) + slOnline + slOffline;
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
// FLEX MESSAGE báo cáo trà
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
    timeZone: 'Asia/Ho_Chi_Minh',
  });

  const headerContents = [
    { type: 'text', text: '🍵 BÁO CÁO TRÀ', color: '#FFFFFF', weight: 'bold', size: 'lg' },
    ...Object.values(SAN_PHAM_TRA).map((ten) => ({
      type: 'text', text: `• ${ten}`, color: '#EAEAEA', size: 'sm', wrap: true,
    })),
    {
      type: 'text',
      text: `${Object.keys(SAN_PHAM_TRA).length} sản phẩm · cập nhật lúc ${thoiGian} · ${tong.soSieuThi} siêu thị · quy đổi 1 thùng = ${QUY_DOI_THUNG} chai`,
      color: '#CFCFCF', size: 'xs', wrap: true, margin: 'md',
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
          type: 'text', text: fmtPct(r.tl), size: 'sm', flex: 2, align: 'end', weight: 'bold',
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
        type: 'box', layout: 'vertical', backgroundColor: MAU_XANH_HEADER, paddingAll: '20px',
        contents: headerContents,
      },
      body: {
        type: 'box', layout: 'vertical', paddingAll: '16px', spacing: 'sm',
        contents: bodyContents,
      },
    },
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
  const [y, m, d] = dateKey.split('-');
  return `${d}/${m}/${y}`;
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

function dongThongTinNgay(icon, nhan, giaTri, dam) {
  return {
    type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: `${icon} ${nhan}`, size: 'sm', flex: 3, color: '#555555' },
      { type: 'text', text: giaTri, size: dam ? 'md' : 'sm', flex: 2, align: 'end', weight: dam ? 'bold' : 'regular', color: dam ? '#22A45D' : '#111111' },
    ],
  };
}

function dongNganhHangDon(ten, giaTri) {
  return {
    type: 'box', layout: 'horizontal', margin: 'sm', contents: [
      { type: 'text', text: ten, size: 'sm', flex: 5, wrap: true, color: '#333333' },
      { type: 'text', text: giaTri, size: 'sm', flex: 3, align: 'end', weight: 'bold', color: '#111111' },
    ],
  };
}

function taoCardBaoCaoTheoNgay(tong, ngayHienThi) {
  const bodyContents = [
    {
      type: 'box', layout: 'vertical', backgroundColor: '#F0F7F2', cornerRadius: 'md', paddingAll: '14px',
      contents: [
        { type: 'text', text: 'TỔNG DOANH THU', size: 'xs', color: '#888888' },
        { type: 'text', text: fmtSo(tong.tongDoanhThu) + ' đ', size: 'xxl', weight: 'bold', color: '#1a1a1a', margin: 'sm' },
      ],
    },
    { type: 'separator', margin: 'lg' },
    dongThongTinNgay('🏬', 'DT Offline', fmtSo(tong.dtOffline) + ' đ', false),
    dongThongTinNgay('🌐', 'DT Online', fmtSo(tong.dtOnline) + ' đ', false),
    dongThongTinNgay('🧾', 'Tổng số bill', fmtSo(tong.soBill), false),
    dongThongTinNgay('💳', 'Giá trị Bill TB', fmtSo(tong.giaTriBillTB) + ' đ', false),
    { type: 'separator', margin: 'lg' },
    { type: 'text', text: '📦 THEO NGÀNH HÀNG', size: 'sm', weight: 'bold', color: '#333333', margin: 'lg' },
  ];

  CARD1_CATEGORY_ORDER.forEach((ten) => {
    bodyContents.push(dongNganhHangDon(ten, fmtSo(tong.byNganh[ten] || 0) + ' đ'));
  });

  return {
    type: 'flex',
    altText: `Báo cáo theo ngày ${ngayHienThi}: Tổng doanh thu ${fmtSo(tong.tongDoanhThu)} đ`,
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#2C4A3B', paddingAll: '20px',
        contents: [
          { type: 'text', text: '📅 BÁO CÁO THEO NGÀY', color: '#FFFFFF', weight: 'bold', size: 'lg' },
          { type: 'text', text: ngayHienThi, color: '#DCEAE1', size: 'sm', margin: 'sm' },
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
  const colDTOffline = timCotTheoTen(headerST, 'Doanh thu offline');
  const colDTOnline = timCotTheoTen(headerST, 'Doanh thu Online');
  const colSoBill = timCotTheoTen(headerST, 'Tổng số bill');

  const ngayMoiNhatST = timNgayMoiNhat(rowsST, colNgayST);
  let dtOffline = 0, dtOnline = 0, soBill = 0;
  for (let i = 1; i < rowsST.length; i++) {
    const row = rowsST[i];
    if (!row || row[colNgayST] === undefined || row[colNgayST] === '') continue;
    if (toDateKey(row[colNgayST]) !== ngayMoiNhatST) continue;
    dtOffline += Number(row[colDTOffline]) || 0;
    dtOnline += Number(row[colDTOnline]) || 0;
    soBill += Number(row[colSoBill]) || 0;
  }
  const tongDoanhThu = dtOffline + dtOnline;
  const giaTriBillTB = soBill > 0 ? tongDoanhThu / soBill : 0;

  const headerNH = rowsNH[0];
  const colNgayNH = timCotTheoTen(headerNH, 'Ngày xuất');
  const colNganhNH = timCotTheoTen(headerNH, 'Ngành hàng BHX');
  const colDoanhThuNH = timCotTheoTen(headerNH, 'Doanh thu');

  const ngayMoiNhatNH = timNgayMoiNhat(rowsNH, colNgayNH);
  const byNganh = {};
  for (let i = 1; i < rowsNH.length; i++) {
    const row = rowsNH[i];
    if (!row || row[colNgayNH] === undefined || row[colNgayNH] === '') continue;
    if (toDateKey(row[colNgayNH]) !== ngayMoiNhatNH) continue;
    const ten = (row[colNganhNH] || '').toString().trim();
    if (!ten) continue;
    byNganh[ten] = (byNganh[ten] || 0) + (Number(row[colDoanhThuNH]) || 0);
  }

  const ngayHienThi = fmtNgayHienThi(ngayMoiNhatST || ngayMoiNhatNH);
  const card1 = taoCardBaoCaoTheoNgay(
    { tongDoanhThu, dtOffline, dtOnline, soBill, giaTriBillTB, byNganh },
    ngayHienThi
  );

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

    const ma = row[colMaSTFresh];
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

  return [card1, ...cardsFresh].slice(0, 5);
}

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
// NẠP FILE NGƯỜI DÙNG GỬI TRỰC TIẾP VÀO GROUP (.xlsx/.xls)
// ---------------------------------------------------------------------------
async function taiNoiDungFileLine(messageId) {
  const stream = await client.getMessageContent(messageId);
  const chunks = [];
  for await (const chunk of stream) chunks.push(chunk);
  return Buffer.concat(chunks);
}

function nhanDangLoaiFile(header) {
  const co = (ten) => header.includes(ten);

  if (co('Mã Model') && co('Tồn kho siêu thị')) {
    return { loai: 'ton', tenTab: GOOGLE_SHEET_TAB_TON };
  }
  if (co('Mã Model') && co('Tổng số lượng') && !co('Tồn kho siêu thị')) {
    return { loai: 'tra_ban', tenTab: GOOGLE_SHEET_TAB_DOANHTHU };
  }
  if (co('Ngày') && co('Mã siêu thị') && co('Doanh thu offline')) {
    return { loai: 'doanhthu_sieuthi', tenTab: GOOGLE_SHEET_TAB_DOANHTHU_SIEUTHI };
  }
  if (co('Ngày') && co('Mã siêu thị') && co('Ngành hàng - Phân tích') && co('Thành tiền phải thu khách hàng (chưa VAT)')) {
    return { loai: 'fresh', tenTab: GOOGLE_SHEET_TAB_FRESH };
  }
  if (co('Ngày xuất') && co('Ngành hàng BHX') && co('Doanh thu')) {
    return { loai: 'doanhthu_nganhhang', tenTab: GOOGLE_SHEET_TAB_DOANHTHU_NGANHHANG };
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

  const nhanDang = nhanDangLoaiFile(header);
  if (!nhanDang) {
    throw new Error(
      `Không nhận diện được loại báo cáo từ file "${fileName || ''}". Kiểm tra lại tiêu đề cột trong file có đúng mẫu không.`
    );
  }

  const sheets = getSheetsClient();
  const destRows = await docTabThanhMangDong(sheets, nhanDang.tenTab);
  const destHeader = destRows[0];

  const rowsToAppend = dataRows.map((row) =>
    destHeader.map((tenCot) => {
      const idx = header.indexOf(tenCot);
      return idx === -1 ? '' : row[idx] ?? '';
    })
  );

  // GHI ĐÈ HOÀN TOÀN - xoá sạch data cũ (giữ hàng tiêu đề) rồi ghi data mới,
  // KHÔNG cộng dồn/nối thêm bất kỳ loại file nào.
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
// LINE BOT
// ---------------------------------------------------------------------------
const app = express();
const client = new line.Client(config);

function laTrigger(text) {
  if (!text) return false;
  const t = text.trim().toLowerCase();
  return TRIGGER_KEYWORDS.some((kw) => t === kw || t.includes(kw));
}

const TRIGGER_NGAY = ['báo cáo ngày', 'bao cao ngay'];
function laTriggerNgay(text) {
  if (!text) return false;
  const t = text.trim().toLowerCase();
  return TRIGGER_NGAY.some((kw) => t === kw || t.includes(kw));
}

app.post('/webhook', line.middleware(config), async (req, res) => {
  res.status(200).end();

  const events = req.body.events || [];
  console.log(`[webhook] nhận ${events.length} event(s)`);

  for (const event of events) {
    if (event.type !== 'message') continue;

    if (event.message.type === 'file') {
      const fileName = event.message.fileName || '';
      if (!/\.(xlsx|xls)$/i.test(fileName)) {
        console.log(`[webhook] file "${fileName}" không phải Excel, bỏ qua`);
        continue;
      }
      console.log(`[webhook] nhận file "${fileName}", đang tải + nạp vào Sheet...`);
      try {
        const buffer = await taiNoiDungFileLine(event.message.id);
        const ketQua = await napFileVaoSheet(fileName, buffer);
        console.log(`[webhook] đã GHI ĐÈ ${ketQua.soDong} dòng vào tab "${ketQua.tenTab}" (loại: ${ketQua.loai})`);

        try {
          let baoCao;
          if (ketQua.loai === 'ton' || ketQua.loai === 'tra_ban') {
            baoCao = await generateTraReport();
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

    if (laTrigger(text)) {
      console.log('[webhook] khớp "Báo cáo trà", đang tạo báo cáo...');
      try {
        const flexMessage = await generateTraReport();
        await client.replyMessage(event.replyToken, flexMessage);
        console.log('[webhook] tạo báo cáo + reply thành công');
      } catch (err) {
        console.error('[webhook] Lỗi tạo báo cáo trà:', err);
        try {
          await client.replyMessage(event.replyToken, {
            type: 'text',
            text: `⚠️ Không tạo được báo cáo trà: ${err.message}`,
          });
        } catch (replyErr) {
          console.error('[webhook] Lỗi luôn cả khi reply lỗi:', replyErr.message);
        }
      }
      continue;
    }

    if (laTriggerNgay(text)) {
      console.log('[webhook] khớp "Báo cáo ngày", đang tạo báo cáo...');
      try {
        const flexMessages = await generateDailyReport();
        await client.replyMessage(event.replyToken, flexMessages);
        console.log('[webhook] tạo báo cáo ngày + reply thành công');
      } catch (err) {
        console.error('[webhook] Lỗi tạo báo cáo ngày:', err);
        try {
          await client.replyMessage(event.replyToken, {
            type: 'text',
            text: `⚠️ Không tạo được báo cáo ngày: ${err.message}`,
          });
        } catch (replyErr) {
          console.error('[webhook] Lỗi luôn cả khi reply lỗi:', replyErr.message);
        }
      }
      continue;
    }

    console.log('[webhook] text không khớp từ khoá trigger nào, bỏ qua');
  }
});

app.get('/health', (req, res) => res.send('ok'));

app.listen(PORT, () => {
  console.log(`LINE bot "Báo cáo trà" đang chạy ở port ${PORT}`);
});
