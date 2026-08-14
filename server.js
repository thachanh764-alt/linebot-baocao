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
 *   3. Trả lời lại đúng định dạng báo cáo.
 *
 * CẦN CHUẨN BỊ TRƯỚC KHI CHẠY (điền vào file .env cùng thư mục):
 * ------------------------------------------------------------
 *   LINE_CHANNEL_ACCESS_TOKEN=...........(lấy trong LINE Developers Console)
 *   LINE_CHANNEL_SECRET=..................(lấy trong LINE Developers Console)
 *   GOOGLE_SERVICE_ACCOUNT_KEY_PATH=./service-account.json
 *      (đường dẫn tới file json service account anh đang có trên máy)
 *   GOOGLE_SHEET_ID=..............(ID của Google Sheet - lấy trong URL:
 *      docs.google.com/spreadsheets/d/<ID_Ở_ĐÂY>/edit)
 *   GOOGLE_SHEET_TAB_TON=TON       (tên tab chứa data Tồn - đổi cho khớp
 *      tên tab thật anh đặt)
 *   GOOGLE_SHEET_TAB_DOANHTHU=DOANHTHU  (tên tab chứa data Doanh thu)
 *   PORT=3000
 *
 * Và phải SHARE Google Sheet đó cho email của service account
 * (vd: linebot-sheets@linebot-baocao.iam.gserviceaccount.com) với quyền
 * Editor (không phải chỉ Viewer, vì bot cần ghi/nạp data từ file gửi vào).
 *
 * QUAN TRỌNG: các tab phải giữ nguyên tên cột giống file gốc ở hàng đầu
 * tiên: tab TỒN cần có cột "Mã Model", "Tên siêu thị", "Tồn kho siêu thị";
 * tab DOANH THU cần có cột "Mã Model", "Tên siêu thị", "Tổng số lượng".
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
  // Cần quyền ghi (spreadsheets - không phải readonly) vì bot còn phải nạp
  // (append) dữ liệu từ file người dùng gửi vào group.
  const scopes = ['https://www.googleapis.com/auth/spreadsheets'];
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
// TÍNH TOÁN + FORMAT BÁO CÁO
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

function taoNoiDungBaoCao(ton, ban) {
  const { rows, tong } = tinhDuLieu(ton, ban);
  const now = new Date();
  const thoiGian = now.toLocaleString('vi-VN', {
    hour: '2-digit', minute: '2-digit', day: '2-digit', month: '2-digit', year: 'numeric',
  });

  const lines = [];
  lines.push('🍵 BÁO CÁO TRÀ');
  Object.values(SAN_PHAM_TRA).forEach((ten) => lines.push(`• ${ten}`));
  lines.push('');
  lines.push(
    `${Object.keys(SAN_PHAM_TRA).length} sản phẩm · cập nhật lúc ${thoiGian} · ` +
    `${tong.soSieuThi} siêu thị · quy đổi 1 thùng = ${QUY_DOI_THUNG} chai`
  );
  lines.push('');
  lines.push(
    `TỔNG TẤT CẢ: Tồn ${fmtSo(tong.ton)} thùng | Bán ${fmtSo(tong.ban)} thùng | TL ${fmtPct(tong.tl)}`
  );
  lines.push('');

  rows.forEach((r) => {
    const icon = r.tl >= 20 ? '🟢' : '🔴';
    lines.push(`${icon} ${r.ten}: Tồn ${fmtSo(r.ton)} | Bán ${fmtSo(r.ban)} | TL ${fmtPct(r.tl)}`);
  });

  return lines.join('\n');
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
// BÁO CÁO DOANH THU / SẢN LƯỢNG (đọc 4 tab thang7_doanhthu, thang8_doanhthu,
// thang7_sanluong, thang8_sanluong) - kích hoạt khi nhắn "Báo cáo" / "Báo cáo DT"
// ---------------------------------------------------------------------------
const GOOGLE_SHEET_TAB_T7_DOANHTHU = process.env.GOOGLE_SHEET_TAB_T7_DOANHTHU || 'thang7_doanhthu.xlsx';
const GOOGLE_SHEET_TAB_T8_DOANHTHU = process.env.GOOGLE_SHEET_TAB_T8_DOANHTHU || 'thang8_doanhthu.xlsx';
const GOOGLE_SHEET_TAB_T7_SANLUONG = process.env.GOOGLE_SHEET_TAB_T7_SANLUONG || 'thang7_sanluong.xlsx';
const GOOGLE_SHEET_TAB_T8_SANLUONG = process.env.GOOGLE_SHEET_TAB_T8_SANLUONG || 'thang8_sanluong.xlsx';

// Nhóm ngành hàng FRESH - còn lại tất cả coi là FMCG
const NGANH_FRESH = ['Thịt', 'Cá (Hải sản)', 'Rau Củ Quả CL', 'Trái cây', 'Trứng'];
function laFresh(nganh) {
  return NGANH_FRESH.includes((nganh || '').trim());
}

// Chuyển giá trị "Ngày xuất" (có thể là serial number của Google Sheets hoặc
// chuỗi dd/mm/yyyy) về 1 khoá ngày thống nhất "yyyy-mm-dd" để đếm số ngày.
function toDateKey(value) {
  if (typeof value === 'number') {
    const epoch = Date.UTC(1899, 11, 30); // mốc ngày 0 của Google Sheets
    const d = new Date(epoch + value * 86400000);
    return d.toISOString().slice(0, 10);
  }
  const s = (value || '').toString().trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return s;
}

function soNgayTrongThang(thang, nam) {
  return new Date(nam, thang, 0).getDate();
}

// Gộp 1 tab (doanh thu hoặc sản lượng) theo Ngành hàng BHX, trả về tổng theo
// từng ngành + tổng tất cả + số ngày có phát sinh dữ liệu trong tab.
function gomTheoNganhHang(rows, tenCotGiaTri) {
  const header = rows[0];
  const colNgay = timCotTheoTen(header, 'Ngày xuất');
  const colNganh = timCotTheoTen(header, 'Ngành hàng BHX');
  const colGiaTri = timCotTheoTen(header, tenCotGiaTri);

  const byNganh = {};
  const ngayKeys = new Set();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const ngay = row[colNgay];
    if (ngay === undefined || ngay === '') continue;
    ngayKeys.add(toDateKey(ngay));

    const nganh = (row[colNganh] || '').toString().trim();
    if (!nganh) continue;
    const giaTri = Number(row[colGiaTri]) || 0;
    byNganh[nganh] = (byNganh[nganh] || 0) + giaTri;
  }

  const tong = Object.values(byNganh).reduce((s, v) => s + v, 0);
  return { byNganh, tong, soNgay: ngayKeys.size };
}

// Định dạng tiền kiểu VN: >=1 tỷ -> "x,xx tỷ", >=1 triệu -> "x,x triệu", còn lại -> "xxx.xxx đ"
function fmtTien(n) {
  const abs = Math.abs(n);
  if (abs >= 1e9) return (n / 1e9).toFixed(2).replace('.', ',') + ' tỷ';
  if (abs >= 1e6) return (n / 1e6).toFixed(1).replace('.', ',') + ' triệu';
  return fmtSo(n) + ' đ';
}

// Định dạng số lượng kiểu VN 1 chữ số thập phân: 19416.4 -> "19.416,4"
function fmtSoLuongVN(n) {
  const rounded = Math.round(n * 10) / 10;
  const parts = rounded.toFixed(1).split('.');
  const nguyen = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return `${nguyen},${parts[1]}`;
}

// "+12.3%" hoặc "-5.0%" kèm mũi tên
function tinhPhanTramTangGiam(giaTriMoi, giaTriCu) {
  if (!giaTriCu) return 0;
  return ((giaTriMoi - giaTriCu) / giaTriCu) * 100;
}
function fmtPctMuiTen(p) {
  const r = Math.round(p * 10) / 10;
  const mui = r >= 0 ? '▲' : '▼';
  const dau = r >= 0 ? '+' : '';
  return `${mui}${dau}${r.toFixed(1)}%`;
}
function mauTangGiam(p) {
  return p >= 0 ? '#27AE60' : '#E74C3C';
}

function taoFlexBaoCaoDoanhThu(t7, t8) {
  const nam = new Date().getFullYear();
  const soNgayThang8 = soNgayTrongThang(8, nam);
  const heSoQuyDoi = t8.soNgay > 0 ? soNgayThang8 / t8.soNgay : 0;

  const duKienT8Tong = t8.tong * heSoQuyDoi;
  const thucTeT7Tong = t7.tong;
  const momTong = tinhPhanTramTangGiam(duKienT8Tong, thucTeT7Tong);

  const tatCaNganh = new Set([...Object.keys(t7.byNganh), ...Object.keys(t8.byNganh)]);
  let freshT8 = 0, freshT7 = 0, fmcgT8 = 0, fmcgT7 = 0;
  const chiTiet = [];
  tatCaNganh.forEach((nganh) => {
    const t7v = t7.byNganh[nganh] || 0;
    const t8vDuKien = (t8.byNganh[nganh] || 0) * heSoQuyDoi;
    if (laFresh(nganh)) { freshT8 += t8vDuKien; freshT7 += t7v; }
    else { fmcgT8 += t8vDuKien; fmcgT7 += t7v; }
    chiTiet.push({ nganh, duKienT8: t8vDuKien, thucTeT7: t7v, mom: tinhPhanTramTangGiam(t8vDuKien, t7v) });
  });

  const momFresh = tinhPhanTramTangGiam(freshT8, freshT7);
  const momFmcg = tinhPhanTramTangGiam(fmcgT8, fmcgT7);

  const card1 = {
    type: 'bubble',
    size: 'giga',
    header: {
      type: 'box', layout: 'vertical', backgroundColor: '#22A45D', paddingAll: '20px',
      contents: [
        { type: 'text', text: '📊 BÁO CÁO DOANH THU', color: '#FFFFFF', weight: 'bold', size: 'lg' },
        { type: 'text', text: `Dự kiến T8 (${t8.soNgay}/${soNgayThang8} ngày) so với T7`, color: '#E8F8EF', size: 'sm', margin: 'sm' },
      ],
    },
    body: {
      type: 'box', layout: 'vertical', paddingAll: '16px', spacing: 'md',
      contents: [
        {
          type: 'box', layout: 'horizontal', contents: [
            { type: 'box', layout: 'vertical', flex: 1, contents: [
              { type: 'text', text: 'DỰ KIẾN T8', size: 'xs', color: '#888888' },
              { type: 'text', text: fmtTien(duKienT8Tong), size: 'md', weight: 'bold' },
            ]},
            { type: 'box', layout: 'vertical', flex: 1, contents: [
              { type: 'text', text: 'THỰC TẾ T7', size: 'xs', color: '#888888' },
              { type: 'text', text: fmtTien(thucTeT7Tong), size: 'md', weight: 'bold' },
            ]},
            { type: 'box', layout: 'vertical', flex: 1, contents: [
              { type: 'text', text: 'MoM T8 - T7', size: 'xs', color: '#888888' },
              { type: 'text', text: fmtPctMuiTen(momTong), size: 'md', weight: 'bold', color: mauTangGiam(momTong) },
              { type: 'text', text: `${momTong >= 0 ? '▲' : '▼'} ${fmtTien(Math.abs(duKienT8Tong - thucTeT7Tong))}`, size: 'xs', color: mauTangGiam(momTong) },
            ]},
          ],
        },
        { type: 'separator' },
        { type: 'text', text: 'DOANH THU THEO NGÀNH HÀNG (DỰ KIẾN T8)', size: 'xs', color: '#888888', weight: 'bold' },
        {
          type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: '🥬 FRESH', size: 'sm', weight: 'bold', flex: 3 },
            { type: 'text', text: fmtTien(freshT8), size: 'sm', flex: 2, align: 'end' },
            { type: 'text', text: fmtPctMuiTen(momFresh), size: 'sm', flex: 2, align: 'end', color: mauTangGiam(momFresh), weight: 'bold' },
          ],
        },
        {
          type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: '🛒 FMCG', size: 'sm', weight: 'bold', flex: 3 },
            { type: 'text', text: fmtTien(fmcgT8), size: 'sm', flex: 2, align: 'end' },
            { type: 'text', text: fmtPctMuiTen(momFmcg), size: 'sm', flex: 2, align: 'end', color: mauTangGiam(momFmcg), weight: 'bold' },
          ],
        },
        { type: 'text', text: '* Số liệu T8 là dự kiến, chiếu từ dữ liệu thực tế các ngày đã phát sinh', size: 'xxs', color: '#AAAAAA', wrap: true, margin: 'md' },
      ],
    },
  };

  const freshRows = chiTiet.filter((c) => laFresh(c.nganh)).sort((a, b) => b.duKienT8 - a.duKienT8);
  const fmcgRows = chiTiet.filter((c) => !laFresh(c.nganh)).sort((a, b) => b.duKienT8 - a.duKienT8);

  function dongNganhHang(c) {
    return {
      type: 'box', layout: 'horizontal', margin: 'sm', contents: [
        { type: 'text', text: c.nganh, size: 'sm', flex: 5, wrap: true },
        { type: 'text', text: fmtTien(c.duKienT8), size: 'sm', flex: 3, align: 'end' },
        { type: 'text', text: fmtPctMuiTen(c.mom), size: 'sm', flex: 2, align: 'end', color: mauTangGiam(c.mom), weight: 'bold' },
      ],
    };
  }

  const card2 = {
    type: 'bubble',
    size: 'giga',
    header: {
      type: 'box', layout: 'vertical', backgroundColor: '#2E86DE', paddingAll: '20px',
      contents: [
        { type: 'text', text: '📋 CHI TIẾT NGÀNH HÀNG', color: '#FFFFFF', weight: 'bold', size: 'lg' },
        { type: 'text', text: 'Tăng/giảm doanh thu dự kiến T8 so với T7', color: '#E4F0FD', size: 'sm', margin: 'sm' },
      ],
    },
    body: {
      type: 'box', layout: 'vertical', paddingAll: '16px', spacing: 'sm',
      contents: [
        {
          type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: '🥬 Ngành FRESH', size: 'sm', weight: 'bold', flex: 5 },
            { type: 'text', text: `${momFresh >= 0 ? '▲' : '▼'}${fmtTien(Math.abs(freshT8 - freshT7))} (${fmtPctMuiTen(momFresh).slice(1)})`, size: 'xs', flex: 5, align: 'end', color: mauTangGiam(momFresh), weight: 'bold' },
          ],
        },
        { type: 'separator', margin: 'sm' },
        ...freshRows.map(dongNganhHang),
        { type: 'separator', margin: 'md' },
        {
          type: 'box', layout: 'horizontal', margin: 'md', contents: [
            { type: 'text', text: '🛒 Ngành FMCG', size: 'sm', weight: 'bold', flex: 5 },
            { type: 'text', text: `${momFmcg >= 0 ? '▲' : '▼'}${fmtTien(Math.abs(fmcgT8 - fmcgT7))} (${fmtPctMuiTen(momFmcg).slice(1)})`, size: 'xs', flex: 5, align: 'end', color: mauTangGiam(momFmcg), weight: 'bold' },
          ],
        },
        { type: 'separator', margin: 'sm' },
        ...fmcgRows.map(dongNganhHang),
      ],
    },
  };

  return { card1, card2, altText: `Báo cáo doanh thu: Dự kiến T8 ${fmtTien(duKienT8Tong)}, MoM ${fmtPctMuiTen(momTong)}` };
}

function taoFlexSanLuong(t7, t8) {
  const nam = new Date().getFullYear();
  const soNgayThang8 = soNgayTrongThang(8, nam);
  const heSoQuyDoi = t8.soNgay > 0 ? soNgayThang8 / t8.soNgay : 0;

  const duKienT8 = t8.tong * heSoQuyDoi;
  const thucTeT7 = t7.tong;
  const mom = tinhPhanTramTangGiam(duKienT8, thucTeT7);

  return {
    type: 'bubble',
    size: 'giga',
    header: {
      type: 'box', layout: 'vertical', backgroundColor: '#E67E22', paddingAll: '20px',
      contents: [
        { type: 'text', text: '📦 SẢN LƯỢNG BÁN', color: '#FFFFFF', weight: 'bold', size: 'lg' },
        { type: 'text', text: `Dự kiến T8 (${t8.soNgay}/${soNgayThang8} ngày) so với T7`, color: '#FCEEE0', size: 'sm', margin: 'sm' },
      ],
    },
    body: {
      type: 'box', layout: 'vertical', paddingAll: '16px', spacing: 'md',
      contents: [
        {
          type: 'box', layout: 'horizontal', contents: [
            { type: 'box', layout: 'vertical', flex: 1, contents: [
              { type: 'text', text: 'DỰ KIẾN T8', size: 'xs', color: '#888888' },
              { type: 'text', text: fmtSoLuongVN(duKienT8), size: 'md', weight: 'bold' },
            ]},
            { type: 'box', layout: 'vertical', flex: 1, contents: [
              { type: 'text', text: 'THỰC TẾ T7', size: 'xs', color: '#888888' },
              { type: 'text', text: fmtSoLuongVN(thucTeT7), size: 'md', weight: 'bold' },
            ]},
            { type: 'box', layout: 'vertical', flex: 1, contents: [
              { type: 'text', text: 'MoM T8 - T7', size: 'xs', color: '#888888' },
              { type: 'text', text: fmtPctMuiTen(mom), size: 'md', weight: 'bold', color: mauTangGiam(mom) },
              { type: 'text', text: `${mom >= 0 ? '▲' : '▼'} ${fmtSoLuongVN(Math.abs(duKienT8 - thucTeT7))}`, size: 'xs', color: mauTangGiam(mom) },
            ]},
          ],
        },
      ],
    },
  };
}

// ---------------------------------------------------------------------------
// BÁO CÁO NGÀY THEO TỪNG SIÊU THỊ (đọc tab baocao_ngay.xlsx)
// - kích hoạt khi nhắn "Báo cáo ngày"
// ---------------------------------------------------------------------------
const GOOGLE_SHEET_TAB_BAOCAO_NGAY = process.env.GOOGLE_SHEET_TAB_BAOCAO_NGAY || 'baocao_ngay.xlsx';
const SO_THE_MOI_CAROUSEL = 12; // giới hạn của LINE: tối đa 12 bubble / carousel
const SO_CAROUSEL_TOI_DA = 5; // giới hạn của LINE: tối đa 5 tin nhắn / lần reply

const TEN_THU_VN = ['Chủ Nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];
function fmtNgayVN(d) {
  return `${TEN_THU_VN[d.getDay()]}, ${d.toLocaleDateString('vi-VN')}`;
}

function dongThongTin(icon, nhan, giaTri, dam) {
  return {
    type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: `${icon} ${nhan}`, size: 'sm', flex: 3, color: '#555555' },
      { type: 'text', text: giaTri, size: dam ? 'md' : 'sm', flex: 2, align: 'end', weight: dam ? 'bold' : 'regular', color: dam ? '#22A45D' : '#111111' },
    ],
  };
}

function taoTheDoanhThuNgay(cuaHang, ngayHienThi) {
  const tongDoanhThu = cuaHang.dtOffline + cuaHang.dtOnline;

  return {
    type: 'flex',
    altText: `Doanh thu ${cuaHang.ten}: ${fmtSo(tongDoanhThu)} đ`,
    contents: {
      type: 'bubble',
      size: 'mega',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#22A45D', paddingAll: '16px',
        contents: [
          { type: 'text', text: '📊 BÁO CÁO DOANH THU', color: '#FFFFFF', weight: 'bold', size: 'md' },
          { type: 'text', text: `📅 ${ngayHienThi}`, color: '#E8F8EF', size: 'xs', margin: 'sm' },
        ],
      },
      body: {
        type: 'box', layout: 'vertical', paddingAll: '16px', spacing: 'sm',
        contents: [
          { type: 'text', text: `🏢 ${cuaHang.ten}`, weight: 'bold', size: 'sm', wrap: true, margin: 'none' },
          { type: 'separator', margin: 'md' },
          dongThongTin('🏬', 'Doanh thu offline', fmtSo(cuaHang.dtOffline) + ' đ', false),
          dongThongTin('🛍️', 'Doanh thu online', fmtSo(cuaHang.dtOnline) + ' đ', false),
          { type: 'separator', margin: 'md' },
          dongThongTin('💰', 'Tổng doanh thu', fmtSo(tongDoanhThu) + ' đ', true),
        ],
      },
    },
  };
}

function taoTheBillNgay(cuaHang, ngayHienThi) {
  const tongDoanhThu = cuaHang.dtOffline + cuaHang.dtOnline;
  const giaTriBill = cuaHang.soBill > 0 ? tongDoanhThu / cuaHang.soBill : 0;

  return {
    type: 'flex',
    altText: `Bill ${cuaHang.ten}: ${fmtSo(cuaHang.soBill)} bill`,
    contents: {
      type: 'bubble',
      size: 'mega',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#2E86DE', paddingAll: '16px',
        contents: [
          { type: 'text', text: '🧾 BÁO CÁO BILL', color: '#FFFFFF', weight: 'bold', size: 'md' },
          { type: 'text', text: `📅 ${ngayHienThi}`, color: '#E4F0FD', size: 'xs', margin: 'sm' },
        ],
      },
      body: {
        type: 'box', layout: 'vertical', paddingAll: '16px', spacing: 'sm',
        contents: [
          { type: 'text', text: `🏢 ${cuaHang.ten}`, weight: 'bold', size: 'sm', wrap: true, margin: 'none' },
          { type: 'separator', margin: 'md' },
          dongThongTin('📋', 'Số lượng bill', fmtSo(cuaHang.soBill), false),
          dongThongTin('🧾', 'Bill online', fmtSo(cuaHang.soBillOnline), false),
          { type: 'separator', margin: 'md' },
          dongThongTin('📈', 'Giá trị bill', fmtSo(giaTriBill) + ' đ', true),
        ],
      },
    },
  };
}

async function generateDailyStoreReport() {
  const sheets = getSheetsClient();
  const rows = await docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_BAOCAO_NGAY);

  const header = rows[0];
  const colNgay = timCotTheoTen(header, 'Ngày');
  const colMaST = timCotTheoTen(header, 'Mã siêu thị');
  const colTenST = timCotTheoTen(header, 'Tên siêu thị');
  const colDTOffline = timCotTheoTen(header, 'Doanh thu offline');
  const colDTOnline = timCotTheoTen(header, 'Doanh thu Online');
  const colSoBill = timCotTheoTen(header, 'Tổng số bill');
  const colSoBillOnline = timCotTheoTen(header, 'Tổng số bill online');

  // Chỉ lấy đúng ngày mới nhất có trong tab (tránh dữ liệu ngày cũ còn sót lại)
  let ngayMoiNhat = null;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row[colNgay] === undefined || row[colNgay] === '') continue;
    const key = toDateKey(row[colNgay]);
    if (ngayMoiNhat === null || key > ngayMoiNhat) ngayMoiNhat = key;
  }

  // Gộp theo mã siêu thị (đề phòng 1 siêu thị có nhiều dòng cùng ngày mới nhất)
  const theoMaSieuThi = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row[colNgay] === undefined || row[colNgay] === '') continue;
    if (toDateKey(row[colNgay]) !== ngayMoiNhat) continue;

    const ma = row[colMaST];
    if (!theoMaSieuThi[ma]) {
      theoMaSieuThi[ma] = {
        ma,
        ten: row[colTenST] || ma,
        dtOffline: 0,
        dtOnline: 0,
        soBill: 0,
        soBillOnline: 0,
      };
    }
    const ch = theoMaSieuThi[ma];
    ch.dtOffline += Number(row[colDTOffline]) || 0;
    ch.dtOnline += Number(row[colDTOnline]) || 0;
    ch.soBill += Number(row[colSoBill]) || 0;
    ch.soBillOnline += Number(row[colSoBillOnline]) || 0;
  }

  const cuaHangs = Object.values(theoMaSieuThi).sort(
    (a, b) => (b.dtOffline + b.dtOnline) - (a.dtOffline + a.dtOnline)
  );

  if (cuaHangs.length === 0) {
    throw new Error(`Tab "${GOOGLE_SHEET_TAB_BAOCAO_NGAY}" chưa có dữ liệu siêu thị nào`);
  }

  const ngayHienThi = fmtNgayVN(new Date());

  // Mỗi siêu thị 1 thẻ Doanh thu + 1 thẻ Bill, gửi thành tin nhắn RIÊNG (không
  // gộp carousel phải vuốt) - LINE giới hạn tối đa 5 tin nhắn / lần reply nên
  // cắt bớt nếu quá nhiều siêu thị.
  const messages = [];
  cuaHangs.forEach((ch) => {
    messages.push(taoTheDoanhThuNgay(ch, ngayHienThi));
    messages.push(taoTheBillNgay(ch, ngayHienThi));
  });

  return messages.slice(0, SO_CAROUSEL_TOI_DA);
}

async function generateRevenueReport() {
  const sheets = getSheetsClient();

  const [rowsT7DT, rowsT8DT, rowsT7SL, rowsT8SL] = await Promise.all([
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_T7_DOANHTHU),
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_T8_DOANHTHU),
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_T7_SANLUONG),
    docTabThanhMangDong(sheets, GOOGLE_SHEET_TAB_T8_SANLUONG),
  ]);

  const t7DT = gomTheoNganhHang(rowsT7DT, 'Doanh thu');
  const t8DT = gomTheoNganhHang(rowsT8DT, 'Doanh thu');
  const t7SL = gomTheoNganhHang(rowsT7SL, 'Sản lượng bán');
  const t8SL = gomTheoNganhHang(rowsT8SL, 'Sản lượng bán');

  const { card1, card2, altText } = taoFlexBaoCaoDoanhThu(t7DT, t8DT);
  const card3 = taoFlexSanLuong(t7SL, t8SL);

  return [
    { type: 'flex', altText: altText.slice(0, 400), contents: card1 },
    { type: 'flex', altText: 'Chi tiết ngành hàng dự kiến T8', contents: card2 },
    { type: 'flex', altText: 'Sản lượng bán dự kiến T8', contents: card3 },
  ];
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
// NẠP FILE NGƯỜI DÙNG GỬI TRỰC TIẾP VÀO GROUP (.xlsx/.xls)
// - Tự nhận diện loại báo cáo qua tiêu đề cột, nối thêm data vào đúng tab,
//   rồi tự trả báo cáo tương ứng.
// ---------------------------------------------------------------------------

// Tải nội dung file người dùng gửi trong LINE về dạng Buffer
async function taiNoiDungFileLine(messageId) {
  const stream = await client.getMessageContent(messageId);
  const chunks = [];
  for await (const chunk of stream) chunks.push(chunk);
  return Buffer.concat(chunks);
}

// Xác định tháng (7 hoặc 8) dựa vào cột "Ngày xuất" xuất hiện nhiều nhất trong file
function xacDinhThangDuLieu(header, dataRows) {
  const idx = header.indexOf('Ngày xuất');
  const dem = {};
  for (const row of dataRows) {
    const v = row[idx];
    if (v === undefined || v === null || v === '') continue;
    const key = toDateKey(v);
    const thang = parseInt(key.slice(5, 7), 10);
    if (!Number.isNaN(thang)) dem[thang] = (dem[thang] || 0) + 1;
  }
  const phoBien = Object.entries(dem).sort((a, b) => b[1] - a[1])[0];
  if (!phoBien) throw new Error('Không xác định được tháng của dữ liệu (cột "Ngày xuất" trống)');
  const thang = parseInt(phoBien[0], 10);
  if (thang !== 7 && thang !== 8) {
    throw new Error(`Hiện chỉ hỗ trợ dữ liệu tháng 7 hoặc tháng 8, file này là tháng ${thang}`);
  }
  return thang;
}

// Nhận diện loại file dựa theo tiêu đề cột -> trả về { loai, tenTab } hoặc null nếu không nhận ra
function nhanDangLoaiFile(header, dataRows) {
  const co = (ten) => header.includes(ten);

  if (co('Mã Model') && co('Tồn kho siêu thị')) {
    return { loai: 'ton', tenTab: GOOGLE_SHEET_TAB_TON };
  }
  if (co('Mã Model') && co('Tổng số lượng') && !co('Tồn kho siêu thị')) {
    return { loai: 'tra_ban', tenTab: GOOGLE_SHEET_TAB_DOANHTHU };
  }
  if (co('Ngày') && co('Mã siêu thị') && co('Doanh thu offline')) {
    return { loai: 'ngay', tenTab: GOOGLE_SHEET_TAB_BAOCAO_NGAY };
  }
  if (co('Ngày xuất') && co('Ngành hàng BHX') && co('Sản lượng bán')) {
    const thang = xacDinhThangDuLieu(header, dataRows);
    return { loai: 'sanluong', tenTab: thang === 7 ? GOOGLE_SHEET_TAB_T7_SANLUONG : GOOGLE_SHEET_TAB_T8_SANLUONG };
  }
  if (co('Ngày xuất') && co('Ngành hàng BHX') && co('Doanh thu') && !co('Sản lượng bán')) {
    const thang = xacDinhThangDuLieu(header, dataRows);
    return { loai: 'doanhthu', tenTab: thang === 7 ? GOOGLE_SHEET_TAB_T7_DOANHTHU : GOOGLE_SHEET_TAB_T8_DOANHTHU };
  }
  return null;
}

// Nạp data vào đúng tab (nối thêm, giữ nguyên data cũ), trả về { loai, tenTab, soDong }
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
  const destRows = await docTabThanhMangDong(sheets, nhanDang.tenTab);
  const destHeader = destRows[0];

  // Sắp lại cột theo đúng thứ tự cột của tab đích (khớp theo TÊN cột, không
  // phụ thuộc thứ tự cột trong file người dùng gửi)
  const rowsToAppend = dataRows.map((row) =>
    destHeader.map((tenCot) => {
      const idx = header.indexOf(tenCot);
      return idx === -1 ? '' : row[idx] ?? '';
    })
  );

  await sheets.spreadsheets.values.append({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: nhanDang.tenTab,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: rowsToAppend },
  });

  return { loai: nhanDang.loai, tenTab: nhanDang.tenTab, soDong: rowsToAppend.length };
}

// ---------------------------------------------------------------------------
// LINE BOT
// ---------------------------------------------------------------------------
const app = express();
const client = new line.Client(config);

// Trả về 'tra' nếu nhắn "báo cáo trà", 'ngay' nếu nhắn "báo cáo ngày",
// 'doanhthu' nếu nhắn "báo cáo"/"báo cáo dt" chung chung
// (kiểm tra các từ khoá cụ thể trước để không bị "báo cáo" tổng quát nuốt mất)
function loaiTrigger(text) {
  if (!text) return null;
  const t = text.trim().toLowerCase();
  if (TRIGGER_KEYWORDS.some((kw) => t === kw || t.includes(kw))) return 'tra';
  if (t.includes('báo cáo ngày') || t.includes('bao cao ngay')) return 'ngay';
  if (t === 'báo cáo' || t === 'bao cao' || t.includes('báo cáo') || t.includes('bao cao')) return 'doanhthu';
  return null;
}

app.post('/webhook', line.middleware(config), async (req, res) => {
  // Trả 200 ngay để LINE không retry, xử lý reply ở dưới (không chặn response)
  res.status(200).end();

  const events = req.body.events || [];
  console.log(`[webhook] nhận ${events.length} event(s)`);

  for (const event of events) {
    console.log(`[webhook] event type=${event.type} messageType=${event.message?.type} text="${event.message?.text}"`);

    if (event.type !== 'message') continue;

    // ---- Người dùng gửi FILE (.xlsx/.xls) vào group -> tự nạp + tự báo cáo ----
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
        console.log(`[webhook] đã nạp ${ketQua.soDong} dòng vào tab "${ketQua.tenTab}" (loại: ${ketQua.loai})`);

        // Nạp xong tự trả báo cáo tương ứng loại file đó
        try {
          let baoCao;
          if (ketQua.loai === 'ton' || ketQua.loai === 'tra_ban') baoCao = await generateTraReport();
          else if (ketQua.loai === 'ngay') baoCao = await generateDailyStoreReport();
          else baoCao = await generateRevenueReport();
          await client.replyMessage(event.replyToken, baoCao);
        } catch (loiBaoCao) {
          // Nạp file đã thành công, chỉ là chưa đủ data để lên báo cáo (vd thiếu tab kia)
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
    const loai = loaiTrigger(event.message.text);
    if (!loai) {
      console.log('[webhook] text không khớp từ khoá trigger, bỏ qua');
      continue;
    }

    console.log(`[webhook] khớp trigger loại="${loai}", đang tạo báo cáo...`);
    try {
      if (loai === 'tra') {
        const flexMessage = await generateTraReport();
        await client.replyMessage(event.replyToken, flexMessage);
      } else if (loai === 'ngay') {
        const flexMessages = await generateDailyStoreReport();
        await client.replyMessage(event.replyToken, flexMessages);
      } else {
        const flexMessages = await generateRevenueReport();
        await client.replyMessage(event.replyToken, flexMessages);
      }
      console.log('[webhook] tạo báo cáo + reply thành công');
    } catch (err) {
      console.error('[webhook] Lỗi tạo báo cáo:', err);
      try {
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: `⚠️ Không tạo được báo cáo: ${err.message}`,
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
