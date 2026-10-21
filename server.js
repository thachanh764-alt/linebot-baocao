/**
 * server.js — LINE Bot linebot-baocao-bhx
 * ========================================
 * ĐÃ XOÁ TOÀN BỘ báo cáo cũ (Báo cáo ngày, Luỹ Kế DT/DT Dự Kiến, Giá Vốn, Huỷ MMKK,
 * SP 1 Đồng, Lượt Bill, Bánh Trung Thu). Bot hiện chỉ còn:
 *
 * 1) Lệnh "DT NGÀNH HÀNG" (trong group PHẢI @tag bot):
 *    - So DT thực tế luỹ kế trong tuần thi đua với MỤC TIÊU theo mức (M2 mặc định, T3 = M3; đổi bằng "ĐẶT MỨC T3 M3").
 *    - 3 cột: MT M2 · DT hiện tại · % đạt — xếp từ cao xuống thấp.
 *    - Gõ thêm ngày để xem tới ngày đó:   "DT NGÀNH HÀNG 07/10"
 *    - Gõ thêm tuần để chọn tuần:         "DT NGÀNH HÀNG T2"
 *
 * 2) Gửi file Excel vào group, bot tự nhận diện:
 *    - File THI ĐUA (vd THI_DUA_FMCG_T10_...xlsx): đọc các tab "TUẦN n · dd/mm–dd/mm ... (NH ...)"
 *      lấy cột M2 của từng siêu thị -> GHI ĐÈ tab MUCTIEU_NGANHHANG.
 *    - File "Doanh Thu Theo Model": gộp DT theo Siêu thị + Ngành hàng -> lưu tab DT_NGANHHANG_NGAY
 *      theo NGÀY (ngày lấy từ tên file _YYYYMMDD_, không có thì lấy hôm nay). Gửi lại cùng ngày
 *      thì chỉ ghi đè đúng ngày + ngành hàng đó, các ngày khác giữ nguyên để cộng luỹ kế tuần.
 *      Nạp xong bot tự trả báo cáo "DT NGÀNH HÀNG".
 *
 * 3) AI (giữ nguyên): @tag bot + câu hỏi tự do -> AI đọc tab liên quan trả lời;
 *    @tag hỏi rồi gửi file/ảnh trong 3 phút -> AI phân tích nhanh, KHÔNG lưu Sheet.
 *
 * Biến môi trường: LINE_CHANNEL_ACCESS_TOKEN, LINE_CHANNEL_SECRET, GOOGLE_SERVICE_ACCOUNT_JSON
 * (hoặc GOOGLE_SERVICE_ACCOUNT_KEY_PATH), GOOGLE_SHEET_ID, ANTHROPIC_API_KEY, PORT.
 * 2 tab MUCTIEU_NGANHHANG và DT_NGANHHANG_NGAY bot TỰ TẠO nếu chưa có.
 */

require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const path = require('path');
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

const TAB_MUCTIEU = process.env.GOOGLE_SHEET_TAB_MUCTIEU_NGANHHANG || 'MUCTIEU_NGANHHANG';
const TAB_DT_NGAY = process.env.GOOGLE_SHEET_TAB_DT_NGANHHANG_NGAY || 'DT_NGANHHANG_NGAY';

const HEADER_MUCTIEU = ['Tuần', 'Từ ngày', 'Đến ngày', 'Tên thi đua', 'Mã ngành hàng', 'Mã siêu thị', 'Tên siêu thị', 'MT M1 (triệu)', 'MT M2 (triệu)', 'MT M3 (triệu)'];
const TAB_CAUHINH = process.env.GOOGLE_SHEET_TAB_CAUHINH || 'CAUHINH';
// Mức thi đua mặc định từng tuần (đổi được bằng lệnh "ĐẶT MỨC T3 M3"). Tuần không ghi = M2.
const MUC_MAC_DINH = { T3: 3 };
// Tuần thi đua ĐÃ TẮT (không báo cáo, không nạp file DT của ngành tuần đó). Mở lại: xoá tuần khỏi danh sách.
const TUAN_TAT = new Set(['T2']); // T2 = Đông mát – Sữa – Kem
const NH_TAT = new Set(['990', '1354', '1355', '1056']); // ngành của tuần đã tắt
// Tab tuần không ghi "(NH ...)" trong tiêu đề -> đoán mã ngành theo tên
const NH_THEO_TEN = [
  [/HÓA PHẨM|HOÁ PHẨM|CHĂM SÓC NHÀ CỬA/i, '1014'],
  [/MỸ PHẨM|CHĂM SÓC CÁ NHÂN/i, '925'],
  [/NHÀ BẾP/i, '1055'],
  [/BÁNH KẸO/i, '1054'],
];
const HEADER_DT_NGAY = ['Ngày', 'Mã siêu thị', 'Tên siêu thị', 'Mã ngành hàng', 'Ngành hàng', 'Doanh thu', 'Giờ xuất'];


const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
// Muốn tiết kiệm chi phí thì đổi thành: 'claude-haiku-4-5-20251001'
const AI_MODEL = process.env.AI_MODEL || 'claude-sonnet-5-5';
const AI_ENABLED = !!process.env.ANTHROPIC_API_KEY;

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

// Đọc tab, nếu chưa có thì trả mảng rỗng (không báo lỗi)
async function docTabAnToan(sheets, tenTab) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: tenTab,
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    return res.data.values || [];
  } catch (err) {
    if (/Unable to parse range|not found/i.test(err.message || '')) return [];
    throw err;
  }
}

// Tự tạo tab nếu chưa có
async function damBaoCoTab(sheets, tenTab) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID, fields: 'sheets.properties.title' });
  const daCo = (meta.data.sheets || []).some((s) => s.properties.title === tenTab);
  if (!daCo) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: tenTab } } }] },
    });
    console.log(`[sheet] đã tự tạo tab mới "${tenTab}"`);
  }
}

// Ghi đè toàn bộ tab (xoá sạch rồi ghi header + data). Ghi RAW để ngày giữ dạng chữ YYYY-MM-DD.
async function ghiDeTab(sheets, tenTab, header, rows) {
  await damBaoCoTab(sheets, tenTab);
  await sheets.spreadsheets.values.clear({ spreadsheetId: GOOGLE_SHEET_ID, range: tenTab });
  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${tenTab}!A1`,
    valueInputOption: 'RAW',
    requestBody: { values: [header, ...rows] },
  });
}

// ---------------------------------------------------------------------------
// TIỆN ÍCH NGÀY / SỐ
// ---------------------------------------------------------------------------
function homNayVN() {
  return new Date(Date.now() + 7 * 3600 * 1000).toISOString().slice(0, 10); // YYYY-MM-DD
}

function taoKeyNgay(y, m, d) {
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

function chuanHoaKeyNgay(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') {
    // serial ngày của Google Sheet/Excel
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return taoKeyNgay(m[1], m[2], m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return taoKeyNgay(m[3], m[2], m[1]);
  return null;
}

function soNgayGiua(keyA, keyB) {
  return Math.round((Date.parse(keyB) - Date.parse(keyA)) / 86400000);
}

function fmtNgay(key) {
  if (!key) return '';
  const [, m, d] = key.split('-');
  return `${d}/${m}`;
}

// 29.007 -> "29,0" ; 1234.5 -> "1.234,5"
function fmtTrieu(n) {
  const r = Math.round((Number(n) || 0) * 10) / 10;
  const [nguyen, le] = r.toFixed(1).split('.');
  return `${nguyen.replace(/\B(?=(\d{3})+(?!\d))/g, '.')},${le}`;
}

function fmtPhanTram(p) {
  return `${(Math.round(p * 10) / 10).toFixed(1).replace('.', ',')}%`;
}

function chuanHoaMaST(v) {
  if (v === null || v === undefined) return '';
  const s = String(v).trim();
  const idx = s.indexOf(' - ');
  return (idx === -1 ? s : s.slice(0, idx)).trim();
}

function tenNganSieuThi(tenDayDu) {
  if (!tenDayDu) return '';
  const s = String(tenDayDu);
  const idx = s.indexOf(' - ');
  return idx === -1 ? s.trim() : s.slice(idx + 3).trim();
}

function rutGonTen(ten, maxLen) {
  if (!ten) return '';
  return ten.length > maxLen ? ten.slice(0, maxLen - 1).trim() + '…' : ten;
}

// ---------------------------------------------------------------------------
// ĐỌC FILE EXCEL
// ---------------------------------------------------------------------------
function docSheetThanhMang(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
}

// ---- File THI ĐUA: tìm các tab có tiêu đề "TUẦN n · dd/mm–dd/mm · ... (NH 1354, 1355, ...)" ----
function docFileThiDua(wb, fileName) {
  const namHienTai = Number(homNayVN().slice(0, 4));
  const ketQua = []; // các dòng MUCTIEU
  const tuanDaNap = [];
  const tuanBoQua = [];

  for (const tenSheet of wb.SheetNames) {
    const rows = docSheetThanhMang(wb.Sheets[tenSheet]);
    const tieuDe = String((rows[0] || [])[0] || '').trim();
    const mTuan = tieuDe.match(/TU[ẦA]N\s*(\d+)/i);
    const mNgay = tieuDe.match(/(\d{1,2})\/(\d{1,2})\s*[–\-—]\s*(\d{1,2})\/(\d{1,2})/);
    if (!mTuan || !mNgay) continue;

    const tuan = `T${mTuan[1]}`;
    const mNH = tieuDe.match(/\(NH\s*([\d,\s]+)/i);
    let maNH = mNH ? mNH[1].split(/[,\s]+/).filter(Boolean).join(',') : '';
    if (!maNH) maNH = [...new Set(NH_THEO_TEN.filter(([re]) => re.test(tieuDe)).map(([, ma]) => ma))].join(',');
    if (!maNH) {
      tuanBoQua.push(`${tuan} (${tenSheet}) – không có mã ngành hàng trong tiêu đề`);
      continue;
    }

    // tìm dòng tiêu đề cột có "Siêu thị" và cột bắt đầu bằng "M2"
    let dongHeader = -1;
    let cotST = -1;
    let cotM2 = -1, cotM1 = -1, cotM3 = -1;
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
      const r = (rows[i] || []).map((c) => String(c || '').trim());
      const iST = r.findIndex((c) => c === 'Siêu thị');
      const iM2 = r.findIndex((c) => /^M2\b/.test(c));
      if (iST !== -1 && iM2 !== -1) {
        dongHeader = i; cotST = iST; cotM2 = iM2;
        cotM1 = r.findIndex((c) => /^M1\b/.test(c));
        cotM3 = r.findIndex((c) => /^M3\b/.test(c));
        break;
      }
    }
    if (dongHeader === -1) {
      tuanBoQua.push(`${tuan} (${tenSheet}) – không có cột M2`);
      continue;
    }

    const thangDau = Number(mNgay[2]);
    const thangCuoi = Number(mNgay[4]);
    const tuNgay = taoKeyNgay(namHienTai, thangDau, mNgay[1]);
    const denNgay = taoKeyNgay(thangCuoi < thangDau ? namHienTai + 1 : namHienTai, thangCuoi, mNgay[3]);
    const tenThiDua = tieuDe.split('·').slice(2).join('·').trim() || tieuDe;

    let soST = 0;
    for (let i = dongHeader + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const st = String(r[cotST] || '').trim();
      if (!st || /^TỔNG/i.test(st)) continue;
      const so = (c) => { const v = c === -1 ? NaN : Number(r[c]); return !v || isNaN(v) ? '' : Math.round(v * 10000) / 10000; };
      const m1 = so(cotM1), m2 = so(cotM2), m3 = so(cotM3);
      if (m2 === '') continue; // ST chưa có mục tiêu -> bỏ
      ketQua.push([tuan, tuNgay, denNgay, tenThiDua, maNH, chuanHoaMaST(st), tenNganSieuThi(st), m1, m2, m3]);
      soST++;
    }
    tuanDaNap.push(`${tuan} ${fmtNgay(tuNgay)}–${fmtNgay(denNgay)} · NH ${maNH} · ${soST} ST`);
  }

  if (ketQua.length === 0) return null;
  return { rows: ketQua, tuanDaNap, tuanBoQua };
}

// ---- File "Doanh Thu Theo Model": gộp DT theo Siêu thị + Ngành hàng ----
function laFileDoanhThuModel(header) {
  const co = (t) => header.includes(t);
  return co('Mã Model') && co('Mã siêu thị') && co('Mã ngành hàng') && co('Tổng doanh thu');
}

function ngayTuTenFile(fileName) {
  const m = String(fileName || '').match(/(20\d{2})(\d{2})(\d{2})/);
  if (m) {
    const key = taoKeyNgay(m[1], m[2], m[3]);
    if (!isNaN(Date.parse(key))) return key;
  }
  return homNayVN();
}

// Giờ xuất file (số giờ thập phân, vd 16.5) lấy từ tên file _YYYYMMDD_HHMMSS
function gioTuTenFile(fileName) {
  const m = String(fileName || '').match(/20\d{6}_(\d{2})(\d{2})\d{2}/);
  return m ? Number(m[1]) + Number(m[2]) / 60 : null;
}

function gopDoanhThuModel(header, dataRows, ngayKey, gioXuat) {
  const iMaST = header.indexOf('Mã siêu thị');
  const iTenST = header.indexOf('Tên siêu thị');
  const iMaNH = header.indexOf('Mã ngành hàng');
  const iTenNH = header.indexOf('Ngành hàng');
  const iDT = header.indexOf('Tổng doanh thu');

  const gop = new Map();
  for (const r of dataRows) {
    const maST = chuanHoaMaST(r[iMaST]);
    const maNH = String(r[iMaNH] ?? '').trim();
    if (!maST || !maNH) continue;
    const key = `${maST}|${maNH}`;
    if (!gop.has(key)) {
      gop.set(key, [ngayKey, maST, tenNganSieuThi(r[iTenST]), maNH, String(r[iTenNH] || '').trim(), 0, gioXuat ?? '']);
    }
    gop.get(key)[5] += Number(r[iDT]) || 0;
  }
  return [...gop.values()].map((d) => { d[5] = Math.round(d[5]); return d; });
}

// ---- Nạp file vào Sheet ----
async function napFileVaoSheet(fileName, buffer, ngayChiDinh, laBaseDmsk, laBaseBnnb, laBaseThst) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sheets = getSheetsClient();

  // 1) File thi đua (nhiều tab tuần)
  const thiDua = docFileThiDua(wb, fileName);
  if (thiDua) {
    await ghiDeTab(sheets, TAB_MUCTIEU, HEADER_MUCTIEU, thiDua.rows);
    return { loai: 'muctieu', tenTab: TAB_MUCTIEU, soDong: thiDua.rows.length, chiTiet: thiDua };
  }

  // 2) File Doanh Thu Theo Model
  const allRows = docSheetThanhMang(wb.Sheets[wb.SheetNames[0]]);
  if (allRows.length < 2) throw new Error('File trống hoặc không đọc được dữ liệu');
  const header = allRows[0].map((h) => (h || '').toString().trim());
  const dataRows = allRows.slice(1).filter((r) => r && r.some((v) => v !== null && v !== ''));

  if (laFileDtSieuThi(header)) return napFileDtSieuThi(sheets, header, dataRows, fileName);
  if (laFileGiaVon(header)) return napFileGiaVon(sheets, header, dataRows, fileName);
  if (laFileFreshBan(header)) return napFileFreshBan(sheets, header, dataRows, fileName);
  if (laFileHuyMmkk(header)) return napFileHuyMmkk(sheets, header, dataRows);
  if (laFileTonModel(header)) return napFileTonSua(sheets, header, dataRows, fileName);
  if (laFileDoanhThuModel(header) && (await laFileBanSua(sheets, header, dataRows))) return napFileBanSua(sheets, header, dataRows, fileName);

  // File Thuỷ hải sản / Trứng của 1 siêu thị -> Báo cáo THS + Trứng
  if (laFileDoanhThuModel(header) && laFileThst(header, dataRows)) {
    return laBaseThst ? napBaseThst(sheets, header, dataRows) : napT10Thst(sheets, header, dataRows, fileName, ngayChiDinh);
  }
  // File Bia nước / Gia vị của 1 siêu thị -> Báo cáo BNNB
  if (laFileDoanhThuModel(header) && laFileBnnb(header, dataRows)) {
    return laBaseBnnb ? napBaseBnnb(sheets, header, dataRows) : napT10Bnnb(sheets, header, dataRows, fileName, ngayChiDinh);
  }

  if (laFileDoanhThuModel(header)) {
    // File chỉ gồm ngành của tuần đã tắt (Đông mát – Sữa – Kem) -> không nạp
    const iNH = header.indexOf('Mã ngành hàng');
    const nhFile = new Set(dataRows.map((r) => String(r[iNH] ?? '').trim()).filter(Boolean));
    if (nhFile.size && [...nhFile].every((nh) => NH_TAT.has(nh))) {
      // File Đông mát – Sữa – Kem -> BC ĐMSK (base T9 nếu vừa gõ "NẠP BASE ĐMSK", ngược lại là luỹ kế T10)
      return laBaseDmsk ? napBaseDmsk(sheets, header, dataRows, fileName) : napT10Dmsk(sheets, header, dataRows, fileName, ngayChiDinh);
    }
    const ngayKey = ngayChiDinh || ngayTuTenFile(fileName);
    // Gửi bù bằng lệnh NẠP NGÀY = file đủ cả ngày -> coi như xuất lúc 23h
    const gioXuat = ngayChiDinh ? 23 : gioTuTenFile(fileName);
    const moi = gopDoanhThuModel(header, dataRows, ngayKey, gioXuat);
    if (moi.length === 0) throw new Error('File không có dòng doanh thu nào');

    // Giữ các ngày khác; chỉ thay đúng (ngày + ngành hàng) có trong file mới
    const nhMoi = new Set(moi.map((d) => d[3]));
    const cu = await docTabAnToan(sheets, TAB_DT_NGAY);
    const giuLai = cu.slice(1).filter((r) => {
      const ngay = chuanHoaKeyNgay(r[0]);
      return !(ngay === ngayKey && nhMoi.has(String(r[3]).trim()));
    });
    await ghiDeTab(sheets, TAB_DT_NGAY, HEADER_DT_NGAY, [...giuLai, ...moi]);
    return { loai: 'dt_nganhhang', tenTab: TAB_DT_NGAY, soDong: moi.length, ngay: ngayKey };
  }

  throw new Error(
    `Không nhận diện được file "${fileName || ''}". Bot hiện nhận file THI ĐUA, file "Doanh Thu Theo Model", file "Báo cáo giá vốn Fresh", file "BC nhập xuất Fresh" và file "Báo cáo theo dõi bán, huỷ, MMKK".`
  );
}

// ---------------------------------------------------------------------------
// BÁO CÁO "DT NGÀNH HÀNG"
// ---------------------------------------------------------------------------
const TRIGGER_DT_NH = ['dt ngành hàng', 'dt nganh hang', 'doanh thu ngành hàng', 'doanh thu nganh hang'];
// Tìm tên lệnh ở BẤT KỲ đâu trong câu (vd "Cho anh số HUỶ MMKK đầu tháng tới nay")
function chuanCau(text) {
  return ' ' + (text || '').normalize('NFC').toLowerCase().replace(/[.,!?:;"'()]/g, ' ').replace(/\s+/g, ' ').trim() + ' ';
}
function coTuKhoa(text, dsTu) {
  const t = chuanCau(text);
  return dsTu.some((k) => t.includes(' ' + k + ' ') || t.includes(' ' + k));
}
function laTriggerDtNganhHang(text) {
  return coTuKhoa(text, TRIGGER_DT_NH);
}

function phanTichThamSo(text) {
  const t = (text || '').normalize('NFC');
  const mTuan = t.match(/\b(?:T|tu[ầa]n\s*)(\d)\b/i);
  const mNgay = t.match(/\b(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?\b/);
  let ngay = null;
  if (mNgay) {
    const nam = mNgay[3] || homNayVN().slice(0, 4);
    ngay = taoKeyNgay(nam, mNgay[2], mNgay[1]);
  }
  return { tuan: mTuan ? `T${mTuan[1]}` : null, ngay };
}

function mauPhanTram(p, chuan = 100) {
  if (p >= chuan) return '#0B8A3E'; // đạt / đúng tiến độ
  if (p >= chuan * 0.9) return '#E08A00'; // gần đạt
  return '#D0312D'; // thiếu
}

function fmtSoPct(p) {
  return (Math.round(p * 10) / 10).toFixed(1).replace('.', ',');
}

async function tinhDuLieuDtNH(text) {
  const sheets = getSheetsClient();
  const rowsMT = (await docTabAnToan(sheets, TAB_MUCTIEU)).slice(1);
  if (rowsMT.length === 0) {
    throw new Error('Chưa có mục tiêu. Anh gửi file THI ĐUA vào group trước giúp em.');
  }
  const rowsDT = (await docTabAnToan(sheets, TAB_DT_NGAY)).slice(1);
  // Tab cũ chỉ có 1 cột M2 (cột H); tab mới có M1/M2/M3 (cột H/I/J)
  const kieuMoi = rowsMT.some((r) => r.length >= 9);
  const mucDaDat = new Map((await docTabAnToan(sheets, TAB_CAUHINH)).slice(1).map((r) => [String(r[0]), Number(r[1])]));
  const ts = phanTichThamSo(text);

  // Danh sách tuần
  const dsTuan = new Map();
  for (const r of rowsMT) {
    const tuan = String(r[0]);
    if (!dsTuan.has(tuan)) {
      dsTuan.set(tuan, {
        tuan, tuNgay: chuanHoaKeyNgay(r[1]), denNgay: chuanHoaKeyNgay(r[2]), ten: String(r[3] || ''),
        maNH: new Set(String(r[4]).split(',').map((s) => s.trim()).filter(Boolean)), mucTieu: [],
      });
    }
    const m = kieuMoi ? { 1: Number(r[7]) || 0, 2: Number(r[8]) || 0, 3: Number(r[9]) || 0 } : { 1: 0, 2: Number(r[7]) || 0, 3: 0 };
    dsTuan.get(tuan).mucTieu.push({ maST: chuanHoaMaST(r[5]), ten: String(r[6] || ''), m });
  }

  for (const tt of TUAN_TAT) dsTuan.delete(tt);
  if (ts.tuan && TUAN_TAT.has(ts.tuan)) throw new Error(`Báo cáo tuần ${ts.tuan.slice(1)} đã tắt.`);
  if (dsTuan.size === 0) throw new Error('Chưa có tuần thi đua nào đang chạy.');

  const dt = rowsDT
    .map((r) => ({ ngay: chuanHoaKeyNgay(r[0]), maST: chuanHoaMaST(r[1]), ten: String(r[2] || ''), maNH: String(r[3]).trim(), dt: Number(r[5]) || 0, gio: r[6] === '' || r[6] == null ? null : Number(r[6]) }))
    .filter((d) => d.ngay);
  const ngayMoiNhat = dt.reduce((a, d) => (a && a > d.ngay ? a : d.ngay), null);

  // Chọn tuần: theo "T2" nếu gõ; theo ngày gõ; không thì theo ngành hàng + ngày của data mới nhất
  let tuanChon = null;
  if (ts.tuan) {
    tuanChon = dsTuan.get(ts.tuan);
    if (!tuanChon) throw new Error(`Không có mục tiêu tuần ${ts.tuan}. Đang có: ${[...dsTuan.keys()].join(', ')}`);
  } else {
    const ngayXet = ts.ngay || ngayMoiNhat || homNayVN();
    const nhMoiNhat = new Set(dt.filter((d) => d.ngay === ngayMoiNhat).map((d) => d.maNH));
    const tuans = [...dsTuan.values()];
    const coNH = tuans.filter((t) => [...t.maNH].some((nh) => nhMoiNhat.has(nh)));
    const ung = coNH.length ? coNH : tuans;
    // Ưu tiên tuần chứa ngày đang xét mà đã có doanh thu đúng ngành của tuần đó
    // (gửi thử file ngành tuần sau thì báo cáo vẫn giữ tuần hiện tại)
    tuanChon =
      tuans.find((t) => t.tuNgay <= ngayXet && ngayXet <= t.denNgay && dt.some((d) => t.maNH.has(d.maNH) && d.ngay >= t.tuNgay && d.ngay <= t.denNgay)) ||
      ung.find((t) => t.tuNgay <= ngayXet && ngayXet <= t.denNgay) ||
      ung.slice().sort((a, b) => Math.abs(soNgayGiua(ngayXet, a.tuNgay)) - Math.abs(soNgayGiua(ngayXet, b.tuNgay)))[0];
  }

  // Ngày tính tới
  const dtTuan = dt.filter((d) => tuanChon.maNH.has(d.maNH));
  let denNgay = ts.ngay && ts.ngay <= tuanChon.denNgay ? ts.ngay : tuanChon.denNgay;
  const ngayTrongTuan = [...new Set(dtTuan.filter((d) => d.ngay >= tuanChon.tuNgay && d.ngay <= denNgay).map((d) => d.ngay))].sort();

  let xemTruoc = false;
  let ngayTinh;
  if (ngayTrongTuan.length > 0) {
    denNgay = ngayTrongTuan[ngayTrongTuan.length - 1];
    ngayTinh = (d) => d.ngay >= tuanChon.tuNgay && d.ngay <= denNgay;
  } else {
    // Chưa có ngày nào thuộc tuần -> xem trước bằng ngày data gần nhất của đúng ngành hàng
    const ganNhat = dtTuan.reduce((a, d) => (a && a > d.ngay ? a : d.ngay), null);
    if (!ganNhat) throw new Error(`Chưa có doanh thu ngành hàng ${[...tuanChon.maNH].join(', ')}. Anh gửi file "Doanh Thu Theo Model" vào group giúp em.`);
    xemTruoc = true;
    denNgay = ganNhat;
    ngayTinh = (d) => d.ngay === ganNhat;
  }

  // Mỗi ngày bot giữ bản file gửi MỚI NHẤT của ngày đó. File có thể là:
  //   (a) LUỸ KẾ từ đầu tuần tới lúc xuất (vd 05/10→06/10), hoặc
  //   (b) CHỈ 1 NGÀY (vd 06/10→06/10).
  // Bot tự nhận biết theo tổng cả file: nếu tổng DT trong file NHỎ HƠN tổng luỹ kế các ngày trước
  // (sai số 5%) thì đó là file 1 ngày; ngược lại là file luỹ kế.
  //  - DT NGÀY  = (a) file − luỹ kế các ngày trước   | (b) chính số trong file
  //  - DT TUẦN  = luỹ kế các ngày trước + DT NGÀY
  const stCoMucTieu = new Map(tuanChon.mucTieu.map((m) => [m.maST, m]));
  const tenST = new Map();
  const tongTheoNgay = (ngay) => {
    const m = new Map();
    for (const d of dtTuan) {
      if (d.ngay !== ngay) continue;
      m.set(d.maST, (m.get(d.maST) || 0) + d.dt);
      if (!tenST.has(d.maST)) tenST.set(d.maST, d.ten);
    }
    return m;
  };
  const cacNgay = xemTruoc ? [denNgay] : ngayTrongTuan.filter((n) => n <= denNgay);
  const luyKe = new Map(); // maST -> luỹ kế tới ngày đang xét
  let dtNgayST = new Map();
  let soNgayDaTinh = 0;
  for (const ngay of cacNgay) {
    const file = tongTheoNgay(ngay);
    // Nhận biết theo TỔNG cả file (các siêu thị có trong file), không theo từng ST
    let tongFile = 0, tongTruoc = 0;
    for (const [ma, v] of file) { tongFile += v; tongTruoc += luyKe.get(ma) || 0; }
    let la1Ngay = false;
    if (tongTruoc > 0) {
      const gio = dtTuan.reduce((g, d) => (d.ngay === ngay && d.gio != null && !isNaN(d.gio) ? Math.max(g ?? 0, d.gio) : g), null);
      if (gio != null) {
        // Ước lượng phần đã bán trong ngày theo giờ xuất (6h mở cửa → 21h ≈ đủ ngày)
        const tbNgay = tongTruoc / soNgayDaTinh;
        const daBan = Math.min(1, Math.max(0.05, (gio - 6) / 15)) * tbNgay;
        la1Ngay = Math.abs(tongFile - daBan) < Math.abs(tongFile - (tongTruoc + daBan));
      } else {
        la1Ngay = tongFile < tongTruoc * 0.95;
      }
    }
    const ngayNay = new Map();
    for (const [ma, v] of file) ngayNay.set(ma, la1Ngay ? v : Math.max(0, v - (luyKe.get(ma) || 0)));
    for (const [ma, v] of ngayNay) luyKe.set(ma, (luyKe.get(ma) || 0) + v);
    dtNgayST = ngayNay;
    soNgayDaTinh++;
  }
  const dtLkST = luyKe;

  // MT NGÀY = M2 tuần ÷ 7 ; MT TUẦN = M2 cả tuần
  const soNgay = xemTruoc ? 1 : Math.min(7, soNgayGiua(tuanChon.tuNgay, denNgay) + 1);
  const chuanTuan = (soNgay / 7) * 100; // % tuần cần có tới hôm nay để kịp tiến độ
  const lamDong = (maST, ten, m2) => {
    const dtNgay = (dtNgayST.get(maST) || 0) / 1e6;
    const dtLk = (dtLkST.get(maST) || 0) / 1e6;
    const mtNgay = m2 / 7;
    return {
      ten, m2, mtNgay, dtNgay, dtLk,
      pNgay: mtNgay > 0 ? (dtNgay / mtNgay) * 100 : 0,
      pTuan: m2 > 0 ? (dtLk / m2) * 100 : 0,
    };
  };
  // Chọn mức: lệnh "ĐẶT MỨC" > mặc định theo tuần > M2. Thiếu mức đã chọn thì lùi về M2.
  const muc = mucDaDat.get(tuanChon.tuan) || MUC_MAC_DINH[tuanChon.tuan] || 2;
  const mucCo = tuanChon.mucTieu.some((mt) => mt.m[muc] > 0) ? muc : 2;
  const bang = tuanChon.mucTieu
    .map((mt) => lamDong(mt.maST, mt.ten, mt.m[mucCo] || mt.m[2]))
    .sort((a, b) => b.pTuan - a.pTuan || b.pNgay - a.pNgay);
  const dsKhongMT = [...dtLkST.keys()]
    .filter((ma) => !stCoMucTieu.has(ma))
    .map((ma) => lamDong(ma, tenST.get(ma), 0))
    .sort((a, b) => b.dtLk - a.dtLk);

  const tong = (arr, k) => arr.reduce((s, x) => s + x[k], 0);
  const tongM2 = tong(bang, 'm2');
  const tongMtNgay = tongM2 / 7;
  // DT tổng KV cộng luôn shop mới (chưa có mục tiêu); MT chỉ có ở 43 ST
  const tongDtNgay = tong(bang, 'dtNgay') + tong(dsKhongMT, 'dtNgay');
  const tongDtLk = tong(bang, 'dtLk') + tong(dsKhongMT, 'dtLk');
  const pNgayKV = tongMtNgay > 0 ? (tongDtNgay / tongMtNgay) * 100 : 0;
  const pTuanKV = tongM2 > 0 ? (tongDtLk / tongM2) * 100 : 0;
  const nhanNgay = `NGÀY ${fmtNgay(denNgay)}`;

  const tenNH = tuanChon.ten || `NH ${[...tuanChon.maNH].join(', ')}`;
  return {
    tuan: tuanChon.tuan, soTuan: tuanChon.tuan.slice(1), tuNgay: tuanChon.tuNgay, denNgayTuan: tuanChon.denNgay,
    denNgay, soNgay, chuanTuan, tenNH, nhanNgay, muc: mucCo,
    dong: [...bang.slice().sort((a, b) => b.pTuan - a.pTuan || b.pNgay - a.pNgay), ...dsKhongMT],
    kv: { dtNgay: tongDtNgay, mtNgay: tongMtNgay, pNgay: pNgayKV, dtLk: tongDtLk, m2: tongM2, pTuan: pTuanKV },
  };
}

// ---------------------------------------------------------------------------
// VẼ BÁO CÁO THÀNH ẢNH (SVG -> PNG bằng @resvg/resvg-js, font Việt đóng gói trong thư mục fonts/)
// ---------------------------------------------------------------------------
const FONT_DIR = path.join(__dirname, 'fonts');
const FONT_FAMILY = 'DejaVu Sans Condensed';
const FONT_FAMILY_SVG = "'DejaVu Sans Condensed', 'DejaVu Sans'";

function escXml(t) {
  return String(t ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function veSvgDtNH(d) {
  const W = 1000;
  const PAD = 24;
  const ROW_H = 30;
  const C = { xanh: '#0B8A3E', xanhDam: '#0B6E35', nenTieuDe: '#E6F4EC', nenChan: '#F3F7F5', vien: '#D5E3DA', chu: '#1a1a1a', phu: '#666666' };

  // Cột: x là mép phải (số) hoặc mép trái (chữ)
  const xTen = PAD + 44;
  const cot = {
    stt: PAD + 30,
    dtNgay: 470, mtNgay: 555, pNgay: 640,
    dtLk: 760, m2: 860, pTuan: W - PAD - 6,
  };
  const xChiaNgay = 372; // bắt đầu nhóm NGÀY
  const xChiaTuan = 662; // đường chia NGÀY | TUẦN

  const t = (x, y, text, o = {}) =>
    `<text x="${x}" y="${y}" font-size="${o.size || 15}" font-weight="${o.bold ? 700 : 400}" fill="${o.fill || C.chu}" text-anchor="${o.anchor || 'start'}">${escXml(text)}</text>`;

  const parts = [];
  let y = 0;

  // Header xanh
  const H_HEAD = 96;
  parts.push(`<rect x="0" y="0" width="${W}" height="${H_HEAD}" fill="${C.xanh}"/>`);
  parts.push(t(PAD, 38, `DT NGÀNH HÀNG · MỨC ${d.muc}`, { size: 26, bold: true, fill: '#FFFFFF' }));
  parts.push(t(PAD, 64, `Tuần ${d.soTuan} · ${fmtNgay(d.tuNgay)}–${fmtNgay(d.denNgayTuan)} · luỹ kế đến ${fmtNgay(d.denNgay)}`, { size: 16, fill: '#E3F5EA' }));
  parts.push(t(PAD, 86, d.tenNH, { size: 13, fill: '#E3F5EA' }));
  y = H_HEAD + 16;

  // 2 ô tổng KV
  const oW = (W - PAD * 2 - 16) / 2;
  const oH = 70;
  const veO = (x, nhan, p, chuan, dt, mt) => {
    parts.push(`<rect x="${x}" y="${y}" width="${oW}" height="${oH}" rx="10" fill="#F2F8F4"/>`);
    parts.push(t(x + 14, y + 24, nhan, { size: 14, fill: '#777777' }));
    parts.push(t(x + 14, y + 56, fmtPhanTram(p), { size: 28, bold: true, fill: mauPhanTram(p, chuan) }));
    parts.push(t(x + oW - 14, y + 56, `${fmtTrieu(dt)} / ${fmtTrieu(mt)} tr`, { size: 15, fill: '#555555', anchor: 'end' }));
  };
  veO(PAD, `KV · ${d.nhanNgay}`, d.kv.pNgay, 100, d.kv.dtNgay, d.kv.mtNgay);
  veO(PAD + oW + 16, `KV · TUẦN (ngày ${d.soNgay}/7)`, d.kv.pTuan, d.chuanTuan, d.kv.dtLk, d.kv.m2);
  y += oH + 16;

  // Tiêu đề bảng 2 tầng
  const H_TD = 56;
  parts.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${H_TD}" rx="6" fill="${C.nenTieuDe}"/>`);
  parts.push(t((xChiaNgay + xChiaTuan) / 2, y + 22, d.nhanNgay, { bold: true, fill: C.xanhDam, anchor: 'middle' }));
  parts.push(t((xChiaTuan + W - PAD) / 2, y + 22, `TUẦN ${d.soNgay}/7`, { bold: true, fill: C.xanhDam, anchor: 'middle' }));
  const yH = y + 46;
  parts.push(t(cot.stt, yH, '#', { bold: true, fill: C.xanhDam, anchor: 'end' }));
  parts.push(t(xTen, yH, 'Siêu thị', { bold: true, fill: C.xanhDam }));
  [['DT', cot.dtNgay], ['MT', cot.mtNgay], ['%', cot.pNgay], ['DT LK', cot.dtLk], ['MT', cot.m2], ['%', cot.pTuan]].forEach(([n, x]) =>
    parts.push(t(x, yH, n, { bold: true, fill: C.xanhDam, anchor: 'end' }))
  );
  const yBangTop = y;
  y += H_TD;

  // Các dòng
  d.dong.forEach((x, k) => {
    const coMT = x.m2 > 0;
    if (k % 2 === 1) parts.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW_H}" fill="${C.nenChan}"/>`);
    const yt = y + 20;
    parts.push(t(cot.stt, yt, k + 1, { fill: C.phu, anchor: 'end' }));
    parts.push(t(xTen, yt, x.ten));
    parts.push(t(cot.dtNgay, yt, fmtTrieu(x.dtNgay), { bold: true, anchor: 'end' }));
    parts.push(t(cot.mtNgay, yt, coMT ? fmtTrieu(x.mtNgay) : '–', { fill: C.phu, anchor: 'end' }));
    parts.push(t(cot.pNgay, yt, coMT ? fmtSoPct(x.pNgay) + '%' : '–', { bold: coMT, fill: coMT ? mauPhanTram(x.pNgay) : C.phu, anchor: 'end' }));
    parts.push(t(cot.dtLk, yt, fmtTrieu(x.dtLk), { bold: true, anchor: 'end' }));
    parts.push(t(cot.m2, yt, coMT ? fmtTrieu(x.m2) : '–', { fill: C.phu, anchor: 'end' }));
    parts.push(t(cot.pTuan, yt, coMT ? fmtSoPct(x.pTuan) + '%' : '–', { bold: coMT, fill: coMT ? mauPhanTram(x.pTuan, d.chuanTuan) : C.phu, anchor: 'end' }));
    y += ROW_H;
  });

  // Đường chia NGÀY | TUẦN và khung bảng
  parts.push(`<line x1="${xChiaTuan}" y1="${yBangTop + 6}" x2="${xChiaTuan}" y2="${y}" stroke="${C.vien}" stroke-width="2"/>`);
  parts.push(`<line x1="${xChiaNgay}" y1="${yBangTop + 6}" x2="${xChiaNgay}" y2="${y}" stroke="${C.vien}" stroke-width="1"/>`);
  parts.push(`<rect x="${PAD}" y="${yBangTop}" width="${W - PAD * 2}" height="${y - yBangTop}" rx="6" fill="none" stroke="${C.vien}"/>`);

  const H = y + PAD;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT_FAMILY_SVG}">` +
    `<rect width="${W}" height="${H}" fill="#FFFFFF"/>` + parts.join('') + `</svg>`;
}

function svgThanhPng(svg, rongPx) {
  const { Resvg } = require('@resvg/resvg-js');
  const fs = require('fs');
  const fontFiles = fs.existsSync(FONT_DIR)
    ? fs.readdirSync(FONT_DIR).filter((f) => /\.(ttf|otf)$/i.test(f)).map((f) => path.join(FONT_DIR, f))
    : [];
  const r = new Resvg(svg, {
    font: { fontFiles, loadSystemFonts: fontFiles.length === 0, defaultFontFamily: 'DejaVu Sans' },
    fitTo: { mode: 'width', value: rongPx },
    background: '#FFFFFF',
  });
  return r.render().asPng();
}

// Bộ nhớ tạm ảnh đã vẽ (để LINE tải về). Mất khi Render khởi động lại -> tự vẽ lại theo tham số trên link.
const khoAnh = new Map(); // id -> { png, preview, t }
function luuAnh(id, png, preview) {
  khoAnh.set(id, { png, preview, t: Date.now() });
  if (khoAnh.size > 30) khoAnh.delete(khoAnh.keys().next().value);
}

const PUBLIC_URL = (process.env.PUBLIC_URL || process.env.RENDER_EXTERNAL_URL || 'https://linebot-baocao.onrender.com').replace(/\/$/, '');

async function veAnhDtNH(text) {
  const d = await tinhDuLieuDtNH(text);
  const svg = veSvgDtNH(d);
  const png = svgThanhPng(svg, 1500);
  const preview = svgThanhPng(svg, 600);
  return { d, png, preview };
}

async function generateDtNganhHangReport(text) {
  const { d, png, preview } = await veAnhDtNH(text);
  const id = `${d.tuan}-${d.denNgay}-${Date.now().toString(36)}`;
  luuAnh(id, png, preview);
  const q = `t=${encodeURIComponent(d.tuan)}&n=${d.denNgay}`;
  return {
    type: 'image',
    originalContentUrl: `${PUBLIC_URL}/anh/${id}.png?${q}`,
    previewImageUrl: `${PUBLIC_URL}/anh/${id}.png?${q}&nho=1`,
  };
}

// ---------------------------------------------------------------------------
// BÁO CÁO "GIÁ VỐN" (Fresh) — file "Báo cáo giá vốn Fresh siêu thị" -> tab GIAVON (ghi đè)
// ---------------------------------------------------------------------------
const TAB_GIAVON = process.env.GOOGLE_SHEET_TAB_GIAVON || 'GIAVON';
const COT_GIAVON = [
  'Tháng', 'Mã siêu thị', 'Tên siêu thị', 'Mã ngành hàng', 'Ngành hàng', 'SL thực nhập hôm nay',
  'Giá vốn cơ bản hôm nay', 'DT FRESH tính giá vốn', 'Giá vốn cơ bản lũy kế đến ngày hôm qua',
  'Tỉ lệ doanh thu trên giá vốn cơ bản', 'LN lũy kế', 'LN TB 3 tháng trước', 'Chênh lệch LN so với 3 tháng trước',
];

// File "BC nhập xuất Fresh" (cột Thành tiền phải thu khách hàng) -> DT bán hôm nay theo ngành giá vốn
const TAB_FRESH_BAN = process.env.GOOGLE_SHEET_TAB_FRESH_BAN || 'FRESH_BAN';
const FRESH_SANG_GIAVON = { // Mã nhóm SP tính doanh thu -> Mã ngành hàng trong file giá vốn
  2090: 1236, 2091: 1236, // Thịt địa phương + nhập khẩu -> Thịt gia cầm gia súc
  2085: 1254, 2086: 1254, // Thủy hải sản nhập khẩu + tập trung
  2087: 2087, 2088: 2088, // Rau Đà Lạt, Rau địa phương
  2092: 2854, 2093: 2855, // Trái cây nhập khẩu -> ngoại, tập trung -> nội
  2094: 2856, // Trứng
};
function laFileFreshBan(header) {
  return header.includes('Thành tiền phải thu khách hàng (chưa VAT)') && header.includes('Mã nhóm SP tính doanh thu') && header.includes('Mã siêu thị');
}
async function napFileFreshBan(sheets, header, dataRows, fileName) {
  const i = (c) => header.indexOf(c);
  const iNgay = i('Ngày'), iST = i('Mã siêu thị'), iNhom = i('Mã nhóm SP tính doanh thu'), iTien = i('Thành tiền phải thu khách hàng (chưa VAT)');
  const gop = new Map();
  for (const r of dataRows) {
    const ma = chuanHoaMaST(r[iST]);
    const nh = FRESH_SANG_GIAVON[Number(r[iNhom])];
    if (!ma || !nh) continue;
    const ngay = chuanHoaKeyNgay(r[iNgay]) || ngayTuTenFile(fileName);
    const key = `${ngay}|${ma}|${nh}`;
    if (!gop.has(key)) gop.set(key, [ngay, ma, nh, 0]);
    gop.get(key)[3] += soGV(r[iTien]);
  }
  const rows = [...gop.values()].map((x) => [x[0], x[1], x[2], Math.round(x[3])]);
  if (rows.length === 0) throw new Error('File nhập xuất Fresh không có dòng doanh thu nào');
  const g = gioTuTenFile(fileName);
  const capNhat = g == null ? '' : `${String(Math.floor(g)).padStart(2, '0')}:${String(Math.round((g % 1) * 60)).padStart(2, '0')}`;
  await ghiDeTab(sheets, TAB_FRESH_BAN, ['Ngày', 'Mã siêu thị', 'Mã ngành giá vốn', 'DT bán', 'Cập nhật'], rows.map((r) => [...r, capNhat]));
  return { loai: 'fresh_ban', tenTab: TAB_FRESH_BAN, soDong: rows.length, dsST: [...new Set(rows.map((r) => r[1]))] };
}

function laFileGiaVon(header) {
  return header.includes('Giá vốn cơ bản hôm nay') && header.includes('DT FRESH tính giá vốn');
}

function soGV(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).replace('%', '').replace(/\s/g, '').replace(/,/g, '');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

async function napFileGiaVon(sheets, header, dataRows, fileName) {
  const idx = COT_GIAVON.map((c) => header.indexOf(c));
  const rows = dataRows
    .filter((r) => r[idx[1]] !== null && r[idx[1]] !== undefined && r[idx[1]] !== '')
    .map((r) => idx.map((i) => (i === -1 ? '' : r[i] ?? '')));
  if (rows.length === 0) throw new Error('File giá vốn không có dòng dữ liệu');
  // Cột cuối: thời điểm xuất file (lấy từ tên file) để in lên báo cáo
  const m = String(fileName || '').match(/(20\d{2})(\d{2})(\d{2})_(\d{2})(\d{2})/);
  const capNhat = m ? `${m[4]}:${m[5]} ${m[3]}/${m[2]}` : '';
  await ghiDeTab(sheets, TAB_GIAVON, [...COT_GIAVON, 'Cập nhật'], rows.map((r) => [...r, capNhat]));
  return { loai: 'giavon', tenTab: TAB_GIAVON, soDong: rows.length, dsST: [...new Set(rows.map((r) => chuanHoaMaST(r[1])))] };
}

async function docDuLieuGiaVon() {
  const rows = (await docTabAnToan(getSheetsClient(), TAB_GIAVON)).slice(1);
  if (rows.length === 0) throw new Error('Chưa có dữ liệu giá vốn. Anh gửi file "Báo cáo giá vốn Fresh" vào group trước giúp em.');
  const theoST = new Map();
  for (const r of rows) {
    const ma = chuanHoaMaST(r[1]);
    if (!ma) continue;
    if (!theoST.has(ma)) {
      theoST.set(ma, { ma, ten: tenNganSieuThi(r[2]) || ma, thang: String(r[0] || ''), capNhat: String(r[13] || ''), nganh: [] });
    }
    theoST.get(ma).nganh.push({
      ten: String(r[4] || ''),
      slNhap: soGV(r[5]), gvHomNay: soGV(r[6]), dt: soGV(r[7]), gvLk: soGV(r[8]),
      maNH: Number(r[3]), tiLe: soGV(r[9]), ln: soGV(r[10]), ln3t: soGV(r[11]), chenh: soGV(r[12]), dtHomNay: null,
    });
  }
  // Ghép DT bán hôm nay (ngày mới nhất trong tab FRESH_BAN)
  const ban = (await docTabAnToan(getSheetsClient(), TAB_FRESH_BAN)).slice(1).filter((r) => r[0] !== undefined && r[0] !== '');
  const ngayBan = ban.reduce((a, r) => { const k = chuanHoaKeyNgay(r[0]); return k && (!a || k > a) ? k : a; }, null);
  for (const r of ban) {
    if (chuanHoaKeyNgay(r[0]) !== ngayBan) continue;
    const st = theoST.get(chuanHoaMaST(r[1]));
    if (!st) continue;
    st.ngayBan = ngayBan;
    st.gioBan = String(r[4] || '');
    const ng = st.nganh.find((x) => x.maNH === Number(r[2]));
    if (ng) ng.dtHomNay = (ng.dtHomNay || 0) + soGV(r[3]);
  }
  return theoST;
}

const TEN_NGAN_GV = { 'Thịt gia cầm gia súc các loại': 'Thịt gia cầm, gia súc', 'Thủy Hải Sản Các Loại': 'Thủy hải sản', 'Trứng gia cầm các loại': 'Trứng gia cầm', 'Trái cây ngoại CL': 'Trái cây ngoại', 'Trái cây nội CL': 'Trái cây nội' };
const tenNganGV = (t) => TEN_NGAN_GV[String(t).trim()] || rutGonTen(String(t), 22);
const trGV = (v) => (Math.round((v / 1e6) * 10) / 10).toFixed(1).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d),)/g, '.');
const trGVdau = (v) => (v > 0.05e6 ? '+' : '') + trGV(v);
const mauLN = (v) => (v < 0 ? '#D0312D' : '#0B8A3E');

function veSvgGiaVon(st) {
  const W = 1100, PAD = 24, ROW = 34;
  const C = { xanh: '#0B8A3E', xanhDam: '#0B6E35', nen: '#E6F4EC', chan: '#F3F7F5', vien: '#D5E3DA', phu: '#666666', do: '#D0312D', vang: '#FFF59D' };
  const t = (x, y, s, o = {}) =>
    `<text x="${x}" y="${y}" font-size="${o.size || 15}" font-weight="${o.bold ? 700 : 400}" fill="${o.fill || '#1a1a1a'}" text-anchor="${o.anchor || 'start'}">${escXml(s)}</text>`;
  const coBan = st.nganh.some((x) => x.dtHomNay != null);
  const ds = st.nganh.slice().sort((a, b) => a.ln - b.ln);
  const sum = (k) => ds.reduce((s, x) => s + (x[k] || 0), 0);
  const tong = { dt: sum('dt'), gvLk: sum('gvLk'), ln: sum('ln'), ln3t: sum('ln3t'), chenh: sum('chenh'), gvHomNay: sum('gvHomNay'), dtHomNay: sum('dtHomNay') };
  const tiLeTong = tong.gvLk > 0 ? (tong.dt / tong.gvLk) * 100 : 0;
  const thang = st.thang.length === 6 ? `T${Number(st.thang.slice(4))}/${st.thang.slice(0, 4)}` : st.thang;
  const pct = (a, b) => (b > 0 ? (a / b) * 100 : 0);
  const fPct = (v) => v.toFixed(1).replace('.', ',') + '%';

  const p = [];
  p.push(`<rect x="0" y="0" width="${W}" height="96" fill="${C.xanh}"/>`);
  p.push(t(PAD, 38, 'GIÁ VỐN FRESH · BÁN HÔM NAY · LỢI NHUẬN', { size: 26, bold: true, fill: '#FFFFFF' }));
  p.push(t(PAD, 64, `${st.ma} · ${st.ten}`, { size: 16, fill: '#E3F5EA' }));
  const capNhat = [`Tháng ${thang}`, st.capNhat ? `giá vốn ${st.capNhat}` : '', coBan ? `bán ngày ${fmtNgay(st.ngayBan)}${st.gioBan ? ' lúc ' + st.gioBan : ''}` : 'chưa có file bán hôm nay', 'ĐVT: triệu'].filter(Boolean).join(' · ');
  p.push(t(PAD, 86, capNhat, { size: 13, fill: '#E3F5EA' }));
  let y = 112;

  const oW = (W - PAD * 2 - 36) / 4, oH = 74;
  const o = (i, nhan, gt, mau, phu) => {
    const x = PAD + i * (oW + 12);
    p.push(`<rect x="${x}" y="${y}" width="${oW}" height="${oH}" rx="10" fill="#F2F8F4"/>`);
    p.push(t(x + 12, y + 22, nhan, { size: 13, fill: '#777777' }));
    p.push(t(x + 12, y + 54, gt, { size: 25, bold: true, fill: mau || '#1a1a1a' }));
    if (phu) p.push(t(x + oW - 12, y + 54, phu, { size: 12, fill: '#777777', anchor: 'end' }));
  };
  o(0, 'Giá vốn nhập hôm nay', trGV(tong.gvHomNay));
  if (coBan) o(1, 'DT bán hôm nay', trGV(tong.dtHomNay), tong.dtHomNay >= tong.gvHomNay ? C.xanh : C.do, `= ${fPct(pct(tong.dtHomNay, tong.gvHomNay))} GV`);
  else o(1, 'DT bán hôm nay', '–', C.phu, 'gửi file nhập xuất');
  o(2, 'LN luỹ kế tháng', trGV(tong.ln), mauLN(tong.ln), `DT/GV ${fPct(tiLeTong)}`);
  o(3, 'So TB 3 tháng trước', trGVdau(tong.chenh), mauLN(tong.chenh), `TB 3T ${trGV(tong.ln3t)}`);
  y += oH + 16;

  // Bảng: nhóm HÔM NAY | nhóm LUỸ KẾ THÁNG
  const cot = { ten: PAD + 10, sl: 300, gvHn: 385, dtHn: 470, chHn: 555, tlHn: 635, dt: 735, gv: 820, tl: 905, ln: 985, ch: W - PAD - 10 };
  const xChia = 660;
  const H_TD = 72;
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${H_TD}" rx="6" fill="${C.nen}"/>`);
  p.push(t((235 + xChia) / 2, y + 20, 'HÔM NAY', { size: 13, bold: true, fill: C.xanhDam, anchor: 'middle' }));
  p.push(t((xChia + W - PAD) / 2, y + 20, 'LUỸ KẾ THÁNG', { size: 13, bold: true, fill: C.xanhDam, anchor: 'middle' }));
  const td = (x, a, b, anchor = 'end') => { p.push(t(x, y + 44, a, { size: 13, bold: true, fill: C.xanhDam, anchor })); if (b) p.push(t(x, y + 62, b, { size: 13, bold: true, fill: C.xanhDam, anchor })); };
  td(cot.ten, 'Ngành hàng', '', 'start');
  td(cot.sl, 'SL nhập', ''); td(cot.gvHn, 'GV nhập', ''); td(cot.dtHn, 'DT bán', ''); td(cot.chHn, 'Bán − GV', ''); td(cot.tlHn, 'Bán/GV', '');
  td(cot.dt, 'DT', ''); td(cot.gv, 'GV', ''); td(cot.tl, 'DT/GV', ''); td(cot.ln, 'LN', ''); td(cot.ch, 'so 3T', '');
  const yTop = y; y += H_TD;
  const dongBang = (x, k, tongDong) => {
    const nen = tongDong ? C.xanh : k % 2 === 1 ? C.chan : null;
    if (nen) p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW}" fill="${nen}"/>`);
    const yt = y + 22, wt = tongDong ? '#FFFFFF' : null;
    const co = x.dtHomNay != null;
    const chHn = co ? x.dtHomNay - x.gvHomNay : null;
    const tlHn = co ? pct(x.dtHomNay, x.gvHomNay) : null;
    p.push(t(cot.ten, yt, tongDong ? 'TỔNG FRESH' : tenNganGV(x.ten), { bold: tongDong || x.ln < 0, fill: wt || '#1a1a1a' }));
    if (!tongDong) p.push(t(cot.sl, yt, (Math.round(x.slNhap * 10) / 10).toFixed(1).replace('.', ','), { fill: C.phu, anchor: 'end' }));
    p.push(t(cot.gvHn, yt, trGV(x.gvHomNay), { fill: wt || C.phu, anchor: 'end', bold: tongDong }));
    p.push(t(cot.dtHn, yt, co ? trGV(x.dtHomNay) : '–', { fill: wt || '#1a1a1a', anchor: 'end', bold: true }));
    p.push(t(cot.chHn, yt, co ? trGVdau(chHn) : '–', { fill: wt || (co ? mauLN(chHn) : C.phu), anchor: 'end', bold: true }));
    p.push(t(cot.tlHn, yt, co && x.gvHomNay > 0 ? fPct(tlHn) : '–', { fill: wt || (co ? (tlHn >= 100 ? C.xanh : C.do) : C.phu), anchor: 'end', bold: true }));
    p.push(t(cot.dt, yt, trGV(x.dt), { anchor: 'end', fill: wt || '#1a1a1a', bold: tongDong }));
    p.push(t(cot.gv, yt, trGV(x.gvLk), { anchor: 'end', fill: wt || '#1a1a1a', bold: tongDong }));
    const tl = x.gvLk > 0 ? pct(x.dt, x.gvLk) : 0;
    p.push(t(cot.tl, yt, fPct(tl), { bold: true, fill: wt || (tl < 100 ? C.do : C.xanh), anchor: 'end' }));
    p.push(t(cot.ln, yt, trGV(x.ln), { bold: true, fill: wt || mauLN(x.ln), anchor: 'end' }));
    p.push(t(cot.ch, yt, trGVdau(x.chenh), { bold: true, fill: wt || mauLN(x.chenh), anchor: 'end' }));
    y += ROW;
  };
  ds.forEach((x, k) => dongBang(x, k, false));
  dongBang({ ...tong, dtHomNay: coBan ? tong.dtHomNay : null }, 0, true);
  p.push(`<line x1="${xChia}" y1="${yTop + 6}" x2="${xChia}" y2="${y}" stroke="${C.vien}" stroke-width="2"/>`);
  p.push(`<line x1="235" y1="${yTop + 6}" x2="235" y2="${y}" stroke="${C.vien}" stroke-width="1"/>`);
  p.push(`<rect x="${PAD}" y="${yTop}" width="${W - PAD * 2}" height="${y - yTop}" rx="6" fill="none" stroke="${C.vien}"/>`);

  // Cảnh báo + việc cần làm
  y += 26;
  const dong = [];
  if (coBan) {
    const thieu = ds.filter((x) => x.dtHomNay != null && x.dtHomNay < x.gvHomNay).sort((a, b) => (a.dtHomNay - a.gvHomNay) - (b.dtHomNay - b.gvHomNay));
    if (thieu.length) dong.push(`● Hôm nay cần bán thêm để bù giá vốn nhập: ${thieu.map((x) => `${tenNganGV(x.ten)} ${trGV(x.gvHomNay - x.dtHomNay)}`).join(' · ')} (tổng ${trGV(thieu.reduce((s, x) => s + x.gvHomNay - x.dtHomNay, 0))} tr)`);
    else dong.push('✔ Hôm nay tất cả ngành đã bán vượt giá vốn nhập.');
  }
  const lo = ds.filter((x) => x.ln < 0);
  const giam = ds.filter((x) => x.chenh < -1e6).sort((a, b) => a.chenh - b.chenh);
  if (lo.length) dong.push(`● Đang LỖ luỹ kế tháng: ${lo.map((x) => `${tenNganGV(x.ten)} (${trGV(x.ln)})`).join(' · ')}`);
  if (giam.length) dong.push(`▼ Giảm mạnh so TB 3 tháng: ${giam.map((x) => `${tenNganGV(x.ten)} (${trGVdau(x.chenh)})`).join(' · ')}`);
  const wrap = (s, n) => { const out = []; let cur = ''; for (const w2 of s.split(' ')) { if ((cur + ' ' + w2).trim().length > n) { out.push(cur.trim()); cur = w2; } else cur += ' ' + w2; } if (cur.trim()) out.push(cur.trim()); return out; };
  const mauDong = (d) => (d.startsWith('✔') ? C.xanh : d.includes('LỖ') || d.startsWith('▼') ? C.do : '#E08A00');
  for (const d of dong) for (const [i, l] of wrap(d, 135).entries()) {
    if (i === 0) { p.push(t(PAD, y, l.charAt(0), { size: 15, fill: mauDong(d), bold: true })); p.push(t(PAD + 20, y, l.slice(2), { size: 14, bold: d.includes('cần bán thêm') })); }
    else p.push(t(PAD + 20, y, l, { size: 14, bold: d.includes('cần bán thêm') }));
    y += 22;
  }
  const H = y + PAD - 6;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT_FAMILY_SVG}"><rect width="${W}" height="${H}" fill="#FFFFFF"/>${p.join('')}</svg>`;
}

const TRIGGER_GIAVON = ['giá vốn', 'gia von'];
function laTriggerGiaVon(text) {
  return coTuKhoa(text, TRIGGER_GIAVON);
}

async function veAnhGiaVon(maST) {
  const theoST = await docDuLieuGiaVon();
  const st = theoST.get(String(maST)) || [...theoST.values()][0];
  const svg = veSvgGiaVon(st);
  return { st, png: svgThanhPng(svg, 1500), preview: svgThanhPng(svg, 600) };
}

// Trả tối đa 5 ảnh (giới hạn của LINE). Gõ "Giá Vốn 8104" để xem riêng 1 siêu thị.
async function generateGiaVonReport(text, chiCacST) {
  const theoST = await docDuLieuGiaVon();
  const m = String(text || '').match(/\b(\d{3,6})\b/);
  let ds = [...theoST.keys()];
  if (m && theoST.has(m[1])) ds = [m[1]];
  else if (chiCacST && chiCacST.length) ds = ds.filter((ma) => chiCacST.includes(ma));
  if (ds.length === 0) throw new Error('Không tìm thấy siêu thị trong dữ liệu giá vốn.');
  const out = [];
  for (const ma of ds.slice(0, 5)) {
    const { png, preview } = await veAnhGiaVon(ma);
    const id = `gv-${ma}-${Date.now().toString(36)}`;
    luuAnh(id, png, preview);
    const q = `loai=gv&st=${ma}`;
    out.push({ type: 'image', originalContentUrl: `${PUBLIC_URL}/anh/${id}.png?${q}`, previewImageUrl: `${PUBLIC_URL}/anh/${id}.png?${q}&nho=1` });
  }
  return out.length === 1 ? out[0] : out;
}

// ---------------------------------------------------------------------------
// BÁO CÁO "HUỶ MMKK" — file "Báo cáo theo dõi số lượng bán, huỷ, hao hụt, mất mát kiểm kê"
// Chỉ lấy 4 ngành Fresh: Rau củ, Thịt, Trái cây, Thủy hải sản. Ghi đè tab HUYMMKK.
// Thành tiền huỷ/MMKK trong file thường = 0 -> ước tính = SL × giá bán gốc TB của chính mã đó
// (giá gốc = (Doanh thu + Tiền giảm giá) ÷ SL bán).
// ---------------------------------------------------------------------------
const TAB_HUYMMKK = process.env.GOOGLE_SHEET_TAB_HUYMMKK || 'HUYMMKK';
const NGANH_MMKK = ['Thịt gia cầm gia súc các loại', 'Rau Củ Các Loại', 'Trái Cây Các Loại', 'Thủy Hải Sản Các Loại'];
const TEN_NGAN_MMKK = { 'Thịt gia cầm gia súc các loại': 'Thịt', 'Rau Củ Các Loại': 'Rau củ', 'Trái Cây Các Loại': 'Trái cây', 'Thủy Hải Sản Các Loại': 'Thủy hải sản' };
const HEADER_HUYMMKK = ['Ngày', 'Mã siêu thị', 'Tên siêu thị', 'Ngành hàng', 'Nhóm hàng', 'Sản phẩm', 'Đơn vị', 'SL bán', 'Doanh thu', 'Tiền giảm giá', 'SL giảm giá', 'SL huỷ', 'SL MMKK', 'TT huỷ (file)', 'TT MMKK (file)'];

function laFileHuyMmkk(header) {
  return header.includes('SL mất mát kiểm kê') && header.includes('SL hủy tồn') && header.includes('Tiền s.thị bán giảm giá(chưa VAT)') && header.includes('Tên sản phẩm');
}

async function napFileHuyMmkk(sheets, header, dataRows) {
  const i = (c) => header.indexOf(c);
  const c = {
    ngay: i('Ngày xuất'), st: i('Mã siêu thị'), tenST: i('Tên siêu thị'), nganh: i('Ngành hàng'), nhom: i('Nhóm hàng'),
    sp: i('Tên sản phẩm'), dv: i('Đơn vị'), sl: i('Tổng SL bán'), dt: i('Doanh thu'), gg: i('Tiền s.thị bán giảm giá(chưa VAT)'),
    slgg: i('SL s.thị bán giảm giá'), huy: i('SL hủy tồn'), mm: i('SL mất mát kiểm kê'), ttHuy: i('Thành tiền SL hủy tồn'), ttMm: i('Thành tiền mất mát kiểm kê'),
  };
  const rows = [];
  for (const r of dataRows) {
    const nganh = String(r[c.nganh] || '').trim();
    if (!NGANH_MMKK.includes(nganh)) continue;
    const ma = chuanHoaMaST(r[c.st]);
    if (!ma) continue;
    rows.push([
      chuanHoaKeyNgay(r[c.ngay]) || '', ma, tenNganSieuThi(r[c.tenST]), nganh, String(r[c.nhom] || ''), String(r[c.sp] || ''), String(r[c.dv] || ''),
      soGV(r[c.sl]), soGV(r[c.dt]), soGV(r[c.gg]), soGV(r[c.slgg]), soGV(r[c.huy]), soGV(r[c.mm]), soGV(r[c.ttHuy]), soGV(r[c.ttMm]),
    ]);
  }
  if (rows.length === 0) throw new Error('File không có dòng nào thuộc 4 ngành Rau củ / Thịt / Trái cây / Thủy hải sản');
  await ghiDeTab(sheets, TAB_HUYMMKK, HEADER_HUYMMKK, rows);
  return { loai: 'huymmkk', tenTab: TAB_HUYMMKK, soDong: rows.length, dsST: [...new Set(rows.map((r) => r[1]))] };
}

async function docDuLieuHuyMmkk() {
  const rows = (await docTabAnToan(getSheetsClient(), TAB_HUYMMKK)).slice(1);
  if (rows.length === 0) throw new Error('Chưa có dữ liệu. Anh gửi file "Báo cáo theo dõi số lượng bán, huỷ, hao hụt, MMKK" vào group trước giúp em.');
  const theoST = new Map();
  for (const r of rows) {
    const ma = chuanHoaMaST(r[1]);
    if (!theoST.has(ma)) theoST.set(ma, { ma, ten: String(r[2] || ma), dong: [] });
    theoST.get(ma).dong.push({
      ngay: chuanHoaKeyNgay(r[0]), nganh: String(r[3]), nhom: String(r[4]), sp: String(r[5]), dv: String(r[6]),
      sl: soGV(r[7]), dt: soGV(r[8]), gg: soGV(r[9]), slgg: soGV(r[10]), huy: soGV(r[11]), mm: soGV(r[12]), ttHuy: soGV(r[13]), ttMm: soGV(r[14]),
    });
  }
  // Giá gốc TB theo sản phẩm (gộp mọi ngày, mọi ST) để ước tính tiền huỷ/MMKK
  const gia = new Map();
  for (const st of theoST.values()) for (const d of st.dong) {
    const g = gia.get(d.sp) || { tien: 0, sl: 0 };
    g.tien += d.dt + d.gg; g.sl += d.sl; gia.set(d.sp, g);
  }
  for (const st of theoST.values()) for (const d of st.dong) {
    const g = gia.get(d.sp); const p = g && g.sl > 0 ? g.tien / g.sl : 0;
    d.vHuy = d.ttHuy || d.huy * p;
    d.vMm = d.ttMm || d.mm * p;
  }
  return theoST;
}

function tenSpDv(sp, dv) {
  const ten = rutGonTen(sp, 38);
  return !dv || /\((KG|TRÁI|BÓ|VỈ|HỘP|TÚI)/i.test(sp) ? ten : `${ten} (${dv})`;
}

function veSvgHuyMmkk(st) {
  const W = 1100, PAD = 24, ROW = 30;
  const C = { xanh: '#0B8A3E', xanhDam: '#0B6E35', nen: '#E6F4EC', chan: '#F3F7F5', vien: '#D5E3DA', phu: '#666666', do: '#D0312D', cam: '#E08A00' };
  const t = (x, y, s, o = {}) =>
    `<text x="${x}" y="${y}" font-size="${o.size || 15}" font-weight="${o.bold ? 700 : 400}" fill="${o.fill || '#1a1a1a'}" text-anchor="${o.anchor || 'start'}">${escXml(s)}</text>`;
  const ds = st.dong;
  const ngays = [...new Set(ds.map((d) => d.ngay).filter(Boolean))].sort();
  const tu = ngays[0], den = ngays[ngays.length - 1];
  const sum = (arr, k) => arr.reduce((s, x) => s + (x[k] || 0), 0);
  const fP = (a, b) => (b > 0 ? ((a / b) * 100).toFixed(1).replace('.', ',') + '%' : '–');
  const mauP = (a, b, nguong) => (b > 0 && (a / b) * 100 >= nguong ? C.do : b > 0 && (a / b) * 100 >= nguong / 2 ? C.cam : '#1a1a1a');
  const fKg = (v) => (Math.round(v * 10) / 10).toFixed(1).replace('.', ',');
  const DT = sum(ds, 'dt'), GG = sum(ds, 'gg'), HUY = sum(ds, 'vHuy'), MM = sum(ds, 'vMm');

  const p = [];
  p.push(`<rect x="0" y="0" width="${W}" height="96" fill="${C.xanh}"/>`);
  p.push(t(PAD, 38, 'HUỶ · MẤT MÁT KIỂM KÊ · GIẢM GIÁ FRESH', { size: 26, bold: true, fill: '#FFFFFF' }));
  p.push(t(PAD, 64, `${st.ma} · ${st.ten}`, { size: 16, fill: '#E3F5EA' }));
  p.push(t(PAD, 86, `Từ ${fmtNgay(tu)} đến ${fmtNgay(den)} (${ngays.length} ngày) · Thịt · Rau củ · Trái cây · Thủy hải sản · ĐVT: triệu`, { size: 13, fill: '#E3F5EA' }));
  let y = 112;

  const oW = (W - PAD * 2 - 36) / 4, oH = 74;
  const o = (i, nhan, gt, mau, phu) => {
    const x = PAD + i * (oW + 12);
    p.push(`<rect x="${x}" y="${y}" width="${oW}" height="${oH}" rx="10" fill="#F2F8F4"/>`);
    p.push(t(x + 12, y + 22, nhan, { size: 13, fill: '#777777' }));
    p.push(t(x + 12, y + 54, gt, { size: 25, bold: true, fill: mau || '#1a1a1a' }));
    if (phu) p.push(t(x + oW - 12, y + 54, phu, { size: 12, fill: '#777777', anchor: 'end' }));
  };
  o(0, 'Doanh thu Fresh', trGV(DT), null, `${fmtNgay(tu)}–${fmtNgay(den)}`);
  o(1, 'Tiền giảm giá xả', trGV(GG), C.cam, `${fP(GG, DT)} DT`);
  o(2, 'Mất mát kiểm kê (ước)', trGV(MM), MM > 0 ? C.do : C.xanh, `${fP(MM, DT)} DT`);
  o(3, 'Huỷ tồn (ước)', trGV(HUY), HUY > 0 ? C.cam : C.xanh, `${fP(HUY, DT)} DT`);
  y += oH + 30;

  // Bảng 1: theo nhóm hàng
  const tieuDe = (s) => { p.push(t(PAD, y, s, { size: 16, bold: true, fill: C.xanhDam })); y += 10; };
  const headRow = (cols) => {
    p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="30" rx="5" fill="${C.nen}"/>`);
    for (const [x, s, a] of cols) p.push(t(x, y + 20, s, { size: 13, bold: true, fill: C.xanhDam, anchor: a || 'end' }));
    y += 30;
  };
  tieuDe('THEO NHÓM HÀNG (xếp theo tổng thất thoát)');
  const nhom = new Map();
  for (const d of ds) {
    const k = `${TEN_NGAN_MMKK[d.nganh] || d.nganh} · ${d.nhom}`;
    const g = nhom.get(k) || { ten: k, dt: 0, gg: 0, vHuy: 0, vMm: 0 };
    g.dt += d.dt; g.gg += d.gg; g.vHuy += d.vHuy; g.vMm += d.vMm; nhom.set(k, g);
  }
  const dsNhom = [...nhom.values()].map((g) => ({ ...g, tt: g.gg + g.vHuy + Math.max(0, g.vMm) })).sort((a, b) => b.tt - a.tt);
  const cN = { ten: PAD + 10, dt: 460, gg: 560, pgg: 650, huy: 740, mm: 840, pmm: 930, tt: W - PAD - 10 };
  headRow([[cN.ten, 'Nhóm hàng', 'start'], [cN.dt, 'DT'], [cN.gg, 'Giảm giá'], [cN.pgg, '%GG'], [cN.huy, 'Huỷ'], [cN.mm, 'MMKK'], [cN.pmm, '%MM'], [cN.tt, 'Thất thoát/DT']]);
  dsNhom.forEach((g, k) => {
    if (k % 2 === 1) p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW}" fill="${C.chan}"/>`);
    const yt = y + 20;
    p.push(t(cN.ten, yt, rutGonTen(g.ten, 40)));
    p.push(t(cN.dt, yt, trGV(g.dt), { anchor: 'end' }));
    p.push(t(cN.gg, yt, trGV(g.gg), { anchor: 'end' }));
    p.push(t(cN.pgg, yt, fP(g.gg, g.dt), { anchor: 'end', bold: true, fill: mauP(g.gg, g.dt, 20) }));
    p.push(t(cN.huy, yt, trGV(g.vHuy), { anchor: 'end', fill: C.phu }));
    p.push(t(cN.mm, yt, trGV(g.vMm), { anchor: 'end', bold: g.vMm > 0.5e6, fill: g.vMm > 0.5e6 ? C.do : '#1a1a1a' }));
    p.push(t(cN.pmm, yt, fP(g.vMm, g.dt), { anchor: 'end', bold: true, fill: mauP(g.vMm, g.dt, 10) }));
    p.push(t(cN.tt, yt, fP(g.tt, g.dt), { anchor: 'end', bold: true, fill: mauP(g.tt, g.dt, 25) }));
    y += ROW;
  });
  y += 26;

  // Bảng 2: top sản phẩm mất mát kiểm kê
  const theoSP = new Map();
  for (const d of ds) {
    const g = theoSP.get(d.sp) || { sp: d.sp, nganh: d.nganh, dv: d.dv, sl: 0, dt: 0, gg: 0, slgg: 0, mm: 0, vMm: 0, ngayMm: new Set() };
    g.sl += d.sl; g.dt += d.dt; g.gg += d.gg; g.slgg += d.slgg; g.mm += d.mm; g.vMm += d.vMm;
    if (d.mm > 0) g.ngayMm.add(d.ngay);
    theoSP.set(d.sp, g);
  }
  const sp = [...theoSP.values()];
  const topMm = sp.filter((x) => x.vMm > 0).sort((a, b) => b.vMm - a.vMm).slice(0, 10);
  if (topMm.length) {
    tieuDe('TOP SẢN PHẨM MẤT MÁT KIỂM KÊ');
    const cS = { ten: PAD + 10, sl: 560, mm: 660, pmm: 760, v: 870, ng: W - PAD - 10 };
    headRow([[cS.ten, 'Sản phẩm', 'start'], [cS.sl, 'SL bán'], [cS.mm, 'SL mất'], [cS.pmm, 'Mất/bán'], [cS.v, 'Tiền (ước)'], [cS.ng, 'Số ngày mất']]);
    topMm.forEach((x, k) => {
      if (k % 2 === 1) p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW}" fill="${C.chan}"/>`);
      const yt = y + 20;
      p.push(t(cS.ten, yt, tenSpDv(x.sp, x.dv)));
      p.push(t(cS.sl, yt, fKg(x.sl), { anchor: 'end', fill: C.phu }));
      p.push(t(cS.mm, yt, fKg(x.mm), { anchor: 'end', bold: true }));
      p.push(t(cS.pmm, yt, fP(x.mm, x.sl), { anchor: 'end', bold: true, fill: mauP(x.mm, x.sl, 30) }));
      p.push(t(cS.v, yt, trGV(x.vMm), { anchor: 'end', bold: true, fill: C.do }));
      p.push(t(cS.ng, yt, `${x.ngayMm.size}/${ngays.length}`, { anchor: 'end', fill: x.ngayMm.size >= 3 ? C.do : C.phu, bold: x.ngayMm.size >= 3 }));
      y += ROW;
    });
    y += 26;
  }

  // Bảng 3: top giảm giá nặng
  const topGg = sp.filter((x) => x.gg > 0).sort((a, b) => b.gg - a.gg).slice(0, 8);
  if (topGg.length) {
    tieuDe('TOP SẢN PHẨM GIẢM GIÁ XẢ (tiền giảm nhiều nhất)');
    const cG = { ten: PAD + 10, sl: 560, slgg: 660, gg: 770, dt: 870, p: W - PAD - 10 };
    headRow([[cG.ten, 'Sản phẩm', 'start'], [cG.sl, 'SL bán'], [cG.slgg, 'SL giảm'], [cG.gg, 'Tiền giảm'], [cG.dt, 'DT'], [cG.p, 'Giảm/giá gốc']]);
    topGg.forEach((x, k) => {
      if (k % 2 === 1) p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW}" fill="${C.chan}"/>`);
      const yt = y + 20;
      p.push(t(cG.ten, yt, tenSpDv(x.sp, x.dv)));
      p.push(t(cG.sl, yt, fKg(x.sl), { anchor: 'end', fill: C.phu }));
      p.push(t(cG.slgg, yt, fKg(x.slgg), { anchor: 'end' }));
      p.push(t(cG.gg, yt, trGV(x.gg), { anchor: 'end', bold: true, fill: C.cam }));
      p.push(t(cG.dt, yt, trGV(x.dt), { anchor: 'end', fill: C.phu }));
      p.push(t(cG.p, yt, fP(x.gg, x.gg + x.dt), { anchor: 'end', bold: true, fill: mauP(x.gg, x.gg + x.dt, 30) }));
      y += ROW;
    });
    y += 26;
  }

  // Theo ngày
  tieuDe('THEO NGÀY (MMKK / giảm giá, triệu)');
  y += 6;
  const wN = (W - PAD * 2) / Math.max(ngays.length, 1);
  ngays.forEach((n, k) => {
    const x = PAD + k * wN;
    const dn = ds.filter((d) => d.ngay === n);
    const mm = sum(dn, 'vMm'), gg = sum(dn, 'gg');
    p.push(`<rect x="${x + 4}" y="${y}" width="${wN - 8}" height="58" rx="8" fill="#F2F8F4"/>`);
    p.push(t(x + wN / 2, y + 18, fmtNgay(n), { size: 13, fill: C.phu, anchor: 'middle' }));
    p.push(t(x + wN / 2, y + 38, `MM ${trGV(mm)}`, { size: 14, bold: true, fill: mm > 1e6 ? C.do : '#1a1a1a', anchor: 'middle' }));
    p.push(t(x + wN / 2, y + 54, `GG ${trGV(gg)}`, { size: 12, fill: C.cam, anchor: 'middle' }));
  });
  y += 76;
  p.push(t(PAD, y, 'Tiền huỷ/MMKK ước tính = số lượng × giá bán gốc trung bình của mã (file không có thành tiền). %GG đỏ ≥ 20%, %MM đỏ ≥ 10%.', { size: 12, fill: C.phu }));
  const H = y + PAD;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT_FAMILY_SVG}"><rect width="${W}" height="${H}" fill="#FFFFFF"/>${p.join('')}</svg>`;
}

const TRIGGER_MMKK = ['mmkk', 'mất mát kiểm kê', 'mat mat kiem ke', 'hao hụt', 'huỷ tồn', 'hủy tồn'];
function laTriggerHuyMmkk(text) {
  const t = chuanCau(text);
  return coTuKhoa(text, TRIGGER_MMKK) || t.trim() === 'huỷ' || t.trim() === 'hủy';
}

async function veAnhHuyMmkk(maST) {
  const theoST = await docDuLieuHuyMmkk();
  const st = theoST.get(String(maST)) || [...theoST.values()][0];
  const svg = veSvgHuyMmkk(st);
  return { st, png: svgThanhPng(svg, 1500), preview: svgThanhPng(svg, 600) };
}

async function generateHuyMmkkReport(text, chiCacST) {
  const theoST = await docDuLieuHuyMmkk();
  const m = String(text || '').match(/\b(\d{3,6})\b/);
  let ds = [...theoST.keys()];
  if (m && theoST.has(m[1])) ds = [m[1]];
  else if (chiCacST && chiCacST.length) ds = ds.filter((ma) => chiCacST.includes(ma));
  const out = [];
  for (const ma of ds.slice(0, 5)) {
    const { png, preview } = await veAnhHuyMmkk(ma);
    const id = `mm-${ma}-${Date.now().toString(36)}`;
    luuAnh(id, png, preview);
    const q = `loai=mm&st=${ma}`;
    out.push({ type: 'image', originalContentUrl: `${PUBLIC_URL}/anh/${id}.png?${q}`, previewImageUrl: `${PUBLIC_URL}/anh/${id}.png?${q}&nho=1` });
  }
  if (!out.length) throw new Error('Không tìm thấy siêu thị trong dữ liệu huỷ MMKK.');
  return out.length === 1 ? out[0] : out;
}

// ---------------------------------------------------------------------------
// BÁO CÁO "BC SỮA" — tồn & bán từng sản phẩm sữa (mặc định 4 mã Celano)
//  - File "BC Tồn Theo Model" (cột Tồn kho siêu thị)      -> tab SUA_TON (ghi đè)
//  - File "Doanh Thu Theo Model" chỉ gồm các mã sữa này -> tab SUA_BAN (ghi đè)
//    (file DT có cả ngành khác thì vẫn đi vào DT NGÀNH HÀNG như cũ)
// ---------------------------------------------------------------------------
const TAB_SUA_TON = 'SUA_TON';
const TAB_SUA_BAN = 'SUA_BAN';
const TAB_SUA_TON_CU = 'SUA_TON_CU'; // tồn của ngày trước để so chênh lệch
const ngayCuaCapNhat = (c) => { const m = String(c || '').match(/(\d{2})\/(\d{2})\s*$/); return m ? `${m[2]}-${m[1]}` : ''; };
const MA_SUA_MAC_DINH = ['2609001221', '2609001223', '2609001656', '2609001657', /* C2 */ '2203000875', '2204000011', '2602001178', '2006000354', /* C2 355ML */ '1607002174'];
// Nhóm hiển thị: tên có "C2" -> TRÀ C2, còn lại -> SỮA CELANO
function nhomSua(ten) {
  const s = String(ten);
  if (!/(^|\s)C2(\s|$)/i.test(s)) return 'SUA';
  return /35[05]\s*ML|360\s*ML/i.test(s) ? 'C2N' : 'C2';
}
const NHOM_SUA = [{ k: 'SUA', ten: 'SỮA CELANO', dv: 'hộp' }, { k: 'C2', ten: 'C2 455ML', dv: 'chai' }, { k: 'C2N', ten: 'C2 355ML', dv: 'chai' }];

function laFileTonModel(header) {
  return header.includes('Mã Model') && header.includes('Tồn kho siêu thị') && header.includes('Mã siêu thị');
}

async function dsMaSua(sheets) {
  const ton = (await docTabAnToan(sheets, TAB_SUA_TON)).slice(1);
  const s = new Set(MA_SUA_MAC_DINH);
  for (const r of ton) s.add(String(r[0]).trim());
  return s;
}

async function laFileBanSua(sheets, header, dataRows) {
  const iMa = header.indexOf('Mã Model');
  if (iMa === -1 || dataRows.length === 0 || dataRows.length > 3000) return false;
  const ma = await dsMaSua(sheets);
  return dataRows.every((r) => ma.has(String(r[iMa] ?? '').trim()));
}

function bayGioVN() {
  const d = new Date(Date.now() + 7 * 3600 * 1000);
  const z = (n) => String(n).padStart(2, '0');
  return `${z(d.getUTCHours())}:${z(d.getUTCMinutes())} ${z(d.getUTCDate())}/${z(d.getUTCMonth() + 1)}`;
}
function capNhatTuTen(fileName) {
  const m = String(fileName || '').match(/(20\d{2})(\d{2})(\d{2})_(\d{2})(\d{2})/);
  return m ? `${m[4]}:${m[5]} ${m[3]}/${m[2]}` : '';
}

async function napFileTonSua(sheets, header, dataRows, fileName) {
  const i = (c) => header.indexOf(c);
  const c = { ma: i('Mã Model'), ten: i('Model'), st: i('Mã siêu thị'), tenST: i('Tên siêu thị'), ton: i('Tồn kho siêu thị'), di: i('Số lượng đang  đi đường') };
  if (c.di === -1) c.di = header.findIndex((h) => /đi đường/i.test(h));
  const rows = dataRows
    .filter((r) => r[c.ma] != null && r[c.st] != null)
    .map((r) => [String(r[c.ma]).trim(), String(r[c.ten] || ''), chuanHoaMaST(r[c.st]), tenNganSieuThi(r[c.tenST]), soGV(r[c.ton]), c.di === -1 ? 0 : soGV(r[c.di]), capNhatTuTen(fileName) || bayGioVN()]);
  if (rows.length > 2000) throw new Error('File tồn quá nhiều mã. Anh lọc đúng các mã sữa cần theo dõi rồi gửi lại giúp em.');
  // Gộp theo mã: chỉ thay các mã có trong file mới, giữ nguyên mã khác (VD file riêng C2 355ML)
  const maMoi = new Set(rows.map((r) => r[0]));
  const giuLai = (await docTabAnToan(sheets, TAB_SUA_TON)).slice(1).filter((r) => !maMoi.has(String(r[0]).trim()));
  rows.unshift(...giuLai);
  await ghiDeTab(sheets, TAB_SUA_TON, ['Mã Model', 'Model', 'Mã siêu thị', 'Tên siêu thị', 'Tồn', 'Đang đi đường', 'Cập nhật'], rows);
  return { loai: 'sua', tenTab: TAB_SUA_TON, soDong: rows.length };
}

async function napFileBanSua(sheets, header, dataRows, fileName) {
  const i = (c) => header.indexOf(c);
  const c = { ma: i('Mã Model'), ten: i('Model'), st: i('Mã siêu thị'), tenST: i('Tên siêu thị'), sl: i('Tổng số lượng'), dt: i('Tổng doanh thu') };
  const rows = dataRows.map((r) => [String(r[c.ma]).trim(), String(r[c.ten] || ''), chuanHoaMaST(r[c.st]), tenNganSieuThi(r[c.tenST]), soGV(r[c.sl]), soGV(r[c.dt]), capNhatTuTen(fileName) || bayGioVN()]);
  const maMoiBan = new Set(rows.map((r) => r[0]));
  rows.unshift(...(await docTabAnToan(sheets, TAB_SUA_BAN)).slice(1).filter((r) => !maMoiBan.has(String(r[0]).trim())));
  await ghiDeTab(sheets, TAB_SUA_BAN, ['Mã Model', 'Model', 'Mã siêu thị', 'Tên siêu thị', 'SL bán', 'Doanh thu', 'Cập nhật'], rows);
  return { loai: 'sua', tenTab: TAB_SUA_BAN, soDong: rows.length };
}

function tenNganSua(ten) {
  let x = String(ten);
  if (nhomSua(x) !== 'SUA') {
    x = (' ' + x + ' ').replace(/\s*CHAI.*$/i, ' ').replace(/\s(NƯỚC|TRÀ|C2|FREEZE|HƯƠNG|VỊ)(?=\s)/gi, ' ').replace(/\s(NƯỚC|TRÀ|C2|FREEZE|HƯƠNG|VỊ)(?=\s)/gi, ' ');
  } else {
    x = x.replace(/^CELANO\s+/i, '').replace(/SỮA YẾN MẠCH VỊ\s*/i, '').replace(/T\.?\s*UỐNG\s*/i, '').replace(/\s*CÓ THẠCH/i, '').replace(/\s*\d+\s*ML.*$/i, '');
  }
  return x.replace(/\s+/g, ' ').trim().toLowerCase().replace(/(^|\s)\S/g, (c) => c.toUpperCase());
}

async function docDuLieuSua() {
  const sheets = getSheetsClient();
  const ton = (await docTabAnToan(sheets, TAB_SUA_TON)).slice(1);
  const ban = (await docTabAnToan(sheets, TAB_SUA_BAN)).slice(1);
  const tonCu = (await docTabAnToan(sheets, TAB_SUA_TON_CU)).slice(1);
  if (!ton.length && !ban.length) throw new Error('Chưa có dữ liệu sữa. Anh gửi file "BC Tồn Theo Model" và file "Doanh Thu Theo Model" của các mã sữa vào group giúp em.');
  const sp = new Map(); // ma -> ten
  const st = new Map(); // maST -> { ten, ton:{}, ban:{} }
  const lay = (ma, ten) => { if (!st.has(ma)) st.set(ma, { ma, ten, ton: {}, di: {}, ban: {}, cu: {} }); return st.get(ma); };
  for (const r of ton) {
    const ma = String(r[0]); sp.set(ma, String(r[1]));
    const x = lay(String(r[2]), String(r[3]));
    x.ton[ma] = (x.ton[ma] || 0) + soGV(r[4]); x.di[ma] = (x.di[ma] || 0) + soGV(r[5]);
  }
  const coCu = false; // đã bỏ cột ± tồn theo yêu cầu
  for (const r of tonCu) {
    const ma = String(r[0]); if (!sp.has(ma)) continue;
    const x = st.get(String(r[2])); if (!x) continue;
    x.cu[ma] = (x.cu[ma] || 0) + soGV(r[4]);
  }
  for (const r of ban) {
    const ma = String(r[0]); if (!sp.has(ma)) sp.set(ma, String(r[1]));
    const x = lay(String(r[2]), String(r[3]));
    x.ban[ma] = (x.ban[ma] || 0) + soGV(r[4]);
  }
  const dsSP = [...sp.entries()].map(([ma, ten]) => ({ ma, ten, ngan: tenNganSua(ten), nhom: nhomSua(ten) }))
    .sort((a, b) => (a.nhom === b.nhom ? a.ngan.localeCompare(b.ngan) : ['SUA', 'C2', 'C2N'].indexOf(a.nhom) - ['SUA', 'C2', 'C2N'].indexOf(b.nhom)));
  return { dsSP, dsST: [...st.values()], capTon: String((ton[0] || [])[6] || ''), capBan: String((ban[0] || [])[6] || ''), coCu, capCu: String((tonCu[0] || [])[6] || '') };
}

function veSvgSua(d) {
  // Ô trên: chi tiết từng sản phẩm (mỗi nhóm 1 hàng). Bảng: mỗi nhóm 1 cột dồn Tồn | Bán.
  const { dsSP } = d;
  const nhom = NHOM_SUA.map((g) => ({ ...g, sp: dsSP.filter((s2) => s2.nhom === g.k) })).filter((g) => g.sp.length);
  const W = 1000, PAD = 24, ROW = 30, xTen = PAD + 44;
  const C = { xanh: '#0B8A3E', xanhDam: '#0B6E35', nen: '#E6F4EC', chan: '#F3F7F5', vien: '#D5E3DA', phu: '#888888', do: '#D0312D', cam: '#E08A00', vang: '#FFC20E' };
  const t = (x, y, s2, o = {}) =>
    `<text x="${x}" y="${y}" font-size="${o.size || 15}" font-weight="${o.bold ? 700 : 400}" fill="${o.fill || '#1a1a1a'}" text-anchor="${o.anchor || 'start'}">${escXml(s2)}</text>`;
  const fs = (v) => String(Math.round(v)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const tongNhom = (x, k, g) => g.sp.reduce((a2, p2) => a2 + (x[k][p2.ma] || 0), 0);
  const rows = d.dsST.map((x) => {
    const o = { ...x, g: {} };
    for (const g of nhom) o.g[g.k] = { ton: tongNhom(x, 'ton', g), ban: tongNhom(x, 'ban', g), cu: tongNhom(x, 'cu', g) };
    o.tBan = nhom.reduce((a2, g) => a2 + o.g[g.k].ban, 0);
    return o;
  }).sort((a2, b2) => b2.tBan - a2.tBan);

  const p = [];
  p.push(`<rect x="0" y="0" width="${W}" height="92" fill="${C.xanh}"/>`);
  p.push(t(PAD, 38, 'BC SỮA · TỒN & BÁN', { size: 26, bold: true, fill: '#FFFFFF' }));
  p.push(t(PAD, 64, nhom.map((g) => `${g.ten} (${g.sp.length} mã)`).join(' · '), { size: 14, fill: '#E3F5EA' }));
  p.push(t(PAD, 84, `Tồn ${d.capTon || '–'}${d.coCu ? ` (so với tồn ${d.capCu})` : ''} · Bán ${d.capBan || '–'} · ĐVT: ${nhom.map((g) => g.dv).join(' / ')}`, { size: 13, fill: '#E3F5EA' }));
  const dau = (v) => (v > 0 ? '+' : v < 0 ? '−' : '') + fs(Math.abs(v));
  { const bw = 260, bx = W - PAD - bw, by = 14;
    p.push(`<rect x="${bx}" y="${by}" width="${bw}" height="64" rx="10" fill="#FFFFFF"/>`);
    p.push(`<rect x="${bx}" y="${by}" width="8" height="64" rx="4" fill="${C.vang}"/>`);
    p.push(t(bx + bw / 2 + 4, by + 24, 'KHU VỰC QLTP', { size: 13, bold: true, fill: C.phu, anchor: 'middle' }));
    p.push(t(bx + bw / 2 + 4, by + 50, 'QUỐC NAM · 40909', { size: 20, bold: true, fill: C.xanh, anchor: 'middle' })); }
  let y = 108;

  // Ô chi tiết từng sản phẩm, mỗi nhóm 1 hàng
  const soO = Math.max(...nhom.map((g) => g.sp.length));
  const oW = (W - PAD * 2 - 12 * (soO - 1)) / soO, oH = 92;
  for (const g of nhom) {
    p.push(t(PAD, y + 14, `${g.ten} · chi tiết từng mã (${g.dv})`, { size: 14, bold: true, fill: C.xanhDam }));
    y += 22;
    g.sp.forEach((s2, k) => {
      const x = PAD + k * (oW + 12);
      const tTon = rows.reduce((a2, r) => a2 + (r.ton[s2.ma] || 0), 0);
      const tBan = rows.reduce((a2, r) => a2 + (r.ban[s2.ma] || 0), 0);
      const het = rows.filter((r) => (r.ton[s2.ma] || 0) <= 0).length;
      p.push(`<rect x="${x}" y="${y}" width="${oW}" height="${oH}" rx="10" fill="#F2F8F4"/>`);
      p.push(t(x + 12, y + 20, rutGonTen(s2.ngan, 24), { size: 13, bold: true, fill: C.xanhDam }));
      p.push(t(x + 12, y + 44, 'Tồn', { size: 13, fill: C.phu }));
      p.push(t(x + oW - 12, y + 44, fs(tTon), { size: 17, bold: true, anchor: 'end' }));
      if (d.coCu) { const tCu = rows.reduce((a2, r) => a2 + (r.cu[s2.ma] || 0), 0); p.push(t(x + 44, y + 44, `(${dau(tTon - tCu)})`, { size: 12, fill: tTon - tCu > 0 ? '#1F6FD1' : C.phu })); }
      p.push(t(x + 12, y + 66, 'Bán', { size: 13, fill: C.phu }));
      p.push(t(x + oW - 12, y + 66, fs(tBan), { size: 17, bold: true, fill: C.xanh, anchor: 'end' }));
      p.push(t(x + 12, y + 84, het ? `${het} ST hết hàng` : 'Không ST nào hết', { size: 11, fill: het ? C.do : C.phu }));
    });
    y += oH + 14;
  }
  y += 4;

  // Bảng: mỗi nhóm 1 cột dồn (Tồn | Bán)
  const xNhom0 = 430, wNhom = (W - PAD - xNhom0) / nhom.length;
  const nCot = d.coCu ? 3 : 2; // Tồn | ± hôm qua | Bán
  const xc = (gi, j) => xNhom0 + gi * wNhom + (wNhom * (j + 1)) / nCot - 14;
  const jBan = nCot - 1;
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="62" rx="6" fill="${C.nen}"/>`);
  nhom.forEach((g, gi) => {
    const x1 = xNhom0 + gi * wNhom;
    p.push(`<rect x="${x1 + 8}" y="${y + 6}" width="${wNhom - 16}" height="22" rx="4" fill="${C.xanh}"/>`);
    p.push(t(x1 + wNhom / 2, y + 22, `${g.ten} (${g.dv})`, { size: 13, bold: true, fill: '#FFFFFF', anchor: 'middle' }));
    p.push(t(xc(gi, 0), y + 50, 'Tồn', { size: 13, bold: true, fill: C.xanhDam, anchor: 'end' }));
    if (d.coCu) p.push(t(xc(gi, 1), y + 50, '± tồn', { size: 13, bold: true, fill: C.xanhDam, anchor: 'end' }));
    p.push(t(xc(gi, jBan), y + 50, 'Bán', { size: 13, bold: true, fill: C.xanhDam, anchor: 'end' }));
  });
  p.push(t(PAD + 30, y + 50, '#', { size: 13, bold: true, fill: C.xanhDam, anchor: 'end' }));
  p.push(t(xTen, y + 50, 'Siêu thị', { size: 13, bold: true, fill: C.xanhDam }));
  const yTop = y; y += 62;
  // dòng tổng KV
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW + 2}" fill="${C.xanh}"/>`);
  p.push(t(xTen, y + 21, 'TỔNG KHU VỰC', { fill: '#FFFFFF', bold: true }));
  nhom.forEach((g, gi) => {
    p.push(t(xc(gi, 0), y + 21, fs(rows.reduce((a2, r) => a2 + r.g[g.k].ton, 0)), { fill: '#FFFFFF', bold: true, anchor: 'end' }));
    if (d.coCu) p.push(t(xc(gi, 1), y + 21, dau(rows.reduce((a2, r) => a2 + r.g[g.k].ton - r.g[g.k].cu, 0)), { fill: '#FFFFFF', bold: true, anchor: 'end' }));
    p.push(t(xc(gi, jBan), y + 21, fs(rows.reduce((a2, r) => a2 + r.g[g.k].ban, 0)), { fill: '#FFFFFF', bold: true, anchor: 'end' }));
  });
  y += ROW + 2;
  rows.forEach((r, i) => {
    if (i % 2 === 1) p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW}" fill="${C.chan}"/>`);
    const yt = y + 20;
    p.push(t(PAD + 30, yt, i + 1, { fill: C.phu, anchor: 'end' }));
    p.push(t(xTen, yt, `${r.ma} · ${rutGonTen(r.ten, 30)}`));
    nhom.forEach((g, gi) => {
      const { ton, ban } = r.g[g.k];
      p.push(t(xc(gi, 0), yt, fs(ton), { anchor: 'end', bold: ton <= 0, fill: ton <= 0 ? C.do : '#1a1a1a' }));
      if (d.coCu) { const dl = ton - r.g[g.k].cu; p.push(t(xc(gi, 1), yt, dl ? dau(dl) : '0', { anchor: 'end', fill: dl > 0 ? '#1F6FD1' : dl < 0 ? '#1a1a1a' : C.phu })); }
      p.push(t(xc(gi, jBan), yt, ban ? fs(ban) : '–', { anchor: 'end', bold: ban > 0, fill: ban > 0 ? C.xanh : C.phu }));
    });
    y += ROW;
  });
  nhom.forEach((g, gi) => { const x = xNhom0 + gi * wNhom; p.push(`<line x1="${x}" y1="${yTop + 30}" x2="${x}" y2="${y}" stroke="#9DBFAA" stroke-width="1.5"/>`); });
  p.push(`<rect x="${PAD}" y="${yTop}" width="${W - PAD * 2}" height="${y - yTop}" rx="6" fill="none" stroke="${C.vien}"/>`);
  y += 24;
  p.push(t(PAD, y, `Mỗi nhóm cộng dồn các mã ở trên · Xếp theo tổng bán giảm dần · Tồn đỏ = hết hàng · "–" = chưa bán${d.coCu ? ' · ± tồn = tồn nay − tồn hôm trước (xanh dương = có nhập)' : ''}.`, { size: 12, fill: C.phu }));
  const H = y + PAD - 4;
  return { svg: `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT_FAMILY_SVG}"><rect width="${W}" height="${H}" fill="#FFFFFF"/>${p.join('')}</svg>`, W };
}

function laTriggerSua(text) {
  return coTuKhoa(text, ['bc sữa', 'bc sua', 'báo cáo sữa', 'bao cao sua', 'bc celano', 'celano']);
}

async function veAnhSua() {
  const d = await docDuLieuSua();
  const { svg, W } = veSvgSua(d);
  return { png: svgThanhPng(svg, Math.round(W * 1.4)), preview: svgThanhPng(svg, 600) };
}

async function generateSuaReport() {
  const { png, preview } = await veAnhSua();
  const id = `sua-${Date.now().toString(36)}`;
  luuAnh(id, png, preview);
  return { type: 'image', originalContentUrl: `${PUBLIC_URL}/anh/${id}.png?loai=sua`, previewImageUrl: `${PUBLIC_URL}/anh/${id}.png?loai=sua&nho=1` };
}

// ---------------------------------------------------------------------------
// BÁO CÁO "BC ĐMSK" — Đông mát – Sữa – Kem (NH 990, 1056, 1354, 1355) T10 so mục tiêu = T9 × 120%
//  - "@bot NẠP BASE ĐMSK" rồi gửi file Doanh Thu Theo Model THÁNG 9 trong 3 phút -> tab DMSK_BASE
//  - Gửi file Doanh Thu Theo Model ĐMSK bình thường = luỹ kế THÁNG 10            -> tab DMSK_T10 (ghi đè)
// ---------------------------------------------------------------------------
const TAB_DMSK_BASE = 'DMSK_BASE';
const TAB_DMSK_T10 = 'DMSK_T10';
const HE_SO_DMSK = 1.2;
const cho_NapBaseDmsk = new Map(); // targetId -> hạn

function tongDtTheoST(header, dataRows) {
  const iST = header.indexOf('Mã siêu thị'), iTen = header.indexOf('Tên siêu thị'), iDT = header.indexOf('Tổng doanh thu');
  const m = new Map();
  for (const r of dataRows) {
    const ma = chuanHoaMaST(r[iST]);
    if (!ma) continue;
    const x = m.get(ma) || { ma, ten: tenNganSieuThi(r[iTen]), dt: 0 };
    x.dt += soGV(r[iDT]);
    m.set(ma, x);
  }
  return [...m.values()];
}

async function napBaseDmsk(sheets, header, dataRows, fileName) {
  const ds = tongDtTheoST(header, dataRows);
  await ghiDeTab(sheets, TAB_DMSK_BASE, ['Mã siêu thị', 'Tên siêu thị', 'DT tháng base', `MT (×${HE_SO_DMSK * 100}%)`, 'Cập nhật'],
    ds.map((x) => [x.ma, x.ten, Math.round(x.dt), Math.round(x.dt * HE_SO_DMSK), capNhatTuTen(fileName)]));
  return { loai: 'dmsk_base', soDong: ds.length, tong: ds.reduce((s, x) => s + x.dt, 0) };
}

// DOANH THU T10 CHO BC ĐMSK — mỗi ngày chỉ cần đổ file 1 NGÀY, bot tự cộng dồn tháng.
//  - DMSK_T10  = mốc luỹ kế (file luỹ kế từ 01/10). Gửi kèm "NẠP NGÀY dd/mm" = luỹ kế đến HẾT ngày đó.
//  - DMSK_NGAY = DT từng ngày (file chỉ 1 ngày; gửi nhiều lần trong ngày thì lấy bản mới nhất).
//  Luỹ kế = mốc + tổng các ngày sau mốc. Cột "Hôm nay" = file ngày mới nhất.
const TAB_DMSK_NGAY = 'DMSK_NGAY';
function soNgayChayDmsk(ngay, gio) {
  const d = Number(ngay.slice(8, 10));
  return Math.max(0.1, d - 1 + (gio == null || gio === '' ? 1 : Math.min(1, Math.max(0, (Number(gio) - 6) / 15))));
}
async function tinhLuyKeDmsk(sheets, tabLK = TAB_DMSK_T10, tabNgay = TAB_DMSK_NGAY) {
  const t10 = (await docTabAnToan(sheets, tabLK)).slice(1);
  const ng = (await docTabAnToan(sheets, tabNgay)).slice(1).filter((r) => r[0]);
  if (!t10.length && !ng.length) return null;
  const m = new Map();
  const add = (ma, ten, v) => { const o = m.get(ma) || { ten, dt: 0 }; o.dt += v; if (!o.ten) o.ten = ten; m.set(ma, o); };
  for (const r of t10) add(chuanHoaMaST(r[0]), String(r[1] || ''), soGV(r[2]));
  const aNgay = t10.length ? String(t10[0][6] || '') : '';
  const aDu = t10.length ? Number(t10[0][7] || 0) === 1 : false;
  const cacNgay = [...new Set(ng.map((r) => String(r[0])))].sort();
  // mốc cũ không ghi ngày -> không cộng thêm ngày nào (tránh cộng trùng)
  const sau = t10.length ? (aNgay ? cacNgay.filter((x) => x > aNgay) : []) : cacNgay;
  for (const r of ng) if (sau.includes(String(r[0]))) add(chuanHoaMaST(r[1]), String(r[2] || ''), soGV(r[3]));
  let ngayMoi = aNgay, gioMoi = aDu ? 23 : t10.length ? t10[0][8] : null, capNhat = t10.length ? String(t10[0][3] || '') : '';
  let soNgay = t10.length && !aNgay ? soGV(t10[0][4]) || 1 : null;
  const ngayCuoi = cacNgay[cacNgay.length - 1];
  if (ngayCuoi && (!aNgay || ngayCuoi >= aNgay) && (sau.length || ngayCuoi === aNgay || !t10.length)) {
    const rr = ng.filter((r) => String(r[0]) === ngayCuoi);
    if (sau.includes(ngayCuoi) || !t10.length) { ngayMoi = ngayCuoi; gioMoi = rr[0][5]; capNhat = String(rr[0][4] || ''); }
  }
  if (ngayMoi) soNgay = soNgayChayDmsk(ngayMoi, gioMoi);
  const keyThang = ngayMoi || homNayVN();
  const ngayThang = t10.length && !aNgay ? soGV(t10[0][5]) || 31 : new Date(Number(keyThang.slice(0, 4)), Number(keyThang.slice(5, 7)), 0).getDate();
  // Hôm nay = file 1 ngày của ngày mới nhất
  let homNay = null, nhanHN = '';
  const ngayHN = ngayMoi && cacNgay.includes(ngayMoi) ? ngayMoi : null;
  if (ngayHN) {
    const rr = ng.filter((r) => String(r[0]) === ngayHN);
    homNay = new Map(rr.map((r) => [chuanHoaMaST(r[1]), soGV(r[3])]));
    nhanHN = String(rr[0][4] || '');
  }
  return { m, soNgay: soNgay || 1, ngayThang, capNhat, homNay, nhanHN, ngayMoi, aNgay, aDu };
}

async function napT10Dmsk(sheets, header, dataRows, fileName, ngayChiDinh, dsSan, tabLK = TAB_DMSK_T10, tabNgay = TAB_DMSK_NGAY) {
  const ds = dsSan || tongDtTheoST(header, dataRows);
  const ngay = ngayChiDinh || ngayTuTenFile(fileName);
  // Tên file không có giờ xuất -> lấy giờ bot nhận file
  const gioVN = () => { const d = new Date(Date.now() + 7 * 3600 * 1000); return d.getUTCHours() + d.getUTCMinutes() / 60; };
  const gio = ngayChiDinh ? 23 : (gioTuTenFile(fileName) ?? gioVN());
  const tong = ds.reduce((s2, x) => s2 + x.dt, 0);
  const cap = ngayChiDinh ? `hết ${ngay.slice(8, 10)}/${ngay.slice(5, 7)}` : capNhatTuTen(fileName) || bayGioVN();
  // File 1 ngày hay luỹ kế? So với nhịp bán TB/ngày đang có
  const hien = await tinhLuyKeDmsk(sheets, tabLK, tabNgay);
  let motNgay = false;
  if (hien) {
    const tongLK = [...hien.m.values()].reduce((s2, x) => s2 + x.dt, 0);
    const tb = tongLK / Math.max(1, hien.soNgay);
    motNgay = tong < tb * 2.5 && tong < tongLK * 0.6;
  }
  // File luỹ kế mà tổng NHỎ HƠN luỹ kế hiện có (đã gồm hôm nay) -> là luỹ kế đến HẾT HÔM QUA
  let ngayMoc = ngay, du = !!ngayChiDinh;
  // File luỹ kế KHÔNG có giờ trong tên (kiểu xuất mới từ 13/10) -> mặc định là luỹ kế đến HẾT HÔM QUA
  if (!motNgay && !ngayChiDinh && gioTuTenFile(fileName) == null) { ngayMoc = taoKeyNgayLuiDmsk(ngay, 1); du = true; }
  else if (!motNgay && hien && !ngayChiDinh) {
    const tongLK = [...hien.m.values()].reduce((s2, x) => s2 + x.dt, 0);
    if (tong < tongLK * 0.995) { ngayMoc = taoKeyNgayLuiDmsk(ngay, 1); du = true; }
  }
  if (motNgay) {
    const cu = (await docTabAnToan(sheets, tabNgay)).slice(1)
      .filter((r) => r[0] && String(r[0]) !== ngay && String(r[0]).slice(0, 7) === ngay.slice(0, 7));
    for (const x of ds) cu.push([ngay, x.ma, x.ten, Math.round(x.dt), cap, gio == null ? '' : Math.round(gio * 100) / 100]);
    await ghiDeTab(sheets, tabNgay, ['Ngày', 'Mã siêu thị', 'Tên siêu thị', 'DT ngày', 'Cập nhật', 'Giờ'], cu);
    return { loai: 'dmsk', soDong: ds.length, motNgay: true, ngay };
  }
  const capMoc = du ? `hết ${ngayMoc.slice(8, 10)}/${ngayMoc.slice(5, 7)}` : cap;
  const gioMoc = du ? 23 : gio;
  const ngayThang = new Date(Number(ngayMoc.slice(0, 4)), Number(ngayMoc.slice(5, 7)), 0).getDate();
  await ghiDeTab(sheets, tabLK,
    ['Mã siêu thị', 'Tên siêu thị', 'DT luỹ kế', 'Cập nhật', 'Số ngày đã chạy', 'Số ngày trong tháng', 'Ngày', 'Đủ ngày', 'Giờ'],
    ds.map((x) => [x.ma, x.ten, Math.round(x.dt), capMoc, Math.round(soNgayChayDmsk(ngayMoc, gioMoc) * 100) / 100, ngayThang, ngayMoc, du ? 1 : 0, gioMoc == null ? '' : Math.round(gioMoc * 100) / 100]));
  return { loai: 'dmsk', soDong: ds.length, motNgay: false, ngay: ngayMoc, du };
}
function taoKeyNgayLuiDmsk(key, n) {
  const t = new Date(key + 'T00:00:00Z'); t.setUTCDate(t.getUTCDate() - n);
  return t.toISOString().slice(0, 10);
}

async function docDuLieuDmsk() {
  const sheets = getSheetsClient();
  const base = (await docTabAnToan(sheets, TAB_DMSK_BASE)).slice(1);
  if (!base.length) throw new Error('Chưa có base tháng 9. Anh @tag em gõ "NẠP BASE ĐMSK" rồi gửi file Doanh Thu Theo Model tháng 9 trong 3 phút.');
  const lk = await tinhLuyKeDmsk(sheets);
  if (!lk) throw new Error('Chưa có doanh thu tháng 10. Anh gửi file Doanh Thu Theo Model Đông mát – Sữa – Kem tháng 10 vào group.');
  const { soNgay, ngayThang } = lk;
  const tienDo = Math.min(1, soNgay / ngayThang);
  const mBase = new Map(base.map((r) => [chuanHoaMaST(r[0]), { ten: String(r[1]), dt9: soGV(r[2]), mt: soGV(r[3]) }]));
  const dong = [];
  for (const [ma, o] of lk.m) {
    const b = mBase.get(ma);
    dong.push({ ma, ten: o.ten || (b && b.ten), dt9: b ? b.dt9 : 0, mt: b ? b.mt : 0, dt: o.dt, p: b && b.mt > 0 ? (o.dt / b.mt) * 100 : null });
  }
  for (const [ma, b] of mBase) if (!lk.m.has(ma)) dong.push({ ma, ten: b.ten, dt9: b.dt9, mt: b.mt, dt: 0, p: 0 });
  for (const x of dong) x.dk = x.p == null ? null : x.p / tienDo; // dự kiến % cuối tháng nếu giữ nhịp
  for (const x of dong) x.hn = lk.homNay ? lk.homNay.get(x.ma) || 0 : null;
  const coMT = dong.filter((x) => x.p != null).sort((a, b) => b.p - a.p);
  const khongMT = dong.filter((x) => x.p == null).sort((a, b) => b.dt - a.dt);
  const tongMT = coMT.reduce((s, x) => s + x.mt, 0), tongDT = dong.reduce((s, x) => s + x.dt, 0), tongDT9 = coMT.reduce((s, x) => s + x.dt9, 0);
  return {
    dong: [...coMT, ...khongMT], tongMT, tongDT, tongDT9, soNgay, ngayThang, tienDo, capNhat: lk.capNhat, capBase: String(base[0][4] || ''),
    coHomNay: !!lk.homNay, nhanHomNay: lk.nhanHN, tongHN: lk.homNay ? dong.reduce((s2, x) => s2 + (x.hn || 0), 0) : 0,
  };
}

function veSvgDmsk(d) {
  const W = 1100, PAD = 24, ROW = 30;
  const C = { xanh: '#0B8A3E', xanhDam: '#0B6E35', nen: '#E6F4EC', chan: '#F3F7F5', vien: '#D5E3DA', phu: '#666666', do: '#D0312D', cam: '#E08A00' };
  const t = (x, y, s, o = {}) =>
    `<text x="${x}" y="${y}" font-size="${o.size || 15}" font-weight="${o.bold ? 700 : 400}" fill="${o.fill || '#1a1a1a'}" text-anchor="${o.anchor || 'start'}">${escXml(s)}</text>`;
  const chuan = d.tienDo * 100;
  const mau = (p, c) => (p >= c ? C.xanh : p >= c * 0.9 ? C.cam : C.do);
  const pct = (v) => (v == null ? '–' : fmtPhanTram(v));
  const pKV = d.tongMT > 0 ? (d.tongDT / d.tongMT) * 100 : 0;
  const p = [];
  p.push(`<rect x="0" y="0" width="${W}" height="96" fill="${C.xanh}"/>`);
  p.push(t(PAD, 38, 'BC ĐMSK · ĐÔNG MÁT – SỮA – KEM T10', { size: 26, bold: true, fill: '#FFFFFF' }));
  // Huy hiệu khu vực góc phải header
  { const bw = 290, bx = W - PAD - bw, by = 16;
    p.push(`<rect x="${bx}" y="${by}" width="${bw}" height="64" rx="10" fill="#FFFFFF"/>`);
    p.push(`<rect x="${bx}" y="${by}" width="8" height="64" rx="4" fill="#FFC20E"/>`);
    p.push(t(bx + bw / 2 + 4, by + 24, 'KHU VỰC QLTP', { size: 13, bold: true, fill: C.phu, anchor: 'middle' }));
    p.push(t(bx + bw / 2 + 4, by + 50, 'QUỐC NAM · 40909', { size: 22, bold: true, fill: C.xanh, anchor: 'middle' })); }
  p.push(t(PAD, 64, `Mục tiêu T10 = DT T9 × ${Math.round(HE_SO_DMSK * 100)}% · NH 990 · 1056 · 1354 · 1355`, { size: 15, fill: '#E3F5EA' }));
  p.push(t(PAD, 86, `DT T10 cập nhật ${d.capNhat} · đã chạy ${String(Math.round(d.soNgay * 10) / 10).replace('.', ',')}/${d.ngayThang} ngày · ĐVT: triệu`, { size: 13, fill: '#E3F5EA' }));
  let y = 112;
  const oW = (W - PAD * 2 - 36) / 4, oH = 74;
  const o = (i, nhan, gt, mauGt, phu) => {
    const x = PAD + i * (oW + 12);
    p.push(`<rect x="${x}" y="${y}" width="${oW}" height="${oH}" rx="10" fill="#F2F8F4"/>`);
    p.push(t(x + 12, y + 22, nhan, { size: 13, fill: '#777777' }));
    p.push(t(x + 12, y + 54, gt, { size: 24, bold: true, fill: mauGt || '#1a1a1a' }));
    if (phu) p.push(t(x + oW - 12, y + 54, phu, { size: 12, fill: '#777777', anchor: 'end' }));
  };
  const nguyen = (v) => String(Math.ceil(v - 1e-9)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  o(0, 'Mục tiêu T10 KV', nguyen(d.tongMT / 1e6), null, `T9 ${nguyen(d.tongDT9 / 1e6)}`);
  o(1, 'DT T10 luỹ kế', nguyen(d.tongDT / 1e6));
  o(2, '% đạt mục tiêu', fmtPhanTram(pKV), mau(pKV, chuan), `chuẩn ${fmtPhanTram(chuan)}`);
  o(3, 'Dự kiến cuối tháng', fmtPhanTram(pKV / d.tienDo), mau(pKV / d.tienDo, 100), 'nếu giữ nhịp');
  y += oH + 16;

  // Cột: MT T10 | DT LK | % đạt | TB/ngày | Cần/ngày còn lại | DK tháng (tr) | DK %
  const conLai = Math.max(0.5, d.ngayThang - d.soNgay);
  // Thứ tự: Hôm nay | Cần/ngày | Dự kiến tháng | Dự kiến % | MT T10 | DT LK T10 | % đạt
  const cot = { stt: PAD + 30, ten: PAD + 44, hn: 420, can: 520, mt: 640, dt: 750, p: 850, dkTr: 975, dk: W - PAD - 10 };
  const vach = [545, 872]; // vạch ngăn 3 nhóm: HÔM NAY | LUỸ KẾ T10 | DỰ KIẾN CUỐI THÁNG
  const nhanHN = d.coHomNay ? (d.nhanHomNay.split(' ')[0] || 'realtime') : 'realtime';
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="70" rx="6" fill="${C.nen}"/>`);
  const nhom = [[340, vach[0], 'HÔM NAY'], [vach[0], vach[1], 'LUỸ KẾ T10'], [vach[1], W - PAD, 'DỰ KIẾN CUỐI THÁNG']];
  for (const [x1, x2, ten] of nhom) { p.push(`<rect x="${x1 + 6}" y="${y + 6}" width="${x2 - x1 - 12}" height="20" rx="4" fill="${C.xanh}"/>`); p.push(t((x1 + x2) / 2, y + 21, ten, { size: 12, bold: true, fill: '#FFFFFF', anchor: 'middle' })); }
  y += 22;
  const td = (x, a1, a2, al) => { p.push(t(x, y + 20, a1, { size: 13, bold: true, fill: C.xanhDam, anchor: al || 'end' })); if (a2) p.push(t(x, y + 38, a2, { size: 13, bold: true, fill: C.xanhDam, anchor: al || 'end' })); };
  td(cot.stt, '#', ''); td(cot.ten, 'Siêu thị', '', 'start');
  td(cot.mt, 'MT', 'T10'); td(cot.dt, 'DT T10', 'luỹ kế'); td(cot.hn, 'Hôm nay', nhanHN); td(cot.p, '%', 'đạt');
  td(cot.can, 'Cần/ngày', 'để đạt'); td(cot.dkTr, 'DT', 'cả tháng'); td(cot.dk, '% đạt', 'cuối tháng');
  const yTop = y - 22; y += 48;
  // dòng tổng KV
  const tbKV = d.tongDT / d.soNgay, canKV = Math.max(0, (d.tongMT - d.dong.filter((x) => x.p != null).reduce((s2, x) => s2 + x.dt, 0)) / conLai);
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW + 2}" fill="${C.xanh}"/>`);
  const w = { fill: '#FFFFFF', bold: true, anchor: 'end' };
  p.push(t(cot.ten, y + 21, 'TỔNG KHU VỰC', { fill: '#FFFFFF', bold: true }));
  p.push(t(cot.mt, y + 21, nguyen(d.tongMT / 1e6), w)); p.push(t(cot.dt, y + 21, nguyen(d.tongDT / 1e6), w)); p.push(t(cot.hn, y + 21, d.coHomNay ? nguyen(d.tongHN / 1e6) : '–', w));
  p.push(t(cot.p, y + 21, fmtPhanTram(pKV), w));
  p.push(t(cot.can, y + 21, nguyen(canKV / 1e6), w)); p.push(t(cot.dkTr, y + 21, nguyen((tbKV * d.ngayThang) / 1e6), w));
  p.push(t(cot.dk, y + 21, fmtPhanTram(pKV / d.tienDo), w));
  y += ROW + 2;
  d.dong.forEach((x, k) => {
    if (k % 2 === 1) p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW}" fill="${C.chan}"/>`);
    const yt = y + 20, coMT = x.p != null;
    const tb = x.dt / d.soNgay, dkTr = tb * d.ngayThang, can = coMT ? Math.max(0, (x.mt - x.dt) / conLai) : null;
    p.push(t(cot.stt, yt, k + 1, { fill: C.phu, anchor: 'end' }));
    p.push(t(cot.ten, yt, rutGonTen(x.ten, 26) + (coMT ? '' : ' (mới)'), { fill: coMT ? '#1a1a1a' : C.phu }));
    p.push(t(cot.mt, yt, coMT ? nguyen(x.mt / 1e6) : '–', { fill: C.phu, anchor: 'end' }));
    p.push(t(cot.dt, yt, nguyen(x.dt / 1e6), { bold: true, anchor: 'end' }));
    p.push(t(cot.hn, yt, x.hn == null ? '–' : nguyen(x.hn / 1e6), { bold: true, fill: x.hn != null && coMT && x.hn >= can ? C.xanh : '#1a1a1a', anchor: 'end' }));
    p.push(t(cot.p, yt, pct(x.p), { bold: true, fill: coMT ? '#1a1a1a' : C.phu, anchor: 'end' }));
    p.push(t(cot.can, yt, coMT ? nguyen(can / 1e6) : '–', { bold: true, fill: coMT ? (can > tb * 1.1 ? C.do : can > tb ? C.cam : C.xanh) : C.phu, anchor: 'end' }));
    p.push(t(cot.dkTr, yt, nguyen(dkTr / 1e6), { anchor: 'end' }));
    p.push(t(cot.dk, yt, pct(x.dk), { bold: true, fill: coMT ? mau(x.dk, 100) : C.phu, anchor: 'end' }));
    y += ROW;
  });
  for (const vx of vach) p.push(`<line x1="${vx}" y1="${yTop + 28}" x2="${vx}" y2="${y}" stroke="#9DBFAA" stroke-width="1.5"/>`);
  p.push(`<rect x="${PAD}" y="${yTop}" width="${W - PAD * 2}" height="${y - yTop}" rx="6" fill="none" stroke="${C.vien}"/>`);
  y += 24;
  p.push(t(PAD, y, `Hôm nay xanh = đã bán đủ Cần/ngày · Cần/ngày = (MT − DT) ÷ ${String(Math.round(conLai * 10) / 10).replace('.', ',')} ngày còn lại (đỏ: cao hơn nhịp đang bán >10%) · DT cả tháng = DT luỹ kế ÷ ${String(Math.round(d.soNgay * 10) / 10).replace('.', ',')} ngày × ${d.ngayThang}.`, { size: 12, fill: C.phu }));
  const H = y + PAD - 4;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT_FAMILY_SVG}"><rect width="${W}" height="${H}" fill="#FFFFFF"/>${p.join('')}</svg>`;
}

function laTriggerDmsk(text) {
  return coTuKhoa(text, ['bc đmsk', 'bc dmsk', 'đmsk', 'dmsk']) && !laLenhNapBaseDmsk(text);
}
function laLenhNapBaseDmsk(text) {
  return /n[ạa]p\s*base/i.test((text || '').normalize('NFC')) && !/(bnnb|ths|trứng|trung|hải sản|hai san)/i.test(text || '');
}

async function veAnhDmsk() {
  const svg = veSvgDmsk(await docDuLieuDmsk());
  return { png: svgThanhPng(svg, 1500), preview: svgThanhPng(svg, 600) };
}

async function generateDmskReport() {
  const { png, preview } = await veAnhDmsk();
  const id = `dmsk-${Date.now().toString(36)}`;
  luuAnh(id, png, preview);
  return { type: 'image', originalContentUrl: `${PUBLIC_URL}/anh/${id}.png?loai=dmsk`, previewImageUrl: `${PUBLIC_URL}/anh/${id}.png?loai=dmsk&nho=1` };
}

// ---------------------------------------------------------------------------
// BÁO CÁO BNNB — Bia Nước (MT = T9 × 125%) + Nhà Bếp/Gia vị (MT = T9 × 120%)
// Cho 1 siêu thị, chia theo nhóm hàng. Mỗi ngày đổ file 1 ngày, bot tự cộng dồn (giống BC ĐMSK).
//  - "NẠP BASE BNNB" rồi gửi file T9 Bia nước + Gia vị (gửi chung hoặc 2 file, trong 3 phút)
//  - File DT Theo Model chỉ gồm NH 1060/1255/1055 của 1 siêu thị -> DT tháng 10
// ---------------------------------------------------------------------------
const BNNB_NHOM = [
  { k: 'BN', ten: 'BIA – NƯỚC', nh: ['1060', '1255'], hs: 1.25, thangBase: 'T7' }, // thi đua Bia–Nước so với T07
  { k: 'GV', ten: 'GIA VỊ (NHÀ BẾP)', nh: ['1055'], hs: 1.2, thangBase: 'T9' },
];
const TAB_BNNB_BASE = 'BNNB_BASE';
const cho_NapBaseBnnb = new Map(); // targetId -> hạn (nhận nhiều file trong 3 phút)
const bnnbTab = (k, loai) => `BNNB_${k}_${loai}`;
function nhomBnnbCuaNH(nh) { return BNNB_NHOM.find((g) => g.nh.includes(String(nh ?? '').trim())); }

function laFileBnnb(header, dataRows) {
  const iNH = header.indexOf('Mã ngành hàng'), iST = header.indexOf('Mã siêu thị');
  if (iNH < 0 || iST < 0) return false;
  const nh = new Set(dataRows.map((r) => String(r[iNH] ?? '').trim()).filter(Boolean));
  const st = new Set(dataRows.map((r) => String(r[iST] ?? '').trim()).filter(Boolean));
  return nh.size > 0 && [...nh].every((x) => nhomBnnbCuaNH(x)) && st.size === 1;
}

function tongBnnbTheoNhom(header, dataRows) {
  const i = (c) => header.indexOf(c);
  const c = { nh: i('Mã ngành hàng'), mn: i('Mã nhóm hàng'), tn: i('Nhóm hàng'), dt: i('Tổng doanh thu'), st: i('Mã siêu thị'), tst: i('Tên siêu thị') };
  const out = new Map();
  for (const r of dataRows) {
    const g = nhomBnnbCuaNH(r[c.nh]); if (!g) continue;
    if (!out.has(g.k)) out.set(g.k, { m: new Map(), st: { ma: chuanHoaMaST(r[c.st]), ten: tenNganSieuThi(r[c.tst]) } });
    const o = out.get(g.k), ma = String(r[c.mn] ?? '').trim();
    const x = o.m.get(ma) || { ma, ten: String(r[c.tn] || ''), dt: 0 };
    x.dt += soGV(r[c.dt]); o.m.set(ma, x);
  }
  for (const o of out.values()) o.ds = [...o.m.values()];
  return out;
}

async function napBaseBnnb(sheets, header, dataRows) {
  const theo = tongBnnbTheoNhom(header, dataRows);
  const rows = (await docTabAnToan(sheets, TAB_BNNB_BASE)).slice(1).filter((r) => !theo.has(String(r[0])));
  const tomTat = [];
  for (const [k, o] of theo) {
    const g = BNNB_NHOM.find((x) => x.k === k);
    for (const x of o.ds) rows.push([k, x.ma, x.ten, Math.round(x.dt), Math.round(x.dt * g.hs), o.st.ma, o.st.ten, bayGioVN()]);
    const tong = o.ds.reduce((a, x) => a + x.dt, 0);
    tomTat.push(`• ${g.ten}: ${g.thangBase || 'T9'} ${fmtTrieu(tong / 1e6)} tr × ${Math.round(g.hs * 100)}% = MT ${fmtTrieu((tong * g.hs) / 1e6)} tr`);
  }
  await ghiDeTab(sheets, TAB_BNNB_BASE, ['Nhóm', 'Mã nhóm hàng', 'Nhóm hàng', 'DT base', 'MT', 'Mã siêu thị', 'Tên siêu thị', 'Cập nhật'], rows);
  const conThieu = BNNB_NHOM.filter((g) => !rows.some((r) => String(r[0]) === g.k)).map((g) => g.ten);
  return { loai: 'bnnb_base', soDong: rows.length, tomTat, conThieu };
}

async function napT10Bnnb(sheets, header, dataRows, fileName, ngayChiDinh) {
  const theo = tongBnnbTheoNhom(header, dataRows);
  let soDong = 0;
  for (const [k, o] of theo) {
    await napT10Dmsk(sheets, header, dataRows, fileName, ngayChiDinh, o.ds, bnnbTab(k, 'LK'), bnnbTab(k, 'NGAY'));
    soDong += o.ds.length;
  }
  return { loai: 'bnnb', soDong };
}

async function docDuLieuBnnb() {
  const sheets = getSheetsClient();
  const base = (await docTabAnToan(sheets, TAB_BNNB_BASE)).slice(1);
  if (!base.length) throw new Error('Chưa có base BNNB. Anh @tag em gõ "NẠP BASE BNNB" rồi gửi file tháng 9 Bia nước + Gia vị trong 3 phút.');
  const st = { ma: String(base[0][5] || ''), ten: String(base[0][6] || '') };
  const nhom = [];
  let capNhat = '', soNgayMax = 0, ngayThang = 31;
  for (const g of BNNB_NHOM) {
    const b = base.filter((r) => String(r[0]) === g.k);
    if (!b.length) continue;
    const lk = await tinhLuyKeDmsk(sheets, bnnbTab(g.k, 'LK'), bnnbTab(g.k, 'NGAY'));
    const soNgay = lk ? lk.soNgay : 0;
    if (lk) { ngayThang = lk.ngayThang; if (soNgay >= soNgayMax) { soNgayMax = soNgay; capNhat = lk.capNhat; } }
    const coHN = !!(lk && lk.homNay);
    const m = new Map(b.map((r) => [chuanHoaMaST(r[1]), { ma: chuanHoaMaST(r[1]), ten: String(r[2]), dt9: soGV(r[3]), mt: soGV(r[4]), dt: 0, hn: coHN ? 0 : null }]));
    if (lk) for (const [ma, o] of lk.m) { const x = m.get(ma) || { ma, ten: o.ten, dt9: 0, mt: 0, dt: 0, hn: coHN ? 0 : null }; x.dt += o.dt; m.set(ma, x); }
    if (coHN) for (const [ma, v] of lk.homNay) { const x = m.get(ma); if (x) x.hn = v; }
    // Gộp các nhóm nhỏ (< 2% MT của ngành) thành "Nhóm khác"
    const tongMT = [...m.values()].reduce((a, x) => a + x.mt, 0);
    // Giữ tối đa 20 nhóm lớn (MT ≥ 1% ngành), còn lại gộp "Các nhóm khác"
    const sx = [...m.values()].sort((a, b2) => b2.mt - a.mt || b2.dt - a.dt);
    const lon = [], khac = { ma: '*', ten: 'Các nhóm khác', dt9: 0, mt: 0, dt: 0, hn: coHN ? 0 : null, so: 0 };
    for (const x of sx) {
      if (lon.length < 20 && (x.mt >= tongMT * 0.01 || (x.mt === 0 && x.dt >= tongMT * 0.01))) lon.push(x);
      else { khac.dt9 += x.dt9; khac.mt += x.mt; khac.dt += x.dt; if (coHN) khac.hn += x.hn || 0; khac.so++; }
    }
    if (khac.so) { khac.ten = `Các nhóm khác (${khac.so} nhóm)`; lon.push(khac); }
    const conLai = Math.max(0.5, ngayThang - soNgay), tienDo = soNgay > 0 ? Math.min(1, soNgay / ngayThang) : 0;
    const tinh = (x) => {
      x.p = x.mt > 0 ? (x.dt / x.mt) * 100 : null;
      x.can = x.mt > 0 ? Math.max(0, (x.mt - x.dt) / conLai) : null;
      x.dkTr = soNgay > 0 ? (x.dt / soNgay) * ngayThang : 0;
      x.dk = x.p != null && tienDo > 0 ? x.p / tienDo : null;
      return x;
    };
    lon.forEach(tinh);
    const tong = tinh({ ten: g.ten, mt: tongMT, dt: lon.reduce((a, x) => a + x.dt, 0), hn: coHN ? lon.reduce((a, x) => a + (x.hn || 0), 0) : null });
    nhom.push({ ...g, dong: lon, tong, soNgay, tienDo, coHN, nhanHN: coHN ? lk.nhanHN : '' });
  }
  if (!nhom.some((g) => g.soNgay > 0)) throw new Error('Chưa có doanh thu tháng 10 cho BNNB. Anh gõ "NẠP NGÀY dd/mm" (ngày hôm qua) rồi gửi file luỹ kế từ 01/10, sau đó mỗi ngày gửi file 1 ngày.');
  const tienDo = Math.min(1, soNgayMax / ngayThang);
  const conLai = Math.max(0.5, ngayThang - soNgayMax);
  const T = { mt: 0, dt: 0, hn: 0, coHN: nhom.some((g) => g.coHN) };
  for (const g of nhom) { T.mt += g.tong.mt; T.dt += g.tong.dt; T.hn += g.tong.hn || 0; }
  T.p = T.mt > 0 ? (T.dt / T.mt) * 100 : 0; T.can = Math.max(0, (T.mt - T.dt) / conLai);
  T.dkTr = soNgayMax > 0 ? (T.dt / soNgayMax) * ngayThang : 0; T.dk = tienDo > 0 ? T.p / tienDo : 0;
  return { st, nhom, T, soNgay: soNgayMax, ngayThang, tienDo, conLai, capNhat };
}

// Bảng "CÁC MỨC ĐẠT" so với base T9: mỗi ngành × mỗi mức (105–125%) -> cần bán thêm bao nhiêu/ngày
const MUC_DAT = [1.05, 1.1, 1.15, 1.2, 1.25];
function veBangMucDat(p, t, C, PAD, W, y, d) {
  const tr = (v) => fmtTrieu(v / 1e6);
  const x0 = PAD + 300, cw = (W - PAD - x0) / MUC_DAT.length, RH = 68;
  y += 34;
  p.push(t(PAD, y + 4, 'CÁC MỨC ĐẠT SO VỚI BASE · mỗi ngày cần bán bao nhiêu để lên mức', { size: 15, bold: true, fill: C.xanhDam }));
  y += 14;
  const yTop = y;
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="34" rx="6" fill="${C.nen}"/>`);
  p.push(t(PAD + 14, y + 22, 'Ngành hàng · dự kiến cuối tháng', { size: 13, bold: true, fill: C.xanhDam }));
  MUC_DAT.forEach((m, i) => p.push(t(x0 + i * cw + cw / 2, y + 22, `${Math.round(m * 100)}%`, { size: 14, bold: true, fill: C.xanhDam, anchor: 'middle' })));
  y += 34;
  d.nhom.forEach((g, k) => {
    const base = g.tong.mt / g.hs, dt = g.tong.dt;
    const conLai = Math.max(0.5, d.ngayThang - g.soNgay);
    const dkPct = base > 0 ? (g.tong.dkTr / base) * 100 : 0;
    p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${RH}" fill="${k % 2 ? C.chan : '#FFFFFF'}"/>`);
    p.push(t(PAD + 14, y + 22, `${g.ten} · base ${g.thangBase || 'T9'} ${tr(base)}`, { size: 14, bold: true }));
    const nhip = g.soNgay > 0 ? dt / g.soNgay : 0;
    const datDuoc = MUC_DAT.filter((m) => dkPct >= m * 100).pop();
    p.push(t(PAD + 14, y + 41, `Đang bán ${tr(nhip)}/ngày · dự kiến ${fmtPhanTram(dkPct)}`, { size: 12, fill: C.phu }));
    p.push(t(PAD + 14, y + 58, datDuoc ? `→ Giữ nhịp đạt mức ${Math.round(datDuoc * 100)}%` : '→ Chưa đạt mức nào', { size: 12, bold: true, fill: datDuoc ? C.xanh : C.do }));
    MUC_DAT.forEach((m, i) => {
      const cx = x0 + i * cw, mt = base * m, can = Math.max(0, mt - dt) / conLai;
      const daDat = dt >= mt, kipNhip = dkPct >= m * 100;
      const nen = daDat ? '#BFE6CC' : kipNhip ? '#E3F4E8' : dkPct >= m * 100 * 0.95 ? '#FFF3D6' : '#FDECEC';
      p.push(`<rect x="${cx + 3}" y="${y + 4}" width="${cw - 6}" height="${RH - 8}" rx="6" fill="${nen}"/>`);
      if (daDat) p.push(t(cx + cw / 2, y + 39, 'ĐÃ ĐẠT', { size: 15, bold: true, fill: C.xanh, anchor: 'middle' }));
      else {
        p.push(t(cx + cw / 2, y + 24, `${tr(can)}/ngày`, { size: 16, bold: true, fill: kipNhip ? C.xanh : C.do, anchor: 'middle' }));
        p.push(t(cx + cw / 2, y + 43, kipNhip ? '✔ giữ nhịp là đạt' : `bán thêm +${tr(can - nhip)}/ngày`, { size: 12, bold: true, fill: kipNhip ? C.xanh : C.do, anchor: 'middle' }));
        p.push(t(cx + cw / 2, y + 59, `MT ${tr(mt)}`, { size: 11, fill: C.phu, anchor: 'middle' }));
      }
    });
    y += RH;
  });
  p.push(`<rect x="${PAD}" y="${yTop}" width="${W - PAD * 2}" height="${y - yTop}" rx="6" fill="none" stroke="${C.vien}"/>`);
  y += 18;
  p.push(t(PAD, y, 'Xanh = giữ nhịp hiện tại là đạt · Vàng = sát mức (thiếu < 5%) · Đỏ = phải bán thêm so với nhịp đang bán · ĐVT: triệu.', { size: 12, fill: C.phu }));
  return y;
}

function veSvgBnnb(d) {
  const W = 1100, PAD = 24, ROW = 30;
  const C = { xanh: '#0B8A3E', xanhDam: '#0B6E35', nen: '#E6F4EC', chan: '#F3F7F5', vien: '#D5E3DA', phu: '#666666', do: '#D0312D', cam: '#E08A00', vang: '#FFC20E', xanhNhat: '#DCEFE3' };
  const t = (x, y, s, o = {}) =>
    `<text x="${x}" y="${y}" font-size="${o.size || 15}" font-weight="${o.bold ? 700 : 400}" fill="${o.fill || '#1a1a1a'}" text-anchor="${o.anchor || 'start'}">${escXml(s)}</text>`;
  const chuan = d.tienDo * 100;
  const mau = (p, c) => (p == null ? C.phu : p >= c ? C.xanh : p >= c * 0.9 ? C.cam : C.do);
  const pct = (v) => (v == null ? '–' : fmtPhanTram(v));
  const tr = (v) => fmtTrieu(v / 1e6);
  const soNg = (v) => String(Math.round(v * 10) / 10).replace('.', ',');
  const p = [];
  p.push(`<rect x="0" y="0" width="${W}" height="96" fill="${C.xanh}"/>`);
  p.push(t(PAD, 38, 'BÁO CÁO BNNB · BIA NƯỚC – GIA VỊ T10', { size: 26, bold: true, fill: '#FFFFFF' }));
  p.push(t(PAD, 64, `Mục tiêu T10: ${d.nhom.map((g) => `${g.k === 'BN' ? 'Bia nước' : 'Gia vị'} = ${g.thangBase || 'T9'} × ${Math.round(g.hs * 100)}%`).join(' · ')}`, { size: 15, fill: '#E3F5EA' }));
  p.push(t(PAD, 86, `Cập nhật ${d.capNhat || '–'} · đã chạy ${soNg(d.soNgay)}/${d.ngayThang} ngày · ĐVT: triệu`, { size: 13, fill: '#E3F5EA' }));
  { const bw = 300, bx = W - PAD - bw, by = 16;
    p.push(`<rect x="${bx}" y="${by}" width="${bw}" height="64" rx="10" fill="#FFFFFF"/>`);
    p.push(`<rect x="${bx}" y="${by}" width="8" height="64" rx="4" fill="${C.vang}"/>`);
    p.push(t(bx + bw / 2 + 4, by + 24, `SIÊU THỊ ${d.st.ma}`, { size: 13, bold: true, fill: C.phu, anchor: 'middle' }));
    p.push(t(bx + bw / 2 + 4, by + 50, rutGonTen(d.st.ten, 26), { size: 18, bold: true, fill: C.xanh, anchor: 'middle' })); }
  let y = 112;

  // 3 ô: từng ngành + tổng
  const o3 = d.nhom.map((g) => ({ ten: g.ten, ...g.tong }));
  const oW = (W - PAD * 2 - 12 * (o3.length - 1)) / o3.length, oH = 104;
  o3.forEach((o, k) => {
    const x = PAD + k * (oW + 12), tong = false;
    p.push(`<rect x="${x}" y="${y}" width="${oW}" height="${oH}" rx="10" fill="${tong ? C.xanhNhat : '#F2F8F4'}"/>`);
    p.push(t(x + 14, y + 22, o.ten, { size: 14, bold: true, fill: C.xanhDam }));
    p.push(t(x + oW - 14, y + 22, `chuẩn ${fmtPhanTram(chuan)}`, { size: 12, fill: C.phu, anchor: 'end' }));
    p.push(t(x + 14, y + 56, pct(o.p), { size: 28, bold: true, fill: mau(o.p, chuan) }));
    p.push(t(x + oW - 14, y + 44, `DT ${tr(o.dt)} / MT ${tr(o.mt)}`, { size: 13, anchor: 'end' }));
    p.push(t(x + oW - 14, y + 64, `Cần ${tr(o.can)}/ngày`, { size: 13, bold: true, fill: C.do, anchor: 'end' }));
    p.push(t(x + 14, y + 90, `Dự kiến ${pct(o.dk)}`, { size: 13, bold: true, fill: mau(o.dk, 100) }));
    if (o.hn != null && (o.coHN !== false)) p.push(t(x + oW - 14, y + 90, `H.nay ${tr(o.hn || 0)}`, { size: 13, fill: C.xanhDam, anchor: 'end' }));
  });
  y += oH + 16;

  // Bảng
  const cot = { stt: PAD + 30, ten: PAD + 44, hn: 425, can: 520, mt: 640, dt: 750, p: 850, dkTr: 975, dk: W - PAD - 10 };
  const vach = [545, 872];
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="70" rx="6" fill="${C.nen}"/>`);
  for (const [x1, x2, ten] of [[340, vach[0], 'HÔM NAY'], [vach[0], vach[1], 'LUỸ KẾ T10'], [vach[1], W - PAD, 'DỰ KIẾN CUỐI THÁNG']]) {
    p.push(`<rect x="${x1 + 6}" y="${y + 6}" width="${x2 - x1 - 12}" height="20" rx="4" fill="${C.xanh}"/>`);
    p.push(t((x1 + x2) / 2, y + 21, ten, { size: 12, bold: true, fill: '#FFFFFF', anchor: 'middle' }));
  }
  const yTop = y; y += 22;
  const td = (x, a1, a2, al) => { p.push(t(x, y + 20, a1, { size: 13, bold: true, fill: C.xanhDam, anchor: al || 'end' })); if (a2) p.push(t(x, y + 38, a2, { size: 13, bold: true, fill: C.xanhDam, anchor: al || 'end' })); };
  const nhanHN = (d.nhom.find((g) => g.coHN) || {}).nhanHN || '';
  td(cot.stt, '#', ''); td(cot.ten, 'Ngành hàng', '', 'start');
  td(cot.hn, 'Hôm nay', (nhanHN.split(' ')[0]) || 'realtime'); td(cot.can, 'Cần/ngày', 'để đạt');
  td(cot.mt, 'MT', 'T10'); td(cot.dt, 'DT T10', 'luỹ kế'); td(cot.p, '%', 'đạt');
  td(cot.dkTr, 'DT', 'cả tháng'); td(cot.dk, '% đạt', 'cuối tháng');
  y += 48;
  const dongSo = (x, yy, o, dam) => {
    const w = dam ? { bold: true, fill: '#FFFFFF', anchor: 'end' } : null;
    const hnTxt = o.hn == null ? '–' : tr(o.hn);
    p.push(t(cot.hn, yy, hnTxt, w || { bold: true, fill: o.hn != null && o.can != null && o.hn >= o.can ? C.xanh : '#1a1a1a', anchor: 'end' }));
    p.push(t(cot.can, yy, o.can == null ? '–' : tr(o.can), w || { bold: true, fill: o.can == null ? C.phu : C.do, anchor: 'end' }));
    p.push(t(cot.mt, yy, o.mt ? tr(o.mt) : '–', w || { fill: C.phu, anchor: 'end' }));
    p.push(t(cot.dt, yy, tr(o.dt), w || { bold: true, anchor: 'end' }));
    p.push(t(cot.p, yy, pct(o.p), w || { bold: true, anchor: 'end' }));
    p.push(t(cot.dkTr, yy, tr(o.dkTr), w || { anchor: 'end' }));
    p.push(t(cot.dk, yy, pct(o.dk), w || { bold: true, fill: mau(o.dk, 100), anchor: 'end' }));
  };
  // Gọn: chỉ 1 dòng tổng cho mỗi ngành (không chi tiết nhóm hàng)
  d.nhom.forEach((g, k) => {
    p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW + 8}" fill="${k % 2 ? C.chan : '#FFFFFF'}"/>`);
    p.push(t(cot.stt, y + 25, k + 1, { fill: C.phu, anchor: 'end' }));
    p.push(t(cot.ten, y + 25, `${g.ten} · ×${Math.round(g.hs * 100)}%`, { bold: true, size: 16 }));
    dongSo(null, y + 25, g.tong, false); y += ROW + 8;
  });
  for (const vx of vach) p.push(`<line x1="${vx}" y1="${yTop + 28}" x2="${vx}" y2="${y}" stroke="#9DBFAA" stroke-width="1.5"/>`);
  p.push(`<rect x="${PAD}" y="${yTop}" width="${W - PAD * 2}" height="${y - yTop}" rx="6" fill="none" stroke="${C.vien}"/>`);
  y += 24;
  p.push(t(PAD, y, `Hôm nay xanh = đã bán đủ Cần/ngày · Cần/ngày = (MT − DT) ÷ ${soNg(d.conLai)} ngày còn lại · DT cả tháng = DT luỹ kế ÷ ${soNg(d.soNgay)} ngày × ${d.ngayThang}.`, { size: 12, fill: C.phu }));
  y = veBangMucDat(p, t, C, PAD, W, y, d);
  const H = y + PAD - 4;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT_FAMILY_SVG}"><rect width="${W}" height="${H}" fill="#FFFFFF"/>${p.join('')}</svg>`;
}

function laTriggerBnnb(text) {
  return coTuKhoa(text, ['báo cáo bnnb', 'bao cao bnnb', 'bc bnnb', 'bnnb']) && !laLenhNapBaseBnnb(text);
}
function laLenhNapBaseBnnb(text) {
  const s = (text || '').normalize('NFC');
  return /n[ạa]p\s*base/i.test(s) && /bnnb/i.test(s);
}
async function veAnhBnnb() {
  const svg = veSvgBnnb(await docDuLieuBnnb());
  return { png: svgThanhPng(svg, 1500), preview: svgThanhPng(svg, 600) };
}
async function generateBnnbReport() {
  const { png, preview } = await veAnhBnnb();
  const id = `bnnb-${Date.now().toString(36)}`;
  luuAnh(id, png, preview);
  return { type: 'image', originalContentUrl: `${PUBLIC_URL}/anh/${id}.png?loai=bnnb`, previewImageUrl: `${PUBLIC_URL}/anh/${id}.png?loai=bnnb&nho=1` };
}

// ---------------------------------------------------------------------------
// BÁO CÁO THS + TRỨNG — Thuỷ hải sản (NH 1254) + Trứng gia cầm (nhóm 2856), MT = T9 × 110%
// Cho 1 siêu thị, chia theo nhóm hàng. Mỗi ngày đổ file 1 ngày, bot tự cộng dồn (giống BC ĐMSK).
//  - "NẠP BASE THS" rồi gửi file T9 Thuỷ hải sản + Trứng (gửi chung hoặc 2 file, trong 3 phút)
//  - File DT Theo Model chỉ gồm NH 1060/1255/1055 của 1 siêu thị -> DT tháng 10
// ---------------------------------------------------------------------------
const THST_NHOM = [
  { k: 'TS', ten: 'THUỶ HẢI SẢN', ngan: 'Thuỷ hải sản', hs: 1.1, khop: (nh, mn) => nh === '1254' },
  { k: 'TR', ten: 'TRỨNG GIA CẦM', ngan: 'Trứng', hs: 1.1, khop: (nh, mn) => mn === '2856' },
  // Thịt nhập khẩu: ĐÃ TẮT theo yêu cầu (21/10). Bật lại: bỏ comment dòng dưới.
  // { k: 'TN', ten: 'THỊT NHẬP KHẨU', ngan: 'Thịt NK', hs: 1.1, khop: (nh, mn, ten) => nh === '1236' && laThitNhapKhau(ten) },
];
// Thịt nhập khẩu: nhận theo tên model (NHẬP KHẨU, NK, MỸ, (ĐỨC), (TBN), (BRAZIL), (NGA)...)
function laThitNhapKhau(ten) {
  const s = ` ${String(ten || '').toUpperCase()} `;
  return /NHẬP KHẨU|[\s(]NK[\s)]|[\s(](MỸ|ĐỨC|TBN|TÂY BAN NHA|BRAZIL|NGA|ÚC|CANADA|BA LAN|HÀ LAN|ĐAN MẠCH|PHÁP|ẤN ĐỘ|ARGENTINA)[\s)]/.test(s);
}
const TAB_THST_BASE = 'THST_BASE';
const cho_NapBaseThst = new Map(); // targetId -> hạn (nhận nhiều file trong 3 phút)
const thstTab = (k, loai) => `THST_${k}_${loai}`;
function nhomThstCua(nh, mn, ten) { return THST_NHOM.find((g) => g.khop(String(nh ?? '').trim(), String(mn ?? '').trim(), ten)); }

function laFileThst(header, dataRows) {
  const iNH = header.indexOf('Mã ngành hàng'), iMN = header.indexOf('Mã nhóm hàng'), iST = header.indexOf('Mã siêu thị');
  if (iNH < 0 || iMN < 0 || iST < 0 || !dataRows.length) return false;
  const st = new Set(dataRows.map((r) => String(r[iST] ?? '').trim()).filter(Boolean));
  const iTen = header.indexOf('Model');
  const trongPhamVi = (r) => { const nh = String(r[iNH] ?? '').trim(); return nh === '1254' || String(r[iMN] ?? '').trim() === '2856'; };
  return st.size === 1 && dataRows.every(trongPhamVi) && dataRows.some((r) => nhomThstCua(r[iNH], r[iMN], r[iTen]));
}

function tongThstTheoNhom(header, dataRows) {
  const i = (c) => header.indexOf(c);
  const c = { nh: i('Mã ngành hàng'), mn: i('Mã nhóm hàng'), tn: i('Nhóm hàng'), dt: i('Tổng doanh thu'), st: i('Mã siêu thị'), tst: i('Tên siêu thị'), md: i('Model') };
  const out = new Map();
  for (const r of dataRows) {
    const g = nhomThstCua(r[c.nh], r[c.mn], r[c.md]); if (!g) continue;
    if (!out.has(g.k)) out.set(g.k, { m: new Map(), st: { ma: chuanHoaMaST(r[c.st]), ten: tenNganSieuThi(r[c.tst]) } });
    const o = out.get(g.k), ma = String(r[c.mn] ?? '').trim();
    const x = o.m.get(ma) || { ma, ten: String(r[c.tn] || ''), dt: 0 };
    x.dt += soGV(r[c.dt]); o.m.set(ma, x);
  }
  for (const o of out.values()) o.ds = [...o.m.values()];
  return out;
}

async function napBaseThst(sheets, header, dataRows) {
  const theo = tongThstTheoNhom(header, dataRows);
  const rows = (await docTabAnToan(sheets, TAB_THST_BASE)).slice(1).filter((r) => !theo.has(String(r[0])));
  const tomTat = [];
  for (const [k, o] of theo) {
    const g = THST_NHOM.find((x) => x.k === k);
    for (const x of o.ds) rows.push([k, x.ma, x.ten, Math.round(x.dt), Math.round(x.dt * g.hs), o.st.ma, o.st.ten, bayGioVN()]);
    const tong = o.ds.reduce((a, x) => a + x.dt, 0);
    tomTat.push(`• ${g.ten}: T9 ${fmtTrieu(tong / 1e6)} tr × ${Math.round(g.hs * 100)}% = MT ${fmtTrieu((tong * g.hs) / 1e6)} tr`);
  }
  await ghiDeTab(sheets, TAB_THST_BASE, ['Nhóm', 'Mã nhóm hàng', 'Nhóm hàng', 'DT base', 'MT', 'Mã siêu thị', 'Tên siêu thị', 'Cập nhật'], rows);
  const conThieu = THST_NHOM.filter((g) => !rows.some((r) => String(r[0]) === g.k)).map((g) => g.ten);
  return { loai: 'thst_base', soDong: rows.length, tomTat, conThieu };
}

async function napT10Thst(sheets, header, dataRows, fileName, ngayChiDinh) {
  const theo = tongThstTheoNhom(header, dataRows);
  let soDong = 0;
  for (const [k, o] of theo) {
    await napT10Dmsk(sheets, header, dataRows, fileName, ngayChiDinh, o.ds, thstTab(k, 'LK'), thstTab(k, 'NGAY'));
    soDong += o.ds.length;
  }
  return { loai: 'thst', soDong };
}

async function docDuLieuThst() {
  const sheets = getSheetsClient();
  const base = (await docTabAnToan(sheets, TAB_THST_BASE)).slice(1);
  if (!base.length) throw new Error('Chưa có base THS + Trứng. Anh @tag em gõ "NẠP BASE THS" rồi gửi file tháng 9 Thuỷ hải sản + Trứng trong 3 phút.');
  const st = { ma: String(base[0][5] || ''), ten: String(base[0][6] || '') };
  const nhom = [];
  let capNhat = '', soNgayMax = 0, ngayThang = 31;
  for (const g of THST_NHOM) {
    const b = base.filter((r) => String(r[0]) === g.k);
    if (!b.length) continue;
    const lk = await tinhLuyKeDmsk(sheets, thstTab(g.k, 'LK'), thstTab(g.k, 'NGAY'));
    const soNgay = lk ? lk.soNgay : 0;
    if (lk) { ngayThang = lk.ngayThang; if (soNgay >= soNgayMax) { soNgayMax = soNgay; capNhat = lk.capNhat; } }
    const coHN = !!(lk && lk.homNay);
    const m = new Map(b.map((r) => [chuanHoaMaST(r[1]), { ma: chuanHoaMaST(r[1]), ten: String(r[2]), dt9: soGV(r[3]), mt: soGV(r[4]), dt: 0, hn: coHN ? 0 : null }]));
    if (lk) for (const [ma, o] of lk.m) { const x = m.get(ma) || { ma, ten: o.ten, dt9: 0, mt: 0, dt: 0, hn: coHN ? 0 : null }; x.dt += o.dt; m.set(ma, x); }
    if (coHN) for (const [ma, v] of lk.homNay) { const x = m.get(ma); if (x) x.hn = v; }
    // Gộp các nhóm nhỏ (< 2% MT của ngành) thành "Nhóm khác"
    const tongMT = [...m.values()].reduce((a, x) => a + x.mt, 0);
    // Giữ tối đa 20 nhóm lớn (MT ≥ 1% ngành), còn lại gộp "Các nhóm khác"
    const sx = [...m.values()].sort((a, b2) => b2.mt - a.mt || b2.dt - a.dt);
    const lon = [], khac = { ma: '*', ten: 'Các nhóm khác', dt9: 0, mt: 0, dt: 0, hn: coHN ? 0 : null, so: 0 };
    for (const x of sx) {
      if (lon.length < 20 && (x.mt >= tongMT * 0.01 || (x.mt === 0 && x.dt >= tongMT * 0.01))) lon.push(x);
      else { khac.dt9 += x.dt9; khac.mt += x.mt; khac.dt += x.dt; if (coHN) khac.hn += x.hn || 0; khac.so++; }
    }
    if (khac.so) { khac.ten = `Các nhóm khác (${khac.so} nhóm)`; lon.push(khac); }
    const conLai = Math.max(0.5, ngayThang - soNgay), tienDo = soNgay > 0 ? Math.min(1, soNgay / ngayThang) : 0;
    const tinh = (x) => {
      x.p = x.mt > 0 ? (x.dt / x.mt) * 100 : null;
      x.can = x.mt > 0 ? Math.max(0, (x.mt - x.dt) / conLai) : null;
      x.dkTr = soNgay > 0 ? (x.dt / soNgay) * ngayThang : 0;
      x.dk = x.p != null && tienDo > 0 ? x.p / tienDo : null;
      return x;
    };
    lon.forEach(tinh);
    const tong = tinh({ ten: g.ten, mt: tongMT, dt: lon.reduce((a, x) => a + x.dt, 0), hn: coHN ? lon.reduce((a, x) => a + (x.hn || 0), 0) : null });
    nhom.push({ ...g, dong: lon, tong, soNgay, tienDo, coHN, nhanHN: coHN ? lk.nhanHN : '' });
  }
  if (!nhom.some((g) => g.soNgay > 0)) throw new Error('Chưa có doanh thu tháng 10 cho THS + Trứng. Anh gõ "NẠP NGÀY dd/mm" (ngày hôm qua) rồi gửi file luỹ kế từ 01/10, sau đó mỗi ngày gửi file 1 ngày.');
  const tienDo = Math.min(1, soNgayMax / ngayThang);
  const conLai = Math.max(0.5, ngayThang - soNgayMax);
  const T = { mt: 0, dt: 0, hn: 0, coHN: nhom.some((g) => g.coHN) };
  for (const g of nhom) { T.mt += g.tong.mt; T.dt += g.tong.dt; T.hn += g.tong.hn || 0; }
  if (!T.coHN) T.hn = null;
  T.p = T.mt > 0 ? (T.dt / T.mt) * 100 : 0; T.can = Math.max(0, (T.mt - T.dt) / conLai);
  T.dkTr = soNgayMax > 0 ? (T.dt / soNgayMax) * ngayThang : 0; T.dk = tienDo > 0 ? T.p / tienDo : 0;
  return { st, nhom, T, soNgay: soNgayMax, ngayThang, tienDo, conLai, capNhat };
}

function veSvgThst(d) {
  const W = 1100, PAD = 24, ROW = 30;
  const C = { xanh: '#0B8A3E', xanhDam: '#0B6E35', nen: '#E6F4EC', chan: '#F3F7F5', vien: '#D5E3DA', phu: '#666666', do: '#D0312D', cam: '#E08A00', vang: '#FFC20E', xanhNhat: '#DCEFE3' };
  const t = (x, y, s, o = {}) =>
    `<text x="${x}" y="${y}" font-size="${o.size || 15}" font-weight="${o.bold ? 700 : 400}" fill="${o.fill || '#1a1a1a'}" text-anchor="${o.anchor || 'start'}">${escXml(s)}</text>`;
  const chuan = d.tienDo * 100;
  const mau = (p, c) => (p == null ? C.phu : p >= c ? C.xanh : p >= c * 0.9 ? C.cam : C.do);
  const pct = (v) => (v == null ? '–' : fmtPhanTram(v));
  const tr = (v) => fmtTrieu(v / 1e6);
  const soNg = (v) => String(Math.round(v * 10) / 10).replace('.', ',');
  const p = [];
  p.push(`<rect x="0" y="0" width="${W}" height="96" fill="${C.xanh}"/>`);
  p.push(t(PAD, 38, 'BÁO CÁO THUỶ HẢI SẢN – TRỨNG T10', { size: 26, bold: true, fill: '#FFFFFF' }));
  p.push(t(PAD, 64, `Mục tiêu T10: ${d.nhom.map((g) => `${g.ngan} = T9 × ${Math.round(g.hs * 100)}%`).join(' · ')}`, { size: 15, fill: '#E3F5EA' }));
  p.push(t(PAD, 86, `Cập nhật ${d.capNhat || '–'} · đã chạy ${soNg(d.soNgay)}/${d.ngayThang} ngày · ĐVT: triệu`, { size: 13, fill: '#E3F5EA' }));
  { const bw = 300, bx = W - PAD - bw, by = 16;
    p.push(`<rect x="${bx}" y="${by}" width="${bw}" height="64" rx="10" fill="#FFFFFF"/>`);
    p.push(`<rect x="${bx}" y="${by}" width="8" height="64" rx="4" fill="${C.vang}"/>`);
    p.push(t(bx + bw / 2 + 4, by + 24, `SIÊU THỊ ${d.st.ma}`, { size: 13, bold: true, fill: C.phu, anchor: 'middle' }));
    p.push(t(bx + bw / 2 + 4, by + 50, rutGonTen(d.st.ten, 26), { size: 18, bold: true, fill: C.xanh, anchor: 'middle' })); }
  let y = 112;

  // 3 ô: từng ngành + tổng
  const o3 = [...d.nhom.map((g) => ({ ten: g.ten, ...g.tong })), { ten: 'TỔNG THS + TRỨNG', ...d.T }];
  const oW = (W - PAD * 2 - 12 * (o3.length - 1)) / o3.length, oH = 104;
  o3.forEach((o, k) => {
    const x = PAD + k * (oW + 12), tong = k === o3.length - 1;
    p.push(`<rect x="${x}" y="${y}" width="${oW}" height="${oH}" rx="10" fill="${tong ? C.xanhNhat : '#F2F8F4'}"/>`);
    p.push(t(x + 14, y + 22, o.ten, { size: 14, bold: true, fill: C.xanhDam }));
    p.push(t(x + oW - 14, y + 22, `chuẩn ${fmtPhanTram(chuan)}`, { size: 12, fill: C.phu, anchor: 'end' }));
    p.push(t(x + 14, y + 56, pct(o.p), { size: 28, bold: true, fill: mau(o.p, chuan) }));
    p.push(t(x + oW - 14, y + 44, `DT ${tr(o.dt)} / MT ${tr(o.mt)}`, { size: 13, anchor: 'end' }));
    p.push(t(x + oW - 14, y + 64, `Cần ${tr(o.can)}/ngày`, { size: 13, bold: true, fill: C.do, anchor: 'end' }));
    p.push(t(x + 14, y + 90, `Dự kiến ${pct(o.dk)}`, { size: 13, bold: true, fill: mau(o.dk, 100) }));
    if (o.hn != null && (o.coHN !== false)) p.push(t(x + oW - 14, y + 90, `H.nay ${tr(o.hn || 0)}`, { size: 13, fill: C.xanhDam, anchor: 'end' }));
  });
  y += oH + 16;

  // Bảng
  const cot = { stt: PAD + 30, ten: PAD + 44, hn: 425, can: 520, mt: 640, dt: 750, p: 850, dkTr: 975, dk: W - PAD - 10 };
  const vach = [545, 872];
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="70" rx="6" fill="${C.nen}"/>`);
  for (const [x1, x2, ten] of [[340, vach[0], 'HÔM NAY'], [vach[0], vach[1], 'LUỸ KẾ T10'], [vach[1], W - PAD, 'DỰ KIẾN CUỐI THÁNG']]) {
    p.push(`<rect x="${x1 + 6}" y="${y + 6}" width="${x2 - x1 - 12}" height="20" rx="4" fill="${C.xanh}"/>`);
    p.push(t((x1 + x2) / 2, y + 21, ten, { size: 12, bold: true, fill: '#FFFFFF', anchor: 'middle' }));
  }
  const yTop = y; y += 22;
  const td = (x, a1, a2, al) => { p.push(t(x, y + 20, a1, { size: 13, bold: true, fill: C.xanhDam, anchor: al || 'end' })); if (a2) p.push(t(x, y + 38, a2, { size: 13, bold: true, fill: C.xanhDam, anchor: al || 'end' })); };
  const nhanHN = (d.nhom.find((g) => g.coHN) || {}).nhanHN || '';
  td(cot.stt, '#', ''); td(cot.ten, 'Ngành hàng', '', 'start');
  td(cot.hn, 'Hôm nay', (nhanHN.split(' ')[0]) || 'realtime'); td(cot.can, 'Cần/ngày', 'để đạt');
  td(cot.mt, 'MT', 'T10'); td(cot.dt, 'DT T10', 'luỹ kế'); td(cot.p, '%', 'đạt');
  td(cot.dkTr, 'DT', 'cả tháng'); td(cot.dk, '% đạt', 'cuối tháng');
  y += 48;
  const dongSo = (x, yy, o, dam) => {
    const w = dam ? { bold: true, fill: '#FFFFFF', anchor: 'end' } : null;
    const hnTxt = o.hn == null ? '–' : tr(o.hn);
    p.push(t(cot.hn, yy, hnTxt, w || { bold: true, fill: o.hn != null && o.can != null && o.hn >= o.can ? C.xanh : '#1a1a1a', anchor: 'end' }));
    p.push(t(cot.can, yy, o.can == null ? '–' : tr(o.can), w || { bold: true, fill: o.can == null ? C.phu : C.do, anchor: 'end' }));
    p.push(t(cot.mt, yy, o.mt ? tr(o.mt) : '–', w || { fill: C.phu, anchor: 'end' }));
    p.push(t(cot.dt, yy, tr(o.dt), w || { bold: true, anchor: 'end' }));
    p.push(t(cot.p, yy, pct(o.p), w || { bold: true, anchor: 'end' }));
    p.push(t(cot.dkTr, yy, tr(o.dkTr), w || { anchor: 'end' }));
    p.push(t(cot.dk, yy, pct(o.dk), w || { bold: true, fill: mau(o.dk, 100), anchor: 'end' }));
  };
  // Chỉ hiện tổng từng ngành + tổng chung (không chi tiết nhóm hàng)
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW + 8}" fill="${C.xanhDam}"/>`);
  p.push(t(cot.ten, y + 25, 'TỔNG THS + TRỨNG', { fill: '#FFFFFF', bold: true, size: 16 }));
  dongSo(null, y + 25, d.T, true); y += ROW + 8;
  d.nhom.forEach((g, k) => {
    p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW + 8}" fill="${k % 2 ? C.chan : '#FFFFFF'}"/>`);
    p.push(t(cot.stt, y + 25, k + 1, { fill: C.phu, anchor: 'end' }));
    p.push(t(cot.ten, y + 25, `${g.ten} · ×${Math.round(g.hs * 100)}%`, { bold: true, size: 16 }));
    dongSo(null, y + 25, g.tong, false); y += ROW + 8;
  });
  for (const vx of vach) p.push(`<line x1="${vx}" y1="${yTop + 28}" x2="${vx}" y2="${y}" stroke="#9DBFAA" stroke-width="1.5"/>`);
  p.push(`<rect x="${PAD}" y="${yTop}" width="${W - PAD * 2}" height="${y - yTop}" rx="6" fill="none" stroke="${C.vien}"/>`);
  y += 24;
  p.push(t(PAD, y, `Hôm nay xanh = đã bán đủ Cần/ngày · Cần/ngày = (MT − DT) ÷ ${soNg(d.conLai)} ngày còn lại · DT cả tháng = DT luỹ kế ÷ ${soNg(d.soNgay)} ngày × ${d.ngayThang}.`, { size: 12, fill: C.phu }));
  const H = y + PAD - 4;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT_FAMILY_SVG}"><rect width="${W}" height="${H}" fill="#FFFFFF"/>${p.join('')}</svg>`;
}

function laTriggerThst(text) {
  return coTuKhoa(text, ['bc ths', 'ths trứng', 'ths trung', 'thủy hải sản', 'thuỷ hải sản', 'thuy hai san', 'bc trứng', 'bc trung']) && !laLenhNapBaseThst(text);
}
function laLenhNapBaseThst(text) {
  const s = (text || '').normalize('NFC');
  return /n[ạa]p\s*base/i.test(s) && /(ths|trứng|trung|hải sản|hai san)/i.test(s);
}
async function veAnhThst() {
  const svg = veSvgThst(await docDuLieuThst());
  return { png: svgThanhPng(svg, 1500), preview: svgThanhPng(svg, 600) };
}
async function generateThstReport() {
  const { png, preview } = await veAnhThst();
  const id = `thst-${Date.now().toString(36)}`;
  luuAnh(id, png, preview);
  return { type: 'image', originalContentUrl: `${PUBLIC_URL}/anh/${id}.png?loai=thst`, previewImageUrl: `${PUBLIC_URL}/anh/${id}.png?loai=thst&nho=1` };
}

// ---------------------------------------------------------------------------
// BÁO CÁO DTLK — Doanh thu luỹ kế theo ngày (file "Doanh Thu Theo Siêu Thị")
//  - Gửi file tháng trước + file tháng này: bot lưu theo ngày (gửi lại ngày nào thì thay ngày đó)
//  - So sánh tháng hiện tại với cùng kỳ tháng trước; "MT DTLK 3500" (triệu) hoặc "MT DTLK 3,5" (tỷ) để đặt mục tiêu tháng
// ---------------------------------------------------------------------------
const TAB_DTST = 'DT_SIEUTHI_NGAY';
const TAB_DTLK_MT = 'DTLK_MT';
const THU_VN = ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'];

function laFileDtSieuThi(header) {
  return ['Ngày', 'Mã siêu thị', 'Doanh thu offline', 'Doanh thu Online', 'Tổng số bill'].every((c) => header.includes(c));
}
function keyNgayExcel(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
  if (v instanceof Date) return new Date(v.getTime() + 7 * 3600 * 1000).toISOString().slice(0, 10);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return taoKeyNgay(m[1], m[2], m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return taoKeyNgay(m[3], m[2], m[1]);
  return '';
}

async function napFileDtSieuThi(sheets, header, dataRows, fileName) {
  const i = (c) => header.indexOf(c);
  const c = { ngay: i('Ngày'), st: i('Mã siêu thị'), ten: i('Tên siêu thị'), off: i('Doanh thu offline'), on: i('Doanh thu Online'), bill: i('Tổng số bill'), billOn: i('Tổng số bill online') };
  const capNhat = capNhatTuTen(fileName) || bayGioVN();
  const moi = dataRows.map((r) => [keyNgayExcel(r[c.ngay]), chuanHoaMaST(r[c.st]), tenNganSieuThi(r[c.ten]), Math.round(soGV(r[c.off])), Math.round(soGV(r[c.on])), soGV(r[c.bill]), c.billOn < 0 ? 0 : soGV(r[c.billOn]), capNhat])
    .filter((r) => r[0] && r[1]);
  const khoa = new Set(moi.map((r) => `${r[0]}|${r[1]}`));
  const cu = (await docTabAnToan(sheets, TAB_DTST)).slice(1).filter((r) => !khoa.has(`${r[0]}|${r[1]}`));
  const rows = [...cu, ...moi].sort((a, b) => String(a[0]).localeCompare(String(b[0])) || String(a[1]).localeCompare(String(b[1])));
  await ghiDeTab(sheets, TAB_DTST, ['Ngày', 'Mã siêu thị', 'Tên siêu thị', 'DT offline', 'DT online', 'Số bill', 'Bill online', 'Cập nhật'], rows);
  const ngay = moi.map((r) => r[0]).sort();
  return { loai: 'dtlk', tenTab: TAB_DTST, soDong: moi.length, tu: ngay[0], den: ngay[ngay.length - 1] };
}

function thangTruoc(ym) {
  const [y, m] = ym.split('-').map(Number);
  return m === 1 ? `${y - 1}-12` : `${y}-${String(m - 1).padStart(2, '0')}`;
}

async function docDuLieuDtlk(maST) {
  const sheets = getSheetsClient();
  let rows = (await docTabAnToan(sheets, TAB_DTST)).slice(1).filter((r) => r[0]);
  if (!rows.length) throw new Error('Chưa có dữ liệu DTLK. Anh gửi file "Doanh Thu Theo Siêu Thị" (tháng trước + tháng này) vào group giúp em.');
  const dsST = [...new Set(rows.map((r) => String(r[1])))];
  if (maST && dsST.includes(maST)) rows = rows.filter((r) => String(r[1]) === maST);
  const motST = new Set(rows.map((r) => String(r[1]))).size === 1;
  const st = motST ? { ma: String(rows[0][1]), ten: String(rows[0][2]) } : { ma: `${dsST.length} ST`, ten: 'Tổng các siêu thị' };
  // gộp theo ngày
  const theoNgay = new Map();
  for (const r of rows) {
    const k = String(r[0]);
    const o = theoNgay.get(k) || { off: 0, on: 0, bill: 0, billOn: 0, cap: '' };
    o.off += soGV(r[3]); o.on += soGV(r[4]); o.bill += soGV(r[5]); o.billOn += soGV(r[6]); if (r[7]) o.cap = String(r[7]);
    theoNgay.set(k, o);
  }
  const keys = [...theoNgay.keys()].sort();
  const thang = keys[keys.length - 1].slice(0, 7), thangT = thangTruoc(thang);
  // Ngày hôm nay (đang bán dở) -> tách ra bảng HÔM NAY REALTIME, không tính vào các ngày đã chốt
  const homNay = homNayVN();
  const ngayCur = keys.filter((k) => k.startsWith(thang) && k !== homNay);
  const soNgayThang = new Date(Number(thang.slice(0, 4)), Number(thang.slice(5, 7)), 0).getDate();
  const dong = []; let lk = 0, lkT = 0, coT = false;
  for (const k of ngayCur) {
    const d = Number(k.slice(8, 10));
    const o = theoNgay.get(k);
    const kT = `${thangT}-${String(d).padStart(2, '0')}`;
    const oT = theoNgay.get(kT) || null;
    if (oT) coT = true;
    const dt = o.off + o.on, dtT = oT ? oT.off + oT.on : null;
    lk += dt; if (dtT != null) lkT += dtT;
    const thu = THU_VN[new Date(k + 'T00:00:00Z').getUTCDay()];
    dong.push({ k, d, thu, dt, on: o.on, bill: o.bill, dtb: o.bill ? dt / o.bill : 0, dtT, billT: oT ? oT.bill : null, lk, lkT: coT ? lkT : null });
  }
  // cả tháng trước
  const ngayT = keys.filter((k) => k.startsWith(thangT));
  const tongThangT = ngayT.reduce((a, k) => a + theoNgay.get(k).off + theoNgay.get(k).on, 0);
  const soNgay = ngayCur.length ? Number(ngayCur[ngayCur.length - 1].slice(8, 10)) : 0;
  const mtRow = (await docTabAnToan(sheets, TAB_DTLK_MT)).slice(1).find((r) => String(r[0]) === thang);
  const mt = mtRow ? soGV(mtRow[1]) * 1e6 : 0;
  const tong = {
    dt: lk, on: dong.reduce((a, x) => a + x.on, 0), bill: dong.reduce((a, x) => a + x.bill, 0),
    dtT: coT ? lkT : null, billT: coT ? dong.reduce((a, x) => a + (x.billT || 0), 0) : null,
  };
  let hn = null;
  if (theoNgay.has(homNay) && homNay.startsWith(thang)) {
    const o = theoNgay.get(homNay), dHN = Number(homNay.slice(8, 10));
    const oT = theoNgay.get(`${thangT}-${String(dHN).padStart(2, '0')}`) || null;
    hn = { k: homNay, thu: THU_VN[new Date(homNay + 'T00:00:00Z').getUTCDay()], dt: o.off + o.on, on: o.on, bill: o.bill, cap: o.cap, dtT: oT ? oT.off + oT.on : null, billT: oT ? oT.bill : null };
  }
  return { st, thang, thangT, dong, tong, soNgay, soNgayThang, tongThangT, coThangT: ngayT.length > 0, mt, hn };
}

function veSvgDtlk(d) {
  const W = 1320, PAD = 24, ROW = 30;
  const C = { xanh: '#0B8A3E', xanhDam: '#0B6E35', nen: '#E6F4EC', chan: '#F3F7F5', vien: '#D5E3DA', phu: '#666666', do: '#D0312D', cam: '#E08A00', vang: '#FFC20E', cuoiTuan: '#FFF6DA' };
  const t = (x, y, s, o = {}) =>
    `<text x="${x}" y="${y}" font-size="${o.size || 15}" font-weight="${o.bold ? 700 : 400}" fill="${o.fill || '#1a1a1a'}" text-anchor="${o.anchor || 'start'}">${escXml(s)}</text>`;
  const tr = (v) => fmtTrieu(v / 1e6);
  const so = (v) => String(Math.round(v)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const nghin = (v) => `${so(v / 1000)}k`;
  const tang = (a, b) => (b ? ((a - b) / b) * 100 : null);
  const ptxt = (v) => (v == null ? '–' : `${v >= 0 ? '+' : '−'}${fmtPhanTram(Math.abs(v))}`);
  const pmau = (v) => (v == null ? C.phu : v >= 0 ? C.xanh : C.do);
  const th = Number(d.thang.slice(5, 7)), thT = Number(d.thangT.slice(5, 7));
  const dd = (k) => `${k.slice(8, 10)}/${k.slice(5, 7)}`;
  const den = d.dong.length ? d.dong[d.dong.length - 1].k : d.thang + '-01';
  const p = [];
  p.push(`<rect x="0" y="0" width="${W}" height="96" fill="${C.xanh}"/>`);
  p.push(t(PAD, 38, `BÁO CÁO DTLK · DOANH THU LUỸ KẾ T${th}`, { size: 26, bold: true, fill: '#FFFFFF' }));
  p.push(t(PAD, 64, `So với T${thT} cùng kỳ (01–${String(d.soNgay).padStart(2, '0')}/${String(thT).padStart(2, '0')}) · DT chưa VAT (offline + online)`, { size: 15, fill: '#E3F5EA' }));
  p.push(t(PAD, 86, `Dữ liệu đến hết ${dd(den)} · ${d.soNgay}/${d.soNgayThang} ngày · ĐVT: triệu`, { size: 13, fill: '#E3F5EA' }));
  { const bw = 300, bx = W - PAD - bw, by = 16;
    p.push(`<rect x="${bx}" y="${by}" width="${bw}" height="64" rx="10" fill="#FFFFFF"/>`);
    p.push(`<rect x="${bx}" y="${by}" width="8" height="64" rx="4" fill="${C.vang}"/>`);
    p.push(t(bx + bw / 2 + 4, by + 24, `SIÊU THỊ ${d.st.ma}`, { size: 13, bold: true, fill: C.phu, anchor: 'middle' }));
    p.push(t(bx + bw / 2 + 4, by + 50, rutGonTen(d.st.ten, 26), { size: 18, bold: true, fill: C.xanh, anchor: 'middle' })); }
  let y = 112;

  // Ô tổng
  const T = d.tong, n = d.soNgay || 1;
  const dk = (T.dt / n) * d.soNgayThang;
  const o = [
    { ten: `DT luỹ kế T${th}`, so: tr(T.dt), phu: T.dtT != null ? `T${thT} cùng kỳ ${tr(T.dtT)}` : '', pct: T.dtT != null ? tang(T.dt, T.dtT) : null },
    { ten: 'TB / ngày', so: tr(T.dt / n), phu: T.dtT != null ? `T${thT} cùng kỳ ${tr(T.dtT / n)}` : '', pct: T.dtT != null ? tang(T.dt, T.dtT) : null },
    { ten: 'Bill / ngày · DT / bill', so: `${so(T.bill / n)} · ${nghin(T.bill ? T.dt / T.bill : 0)}`, phu: T.billT ? `T${thT}: ${so(T.billT / n)} · ${nghin(T.dtT / T.billT)}` : '', pct: T.billT ? tang(T.bill, T.billT) : null, nhanPct: 'bill' },
    { ten: `Dự kiến cả T${th}`, so: tr(dk), phu: d.coThangT ? `Cả T${thT} ${tr(d.tongThangT)}` : 'nếu giữ nhịp', pct: d.coThangT ? tang(dk, d.tongThangT) : null },
  ];
  if (d.mt > 0) {
    const conLai = Math.max(1, d.soNgayThang - d.soNgay);
    o.push({ ten: `Mục tiêu T${th} ${tr(d.mt)}`, so: fmtPhanTram((T.dt / d.mt) * 100), phu: `Cần ${tr(Math.max(0, d.mt - T.dt) / conLai)}/ngày`, pct: null, mau: (T.dt / d.mt) >= d.soNgay / d.soNgayThang ? C.xanh : C.do });
  }
  const oW = (W - PAD * 2 - 12 * (o.length - 1)) / o.length, oH = 92;
  o.forEach((x, k) => {
    const ox = PAD + k * (oW + 12);
    p.push(`<rect x="${ox}" y="${y}" width="${oW}" height="${oH}" rx="10" fill="#F2F8F4"/>`);
    p.push(t(ox + 14, y + 22, x.ten, { size: 13, fill: C.phu }));
    p.push(t(ox + 14, y + 52, x.so, { size: o.length > 4 ? 21 : 24, bold: true, fill: x.mau || '#1a1a1a' }));
    if (x.pct != null) p.push(t(ox + oW - 14, y + 52, ptxt(x.pct), { size: 15, bold: true, fill: pmau(x.pct), anchor: 'end' }));
    if (x.phu) p.push(t(ox + 14, y + 78, x.phu, { size: 12, fill: x.mau && !x.pct ? C.do : C.phu }));
  });
  y += oH + 16;

  // Ô HÔM NAY REALTIME
  if (d.hn) {
    const h = d.hn, tb = T.dt / n, H0 = 112;
    p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${H0}" rx="10" fill="#FFF8E6" stroke="${C.vang}" stroke-width="2"/>`);
    p.push(`<rect x="${PAD}" y="${y}" width="250" height="${H0}" rx="10" fill="${C.cam}"/>`);
    p.push(t(PAD + 18, y + 30, 'HÔM NAY · REALTIME', { size: 15, bold: true, fill: '#FFFFFF' }));
    p.push(t(PAD + 18, y + 58, `${h.thu} ${dd(h.k)}`, { size: 26, bold: true, fill: '#FFFFFF' }));
    p.push(t(PAD + 18, y + 86, `cập nhật ${h.cap ? h.cap.split(' ')[0] : '–'}`, { size: 14, fill: '#FFFFFF' }));
    const conLai = Math.max(1, d.soNgayThang - d.soNgay);
    const can = d.mt > 0 ? Math.max(0, d.mt - T.dt) / conLai : null;
    const o2 = [
      { ten: 'DT hôm nay', so: tr(h.dt), dam: true },
      { ten: 'Số bill · DT/bill', so: `${so(h.bill)} · ${nghin(h.bill ? h.dt / h.bill : 0)}` },
      { ten: 'Trong đó online', so: tr(h.on) },
      { ten: `So TB/ngày T${th} (${tr(tb)})`, so: fmtPhanTram(tb ? (h.dt / tb) * 100 : 0), mau: h.dt >= tb ? C.xanh : C.cam },
      { ten: h.dtT != null ? `T${thT} ${dd(`${d.thangT}-${h.k.slice(8, 10)}`)} cả ngày ${tr(h.dtT)}` : `T${thT} cùng ngày`, so: h.dtT != null ? fmtPhanTram((h.dt / h.dtT) * 100) : '–', mau: h.dtT != null && h.dt >= h.dtT ? C.xanh : C.cam },
      can != null ? { ten: `Cần/ngày (MT) ${tr(can)}`, so: fmtPhanTram(can ? (h.dt / can) * 100 : 100), mau: h.dt >= can ? C.xanh : C.cam } : { ten: `LK tính cả hôm nay`, so: tr(T.dt + h.dt) },
    ];
    const x0 = PAD + 266, cw = (W - PAD - 12 - x0) / o2.length;
    o2.forEach((c, k) => {
      const cx = x0 + k * cw;
      if (k) p.push(`<line x1="${cx - 6}" y1="${y + 18}" x2="${cx - 6}" y2="${y + H0 - 18}" stroke="#F0D9A0"/>`);
      p.push(t(cx + 4, y + 36, rutGonTen(c.ten, 26), { size: 12, fill: C.phu }));
      p.push(t(cx + 4, y + 74, c.so, { size: c.dam ? 30 : 22, bold: true, fill: c.mau || (c.dam ? C.cam : '#1a1a1a') }));
    });
    y += H0 + 16;
  }

  // 2 bảng cạnh nhau: HIỆN TẠI (theo ngày) | LUỸ KẾ
  const L = { x0: PAD, x1: PAD + 744 }, R = { x0: PAD + 760, x1: W - PAD };
  const cL = { ngay: L.x0 + 14, dt: L.x0 + 200, bill: L.x0 + 270, dtb: L.x0 + 350, on: L.x0 + 440, dtT: L.x0 + 548, billT: L.x0 + 622, pct: L.x1 - 12 };
  const cR = { ngay: R.x0 + 14, lk: R.x0 + 190, lkT: R.x0 + 300, pct: R.x0 + 400, dk: R.x1 - 12 };
  const yBang = y;
  const dauBang = (B, ten, phu) => {
    p.push(`<rect x="${B.x0}" y="${y}" width="${B.x1 - B.x0}" height="72" rx="6" fill="${C.nen}"/>`);
    p.push(`<rect x="${B.x0 + 6}" y="${y + 6}" width="${B.x1 - B.x0 - 12}" height="22" rx="4" fill="${C.xanh}"/>`);
    p.push(t((B.x0 + B.x1) / 2, y + 22, ten, { size: 13, bold: true, fill: '#FFFFFF', anchor: 'middle' }));
  };
  const td = (x, a1, a2, al) => { p.push(t(x, y + 48, a1, { size: 13, bold: true, fill: C.xanhDam, anchor: al || 'end' })); if (a2) p.push(t(x, y + 65, a2, { size: 13, bold: true, fill: C.xanhDam, anchor: al || 'end' })); };
  dauBang(L, `HIỆN TẠI · DOANH THU TỪNG NGÀY T${th} so với T${thT} cùng ngày`);
  dauBang(R, `LUỸ KẾ T${th} so với T${thT} cùng kỳ`);
  td(cL.ngay, 'Ngày', '', 'start'); td(cL.dt, 'DT', 'ngày'); td(cL.bill, 'Số', 'bill'); td(cL.dtb, 'DT/', 'bill'); td(cL.on, 'Trong đó', 'online');
  td(cL.dtT, 'DT', `T${thT}`); td(cL.billT, 'Bill', `T${thT}`); td(cL.pct, '± DT', 'ngày');
  td(cR.ngay, 'Đến', 'ngày', 'start'); td(cR.lk, 'LK', `T${th}`); td(cR.lkT, 'LK', `T${thT}`); td(cR.pct, '±', 'LK'); td(cR.dk, 'Dự kiến', 'cả tháng');
  y += 72;
  // dòng tổng mỗi bảng
  p.push(`<rect x="${L.x0}" y="${y}" width="${L.x1 - L.x0}" height="${ROW + 2}" fill="${C.xanh}"/>`);
  p.push(`<rect x="${R.x0}" y="${y}" width="${R.x1 - R.x0}" height="${ROW + 2}" fill="${C.xanh}"/>`);
  const w = { fill: '#FFFFFF', bold: true, anchor: 'end' };
  p.push(t(cL.ngay, y + 21, `TỔNG ${d.soNgay} NGÀY`, { fill: '#FFFFFF', bold: true }));
  p.push(t(cL.dt, y + 21, tr(T.dt), w)); p.push(t(cL.bill, y + 21, so(T.bill), w)); p.push(t(cL.dtb, y + 21, nghin(T.bill ? T.dt / T.bill : 0), w)); p.push(t(cL.on, y + 21, tr(T.on), w));
  p.push(t(cL.dtT, y + 21, T.dtT != null ? tr(T.dtT) : '–', w)); p.push(t(cL.billT, y + 21, T.billT != null ? so(T.billT) : '–', w)); p.push(t(cL.pct, y + 21, ptxt(tang(T.dt, T.dtT)), w));
  p.push(t(cR.ngay, y + 21, `HẾT ${dd(den)}`, { fill: '#FFFFFF', bold: true }));
  p.push(t(cR.lk, y + 21, tr(T.dt), w)); p.push(t(cR.lkT, y + 21, T.dtT != null ? tr(T.dtT) : '–', w)); p.push(t(cR.pct, y + 21, ptxt(tang(T.dt, T.dtT)), w)); p.push(t(cR.dk, y + 21, tr(dk), w));
  y += ROW + 2;
  const yDong0 = y;
  d.dong.forEach((x, k) => {
    const cuoiTuan = x.thu === 'T7' || x.thu === 'CN';
    for (const B of [L, R]) {
      if (cuoiTuan) p.push(`<rect x="${B.x0}" y="${y}" width="${B.x1 - B.x0}" height="${ROW}" fill="${C.cuoiTuan}"/>`);
      else if (k % 2 === 1) p.push(`<rect x="${B.x0}" y="${y}" width="${B.x1 - B.x0}" height="${ROW}" fill="${C.chan}"/>`);
    }
    const yt = y + 20, pd = tang(x.dt, x.dtT), pl = x.lkT != null ? tang(x.lk, x.lkT) : null;
    const nhanNgay = { bold: cuoiTuan, fill: cuoiTuan ? C.cam : '#1a1a1a' };
    p.push(t(cL.ngay, yt, `${x.thu} ${dd(x.k)}`, nhanNgay));
    p.push(t(cL.dt, yt, tr(x.dt), { bold: true, anchor: 'end' }));
    p.push(t(cL.bill, yt, so(x.bill), { anchor: 'end' }));
    p.push(t(cL.dtb, yt, nghin(x.dtb), { anchor: 'end', fill: C.phu }));
    p.push(t(cL.on, yt, tr(x.on), { anchor: 'end', fill: C.phu }));
    p.push(t(cL.dtT, yt, x.dtT != null ? tr(x.dtT) : '–', { anchor: 'end', fill: C.phu }));
    p.push(t(cL.billT, yt, x.billT != null ? so(x.billT) : '–', { anchor: 'end', fill: C.phu }));
    p.push(t(cL.pct, yt, ptxt(pd), { bold: true, fill: pmau(pd), anchor: 'end' }));
    p.push(t(cR.ngay, yt, `${x.thu} ${dd(x.k)}`, nhanNgay));
    p.push(t(cR.lk, yt, tr(x.lk), { bold: true, anchor: 'end' }));
    p.push(t(cR.lkT, yt, x.lkT != null ? tr(x.lkT) : '–', { anchor: 'end', fill: C.phu }));
    p.push(t(cR.pct, yt, ptxt(pl), { bold: true, fill: pmau(pl), anchor: 'end' }));
    p.push(t(cR.dk, yt, tr((x.lk / x.d) * d.soNgayThang), { anchor: 'end', fill: C.phu }));
    y += ROW;
  });
  for (const B of [L, R]) p.push(`<rect x="${B.x0}" y="${yBang}" width="${B.x1 - B.x0}" height="${y - yBang}" rx="6" fill="none" stroke="${C.vien}"/>`);
  y += 24;
  p.push(t(PAD, y, `So cùng ngày trong tháng (01/${String(th).padStart(2, '0')} ↔ 01/${String(thT).padStart(2, '0')}) · Nền vàng = thứ 7, CN · Dự kiến cả tháng = LK ÷ số ngày × ${d.soNgayThang}.`, { size: 12, fill: C.phu }));
  const H = y + PAD - 4;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT_FAMILY_SVG}"><rect width="${W}" height="${H}" fill="#FFFFFF"/>${p.join('')}</svg>`;
}

function laTriggerDtlk(text) {
  return coTuKhoa(text, ['dtlk', 'bc dtlk', 'báo cáo dtlk', 'doanh thu luỹ kế', 'doanh thu lũy kế', 'doanh thu luy ke']) && !laLenhMtDtlk(text);
}
function laLenhMtDtlk(text) {
  return /mt\s*dtlk\s*[\d.,]+/i.test((text || '').normalize('NFC'));
}
async function datMtDtlk(text) {
  const m = (text || '').normalize('NFC').match(/mt\s*dtlk\s*([\d.,]+)\s*(tỷ|ty|tr|triệu)?/i);
  if (!m) return null;
  let v = Number(m[1].replace(/\.(?=\d{3}\b)/g, '').replace(',', '.'));
  if (!isFinite(v) || v <= 0) return null;
  const donVi = (m[2] || '').toLowerCase();
  if (donVi.startsWith('t') && /tỷ|ty/.test(donVi)) v *= 1000; else if (!donVi && v < 100) v *= 1000; // "3,5" = 3,5 tỷ
  const sheets = getSheetsClient();
  const rows = (await docTabAnToan(sheets, TAB_DTST)).slice(1);
  const thang = rows.length ? rows.map((r) => String(r[0])).sort().pop().slice(0, 7) : homNayVN().slice(0, 7);
  const cu = (await docTabAnToan(sheets, TAB_DTLK_MT)).slice(1).filter((r) => String(r[0]) !== thang);
  await ghiDeTab(sheets, TAB_DTLK_MT, ['Tháng', 'MT (triệu)'], [...cu, [thang, Math.round(v * 10) / 10]]);
  return { type: 'text', text: `✅ Đã đặt mục tiêu DTLK tháng ${thang.slice(5, 7)}/${thang.slice(0, 4)} = ${fmtTrieu(v)} tr. Gõ "DTLK" để xem.` };
}
async function veAnhDtlk(maST) {
  const svg = veSvgDtlk(await docDuLieuDtlk(maST));
  return { png: svgThanhPng(svg, 1800), preview: svgThanhPng(svg, 600) };
}
async function generateDtlkReport(text) {
  const m = String(text || '').match(/\b(\d{4,5})\b/);
  const maST = m ? m[1] : '';
  const { png, preview } = await veAnhDtlk(maST);
  const id = `dtlk-${Date.now().toString(36)}`;
  luuAnh(id, png, preview);
  const q = `loai=dtlk${maST ? `&st=${maST}` : ''}`;
  return { type: 'image', originalContentUrl: `${PUBLIC_URL}/anh/${id}.png?${q}`, previewImageUrl: `${PUBLIC_URL}/anh/${id}.png?${q}&nho=1` };
}

// ---- Ảnh vuông "DT HÔM NAY" cho nhân viên: ít số, dễ đọc ----
function veSvgDtHomNay(d) {
  // Bảng kẻ ô vuông, ít dòng: Chỉ tiêu | Hôm nay | T9 cùng ngày | So sánh
  const W = 1000, PAD = 28, ROW = 64;
  const C = { xanh: '#0B8A3E', xanhDam: '#0B6E35', nen: '#E6F4EC', vien: '#9DBFAA', phu: '#666666', do: '#D0312D', cam: '#E08A00', vang: '#FFC20E', nhat: '#F6FAF7' };
  const t = (x, y, s, o = {}) =>
    `<text x="${x}" y="${y}" font-size="${o.size || 22}" font-weight="${o.bold ? 700 : 400}" fill="${o.fill || '#1a1a1a'}" text-anchor="${o.anchor || 'start'}">${escXml(s)}</text>`;
  const tr = (v) => fmtTrieu(v / 1e6);
  const so = (v) => String(Math.round(v)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const k = (v) => `${so(v / 1000)}k`;
  const h = d.hn, T = d.tong, n = d.soNgay || 1;
  const th = Number(d.thang.slice(5, 7)), thT = Number(d.thangT.slice(5, 7));
  const dd = (x) => `${x.slice(8, 10)}/${x.slice(5, 7)}`;
  const ngayT = `${d.thangT}-${h.k.slice(8, 10)}`;
  const tb = T.dt / n, conLai = Math.max(1, d.soNgayThang - d.soNgay);
  const mucTieu = d.mt > 0 ? Math.max(0, d.mt - T.dt) / conLai : tb;
  const pDat = mucTieu > 0 ? (h.dt / mucTieu) * 100 : 0;
  const pct = (a, b) => (b ? fmtPhanTram((a / b) * 100) : '–');
  const lk = T.dt + h.dt, lkT = T.dtT != null && h.dtT != null ? T.dtT + h.dtT : null;
  const p = [];
  // Header
  p.push(`<rect x="0" y="0" width="${W}" height="118" fill="${C.xanh}"/>`);
  p.push(t(PAD, 52, 'DOANH THU HÔM NAY', { size: 34, bold: true, fill: '#FFFFFF' }));
  p.push(t(PAD, 92, `${h.thu} ${dd(h.k)} · cập nhật ${h.cap ? h.cap.split(' ')[0] : '–'} · ĐVT: triệu`, { size: 21, fill: '#E3F5EA' }));
  { const bw = 330, bx = W - PAD - bw, by = 22;
    p.push(`<rect x="${bx}" y="${by}" width="${bw}" height="74" rx="10" fill="#FFFFFF"/>`);
    p.push(`<rect x="${bx}" y="${by}" width="8" height="74" rx="4" fill="${C.vang}"/>`);
    p.push(t(bx + bw / 2 + 4, by + 28, `SIÊU THỊ ${d.st.ma}`, { size: 16, bold: true, fill: C.phu, anchor: 'middle' }));
    p.push(t(bx + bw / 2 + 4, by + 58, rutGonTen(d.st.ten, 26), { size: 21, bold: true, fill: C.xanh, anchor: 'middle' })); }
  let y = 146;
  // Bảng
  const x = [PAD, PAD + 300, PAD + 520, PAD + 740, W - PAD]; // mép các cột
  const giua = (i) => (x[i] + x[i + 1]) / 2;
  const tb0 = y;
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW}" fill="${C.xanhDam}"/>`);
  [['CHỈ TIÊU', ''], ['HÔM NAY', dd(h.k)], [`T${thT} CÙNG NGÀY`, `${dd(ngayT)} · cả ngày`], ['SO VỚI', `T${thT} cùng ngày`]].forEach(([a1, a2], i) => {
    p.push(t(i === 0 ? x[0] + 22 : giua(i), y + (a2 ? 28 : 40), a1, { size: 20, bold: true, fill: '#FFFFFF', anchor: i === 0 ? 'start' : 'middle' }));
    if (a2) p.push(t(giua(i), y + 52, a2, { size: 16, fill: '#D6EEDD', anchor: 'middle' }));
  });
  y += ROW;
  const dong = [
    { ten: 'Doanh thu', hn: tr(h.dt), t9: h.dtT != null ? tr(h.dtT) : '–', ss: h.dtT != null ? pct(h.dt, h.dtT) : '–', dam: true },
    { ten: 'Số bill', hn: so(h.bill), t9: h.billT != null ? so(h.billT) : '–', ss: h.billT ? pct(h.bill, h.billT) : '–' },
    { ten: 'Tiền TB / bill', hn: k(h.bill ? h.dt / h.bill : 0), t9: h.billT ? k(h.dtT / h.billT) : '–', ss: h.billT && h.bill ? pct(h.dt / h.bill, h.dtT / h.billT) : '–' },
    { ten: 'Online', hn: tr(h.on), t9: '', ss: '' },
    { ten: `Luỹ kế T${th} (đến hôm nay)`, hn: tr(lk), t9: lkT != null ? tr(lkT) : '–', ss: lkT != null ? pct(lk, lkT) : '–' },
  ];
  dong.forEach((r, i) => {
    p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW}" fill="${i % 2 ? C.nhat : '#FFFFFF'}"/>`);
    p.push(t(x[0] + 22, y + 41, r.ten, { size: 22, bold: r.dam }));
    p.push(t(giua(1), y + 42, r.hn, { size: r.dam ? 32 : 26, bold: true, fill: r.dam ? C.xanhDam : '#1a1a1a', anchor: 'middle' }));
    p.push(t(giua(2), y + 42, r.t9, { size: 24, fill: C.phu, anchor: 'middle' }));
    p.push(t(giua(3), y + 42, r.ss, { size: 24, bold: true, anchor: 'middle' }));
    y += ROW;
  });
  // Kẻ ô
  for (let yy = tb0 + ROW; yy < y - 1; yy += ROW) p.push(`<line x1="${PAD}" y1="${yy}" x2="${W - PAD}" y2="${yy}" stroke="${C.vien}"/>`);
  for (let i = 1; i < x.length - 1; i++) p.push(`<line x1="${x[i]}" y1="${tb0}" x2="${x[i]}" y2="${y}" stroke="${C.vien}"/>`);
  p.push(`<rect x="${PAD}" y="${tb0}" width="${W - PAD * 2}" height="${y - tb0}" fill="none" stroke="${C.xanhDam}" stroke-width="2"/>`);
  y += 34;
  p.push(t(PAD, y, `Hôm nay đang bán dở nên % so với T${thT} (cả ngày) sẽ tăng dần đến cuối ngày.`, { size: 17, fill: C.phu }));
  const H = y + 24;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT_FAMILY_SVG}"><rect width="${W}" height="${H}" fill="#FFFFFF"/>${p.join('')}</svg>`;
}
function laTriggerDtHomNay(text) {
  return coTuKhoa(text, ['dt hôm nay', 'dt hom nay', 'doanh thu hôm nay', 'doanh thu hom nay', 'bc hôm nay', 'hôm nay realtime']);
}
async function veAnhDtHomNay(maST) {
  const d = await docDuLieuDtlk(maST);
  if (!d.hn) throw new Error('Chưa có số hôm nay. Anh xuất file "Doanh Thu Theo Siêu Thị" chọn đến ngày hôm nay rồi gửi vào group.');
  const svg = veSvgDtHomNay(d);
  return { png: svgThanhPng(svg, 1400), preview: svgThanhPng(svg, 600) };
}
async function generateDtHomNayReport(text) {
  const m = String(text || '').match(/\b(\d{4,5})\b/);
  const maST = m ? m[1] : '';
  const { png, preview } = await veAnhDtHomNay(maST);
  const id = `hn-${Date.now().toString(36)}`;
  luuAnh(id, png, preview);
  const q = `loai=hn${maST ? `&st=${maST}` : ''}`;
  return { type: 'image', originalContentUrl: `${PUBLIC_URL}/anh/${id}.png?${q}`, previewImageUrl: `${PUBLIC_URL}/anh/${id}.png?${q}&nho=1` };
}

// "ĐẶT MỨC T3 M3" (hoặc "ĐẶT MỨC 3" = áp cho tuần đang chạy theo ngày hôm nay)
async function datMucThiDua(text) {
  const t = text.normalize('NFC').toUpperCase();
  const m = t.match(/Đ[ẶA]T\s*M[ỨU]C\s*(?:(T\d)\s*)?M?\s*([123])\b/);
  if (!m) return null;
  const sheets = getSheetsClient();
  let tuan = m[1];
  if (!tuan) {
    const homNay = homNayVN();
    const r = (await docTabAnToan(sheets, TAB_MUCTIEU)).slice(1).find((x) => chuanHoaKeyNgay(x[1]) <= homNay && homNay <= chuanHoaKeyNgay(x[2]));
    if (!r) throw new Error('Không xác định được tuần đang chạy, anh gõ rõ tuần, vd "ĐẶT MỨC T3 M3".');
    tuan = String(r[0]);
  }
  const cu = (await docTabAnToan(sheets, TAB_CAUHINH)).slice(1).filter((r) => String(r[0]) !== tuan);
  await ghiDeTab(sheets, TAB_CAUHINH, ['Tuần', 'Mức'], [...cu, [tuan, Number(m[2])]]);
  return { type: 'text', text: `✅ Đã đặt tuần ${tuan.slice(1)} chạy theo MỨC ${m[2]} (M${m[2]}). Gõ "DT NGÀNH HÀNG" để xem.` };
}

async function chayLenh(text) {
  text = (text || '').normalize('NFC');
  const datMuc = await datMucThiDua(text);
  if (datMuc) return { ten: 'Đặt mức', ket: datMuc };
  if (laLenhMtDtlk(text)) { const kq = await datMtDtlk(text); if (kq) return { ten: 'MT DTLK', ket: kq }; }
  if (laTriggerDtHomNay(text)) return { ten: 'DT Hôm nay', ket: await generateDtHomNayReport(text) };
  if (laTriggerDtlk(text)) return { ten: 'Báo Cáo DTLK', ket: await generateDtlkReport(text) };
  if (laTriggerDtNganhHang(text)) return { ten: 'DT Ngành Hàng', ket: await generateDtNganhHangReport(text) };
  if (laTriggerGiaVon(text)) return { ten: 'Giá Vốn', ket: await generateGiaVonReport(text) };
  if (laTriggerHuyMmkk(text)) return { ten: 'Huỷ MMKK', ket: await generateHuyMmkkReport(text) };
  if (laTriggerSua(text)) return { ten: 'BC Sữa', ket: await generateSuaReport() };
  if (laTriggerThst(text)) return { ten: 'BC THS Trứng', ket: await generateThstReport() };
  if (laTriggerBnnb(text)) return { ten: 'Báo Cáo BNNB', ket: await generateBnnbReport() };
  if (laTriggerDmsk(text)) return { ten: 'BC ĐMSK', ket: await generateDmskReport() };
  return null;
}

// ---------------------------------------------------------------------------
// AI (CLAUDE) — đoán tab liên quan + phân tích bảng dữ liệu / ảnh
// ---------------------------------------------------------------------------
const KNOWN_TABS = [
  { name: TAB_MUCTIEU, desc: 'Mục tiêu thi đua theo tuần (M1, M2, M3) của từng siêu thị theo nhóm ngành hàng' },
  { name: TAB_DT_NGAY, desc: 'Doanh thu theo ngày của từng siêu thị theo từng ngành hàng' },
  { name: TAB_FRESH_BAN, desc: 'Doanh thu bán Fresh hôm nay theo siêu thị, theo ngành giá vốn' },
  { name: TAB_SUA_TON, desc: 'Tồn kho các mã sữa theo dõi theo siêu thị' },
  { name: TAB_SUA_BAN, desc: 'Số lượng bán các mã sữa theo dõi theo siêu thị' },
  { name: TAB_HUYMMKK, desc: 'Số lượng bán, giảm giá, huỷ, mất mát kiểm kê từng sản phẩm Fresh theo ngày' },
  { name: TAB_GIAVON, desc: 'Giá vốn, doanh thu, lợi nhuận luỹ kế Fresh theo ngành hàng, so với TB 3 tháng trước' },
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
      { role: 'user', content: `Danh sách tab hiện có:\n${tabList}\n\nCâu hỏi của quản lý: "${question}"\n\nChọn 1 tab phù hợp nhất. Nếu không tab nào phù hợp, trả "tab": null.` },
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
    rows.length - 1 > maxRows ? `\n(Lưu ý: dữ liệu có ${rows.length - 1} dòng, chỉ gửi ${maxRows} dòng đầu để phân tích)` : '';
  const csvText = [header, ...dataRows].map((r) => r.map((c) => (c === undefined || c === null ? '' : c)).join(',')).join('\n');

  const msg = await anthropic.messages.create({
    model: AI_MODEL,
    max_tokens: 1024,
    system:
      'Bạn là trợ lý phân tích dữ liệu bán lẻ cho quản lý cửa hàng Bách Hoá Xanh. ' +
      'Trả lời ngắn gọn, đi thẳng vào số liệu và nhận xét thực tế, dùng tiếng Việt, ' +
      'dùng gạch đầu dòng cho dễ đọc trên LINE (không dùng bảng markdown).',
    messages: [{ role: 'user', content: `Nguồn dữ liệu: ${sourceLabel}${truncatedNote}\n\nDữ liệu (CSV):\n${csvText}\n\nCâu hỏi: ${question}` }],
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
  const ws = wb.Sheets[sheetName || wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
}

// ---------------------------------------------------------------------------
// LINE BOT
// ---------------------------------------------------------------------------
const app = express();
const client = new line.Client(config);
let botUserId = null;

async function taiNoiDungFileLine(messageId) {
  const res = await fetch(`https://api-data.line.me/v2/bot/message/${messageId}/content`, {
    headers: { Authorization: `Bearer ${config.channelAccessToken}` },
  });
  if (!res.ok) throw new Error(`Tai file LINE that bai: ${res.status} ${await res.text()}`);
  return Buffer.from(await res.arrayBuffer());
}

function chiTietLoiLine(err) {
  try {
    return JSON.stringify(
      (err && err.originalError && err.originalError.response && err.originalError.response.data) ||
        (err && err.response && err.response.data) || { message: err && err.message }
    );
  } catch (e) {
    return String(err && err.message);
  }
}

// Trả lời: thử reply, lỗi thì push thẳng
async function guiTraLoi(event, targetId, message) {
  try {
    await client.replyMessage(event.replyToken, message);
  } catch (err) {
    console.error('[webhook] Reply thất bại:', chiTietLoiLine(err), '-> thử push');
    try {
      await client.pushMessage(targetId, message);
    } catch (pushErr) {
      console.error('[webhook] Push cũng thất bại:', chiTietLoiLine(pushErr));
      console.error('[webhook] Kích thước JSON tin nhắn (ký tự):', JSON.stringify(message).length);
    }
  }
}

function extractMentionQuestion(event) {
  const mention = event.message?.mention;
  if (!mention || !Array.isArray(mention.mentionees)) return null;
  const isTaggingBot = mention.mentionees.some((m) => m.isSelf === true || (botUserId && m.userId === botUserId));
  if (!isTaggingBot) return null;
  let question = event.message.text;
  const sorted = [...mention.mentionees].sort((a, b) => b.index - a.index);
  for (const m of sorted) question = question.slice(0, m.index) + question.slice(m.index + m.length);
  return question.trim();
}

// Nhớ tạm "vừa @tag hỏi gì" để file/ảnh gửi tiếp trong 3 phút được AI phân tích (không lưu Sheet)
// Lệnh "NẠP NGÀY dd/mm": file Excel gửi tiếp theo (trong 3 phút) được ghi vào đúng ngày đó
// (dùng khi gửi bù file của ngày trước — mặc định bot lấy ngày theo giờ xuất file).
const cho_NapNgay = new Map(); // targetId -> { ngay, expiresAt }
function laLenhNapNgay(text) {
  const t = (text || '').normalize('NFC').toLowerCase().trim();
  const m = t.match(/^n[ạa]p\s*(?:file\s*)?ng[àa]y\s*(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?/);
  if (!m) return null;
  const key = taoKeyNgay(m[3] || homNayVN().slice(0, 4), m[2], m[1]);
  return isNaN(Date.parse(key)) ? null : key;
}

const cho_AI_PhanTich = new Map();
const THOI_GIAN_CHO_MS = 3 * 60 * 1000;
function datCoDangChoFile(targetId, question) {
  cho_AI_PhanTich.set(targetId, { question, expiresAt: Date.now() + THOI_GIAN_CHO_MS });
}
function layVaXoaCoDangCho(targetId) {
  const info = cho_AI_PhanTich.get(targetId);
  if (!info) return null;
  cho_AI_PhanTich.delete(targetId);
  return Date.now() > info.expiresAt ? null : info;
}

async function traLoiCauHoiAI(event, targetId, question) {
  if (!AI_ENABLED) {
    console.log('[webhook] câu hỏi tự do nhưng AI đang tắt -> gửi danh sách lệnh');
    await guiTraLoi(event, targetId, {
      type: 'text',
      text:
        'Em chưa hiểu lệnh này ạ. Anh @tag em và gõ 1 trong các lệnh:\n' +
        '• DT NGÀNH HÀNG (thêm T3 hoặc ngày, vd "DT NGÀNH HÀNG T3")\n' +
        '• GIÁ VỐN\n' +
        '• HUỶ MMKK\n' +
        '• BC SỮA\n' +
        '• BC ĐMSK (NẠP BASE ĐMSK để nạp tháng 9)\n' +
        '• Báo Cáo BNNB (NẠP BASE BNNB để nạp tháng 9)\n' +
        '• BC THS TRỨNG (NẠP BASE THS để nạp tháng 9)\n' +
        '• DTLK (doanh thu luỹ kế; MT DTLK 3500 để đặt mục tiêu)\n' +
        '• DT HÔM NAY (ảnh vuông số hôm nay cho nhân viên)\n' +
        '• ĐẶT MỨC T3 M3\n' +
        '• NẠP NGÀY 05/10 (rồi gửi file bù trong 3 phút)\n' +
        'Hoặc gửi file Excel vào group, em tự nhận và trả báo cáo.',
    });
    return;
  }
  try {
    await client.replyMessage(event.replyToken, { type: 'text', text: 'Anh chờ chút, em đang tra cứu và phân tích...' });
    const picked = await pickRelevantTab(question);
    if (!picked.tab) {
      await client.pushMessage(targetId, { type: 'text', text: 'Em chưa xác định được câu hỏi này liên quan tab dữ liệu nào, anh hỏi cụ thể hơn giúp em nhé.' });
      return;
    }
    const rows = await docTabThanhMangDong(getSheetsClient(), picked.tab);
    const answer = await analyzeTableData(question, picked.tab, rows);
    await client.pushMessage(targetId, { type: 'text', text: answer });
  } catch (err) {
    console.error('[webhook] Lỗi phân tích câu hỏi:', err);
    try {
      await client.pushMessage(targetId, { type: 'text', text: `❌ Em gặp lỗi khi phân tích: ${err.message}` });
    } catch (e2) {
      console.error('[webhook] Lỗi luôn cả khi push lỗi:', chiTietLoiLine(e2));
    }
  }
}

app.post('/webhook', line.middleware(config), async (req, res) => {
  res.status(200).end();
  const events = req.body.events || [];
  console.log(`[webhook] nhận ${events.length} event(s)`);

  for (const event of events) {
    if (event.type !== 'message') continue;
    const isGroup = event.source?.type === 'group' || event.source?.type === 'room';
    const targetId = event.source?.groupId || event.source?.roomId || event.source?.userId;

    // ---------------- ẢNH -> AI vision ----------------
    if (event.message.type === 'image') {
      if (!AI_ENABLED) continue;
      try {
        const buffer = await taiNoiDungFileLine(event.message.id);
        const cauHoiCho = isGroup ? layVaXoaCoDangCho(targetId) : null;
        const answer = await analyzeImage(cauHoiCho?.question || '', buffer, 'image/jpeg');
        await guiTraLoi(event, targetId, { type: 'text', text: answer });
      } catch (err) {
        console.error('[webhook] Lỗi phân tích ảnh:', err);
        await guiTraLoi(event, targetId, { type: 'text', text: `❌ Em xem ảnh này bị lỗi: ${err.message}` });
      }
      continue;
    }

    // ---------------- FILE EXCEL ----------------
    if (event.message.type === 'file') {
      const fileName = event.message.fileName || '';
      if (!/\.(xlsx|xls)$/i.test(fileName)) continue;

      // Có lệnh "NẠP NGÀY dd/mm" trước đó -> ghi file vào đúng ngày chỉ định
      let ngayChiDinh = null;
      const napNgay = cho_NapNgay.get(targetId);
      if (napNgay) {
        cho_NapNgay.delete(targetId);
        if (Date.now() <= napNgay.expiresAt) ngayChiDinh = napNgay.ngay;
      }

      let laBaseDmsk = false;
      const hanBase = cho_NapBaseDmsk.get(targetId);
      if (hanBase) { cho_NapBaseDmsk.delete(targetId); laBaseDmsk = Date.now() <= hanBase; }
      // Base BNNB: giữ cờ trong 3 phút để nhận được cả 2 file (Bia nước + Gia vị)
      const hanBnnb = cho_NapBaseBnnb.get(targetId);
      const laBaseBnnb = !!hanBnnb && Date.now() <= hanBnnb;
      if (hanBnnb && !laBaseBnnb) cho_NapBaseBnnb.delete(targetId);
      const hanThst = cho_NapBaseThst.get(targetId);
      const laBaseThst = !!hanThst && Date.now() <= hanThst;
      if (hanThst && !laBaseThst) cho_NapBaseThst.delete(targetId);

      const cauHoiCho = !ngayChiDinh && !laBaseDmsk && !laBaseBnnb && !laBaseThst && isGroup ? layVaXoaCoDangCho(targetId) : null;
      if (cauHoiCho && AI_ENABLED) {
        try {
          const buffer = await taiNoiDungFileLine(event.message.id);
          const rows = readExcelBufferAsRows(buffer);
          const answer = await analyzeTableData(
            cauHoiCho.question || 'Tóm tắt và nhận xét nhanh các điểm đáng chú ý trong file này giúp anh.',
            `File Excel đính kèm: ${fileName}`,
            rows
          );
          await guiTraLoi(event, targetId, { type: 'text', text: answer });
        } catch (err) {
          console.error('[webhook] Lỗi AI phân tích file:', err);
          await guiTraLoi(event, targetId, { type: 'text', text: `❌ Em đọc file này bị lỗi: ${err.message}` });
        }
        continue;
      }

      console.log(`[webhook] nhận file "${fileName}", đang nạp vào Sheet...`);
      try {
        const buffer = await taiNoiDungFileLine(event.message.id);
        const kq = await napFileVaoSheet(fileName, buffer, ngayChiDinh, laBaseDmsk, laBaseBnnb, laBaseThst);
        console.log(`[webhook] đã nạp ${kq.soDong} dòng vào tab "${kq.tenTab}" (loại: ${kq.loai})`);

        if (kq.loai === 'muctieu') {
          const ct = kq.chiTiet;
          let msg = `✅ Đã nạp MỤC TIÊU THI ĐUA (M1/M2/M3):\n• ${ct.tuanDaNap.join('\n• ')}`;
          if (ct.tuanBoQua.length) msg += `\n\nBỏ qua:\n• ${ct.tuanBoQua.join('\n• ')}`;
          msg += '\n\nGửi file "Doanh Thu Theo Model" rồi @bot gõ "DT NGÀNH HÀNG" để xem báo cáo.';
          await guiTraLoi(event, targetId, { type: 'text', text: msg });
          continue;
        }

        if (kq.loai === 'dmsk_base') {
          await guiTraLoi(event, targetId, { type: 'text', text: `✅ Đã nạp BASE tháng 9: ${kq.soDong} siêu thị, tổng ${fmtTrieu(kq.tong / 1e6)} tr.\nMục tiêu T10 = ×${Math.round(HE_SO_DMSK * 100)}% = ${fmtTrieu((kq.tong * HE_SO_DMSK) / 1e6)} tr.\nGiờ anh gửi file tháng 10 rồi gõ "BC ĐMSK".` });
          continue;
        }
        if (kq.loai === 'dtlk') {
          try {
            const anh = [await generateDtlkReport('')];
            try { anh.unshift(await generateDtHomNayReport('')); } catch (_) { /* file không có ngày hôm nay */ }
            await guiTraLoi(event, targetId, anh.length > 1 ? anh : anh[0]);
          } catch (e) {
            await guiTraLoi(event, targetId, { type: 'text', text: `✅ Đã nạp DT theo siêu thị (${kq.soDong} dòng).\n⚠️ ${e.message}` });
          }
          continue;
        }
        if (kq.loai === 'thst_base') {
          await guiTraLoi(event, targetId, { type: 'text', text: `✅ Đã nạp BASE THS + Trứng tháng 9:\n${kq.tomTat.join('\n')}${kq.conThieu.length ? `\n\n⏳ Còn thiếu base: ${kq.conThieu.join(', ')} — anh gửi tiếp trong 3 phút.` : '\n\nGiờ anh gõ "NẠP NGÀY dd/mm" (ngày hôm qua) rồi gửi file luỹ kế T10, sau đó mỗi ngày gửi file 1 ngày và gõ "BC THS TRỨNG".'}` });
          continue;
        }
        if (kq.loai === 'thst') {
          try {
            await guiTraLoi(event, targetId, await generateThstReport());
          } catch (e) {
            await guiTraLoi(event, targetId, { type: 'text', text: `✅ Đã nạp DT THS + Trứng (${kq.soDong} nhóm hàng).\n⚠️ ${e.message}` });
          }
          continue;
        }
        if (kq.loai === 'bnnb_base') {
          await guiTraLoi(event, targetId, { type: 'text', text: `✅ Đã nạp BASE BNNB tháng 9:\n${kq.tomTat.join('\n')}${kq.conThieu.length ? `\n\n⏳ Còn thiếu base: ${kq.conThieu.join(', ')} — anh gửi tiếp trong 3 phút.` : '\n\nGiờ anh gõ "NẠP NGÀY dd/mm" (ngày hôm qua) rồi gửi file luỹ kế T10, sau đó mỗi ngày gửi file 1 ngày và gõ "Báo Cáo BNNB".'}` });
          continue;
        }
        if (kq.loai === 'bnnb') {
          try {
            await guiTraLoi(event, targetId, await generateBnnbReport());
          } catch (e) {
            await guiTraLoi(event, targetId, { type: 'text', text: `✅ Đã nạp DT BNNB (${kq.soDong} nhóm hàng).\n⚠️ ${e.message}` });
          }
          continue;
        }
        if (kq.loai === 'dmsk') {
          try {
            await guiTraLoi(event, targetId, await generateDmskReport());
          } catch (e) {
            await guiTraLoi(event, targetId, { type: 'text', text: `✅ Đã nạp DT tháng 10 (${kq.soDong} siêu thị).\n⚠️ ${e.message}` });
          }
          continue;
        }

        if (kq.loai === 'bo_qua') {
          await guiTraLoi(event, targetId, { type: 'text', text: kq.thongBao });
          continue;
        }

        if (kq.loai === 'sua') {
          try {
            await guiTraLoi(event, targetId, await generateSuaReport());
          } catch (e) {
            await guiTraLoi(event, targetId, { type: 'text', text: `✅ Đã nạp ${kq.soDong} dòng vào ${kq.tenTab}.\n⚠️ Chưa tạo được báo cáo: ${e.message}` });
          }
          continue;
        }

        if (kq.loai === 'huymmkk') {
          try {
            await guiTraLoi(event, targetId, await generateHuyMmkkReport('', kq.dsST));
          } catch (e) {
            await guiTraLoi(event, targetId, { type: 'text', text: `✅ Đã nạp huỷ MMKK (${kq.soDong} dòng).\n⚠️ Chưa tạo được báo cáo: ${e.message}` });
          }
          continue;
        }

        if (kq.loai === 'giavon' || kq.loai === 'fresh_ban') {
          try {
            await guiTraLoi(event, targetId, await generateGiaVonReport('', kq.dsST));
          } catch (e) {
            await guiTraLoi(event, targetId, { type: 'text', text: `✅ Đã nạp giá vốn (${kq.soDong} dòng).\n⚠️ Chưa tạo được báo cáo: ${e.message}` });
          }
          continue;
        }

        if (ngayChiDinh) {
          await client.pushMessage(targetId, { type: 'text', text: `✅ Đã ghi file vào NGÀY ${fmtNgay(kq.ngay)} (${kq.soDong} dòng). Báo cáo mới nhất:` }).catch(() => {});
        }
        try {
          const baoCao = await generateDtNganhHangReport('');
          await guiTraLoi(event, targetId, baoCao);
        } catch (loiBaoCao) {
          console.error('[webhook] nạp file OK nhưng chưa tạo được báo cáo:', loiBaoCao.message);
          await guiTraLoi(event, targetId, {
            type: 'text',
            text: `✅ Đã nạp DT ngày ${fmtNgay(kq.ngay)} (${kq.soDong} dòng ST × ngành hàng).\n⚠️ Chưa tạo được báo cáo: ${loiBaoCao.message}`,
          });
        }
      } catch (err) {
        console.error('[webhook] Lỗi nạp file:', err);
        await guiTraLoi(event, targetId, { type: 'text', text: `❌ Lỗi nạp file: ${err.message}` });
      }
      continue;
    }

    // ---------------- TEXT ----------------
    if (event.message.type !== 'text') continue;
    let question = event.message.text;

    if (isGroup) {
      question = extractMentionQuestion(event);
      if (question === null) continue; // group: không @tag bot -> bỏ qua
    }

    if (laLenhNapBaseThst(question)) {
      cho_NapBaseThst.set(targetId, Date.now() + THOI_GIAN_CHO_MS);
      cho_AI_PhanTich.delete(targetId);
      await guiTraLoi(event, targetId, { type: 'text', text: '📥 OK anh, các file Doanh Thu Theo Model tháng 9 (Thuỷ hải sản + Trứng) anh gửi trong 3 phút tới em sẽ lưu làm BASE Thuỷ hải sản + Trứng (× 110% = mục tiêu T10).' });
      continue;
    }

    if (laLenhNapBaseBnnb(question)) {
      cho_NapBaseBnnb.set(targetId, Date.now() + THOI_GIAN_CHO_MS);
      cho_AI_PhanTich.delete(targetId);
      await guiTraLoi(event, targetId, { type: 'text', text: '📥 OK anh, file Doanh Thu Theo Model anh gửi trong 3 phút tới em sẽ lưu làm BASE BNNB: Bia nước = tháng 7 (× 125%), Gia vị = tháng 9 (× 120%). Gửi file nào thì chỉ thay base ngành đó.' });
      continue;
    }

    if (laLenhNapBaseDmsk(question)) {
      cho_NapBaseDmsk.set(targetId, Date.now() + THOI_GIAN_CHO_MS);
      cho_AI_PhanTich.delete(targetId);
      await guiTraLoi(event, targetId, { type: 'text', text: '📥 OK anh, file Doanh Thu Theo Model Đông mát – Sữa – Kem anh gửi tiếp theo (trong 3 phút) em sẽ lưu làm BASE tháng 9 (× 120% = mục tiêu T10).' });
      continue;
    }

    const ngayNap = laLenhNapNgay(question);
    if (ngayNap) {
      cho_NapNgay.set(targetId, { ngay: ngayNap, expiresAt: Date.now() + THOI_GIAN_CHO_MS });
      cho_AI_PhanTich.delete(targetId);
      await guiTraLoi(event, targetId, {
        type: 'text',
        text: `📥 OK anh, file Excel anh gửi tiếp theo (trong 3 phút) em sẽ ghi vào NGÀY ${fmtNgay(ngayNap)}.\nNhớ xuất file luỹ kế từ đầu tuần tới hết ngày ${fmtNgay(ngayNap)}.`,
      });
      continue;
    }

    let ketQua = null;
    try {
      ketQua = await chayLenh(question);
    } catch (err) {
      console.error('[webhook] Lỗi tạo báo cáo:', err);
      await guiTraLoi(event, targetId, { type: 'text', text: `⚠️ Không tạo được báo cáo: ${err.message}` });
      continue;
    }
    if (ketQua) {
      await guiTraLoi(event, targetId, ketQua.ket);
      console.log(`[webhook] khớp lệnh "${ketQua.ten}", đã trả báo cáo`);
      continue;
    }

    // Không khớp lệnh -> câu hỏi tự do cho AI
    if (isGroup) datCoDangChoFile(targetId, question);
    await traLoiCauHoiAI(event, targetId, question);
  }
});

// Link ảnh báo cáo cho LINE tải về
app.get('/anh/:id.png', async (req, res) => {
  try {
    const nho = req.query.nho === '1';
    let anh = khoAnh.get(req.params.id);
    if (!anh) {
      // Render vừa khởi động lại -> vẽ lại theo tham số ghi trên link
      let ve;
      if (req.query.loai === 'gv') {
        ve = await veAnhGiaVon(String(req.query.st || ''));
      } else if (req.query.loai === 'hn') {
        ve = await veAnhDtHomNay(String(req.query.st || ''));
      } else if (req.query.loai === 'dtlk') {
        ve = await veAnhDtlk(String(req.query.st || ''));
      } else if (req.query.loai === 'thst') {
        ve = await veAnhThst();
      } else if (req.query.loai === 'bnnb') {
        ve = await veAnhBnnb();
      } else if (req.query.loai === 'dmsk') {
        ve = await veAnhDmsk();
      } else if (req.query.loai === 'sua') {
        ve = await veAnhSua();
      } else if (req.query.loai === 'mm') {
        ve = await veAnhHuyMmkk(String(req.query.st || ''));
      } else {
        const lenh = `dt ngành hàng ${req.query.t || ''} ${req.query.n ? fmtNgay(String(req.query.n)) : ''}`;
        ve = await veAnhDtNH(lenh);
      }
      luuAnh(req.params.id, ve.png, ve.preview);
      anh = khoAnh.get(req.params.id);
    }
    res.set('Content-Type', 'image/png');
    res.set('Cache-Control', 'public, max-age=86400');
    res.send(nho ? anh.preview : anh.png);
  } catch (err) {
    console.error('[anh] lỗi vẽ ảnh:', err);
    res.status(500).send('error');
  }
});

app.get('/health', (req, res) => res.send('ok'));

if (require.main === module) {
  (async () => {
    try {
      const info = await client.getBotInfo();
      botUserId = info.userId;
      console.log('[startup] Bot userId:', botUserId);
    } catch (err) {
      console.error('[startup] Không lấy được botUserId:', err.message);
    }
    app.listen(PORT, () => console.log(`LINE bot đang chạy ở port ${PORT}`));
  })();
}

module.exports = { docDuLieuThst, veSvgThst, laTriggerThst, laLenhNapBaseThst, veSvgDtHomNay, laTriggerDtHomNay, docDuLieuDtlk, veSvgDtlk, laTriggerDtlk, laLenhMtDtlk, datMtDtlk, docDuLieuBnnb, veSvgBnnb, laTriggerBnnb, laLenhNapBaseBnnb, docDuLieuDmsk, veSvgDmsk, laTriggerDmsk, laLenhNapBaseDmsk, docDuLieuSua, veSvgSua, laTriggerSua, laTriggerDtNganhHang, laTriggerGiaVon, laTriggerHuyMmkk, docDuLieuHuyMmkk, veSvgHuyMmkk, datMucThiDua, veSvgGiaVon, docDuLieuGiaVon, generateGiaVonReport, gioTuTenFile, napFileVaoSheet, generateDtNganhHangReport, tinhDuLieuDtNH, veSvgDtNH, docFileThiDua, gopDoanhThuModel };
