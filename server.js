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
async function napFileVaoSheet(fileName, buffer, ngayChiDinh) {
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

  if (laFileGiaVon(header)) return napFileGiaVon(sheets, header, dataRows, fileName);
  if (laFileFreshBan(header)) return napFileFreshBan(sheets, header, dataRows, fileName);
  if (laFileHuyMmkk(header)) return napFileHuyMmkk(sheets, header, dataRows);
  if (laFileTonModel(header)) return napFileTonSua(sheets, header, dataRows, fileName);
  if (laFileDoanhThuModel(header) && (await laFileBanSua(sheets, header, dataRows))) return napFileBanSua(sheets, header, dataRows, fileName);

  if (laFileDoanhThuModel(header)) {
    // File chỉ gồm ngành của tuần đã tắt (Đông mát – Sữa – Kem) -> không nạp
    const iNH = header.indexOf('Mã ngành hàng');
    const nhFile = new Set(dataRows.map((r) => String(r[iNH] ?? '').trim()).filter(Boolean));
    if (nhFile.size && [...nhFile].every((nh) => NH_TAT.has(nh))) {
      return { loai: 'bo_qua', soDong: 0, thongBao: 'ℹ️ Báo cáo Đông mát – Sữa – Kem đã tắt nên em không nạp file này.' };
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
const MA_SUA_MAC_DINH = ['2609001221', '2609001223', '2609001656', '2609001657'];

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
    .map((r) => [String(r[c.ma]).trim(), String(r[c.ten] || ''), chuanHoaMaST(r[c.st]), tenNganSieuThi(r[c.tenST]), soGV(r[c.ton]), c.di === -1 ? 0 : soGV(r[c.di]), capNhatTuTen(fileName)]);
  if (rows.length > 2000) throw new Error('File tồn quá nhiều mã. Anh lọc đúng các mã sữa cần theo dõi rồi gửi lại giúp em.');
  await ghiDeTab(sheets, TAB_SUA_TON, ['Mã Model', 'Model', 'Mã siêu thị', 'Tên siêu thị', 'Tồn', 'Đang đi đường', 'Cập nhật'], rows);
  return { loai: 'sua', tenTab: TAB_SUA_TON, soDong: rows.length };
}

async function napFileBanSua(sheets, header, dataRows, fileName) {
  const i = (c) => header.indexOf(c);
  const c = { ma: i('Mã Model'), ten: i('Model'), st: i('Mã siêu thị'), tenST: i('Tên siêu thị'), sl: i('Tổng số lượng'), dt: i('Tổng doanh thu') };
  const rows = dataRows.map((r) => [String(r[c.ma]).trim(), String(r[c.ten] || ''), chuanHoaMaST(r[c.st]), tenNganSieuThi(r[c.tenST]), soGV(r[c.sl]), soGV(r[c.dt]), capNhatTuTen(fileName)]);
  await ghiDeTab(sheets, TAB_SUA_BAN, ['Mã Model', 'Model', 'Mã siêu thị', 'Tên siêu thị', 'SL bán', 'Doanh thu', 'Cập nhật'], rows);
  return { loai: 'sua', tenTab: TAB_SUA_BAN, soDong: rows.length };
}

function tenNganSua(ten) {
  return String(ten)
    .replace(/^CELANO\s+/i, '')
    .replace(/SỮA YẾN MẠCH VỊ\s*/i, '')
    .replace(/T\.?\s*UỐNG\s*/i, '')
    .replace(/\s*CÓ THẠCH/i, '')
    .replace(/\s*\d+\s*ML.*$/i, '')
    .trim()
    .toLowerCase()
    .replace(/(^|\s)\S/g, (x) => x.toUpperCase());
}

async function docDuLieuSua() {
  const sheets = getSheetsClient();
  const ton = (await docTabAnToan(sheets, TAB_SUA_TON)).slice(1);
  const ban = (await docTabAnToan(sheets, TAB_SUA_BAN)).slice(1);
  if (!ton.length && !ban.length) throw new Error('Chưa có dữ liệu sữa. Anh gửi file "BC Tồn Theo Model" và file "Doanh Thu Theo Model" của các mã sữa vào group giúp em.');
  const sp = new Map(); // ma -> ten
  const st = new Map(); // maST -> { ten, ton:{}, ban:{} }
  const lay = (ma, ten) => { if (!st.has(ma)) st.set(ma, { ma, ten, ton: {}, di: {}, ban: {} }); return st.get(ma); };
  for (const r of ton) {
    const ma = String(r[0]); sp.set(ma, String(r[1]));
    const x = lay(String(r[2]), String(r[3]));
    x.ton[ma] = (x.ton[ma] || 0) + soGV(r[4]); x.di[ma] = (x.di[ma] || 0) + soGV(r[5]);
  }
  for (const r of ban) {
    const ma = String(r[0]); if (!sp.has(ma)) sp.set(ma, String(r[1]));
    const x = lay(String(r[2]), String(r[3]));
    x.ban[ma] = (x.ban[ma] || 0) + soGV(r[4]);
  }
  const dsSP = [...sp.entries()].sort((a, b) => a[1].localeCompare(b[1])).map(([ma, ten]) => ({ ma, ten, ngan: tenNganSua(ten) }));
  return { dsSP, dsST: [...st.values()], capTon: String((ton[0] || [])[6] || ''), capBan: String((ban[0] || [])[6] || '') };
}

function veSvgSua(d) {
  const { dsSP } = d;
  const nSP = dsSP.length;
  const PAD = 24, ROW = 30, xTen = PAD + 44, wTen = 300, wCot = 66;
  const W = Math.max(1000, xTen + wTen + (nSP + 1) * wCot * 2 + PAD + 10);
  const C = { xanh: '#0B8A3E', xanhDam: '#0B6E35', nen: '#E6F4EC', chan: '#F3F7F5', vien: '#D5E3DA', phu: '#888888', do: '#D0312D', cam: '#E08A00' };
  const t = (x, y, s, o = {}) =>
    `<text x="${x}" y="${y}" font-size="${o.size || 15}" font-weight="${o.bold ? 700 : 400}" fill="${o.fill || '#1a1a1a'}" text-anchor="${o.anchor || 'start'}">${escXml(s)}</text>`;
  const fs = (v) => fmtTrieu(v).replace(/,0$/, '');
  const tong = (x, k) => dsSP.reduce((s, p) => s + (x[k][p.ma] || 0), 0);
  const rows = d.dsST.map((x) => ({ ...x, tTon: tong(x, 'ton'), tBan: tong(x, 'ban') })).sort((a, b) => b.tBan - a.tBan || b.tTon - a.tTon);

  const p = [];
  p.push(`<rect x="0" y="0" width="${W}" height="92" fill="${C.xanh}"/>`);
  p.push(t(PAD, 38, 'BC SỮA · TỒN & BÁN THEO SIÊU THỊ', { size: 26, bold: true, fill: '#FFFFFF' }));
  p.push(t(PAD, 64, `${nSP} sản phẩm: ${dsSP.map((s) => s.ngan).join(' · ')}${/CELANO/i.test(dsSP.map((s) => s.ten).join('')) ? ' (Celano)' : ''}`, { size: 14, fill: '#E3F5EA' }));
  p.push(t(PAD, 84, `Tồn cập nhật ${d.capTon || '–'} · Bán theo file DT ${d.capBan || '–'} · ĐVT: hộp`, { size: 13, fill: '#E3F5EA' }));
  let y = 108;

  // Ô tổng mỗi sản phẩm
  const oW = (W - PAD * 2 - 12 * nSP) / (nSP + 1), oH = 92;
  [...dsSP, { ma: '*', ngan: 'TỔNG' }].forEach((s, k) => {
    const x = PAD + k * (oW + 12);
    const tTon = s.ma === '*' ? rows.reduce((a, r) => a + r.tTon, 0) : rows.reduce((a, r) => a + (r.ton[s.ma] || 0), 0);
    const tBan = s.ma === '*' ? rows.reduce((a, r) => a + r.tBan, 0) : rows.reduce((a, r) => a + (r.ban[s.ma] || 0), 0);
    const het = s.ma === '*' ? null : rows.filter((r) => (r.ton[s.ma] || 0) <= 0).length;
    p.push(`<rect x="${x}" y="${y}" width="${oW}" height="${oH}" rx="10" fill="${s.ma === '*' ? '#E6F4EC' : '#F2F8F4'}"/>`);
    // Tồn và Bán xuống 2 dòng riêng để không bị dính chữ
    p.push(t(x + 12, y + 20, rutGonTen(s.ngan, 22), { size: 13, bold: true, fill: C.xanhDam }));
    p.push(t(x + 12, y + 44, 'Tồn', { size: 13, fill: C.phu }));
    p.push(t(x + oW - 12, y + 44, fs(tTon), { size: 17, bold: true, anchor: 'end' }));
    p.push(t(x + 12, y + 66, 'Bán', { size: 13, fill: C.phu }));
    p.push(t(x + oW - 12, y + 66, fs(tBan), { size: 17, bold: true, fill: C.xanh, anchor: 'end' }));
    if (het != null) p.push(t(x + 12, y + 84, het ? `${het} ST hết hàng` : 'Không ST nào hết', { size: 11, fill: het ? C.do : C.phu }));
  });
  y += oH + 16;

  // Tiêu đề bảng 2 tầng
  const xCot = (k, j) => xTen + wTen + k * wCot * 2 + j * wCot + wCot - 8; // mép phải
  p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="54" rx="6" fill="${C.nen}"/>`);
  [...dsSP, { ngan: 'TỔNG' }].forEach((s, k) => {
    const x0 = xTen + wTen + k * wCot * 2;
    p.push(t(x0 + wCot, y + 21, rutGonTen(s.ngan, 16), { size: 13, bold: true, fill: C.xanhDam, anchor: 'middle' }));
    p.push(t(xCot(k, 0), y + 44, 'Tồn', { size: 13, bold: true, fill: C.xanhDam, anchor: 'end' }));
    p.push(t(xCot(k, 1), y + 44, 'Bán', { size: 13, bold: true, fill: C.xanhDam, anchor: 'end' }));
  });
  p.push(t(PAD + 30, y + 44, '#', { size: 13, bold: true, fill: C.xanhDam, anchor: 'end' }));
  p.push(t(xTen, y + 44, 'Siêu thị', { size: 13, bold: true, fill: C.xanhDam }));
  const yTop = y; y += 54;

  rows.forEach((r, i) => {
    if (i % 2 === 1) p.push(`<rect x="${PAD}" y="${y}" width="${W - PAD * 2}" height="${ROW}" fill="${C.chan}"/>`);
    const yt = y + 20;
    p.push(t(PAD + 30, yt, i + 1, { fill: C.phu, anchor: 'end' }));
    p.push(t(xTen, yt, `${r.ma} · ${rutGonTen(r.ten, 30)}`));
    dsSP.forEach((s, k) => {
      const ton = r.ton[s.ma] || 0, ban = r.ban[s.ma] || 0;
      p.push(t(xCot(k, 0), yt, fs(ton), { anchor: 'end', bold: ton <= 0, fill: ton <= 0 ? C.do : '#1a1a1a' }));
      p.push(t(xCot(k, 1), yt, ban ? fs(ban) : '–', { anchor: 'end', bold: ban > 0, fill: ban > 0 ? C.xanh : C.phu }));
    });
    p.push(t(xCot(nSP, 0), yt, fs(r.tTon), { anchor: 'end', bold: true }));
    p.push(t(xCot(nSP, 1), yt, r.tBan ? fs(r.tBan) : '–', { anchor: 'end', bold: true, fill: r.tBan > 0 ? C.xanh : C.phu }));
    y += ROW;
  });
  for (let k = 0; k <= nSP; k++) {
    const x = xTen + wTen + k * wCot * 2 - 4;
    p.push(`<line x1="${x}" y1="${yTop + 6}" x2="${x}" y2="${y}" stroke="${C.vien}" stroke-width="${k === nSP ? 2 : 1}"/>`);
  }
  p.push(`<rect x="${PAD}" y="${yTop}" width="${W - PAD * 2}" height="${y - yTop}" rx="6" fill="none" stroke="${C.vien}"/>`);
  y += 24;
  p.push(t(PAD, y, 'Xếp theo tổng bán giảm dần · Tồn đỏ = hết hàng (0) · Bán "–" = chưa bán.', { size: 12, fill: C.phu }));
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
  if (laTriggerDtNganhHang(text)) return { ten: 'DT Ngành Hàng', ket: await generateDtNganhHangReport(text) };
  if (laTriggerGiaVon(text)) return { ten: 'Giá Vốn', ket: await generateGiaVonReport(text) };
  if (laTriggerHuyMmkk(text)) return { ten: 'Huỷ MMKK', ket: await generateHuyMmkkReport(text) };
  if (laTriggerSua(text)) return { ten: 'BC Sữa', ket: await generateSuaReport() };
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

      const cauHoiCho = !ngayChiDinh && isGroup ? layVaXoaCoDangCho(targetId) : null;
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
        const kq = await napFileVaoSheet(fileName, buffer, ngayChiDinh);
        console.log(`[webhook] đã nạp ${kq.soDong} dòng vào tab "${kq.tenTab}" (loại: ${kq.loai})`);

        if (kq.loai === 'muctieu') {
          const ct = kq.chiTiet;
          let msg = `✅ Đã nạp MỤC TIÊU THI ĐUA (M1/M2/M3):\n• ${ct.tuanDaNap.join('\n• ')}`;
          if (ct.tuanBoQua.length) msg += `\n\nBỏ qua:\n• ${ct.tuanBoQua.join('\n• ')}`;
          msg += '\n\nGửi file "Doanh Thu Theo Model" rồi @bot gõ "DT NGÀNH HÀNG" để xem báo cáo.';
          await guiTraLoi(event, targetId, { type: 'text', text: msg });
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

module.exports = { docDuLieuSua, veSvgSua, laTriggerSua, laTriggerDtNganhHang, laTriggerGiaVon, laTriggerHuyMmkk, docDuLieuHuyMmkk, veSvgHuyMmkk, datMucThiDua, veSvgGiaVon, docDuLieuGiaVon, generateGiaVonReport, gioTuTenFile, napFileVaoSheet, generateDtNganhHangReport, tinhDuLieuDtNH, veSvgDtNH, docFileThiDua, gopDoanhThuModel };
