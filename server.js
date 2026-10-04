/**
 * server.js — LINE Bot linebot-baocao-bhx
 * ========================================
 * ĐÃ XOÁ TOÀN BỘ báo cáo cũ (Báo cáo ngày, Luỹ Kế DT/DT Dự Kiến, Giá Vốn, Huỷ MMKK,
 * SP 1 Đồng, Lượt Bill, Bánh Trung Thu). Bot hiện chỉ còn:
 *
 * 1) Lệnh "DT NGÀNH HÀNG" (trong group PHẢI @tag bot):
 *    - So DT thực tế luỹ kế trong tuần thi đua với MỤC TIÊU MỨC 2 (M2) của từng siêu thị.
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

const HEADER_MUCTIEU = ['Tuần', 'Từ ngày', 'Đến ngày', 'Tên thi đua', 'Mã ngành hàng', 'Mã siêu thị', 'Tên siêu thị', 'MT M2 (triệu)'];
const HEADER_DT_NGAY = ['Ngày', 'Mã siêu thị', 'Tên siêu thị', 'Mã ngành hàng', 'Ngành hàng', 'Doanh thu'];


const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
// Muốn tiết kiệm chi phí thì đổi thành: 'claude-haiku-4-5-20251001'
const AI_MODEL = 'claude-sonnet-5';
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
    if (!mNH) {
      tuanBoQua.push(`${tuan} (${tenSheet}) – không có mã ngành hàng trong tiêu đề`);
      continue;
    }
    const maNH = mNH[1].split(/[,\s]+/).filter(Boolean).join(',');

    // tìm dòng tiêu đề cột có "Siêu thị" và cột bắt đầu bằng "M2"
    let dongHeader = -1;
    let cotST = -1;
    let cotM2 = -1;
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
      const r = (rows[i] || []).map((c) => String(c || '').trim());
      const iST = r.findIndex((c) => c === 'Siêu thị');
      const iM2 = r.findIndex((c) => /^M2\b/.test(c));
      if (iST !== -1 && iM2 !== -1) { dongHeader = i; cotST = iST; cotM2 = iM2; break; }
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
      const m2 = Number(r[cotM2]);
      if (!m2 || isNaN(m2)) continue; // ST chưa có mục tiêu -> bỏ
      ketQua.push([tuan, tuNgay, denNgay, tenThiDua, maNH, chuanHoaMaST(st), tenNganSieuThi(st), Math.round(m2 * 10000) / 10000]);
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

function gopDoanhThuModel(header, dataRows, ngayKey) {
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
      gop.set(key, [ngayKey, maST, tenNganSieuThi(r[iTenST]), maNH, String(r[iTenNH] || '').trim(), 0]);
    }
    gop.get(key)[5] += Number(r[iDT]) || 0;
  }
  return [...gop.values()].map((d) => { d[5] = Math.round(d[5]); return d; });
}

// ---- Nạp file vào Sheet ----
async function napFileVaoSheet(fileName, buffer) {
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

  if (laFileDoanhThuModel(header)) {
    const ngayKey = ngayTuTenFile(fileName);
    const moi = gopDoanhThuModel(header, dataRows, ngayKey);
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
    `Không nhận diện được file "${fileName || ''}". Bot hiện chỉ nhận file THI ĐUA và file "Doanh Thu Theo Model".`
  );
}

// ---------------------------------------------------------------------------
// BÁO CÁO "DT NGÀNH HÀNG"
// ---------------------------------------------------------------------------
const TRIGGER_DT_NH = ['dt ngành hàng', 'dt nganh hang', 'doanh thu ngành hàng', 'doanh thu nganh hang'];
function laTriggerDtNganhHang(text) {
  const t = (text || '').normalize('NFC').toLowerCase().trim();
  return TRIGGER_DT_NH.some((k) => t.startsWith(k));
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

// 1 ô số trong bảng (viết gọn để 1 thẻ chứa đủ ~50 siêu thị — LINE giới hạn dung lượng thẻ)
// Dòng: Siêu thị | NGÀY: DT · MT · % | TUẦN: DT LK · MT · %
// Viết thật gọn (không flex/align cho ô số) vì LINE giới hạn 1 thẻ tối đa 30KB mà bảng ~50 siêu thị.
// Ô số mặc định flex 1 (cột đều nhau), tên siêu thị flex 3.
function dongBang2(ten, x, chuanTuan, chan) {
  const coMT = x.m2 > 0;
  const so = (text) => ({ type: 'text', text, size: 'xxs' });
  const pct = (p, chuan) => (coMT ? { type: 'text', text: fmtSoPct(p), size: 'xxs', color: mauPhanTram(p, chuan) } : so('–'));
  const o = {
    type: 'box', layout: 'horizontal',
    contents: [
      { type: 'text', text: rutGonTen(ten, 18), size: 'xxs', flex: 3 },
      so(fmtTrieu(x.dtNgay)),
      so(coMT ? fmtTrieu(x.mtNgay) : '–'),
      pct(x.pNgay, 100),
      { type: 'separator' },
      so(fmtTrieu(x.dtLk)),
      so(coMT ? fmtTrieu(x.m2) : '–'),
      pct(x.pTuan, chuanTuan),
    ],
  };
  if (chan) o.backgroundColor = '#F3F7F5';
  return o;
}

function fmtSoPct(p) {
  return (Math.round(p * 10) / 10).toFixed(1).replace('.', ',');
}

function tieuDeBang2(nhanNgay, nhanTuan) {
  const c = (text, flex) => ({ type: 'text', text, size: 'xxs', color: '#0B6E35', flex: flex || 1, weight: 'bold' });
  return {
    type: 'box', layout: 'vertical', margin: 'md', backgroundColor: '#E6F4EC', cornerRadius: 'sm', paddingAll: '4px',
    contents: [
      {
        type: 'box', layout: 'horizontal', contents: [
          { type: 'filler', flex: 3 },
          { type: 'text', text: nhanNgay, size: 'xxs', color: '#0B6E35', weight: 'bold', align: 'center', flex: 3 },
          { type: 'separator' },
          { type: 'text', text: nhanTuan, size: 'xxs', color: '#0B6E35', weight: 'bold', align: 'center', flex: 3 },
        ],
      },
      {
        type: 'box', layout: 'horizontal', margin: 'xs', contents: [
          c('Siêu thị', 3), c('DT'), c('MT'), c('%'),
          { type: 'separator' },
          c('DT LK'), c('MT'), c('%'),
        ],
      },
    ],
  };
}

async function generateDtNganhHangReport(text) {
  const sheets = getSheetsClient();
  const rowsMT = (await docTabAnToan(sheets, TAB_MUCTIEU)).slice(1);
  if (rowsMT.length === 0) {
    throw new Error('Chưa có mục tiêu. Anh gửi file THI ĐUA vào group trước giúp em.');
  }
  const rowsDT = (await docTabAnToan(sheets, TAB_DT_NGAY)).slice(1);
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
    dsTuan.get(tuan).mucTieu.push({ maST: chuanHoaMaST(r[5]), ten: String(r[6] || ''), m2: Number(r[7]) || 0 });
  }

  const dt = rowsDT
    .map((r) => ({ ngay: chuanHoaKeyNgay(r[0]), maST: chuanHoaMaST(r[1]), ten: String(r[2] || ''), maNH: String(r[3]).trim(), dt: Number(r[5]) || 0 }))
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
    tuanChon =
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

  // Cộng DT theo siêu thị: luỹ kế tuần + riêng ngày đang xem
  const stCoMucTieu = new Map(tuanChon.mucTieu.map((m) => [m.maST, m]));
  const dtLkST = new Map();
  const dtNgayST = new Map();
  const tenST = new Map();
  for (const d of dtTuan) {
    if (!ngayTinh(d)) continue;
    dtLkST.set(d.maST, (dtLkST.get(d.maST) || 0) + d.dt);
    if (d.ngay === denNgay) dtNgayST.set(d.maST, (dtNgayST.get(d.maST) || 0) + d.dt);
    if (!tenST.has(d.maST)) tenST.set(d.maST, d.ten);
  }

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
  const bang = tuanChon.mucTieu
    .map((mt) => lamDong(mt.maST, mt.ten, mt.m2))
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

  const oKV = (nhan, dt, mt, p, chuan) => ({
    type: 'box', layout: 'vertical', flex: 1, backgroundColor: '#F2F8F4', cornerRadius: 'md', paddingAll: '8px',
    contents: [
      { type: 'text', text: nhan, size: 'xxs', color: '#777777' },
      {
        type: 'box', layout: 'baseline', spacing: 'sm', contents: [
          { type: 'text', text: fmtPhanTram(p), size: 'lg', weight: 'bold', color: mauPhanTram(p, chuan), flex: 0 },
          { type: 'text', text: `${fmtTrieu(dt)} / ${fmtTrieu(mt)}`, size: 'xxs', color: '#555555' },
        ],
      },
    ],
  });

  const body = [
    {
      type: 'box', layout: 'horizontal', spacing: 'sm', contents: [
        oKV(`KV · ${nhanNgay}`, tongDtNgay, tongMtNgay, pNgayKV, 100),
        oKV(`KV · TUẦN (ngày ${soNgay}/7)`, tongDtLk, tongM2, pTuanKV, chuanTuan),
      ],
    },
    tieuDeBang2(nhanNgay, `TUẦN ${soNgay}/7`),
    // Shop mới chưa có mục tiêu gộp chung vào cuối bảng (chỉ có DT, cột MT/% để –)
    { type: 'box', layout: 'vertical', spacing: 'xs', contents: [...bang, ...dsKhongMT].map((x, k) => dongBang2(x.ten, x, chuanTuan, k % 2 === 1)) },
  ];

  const contents = {
    type: 'bubble', size: 'giga',
    header: {
      type: 'box', layout: 'vertical', backgroundColor: '#0B8A3E', paddingAll: '14px',
      contents: [
        { type: 'text', text: '📊 DT NGÀNH HÀNG · MỨC 2', color: '#FFFFFF', weight: 'bold', size: 'md' },
        { type: 'text', text: `Tuần ${tuanChon.tuan.slice(1)} · ${fmtNgay(tuanChon.tuNgay)}–${fmtNgay(tuanChon.denNgay)} · luỹ kế đến ${fmtNgay(denNgay)}`, color: '#E3F5EA', size: 'xs', margin: 'xs' },
        { type: 'text', text: tuanChon.ten || `NH ${[...tuanChon.maNH].join(', ')}`, color: '#E3F5EA', size: 'xxs', wrap: true, margin: 'xs' },
      ],
    },
    body: { type: 'box', layout: 'vertical', paddingAll: '10px', contents: body },
  };
  return {
    type: 'flex',
    altText: `DT Ngành Hàng ${fmtNgay(denNgay)}: ngày ${fmtPhanTram(pNgayKV)} · tuần ${fmtPhanTram(pTuanKV)} M2`,
    contents,
  };
}

async function chayLenh(text) {
  text = (text || '').normalize('NFC');
  if (laTriggerDtNganhHang(text)) return { ten: 'DT Ngành Hàng', ket: await generateDtNganhHangReport(text) };
  return null;
}

// ---------------------------------------------------------------------------
// AI (CLAUDE) — đoán tab liên quan + phân tích bảng dữ liệu / ảnh
// ---------------------------------------------------------------------------
const KNOWN_TABS = [
  { name: TAB_MUCTIEU, desc: 'Mục tiêu thi đua theo tuần (mức 2) của từng siêu thị theo nhóm ngành hàng' },
  { name: TAB_DT_NGAY, desc: 'Doanh thu theo ngày của từng siêu thị theo từng ngành hàng' },
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
    console.log('[webhook] câu hỏi tự do nhưng AI đang tắt -> bỏ qua');
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

      const cauHoiCho = isGroup ? layVaXoaCoDangCho(targetId) : null;
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
        const kq = await napFileVaoSheet(fileName, buffer);
        console.log(`[webhook] đã nạp ${kq.soDong} dòng vào tab "${kq.tenTab}" (loại: ${kq.loai})`);

        if (kq.loai === 'muctieu') {
          const ct = kq.chiTiet;
          let msg = `✅ Đã nạp MỤC TIÊU THI ĐUA (mức 2):\n• ${ct.tuanDaNap.join('\n• ')}`;
          if (ct.tuanBoQua.length) msg += `\n\nBỏ qua:\n• ${ct.tuanBoQua.join('\n• ')}`;
          msg += '\n\nGửi file "Doanh Thu Theo Model" rồi @bot gõ "DT NGÀNH HÀNG" để xem báo cáo.';
          await guiTraLoi(event, targetId, { type: 'text', text: msg });
          continue;
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

module.exports = { napFileVaoSheet, generateDtNganhHangReport, docFileThiDua, gopDoanhThuModel };
