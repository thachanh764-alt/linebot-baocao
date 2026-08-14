require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const XLSX = require('xlsx');
const { google } = require('googleapis');
const config = { channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN, channelSecret: process.env.LINE_CHANNEL_SECRET };
if (!config.channelAccessToken || !config.channelSecret) { console.error('Thiếu token trong .env'); process.exit(1); }
const client = new line.messagingApi.MessagingApiClient({ channelAccessToken: config.channelAccessToken });
const app = express();
const REPORT_KEYWORD = (process.env.REPORT_KEYWORD || 'báo cáo').toLowerCase();
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_TABS = { doanhThuT7: 'thang7_doanhthu.xlsx', doanhThuT8: 'thang8_doanhthu.xlsx', sanLuongT7: 'thang7_sanluong.xlsx', sanLuongT8: 'thang8_sanluong.xlsx' };
const DAYS_IN_AUGUST = 31;
const DAILY_TAB = 'baocao_ngay.xlsx';
const DAILY_KEYWORD = 'báo cáo dt';
const BANGIAOCA_TAB = 'bangiaoca.xlsx';
const BANGIAOCA_KEYWORD = 'bc bàn giao ca';
const WEEKDAY_VN = ['Chủ Nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];
function norm(s) { return (s == null ? '' : String(s)).normalize('NFC').replace(/\s+/g, ' ').trim().toLowerCase(); }
const FRESH_CATEGORIES = ['Rau Củ Quả CL', 'Thịt', 'Cá (Hải sản)', 'Trái cây', 'Trứng'].map(norm);
const isFresh = (cat) => FRESH_CATEGORIES.includes(norm(cat));

let sheetsApiClient = null;
function getSheetsApi() {
  if (sheetsApiClient) return sheetsApiClient;
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!raw) throw new Error('Thieu bien moi truong GOOGLE_SERVICE_ACCOUNT_JSON');
  const credentials = JSON.parse(raw);
  const auth = new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  sheetsApiClient = google.sheets({ version: 'v4', auth });
  return sheetsApiClient;
}
async function writeSheetTab(sheetName, rows2D) {
  const api = getSheetsApi();
  const quoted = `'${sheetName}'`;
  await api.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: quoted });
  await api.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${quoted}!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: rows2D },
  });
}

async function fetchSheetTable(sheetName) {
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Không tải được tab "${sheetName}" (HTTP ${res.status})`);
  const text = await res.text();
  const s = text.indexOf('{'), e = text.lastIndexOf('}');
  if (s === -1 || e === -1) throw new Error(`Không đọc được tab "${sheetName}"`);
  const json = JSON.parse(text.substring(s, e + 1));
  if (!json.table) throw new Error(`Tab "${sheetName}" không có dữ liệu`);
  const cols = (json.table.cols || []).map((c) => (c && c.label) ? c.label : '');
  const rows = (json.table.rows || []).map((r) => (r.c || []).map((c) => (c ? c.v : null)));
  return { cols, rows };
}

function parseGvizDate(value) {
  if (value == null) return null;
  if (typeof value === 'string') {
    const m1 = value.match(/^Date\((\d+),(\d+),(\d+)/);
    if (m1) return `${m1[1]}-${String(Number(m1[2]) + 1).padStart(2, '0')}-${String(m1[3]).padStart(2, '0')}`;
    const m2 = value.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m2) return `${m2[3]}-${m2[2].padStart(2, '0')}-${m2[1].padStart(2, '0')}`;
    return value.trim();
  }
  return String(value);
}
function findColIndex(cols, target) {
  const t = norm(target);
  return cols.findIndex((h) => norm(h) === t);
}
async function readAgg(sheetName, categoryCol, valueCol) {
  const { cols, rows } = await fetchSheetTable(sheetName);
  let idx = { date: findColIndex(cols, 'Ngày xuất'), category: findColIndex(cols, categoryCol), value: findColIndex(cols, valueCol) };
  if ((idx.date === -1 || idx.category === -1 || idx.value === -1) && rows.length) {
    for (let r = 0; r < Math.min(rows.length, 5); r++) {
      const h = rows[r];
      const test = { date: findColIndex(h, 'Ngày xuất'), category: findColIndex(h, categoryCol), value: findColIndex(h, valueCol) };
      if (test.date !== -1 && test.category !== -1 && test.value !== -1) {
        idx = test;
        let total = 0; const byCategory = {}; const daysSet = new Set();
        for (let rr = r + 1; rr < rows.length; rr++) {
          const row = rows[rr];
          if (!row || row[idx.category] == null) continue;
          const cat = String(row[idx.category]).normalize('NFC').trim();
          const val = Number(row[idx.value]) || 0;
          total += val; byCategory[cat] = (byCategory[cat] || 0) + val;
          const dayKey = parseGvizDate(row[idx.date]);
          if (dayKey) daysSet.add(dayKey);
        }
        return { total, byCategory, days: daysSet.size };
      }
    }
  }
  if (idx.date === -1 || idx.category === -1 || idx.value === -1) {
    throw new Error(`Thiếu cột trong tab "${sheetName}" (cần "Ngày xuất", "${categoryCol}", "${valueCol}"). Cột hiện có: ${cols.map((h) => `"${h}"`).join(', ')}`);
  }
  let total = 0; const byCategory = {}; const daysSet = new Set();
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row[idx.category] == null) continue;
    const cat = String(row[idx.category]).normalize('NFC').trim();
    const val = Number(row[idx.value]) || 0;
    total += val; byCategory[cat] = (byCategory[cat] || 0) + val;
    const dayKey = parseGvizDate(row[idx.date]);
    if (dayKey) daysSet.add(dayKey);
  }
  return { total, byCategory, days: daysSet.size };
}
const readDoanhThu = (s) => readAgg(s, 'Ngành hàng BHX', 'Doanh thu');
const readSanLuong = (s) => readAgg(s, 'Ngành hàng BHX', 'Sản lượng bán');
function projectFullMonth(v, d, total = DAYS_IN_AUGUST) { return d ? (v / d) * total : 0; }
function pctChange(c, p) { return p ? ((c - p) / p) * 100 : null; }
function sumByGroup(byCat, fn) { return Object.entries(byCat).filter(([c]) => fn(c)).reduce((a, [, v]) => a + v, 0); }
async function buildReport() {
  const [dtT7, dtT8, slT7, slT8] = await Promise.all([readDoanhThu(SHEET_TABS.doanhThuT7), readDoanhThu(SHEET_TABS.doanhThuT8), readSanLuong(SHEET_TABS.sanLuongT7), readSanLuong(SHEET_TABS.sanLuongT8)]);
  const dtT7Total = dtT7.total, dtT8Proj = projectFullMonth(dtT8.total, dtT8.days);
  const cats = new Set([...Object.keys(dtT7.byCategory), ...Object.keys(dtT8.byCategory)]);
  const rows = [...cats].map((cat) => { const t7 = dtT7.byCategory[cat] || 0, t8 = projectFullMonth(dtT8.byCategory[cat] || 0, dtT8.days); return { category: cat, group: isFresh(cat) ? 'FRESH' : 'FMCG', t7, t8Projected: t8, diff: t8 - t7, pct: pctChange(t8, t7) }; });
  const freshRows = rows.filter((r) => r.group === 'FRESH').sort((a, b) => b.t8Projected - a.t8Projected);
  const fmcgRows = rows.filter((r) => r.group === 'FMCG').sort((a, b) => b.t8Projected - a.t8Projected);
  const freshT7 = sumByGroup(dtT7.byCategory, isFresh), freshT8 = projectFullMonth(sumByGroup(dtT8.byCategory, isFresh), dtT8.days);
  const fmcgT7 = sumByGroup(dtT7.byCategory, (c) => !isFresh(c)), fmcgT8 = projectFullMonth(sumByGroup(dtT8.byCategory, (c) => !isFresh(c)), dtT8.days);
  const slT7Total = slT7.total, slT8Proj = projectFullMonth(slT8.total, slT8.days);
  return { meta: { d8: dtT8.days, s8: slT8.days, total: DAYS_IN_AUGUST },
    doanhThu: { t7Total: dtT7Total, t8Projected: dtT8Proj, diff: dtT8Proj - dtT7Total, pct: pctChange(dtT8Proj, dtT7Total),
      fresh: { t7Total: freshT7, t8Projected: freshT8, diff: freshT8 - freshT7, pct: pctChange(freshT8, freshT7), rows: freshRows },
      fmcg: { t7Total: fmcgT7, t8Projected: fmcgT8, diff: fmcgT8 - fmcgT7, pct: pctChange(fmcgT8, fmcgT7), rows: fmcgRows } },
    sanLuong: { t7Total: slT7Total, t8Projected: slT8Proj, diff: slT8Proj - slT7Total, pct: pctChange(slT8Proj, slT7Total) } };
}

function fmtVND(n) { return (n < 0 ? '-' : '') + Math.abs(Math.round(n)).toLocaleString('vi-VN') + ' đ'; }
function fmtShort(n) {
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 1e9) return sign + (abs / 1e9).toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' tỷ';
  if (abs >= 1e6) return sign + (abs / 1e6).toLocaleString('vi-VN', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' triệu';
  return sign + Math.round(abs).toLocaleString('vi-VN') + ' đ';
}
function fmtNum(n, d = 0) { return (n < 0 ? '-' : '') + Math.abs(n).toLocaleString('vi-VN', { minimumFractionDigits: d, maximumFractionDigits: d }); }
function fmtPct(p) { if (p == null || !isFinite(p)) return 'N/A'; return `${p > 0 ? '+' : ''}${p.toFixed(1)}%`; }
const cColor = (v) => (v > 0 ? '#06C755' : v < 0 ? '#FF3B30' : '#8C8C8C');
const cArrow = (v) => (v > 0 ? '▲' : v < 0 ? '▼' : '—');

function tableHeaderRow() {
  return { type: 'box', layout: 'horizontal', margin: 'md', contents: [
    { type: 'text', text: 'Ngành hàng', size: 'xxs', color: '#999999', flex: 5, weight: 'bold' },
    { type: 'text', text: 'DK T8', size: 'xxs', color: '#999999', flex: 3, align: 'end', weight: 'bold' },
    { type: 'text', text: 'MoM', size: 'xxs', color: '#999999', flex: 3, align: 'end', weight: 'bold' },
  ] };
}
function catRow(r) {
  return { type: 'box', layout: 'horizontal', margin: 'sm', contents: [
    { type: 'text', text: r.category, size: 'xs', color: '#333333', flex: 5, wrap: true },
    { type: 'text', text: fmtShort(r.t8Projected), size: 'xs', color: '#333333', flex: 3, align: 'end' },
    { type: 'text', text: `${cArrow(r.diff)} ${fmtPct(r.pct)}`, size: 'xs', color: cColor(r.diff), flex: 3, align: 'end', weight: 'bold' },
  ] };
}
function groupSec(title, g) {
  return { type: 'box', layout: 'vertical', margin: 'lg', contents: [
    { type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: title, weight: 'bold', size: 'md', color: '#111111', flex: 3 },
      { type: 'text', text: `${cArrow(g.diff)} ${fmtShort(g.diff)} (${fmtPct(g.pct)})`, size: 'xs', color: cColor(g.diff), align: 'end', flex: 5, weight: 'bold' },
    ] },
    { type: 'separator', margin: 'sm' },
    tableHeaderRow(),
    { type: 'separator', margin: 'xs' },
    ...g.rows.map(catRow),
  ] };
}

function bubbleTongQuan(report) {
  const { doanhThu: d, meta: m } = report;
  return { type: 'bubble', size: 'giga',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#06C755', paddingAll: '16px', contents: [
      { type: 'text', text: 'BÁO CÁO DOANH THU', color: '#FFFFFF', weight: 'bold', size: 'lg' },
      { type: 'text', text: `Dự kiến T8 (${m.d8}/${m.total} ngày) so với T7`, color: '#E8FFF0', size: 'xs', margin: 'sm' },
    ] },
    body: { type: 'box', layout: 'vertical', contents: [
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'box', layout: 'vertical', flex: 1, contents: [
          { type: 'text', text: 'DỰ KIẾN T8', size: 'xxs', color: '#999999' },
          { type: 'text', text: fmtShort(d.t8Projected), size: 'lg', weight: 'bold', color: '#06C755', wrap: true },
        ] },
        { type: 'box', layout: 'vertical', flex: 1, contents: [
          { type: 'text', text: 'THỰC TẾ T7', size: 'xxs', color: '#999999' },
          { type: 'text', text: fmtShort(d.t7Total), size: 'lg', weight: 'bold', color: '#111111', wrap: true },
        ] },
        { type: 'box', layout: 'vertical', flex: 1, contents: [
          { type: 'text', text: 'MoM T8 - T7', size: 'xxs', color: '#999999' },
          { type: 'text', text: fmtPct(d.pct), size: 'lg', weight: 'bold', color: cColor(d.diff) },
          { type: 'text', text: `${cArrow(d.diff)} ${fmtShort(d.diff)}`, size: 'xs', color: cColor(d.diff) },
        ] },
      ] },
      { type: 'separator', margin: 'lg' },
      { type: 'text', text: 'DOANH THU THEO NGÀNH HÀNG (DỰ KIẾN T8)', size: 'xxs', color: '#999999', weight: 'bold', margin: 'lg' },
      { type: 'box', layout: 'horizontal', margin: 'md', contents: [
        { type: 'text', text: '🥬 FRESH', size: 'sm', color: '#333333', flex: 3 },
        { type: 'text', text: fmtShort(d.fresh.t8Projected), size: 'sm', color: '#111111', weight: 'bold', flex: 3, align: 'end' },
        { type: 'text', text: `${cArrow(d.fresh.diff)} ${fmtPct(d.fresh.pct)}`, size: 'xs', color: cColor(d.fresh.diff), flex: 2, align: 'end', weight: 'bold' },
      ] },
      { type: 'box', layout: 'horizontal', margin: 'sm', contents: [
        { type: 'text', text: '🛒 FMCG', size: 'sm', color: '#333333', flex: 3 },
        { type: 'text', text: fmtShort(d.fmcg.t8Projected), size: 'sm', color: '#111111', weight: 'bold', flex: 3, align: 'end' },
        { type: 'text', text: `${cArrow(d.fmcg.diff)} ${fmtPct(d.fmcg.pct)}`, size: 'xs', color: cColor(d.fmcg.diff), flex: 2, align: 'end', weight: 'bold' },
      ] },
    ] },
    footer: { type: 'box', layout: 'vertical', contents: [
      { type: 'text', text: '* Số liệu T8 là dự kiến, chiếu từ dữ liệu thực tế các ngày đã phát sinh', size: 'xxs', color: '#AAAAAA', wrap: true },
    ] },
  };
}

function bubbleChiTiet(report) {
  const { doanhThu: d } = report;
  return { type: 'bubble', size: 'giga',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#1E90FF', paddingAll: '16px', contents: [
      { type: 'text', text: 'CHI TIẾT NGÀNH HÀNG', color: '#FFFFFF', weight: 'bold', size: 'lg' },
      { type: 'text', text: 'Tăng/giảm doanh thu dự kiến T8 so với T7', color: '#E6F2FF', size: 'xs', margin: 'sm' },
    ] },
    body: { type: 'box', layout: 'vertical', contents: [
      groupSec('🥬 Ngành FRESH', d.fresh),
      groupSec('🛒 Ngành FMCG', d.fmcg),
    ] },
  };
}

function bubbleSanLuong(report) {
  const { sanLuong: s, meta: m } = report;
  return { type: 'bubble', size: 'giga',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#FF9500', paddingAll: '16px', contents: [
      { type: 'text', text: 'SẢN LƯỢNG BÁN', color: '#FFFFFF', weight: 'bold', size: 'lg' },
      { type: 'text', text: `Dự kiến T8 (${m.s8}/${m.total} ngày) so với T7`, color: '#FFF3E0', size: 'xs', margin: 'sm' },
    ] },
    body: { type: 'box', layout: 'vertical', contents: [
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'box', layout: 'vertical', flex: 1, contents: [
          { type: 'text', text: 'DỰ KIẾN T8', size: 'xxs', color: '#999999' },
          { type: 'text', text: fmtNum(s.t8Projected, 1), size: 'lg', weight: 'bold', color: '#FF9500', wrap: true },
        ] },
        { type: 'box', layout: 'vertical', flex: 1, contents: [
          { type: 'text', text: 'THỰC TẾ T7', size: 'xxs', color: '#999999' },
          { type: 'text', text: fmtNum(s.t7Total, 1), size: 'lg', weight: 'bold', color: '#111111', wrap: true },
        ] },
        { type: 'box', layout: 'vertical', flex: 1, contents: [
          { type: 'text', text: 'MoM T8 - T7', size: 'xxs', color: '#999999' },
          { type: 'text', text: fmtPct(s.pct), size: 'lg', weight: 'bold', color: cColor(s.diff) },
          { type: 'text', text: `${cArrow(s.diff)} ${fmtNum(s.diff, 1)}`, size: 'xs', color: cColor(s.diff) },
        ] },
      ] },
    ] },
  };
}

function buildFlexMessages(report) {
  return [
    { type: 'flex', altText: 'Tong quan doanh thu T8 vs T7', contents: bubbleTongQuan(report) },
    { type: 'flex', altText: 'Chi tiet nganh hang FRESH/FMCG', contents: bubbleChiTiet(report) },
    { type: 'flex', altText: 'San luong ban T8 vs T7', contents: bubbleSanLuong(report) },
  ];
}

function keyToDateInfo(key) {
  const [y, m, d] = key.split('-').map(Number);
  const dateObj = new Date(Date.UTC(y, m - 1, d));
  return { weekday: WEEKDAY_VN[dateObj.getUTCDay()], display: `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}` };
}

async function readDailyRevenue() {
  const { cols, rows } = await fetchSheetTable(DAILY_TAB);
  const idx = {
    date: findColIndex(cols, 'Ngày'),
    store: findColIndex(cols, 'Tên siêu thị'),
    offline: findColIndex(cols, 'Doanh thu offline'),
    online: findColIndex(cols, 'Doanh thu Online'),
    bill: findColIndex(cols, 'Tổng số bill'),
  };
  const billOnlineIdx = findColIndex(cols, 'Tổng số bill online');
  if (Object.values(idx).some((v) => v === -1)) {
    throw new Error(`Thiếu cột trong tab "${DAILY_TAB}" (cần "Ngày", "Tên siêu thị", "Doanh thu offline", "Doanh thu Online", "Tổng số bill"). Cột hiện có: ${cols.map((h) => `"${h}"`).join(', ')}`);
  }
  const parsed = [];
  const daySet = new Set();
  for (const row of rows) {
    if (!row || row[idx.store] == null) continue;
    const dayKey = parseGvizDate(row[idx.date]);
    if (!dayKey) continue;
    daySet.add(dayKey);
    parsed.push({ row, dayKey });
  }
  if (!daySet.size) return { items: [], dateInfo: null };
  const latestKey = [...daySet].sort().pop();
  const items = [];
  for (const { row, dayKey } of parsed) {
    if (dayKey !== latestKey) continue;
    const offline = Number(row[idx.offline]) || 0;
    const online = Number(row[idx.online]) || 0;
    const bill = Number(row[idx.bill]) || 0;
    const billOnline = billOnlineIdx !== -1 ? (Number(row[billOnlineIdx]) || 0) : null;
    items.push({ store: String(row[idx.store]).normalize('NFC').trim(), offline, online, bill, billOnline, total: offline + online });
  }
  return { items, dateInfo: keyToDateInfo(latestKey) };
}

function rowLine(icon, label, value) {
  return { type: 'box', layout: 'horizontal', margin: 'md', contents: [
    { type: 'text', text: `${icon} ${label}`, size: 'sm', color: '#666666', flex: 3 },
    { type: 'text', text: value, size: 'sm', color: '#111111', flex: 3, align: 'end', weight: 'bold' },
  ] };
}

function bubbleDailyStore(item, dateInfo) {
  const { weekday, display } = dateInfo;
  const avgBill = item.bill ? item.total / item.bill : 0;
  return { type: 'bubble', size: 'giga',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#06C755', paddingAll: '16px', contents: [
      { type: 'text', text: '📊 BÁO CÁO DOANH THU', color: '#FFFFFF', weight: 'bold', size: 'lg' },
      { type: 'text', text: `📅 ${weekday}, ${display}`, color: '#E8FFF0', size: 'sm', margin: 'sm' },
    ] },
    body: { type: 'box', layout: 'vertical', contents: [
      { type: 'text', text: `🏬 ${item.store}`, weight: 'bold', size: 'md', color: '#111111', wrap: true },
      { type: 'separator', margin: 'md' },
      rowLine('🏢', 'Doanh thu offline', fmtVND(item.offline)),
      rowLine('🛍️', 'Doanh thu online', fmtVND(item.online)),
      rowLine('🧾', 'Số lượng bill', fmtNum(item.bill)),
      ...(item.billOnline != null ? [rowLine('💻', 'Bill online', fmtNum(item.billOnline))] : []),
      rowLine('📈', 'Giá trị bill', fmtVND(avgBill)),
      { type: 'separator', margin: 'lg' },
      { type: 'box', layout: 'horizontal', margin: 'lg', contents: [
        { type: 'text', text: '💰 Tổng doanh thu', weight: 'bold', size: 'md', color: '#111111', flex: 3 },
        { type: 'text', text: fmtVND(item.total), weight: 'bold', size: 'md', color: '#06C755', flex: 3, align: 'end' },
      ] },
    ] },
  };
}

async function buildDailyMessages() {
  const { items, dateInfo } = await readDailyRevenue();
  if (!items.length || !dateInfo) {
    return [{ type: 'text', text: `Không tìm thấy dữ liệu hợp lệ trong tab "${DAILY_TAB}". Kiểm tra lại cột "Ngày" đã đúng định dạng chưa.` }];
  }
  return items.map((item) => ({ type: 'flex', altText: `Báo cáo doanh thu ${item.store}`, contents: bubbleDailyStore(item, dateInfo) }));
}

async function readBanGiaoCa() {
  const { cols, rows } = await fetchSheetTable(BANGIAOCA_TAB);
  const idx = {
    store: findColIndex(cols, 'Tên siêu thị'),
    batch: findColIndex(cols, 'Mã đợt châm hàng'),
    slTonTrenKe: findColIndex(cols, 'SL tồn trên kệ'),
    trangThaiCham: findColIndex(cols, 'Trạng thái đợt châm hàng'),
  };
  if (Object.values(idx).some((v) => v === -1)) {
    throw new Error(`Thiếu cột trong tab "${BANGIAOCA_TAB}" (cần "Tên siêu thị", "Mã đợt châm hàng", "SL tồn trên kệ", "Trạng thái đợt châm hàng"). Cột hiện có: ${cols.map((h) => `"${h}"`).join(', ')}`);
  }
  const byStore = new Map();
  for (const row of rows) {
    if (!row || row[idx.store] == null) continue;
    const store = String(row[idx.store]).normalize('NFC').trim();
    if (!store) continue;
    if (!byStore.has(store)) byStore.set(store, { sku: 0, trongKe: 0, batches: new Map() });
    const s = byStore.get(store);
    s.sku += 1;
    const slTonTrenKe = Number(row[idx.slTonTrenKe]) || 0;
    if (slTonTrenKe === 0) s.trongKe += 1;
    const batchId = row[idx.batch];
    const isDone = norm(row[idx.trangThaiCham]) === norm('Đã châm hàng');
    if (batchId != null) {
      const cur = s.batches.get(batchId);
      s.batches.set(batchId, cur === undefined ? isDone : (cur && isDone));
    }
  }
  const items = [];
  let sumQuetKe = 0, sumSku = 0, sumTrongKe = 0, sumChuaHT = 0;
  for (const [name, s] of byStore) {
    const quetKe = s.batches.size;
    const chuaHT = [...s.batches.values()].filter((done) => !done).length;
    const tyLeTrongKe = s.sku ? s.trongKe / s.sku : 0;
    items.push({ name, quetKe, sku: s.sku, chuaHT, trongKe: s.trongKe, tyLeTrongKe });
    sumQuetKe += quetKe; sumSku += s.sku; sumTrongKe += s.trongKe; sumChuaHT += chuaHT;
  }
  items.sort((a, b) => a.name.localeCompare(b.name, 'vi'));
  const totalRow = items.length ? { name: 'TOTAL', quetKe: sumQuetKe, sku: sumSku, chuaHT: sumChuaHT, trongKe: sumTrongKe, tyLeTrongKe: sumSku ? sumTrongKe / sumSku : 0 } : null;
  return { items, totalRow };
}

function bcHeaderRow() {
  return { type: 'box', layout: 'horizontal', paddingAll: '6px', backgroundColor: '#F0F1F3', cornerRadius: '6px', contents: [
    { type: 'text', text: 'Siêu thị', size: 'xxs', color: '#666666', weight: 'bold', flex: 5 },
    { type: 'text', text: 'Quét kệ', size: 'xxs', color: '#666666', weight: 'bold', flex: 2, align: 'end' },
    { type: 'text', text: 'SKU', size: 'xxs', color: '#666666', weight: 'bold', flex: 2, align: 'end' },
    { type: 'text', text: 'Trống kệ', size: 'xxs', color: '#666666', weight: 'bold', flex: 2, align: 'end' },
    { type: 'text', text: 'Tỷ lệ', size: 'xxs', color: '#666666', weight: 'bold', flex: 2, align: 'end' },
    { type: 'text', text: 'Chưa HT', size: 'xxs', color: '#666666', weight: 'bold', flex: 2, align: 'end' },
  ] };
}
function bcRow(item, index) {
  const badQuetKe = item.quetKe < 2;
  const badSku = item.sku < 300;
  const badChuaHT = item.chuaHT > 0;
  const c = (bad) => (bad ? '#FF3B30' : '#333333');
  const zebra = index % 2 === 0 ? '#FFFFFF' : '#FAFAFA';
  return [{ type: 'box', layout: 'horizontal', paddingAll: '6px', backgroundColor: zebra, cornerRadius: '4px', contents: [
    { type: 'text', text: item.name, size: 'xxs', color: '#333333', flex: 5, wrap: true, gravity: 'center' },
    { type: 'text', text: fmtNum(item.quetKe), size: 'xxs', color: c(badQuetKe), weight: badQuetKe ? 'bold' : 'regular', flex: 2, align: 'end', gravity: 'center' },
    { type: 'text', text: fmtNum(item.sku), size: 'xxs', color: c(badSku), weight: badSku ? 'bold' : 'regular', flex: 2, align: 'end', gravity: 'center' },
    { type: 'text', text: item.trongKe != null ? fmtNum(item.trongKe) : '—', size: 'xxs', color: '#666666', flex: 2, align: 'end', gravity: 'center' },
    { type: 'text', text: item.tyLeTrongKe != null ? (item.tyLeTrongKe * 100).toFixed(1) + '%' : '—', size: 'xxs', color: '#666666', flex: 2, align: 'end', gravity: 'center' },
    { type: 'text', text: fmtNum(item.chuaHT), size: 'xxs', color: c(badChuaHT), weight: badChuaHT ? 'bold' : 'regular', flex: 2, align: 'end', gravity: 'center' },
  ] }];
}

function buildBanGiaoCaBubble(items, totalRow) {
  const soanBad = items.filter((i) => i.quetKe < 2 || i.sku < 300 || i.chuaHT > 0).length;
  const body = [bcHeaderRow(), ...items.flatMap((item, i) => bcRow(item, i))];
  if (totalRow) {
    body.push({ type: 'separator', margin: 'md' });
    body.push({ type: 'box', layout: 'horizontal', margin: 'md', paddingAll: '6px', backgroundColor: '#F0F1F3', cornerRadius: '6px', contents: [
      { type: 'text', text: 'TOTAL', size: 'xs', color: '#111111', weight: 'bold', flex: 5 },
      { type: 'text', text: fmtNum(totalRow.quetKe), size: 'xs', color: '#111111', weight: 'bold', flex: 2, align: 'end' },
      { type: 'text', text: fmtNum(totalRow.sku), size: 'xs', color: '#111111', weight: 'bold', flex: 2, align: 'end' },
      { type: 'text', text: fmtNum(totalRow.trongKe), size: 'xs', color: '#111111', weight: 'bold', flex: 2, align: 'end' },
      { type: 'text', text: (totalRow.tyLeTrongKe * 100).toFixed(1) + '%', size: 'xs', color: '#111111', weight: 'bold', flex: 2, align: 'end' },
      { type: 'text', text: fmtNum(totalRow.chuaHT), size: 'xs', color: '#111111', weight: 'bold', flex: 2, align: 'end' },
    ] });
  }
  return { type: 'bubble', size: 'giga',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#E63946', paddingAll: '18px', contents: [
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: '📋', size: 'xl', flex: 0 },
        { type: 'text', text: 'BÁO CÁO BÀN GIAO CA', color: '#FFFFFF', weight: 'bold', size: 'lg', margin: 'sm' },
      ] },
      { type: 'text', text: `${items.length} siêu thị  ·  ${soanBad} cần chú ý`, color: '#FFE0E0', size: 'xs', margin: 'sm' },
    ] },
    body: { type: 'box', layout: 'vertical', spacing: 'xs', paddingAll: '12px', contents: body },
    footer: { type: 'box', layout: 'vertical', paddingAll: '12px', contents: [
      { type: 'separator', margin: 'none' },
      { type: 'text', text: '● Đỏ: Quét kệ < 2, SKU < 300, hoặc Chưa HT > 0', size: 'xxs', color: '#FF3B30', wrap: true, margin: 'md' },
    ] },
  };
}

async function buildBanGiaoCaMessages() {
  const { items, totalRow } = await readBanGiaoCa();
  if (!items.length) {
    return [{ type: 'text', text: `Không có dữ liệu trong tab "${BANGIAOCA_TAB}".` }];
  }
  return [{ type: 'flex', altText: 'Báo cáo Bàn Giao Ca', contents: buildBanGiaoCaBubble(items, totalRow) }];
}

function hasHeader(headers, target) {
  const t = norm(target);
  return headers.some((h) => norm(h) === t);
}
function detectMonthFromRows(headers, rows) {
  const dateIdx = headers.findIndex((h) => norm(h) === norm('Ngày xuất'));
  if (dateIdx === -1) return null;
  for (const row of rows) {
    const v = row[dateIdx];
    let d = null;
    if (v instanceof Date) d = v;
    else if (typeof v === 'string') {
      const m = v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    }
    if (d && !isNaN(d.getTime())) return d.getMonth() + 1;
  }
  return null;
}
function detectTargetTab(headers, rows) {
  if (hasHeader(headers, 'Mã đợt châm hàng') && hasHeader(headers, 'SL tồn trên kệ')) {
    return { tab: BANGIAOCA_TAB, label: 'BC Bàn Giao Ca' };
  }
  if (hasHeader(headers, 'Doanh thu offline') && hasHeader(headers, 'Doanh thu Online')) {
    return { tab: DAILY_TAB, label: 'Báo Cáo DT (theo ngày)' };
  }
  if (hasHeader(headers, 'Sản lượng bán') && hasHeader(headers, 'Ngành hàng BHX') && !hasHeader(headers, 'Doanh thu offline')) {
    const month = detectMonthFromRows(headers, rows);
    if (!month) return null;
    return { tab: month === 7 ? SHEET_TABS.sanLuongT7 : SHEET_TABS.sanLuongT8, label: `Sản lượng tháng ${month}` };
  }
  if (hasHeader(headers, 'Doanh thu') && hasHeader(headers, 'Ngành hàng BHX')) {
    const month = detectMonthFromRows(headers, rows);
    if (!month) return null;
    return { tab: month === 7 ? SHEET_TABS.doanhThuT7 : SHEET_TABS.doanhThuT8, label: `Doanh thu tháng ${month}` };
  }
  return null;
}
function cellToWritable(cell) {
  if (cell instanceof Date && !isNaN(cell.getTime())) {
    const d = String(cell.getDate()).padStart(2, '0');
    const m = String(cell.getMonth() + 1).padStart(2, '0');
    const y = cell.getFullYear();
    return `${d}/${m}/${y}`;
  }
  return cell == null ? '' : cell;
}
async function handleFileMessage(event) {
  const messageId = event.message.id;
  const fileName = event.message.fileName || 'file';
  const res = await fetch(`https://api-data.line.me/v2/bot/message/${messageId}/content`, {
    headers: { Authorization: `Bearer ${config.channelAccessToken}` },
  });
  if (!res.ok) throw new Error(`Không tải được file từ LINE (HTTP ${res.status})`);
  const buf = Buffer.from(await res.arrayBuffer());
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!rows.length) throw new Error(`File "${fileName}" trống, không có dữ liệu.`);
  const headers = rows[0];
  const dataRows = rows.slice(1).filter((r) => r && r.some((c) => c != null && c !== ''));
  const target = detectTargetTab(headers, dataRows);
  if (!target) {
    throw new Error(`Không nhận diện được loại báo cáo từ file "${fileName}". Kiểm tra lại tiêu đề cột trong file có đúng mẫu không.`);
  }
  const writable = [headers, ...dataRows].map((row) => row.map(cellToWritable));
  await writeSheetTab(target.tab, writable);
  return { label: target.label, tab: target.tab, rowCount: dataRows.length };
}

app.get('/', (req, res) => res.send('Bot đang chạy'));
app.post('/webhook', line.middleware(config), (req, res) => {
  res.sendStatus(200);
  const events = req.body.events || [];
  events.forEach((event) => { handleEvent(event).catch((e) => console.error('Loi xu ly event:', e)); });
});
async function handleEvent(event) {
  if (event.type !== 'message') return;

  if (event.message.type === 'file') {
    try {
      const result = await handleFileMessage(event);
      await client.replyMessage({ replyToken: event.replyToken, messages: [{ type: 'text', text: `✅ Đã nạp ${result.rowCount} dòng vào tab "${result.tab}" (${result.label}).\nGõ lệnh báo cáo tương ứng để xem kết quả mới.` }] });
    } catch (e) {
      await client.replyMessage({ replyToken: event.replyToken, messages: [{ type: 'text', text: `❌ Lỗi nạp file: ${e.message}`.slice(0, 800) }] });
    }
    return;
  }

  if (event.message.type !== 'text') return;
  const text = event.message.text.trim().toLowerCase();

  if (text === BANGIAOCA_KEYWORD || text === 'bc ban giao ca' || text === '/bcbangiaoca') {
    try { const messages = await buildBanGiaoCaMessages(); await client.replyMessage({ replyToken: event.replyToken, messages }); }
    catch (e) { await client.replyMessage({ replyToken: event.replyToken, messages: [{ type: 'text', text: `Lỗi: ${e.message}`.slice(0, 800) }] }); }
    return;
  }

  if (text === DAILY_KEYWORD || text === 'bao cao dt' || text === '/baocaodt') {
    try { const messages = await buildDailyMessages(); await client.replyMessage({ replyToken: event.replyToken, messages }); }
    catch (e) { await client.replyMessage({ replyToken: event.replyToken, messages: [{ type: 'text', text: `Lỗi: ${e.message}`.slice(0, 800) }] }); }
    return;
  }

  if (text === REPORT_KEYWORD || text === '/baocao' || text === 'baocao') {
    try { const report = await buildReport(); await client.replyMessage({ replyToken: event.replyToken, messages: buildFlexMessages(report) }); }
    catch (e) { await client.replyMessage({ replyToken: event.replyToken, messages: [{ type: 'text', text: `Lỗi: ${e.message}`.slice(0, 800) }] }); }
  }
}
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Chạy tại cổng ${PORT}`));
