/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  GUDANG ELEKTRIK — INVENTORY MANAGEMENT SYSTEM                  ║
 * ║  Backend Server (Google Apps Script)                             ║
 * ║  Database Structure (10 Kolom):                                 ║
 * ║  [Rak, Item Code, Spesifikasi, Stok Awal, Masuk (In),           ║
 * ║   Keluar (Out), Stok Akhir, Satuan, Minimal Stok, Remarks]       ║
 * ╚═══════════════════════════════════════════════════════════════════╝
 */

// ===================================================================
// KONFIGURASI SISTEM
// ===================================================================

const SS_ID       = ''; // Kosongkan jika script terpasang langsung di spreadsheet (Extensions > Apps Script)
const DB_SHEET    = 'database';
const DEFAULT_PIN = '100700';
const TIMEZONE    = 'Asia/Jakarta';

const DB_HEADERS = [
  'Rak', 'Item Code', 'Spesifikasi', 'Stok Awal', 'Masuk (In)',
  'Keluar (Out)', 'Stok Akhir', 'Satuan', 'Minimal Stok', 'Remarks'
];

const ELECTRICAL_SECTIONS = [
  'Finish Mill', 'Kiln', 'Coal Mill', 'Raw Mill', 'CCP', 'Substation', 'MCC Room', 'Utility', 'Workshop Elektrik'
];

const STANDARD_UNITS = ['PCS', 'SET', 'UNIT', 'MTR', 'ROLL', 'BOX', 'CAN', 'PAIR', 'LTR', 'KG'];

// ===================================================================
// HELPER UTILITY
// ===================================================================

function getSpreadsheet() {
  if (SS_ID && SS_ID !== 'PASTE_SPREADSHEET_ID_HERE') {
    try {
      return SpreadsheetApp.openById(SS_ID);
    } catch(e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((acc, h, i) => {
    const key = h.toString().toLowerCase().trim().replace(/ /g, '_');
    acc[key] = i;
    return acc;
  }, {});
}

function findColIdx_(keys, candidates) {
  if (!keys || !keys.length) return -1;
  const normKeys = keys.map(k => String(k || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
  // 1. Exact match
  for (let c of candidates) {
    const target = c.toLowerCase().replace(/[^a-z0-9]/g, '');
    const idx = normKeys.indexOf(target);
    if (idx !== -1) return idx;
  }
  // 2. Contains match (target inside normKey, only if target >= 3 characters)
  for (let c of candidates) {
    const target = c.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (target.length < 3) continue;
    for (let i = 0; i < normKeys.length; i++) {
      if (normKeys[i] && normKeys[i].includes(target)) {
        return i;
      }
    }
  }
  return -1;
}

function parseHtmlDate_(s) {
  if (!s || s === '') return '';
  if (Object.prototype.toString.call(s) === '[object Date]') return s;
  s = String(s).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function formatDate_(val, includeTime = false) {
  if (!val) return '';
  const d = val instanceof Date ? val : new Date(val);
  if (isNaN(d.getTime())) return String(val);
  const format = includeTime ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd';
  return Utilities.formatDate(d, TIMEZONE, format);
}

function fmtNum(n) {
  return Number(n || 0).toLocaleString('id-ID');
}

function parseAngka_(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  let text = String(value).trim().replace(/\s/g, '');
  if (text.startsWith('#')) return 0;
  if (text.includes('.') && text.includes(',')) text = text.replace(/\./g, '').replace(',', '.');
  else if (text.includes(',')) text = text.replace(',', '.');
  const angka = parseFloat(text);
  return isNaN(angka) ? 0 : angka;
}

function escapeTg(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/\\/g, '\\\\')
    .replace(/\*/g, '\\*')
    .replace(/\`/g, '\\`')
    .replace(/\[/g, '\\[')
    .replace(/\]/g, '\\]')
    .replace(/_/g, '\\_');
}

function escapeHtml_(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getAdminPin_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || DEFAULT_PIN;
}

function changeAdminPin(p, newPinArg) {
  let oldPin, newPin;
  if (typeof p === 'object' && p !== null) {
    oldPin = p.oldPin;
    newPin = p.newPin;
  } else {
    oldPin = p;
    newPin = newPinArg;
  }
  const currentPin = getAdminPin_();
  if (!oldPin || oldPin !== currentPin) {
    return { success: false, message: 'PIN lama tidak sesuai!' };
  }
  if (!newPin || newPin.length !== 6 || !/^\d{6}$/.test(newPin)) {
    return { success: false, message: 'PIN baru harus terdiri dari 6 digit angka!' };
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', newPin);
  return { success: true, message: 'PIN administrator berhasil diperbarui!' };
}

// ===================================================================
// WEB APP ENTRY POINT
// ===================================================================

function doGet(e) {
  if (e && e.parameter) {
    if (e.parameter.page === 'bincard') {
      return renderBincardPrintPage(e.parameter);
    }
    if (e.parameter.page === 'manifest') {
      return getManifestJson_();
    }
    if (e.parameter.page === 'sw') {
      return getServiceWorkerJs_();
    }
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Gudang Elektrik — Inventory System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-status-bar-style', 'black-translucent')
    .addMetaTag('apple-mobile-web-app-title', 'Gudang Elektrik')
    .addMetaTag('application-name', 'Gudang Elektrik')
    .addMetaTag('theme-color', '#0F172A');
}

function getManifestJson_() {
  const manifest = {
    name: "Gudang Elektrik — Inventory Management",
    short_name: "Gudang Elektrik",
    description: "Sistem Manajemen Stok & Mutasi Real-Time Gudang Elektrik",
    start_url: "https://script.google.com/macros/s/AKfycbzMZQ1Qngh1r4ddx-adIPl1-UTiIgrb219ErnyYlXQiVVuL8mSoV5HVez7BgXosFpuG/exec",
    scope: "/",
    display: "standalone",
    orientation: "any",
    background_color: "#0F172A",
    theme_color: "#0F172A",
    icons: [
      {
        src: "https://cdn-icons-png.flaticon.com/512/2933/2933245.png",
        sizes: "192x192",
        type: "image/png",
        purpose: "any maskable"
      },
      {
        src: "https://cdn-icons-png.flaticon.com/512/2933/2933245.png",
        sizes: "512x512",
        type: "image/png",
        purpose: "any maskable"
      }
    ]
  };
  return ContentService.createTextOutput(JSON.stringify(manifest))
    .setMimeType(ContentService.MimeType.JSON);
}

function getServiceWorkerJs_() {
  const swCode = `
const CACHE_NAME = 'elk-cache-v1';
const OFFLINE_URLS = ['https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap'];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(OFFLINE_URLS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(
      keys.map((k) => { if (k !== CACHE_NAME) return caches.delete(k); })
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
`;
  return ContentService.createTextOutput(swCode)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ===================================================================
// ENSURE SHEETS & SEEDING (1.006 ITEMS)
// ===================================================================

function ensureSheets() {
  const ss = getSpreadsheet();

  const configs = [
    {
      name: DB_SHEET,
      headers: DB_HEADERS
    },
    {
      name: 'mutasi',
      headers: [
        'tx_id', 'timestamp', 'type', 'nik', 'seksi', 'rak',
        'item_code', 'spesifikasi', 'qty', 'satuan',
        'keterangan', 'stok_before', 'stok_after', 'source',
        'tanggal_pengambilan_barang', 'tanggal_penerimaan_barang'
      ]
    },
    {
      name: 'audit',
      headers: [
        'adj_id', 'timestamp', 'user_role', 'method', 'rak',
        'item_code', 'spesifikasi', 'reason',
        'stok_before', 'stok_after', 'selisih', 'source'
      ]
    },
    {
      name: 'stok_snapshot',
      headers: [
        'snapshot_id', 'timestamp', 'rak', 'item_code', 'spesifikasi',
        'stok_qty', 'minimal_stok', 'satuan', 'remarks',
        'triggered_by', 'notes'
      ]
    },
    {
      name: 'log_laporan',
      headers: [
        'timestamp', 'tipe_laporan', 'saluran', 'format_dokumen',
        'penerima', 'status', 'detail_ringkasan', 'dieksekusi_oleh'
      ]
    }
  ];

  configs.forEach(conf => {
    let sh = ss.getSheetByName(conf.name);
    if (!sh) {
      sh = ss.insertSheet(conf.name);
      const range = sh.getRange(1, 1, 1, conf.headers.length);
      range.setValues([conf.headers]);
      range.setFontWeight('bold').setBackground('#0F172A').setFontColor('#FFFFFF')
           .setBorder(true, true, true, true, true, true);
      sh.setRowHeight(1, 34);
      sh.setFrozenRows(1);
    }
  });
}

function seedInitialStock(force = false) {
  ensureSheets();
  const ss = getSpreadsheet();
  const dbSheet = ss.getSheetByName(DB_SHEET);
  const lastRow = dbSheet.getLastRow();

  if (lastRow > 1 && !force) {
    return {
      success: false,
      message: 'Sheet database sudah memiliki ' + (lastRow - 1) + ' data item. Gunakan mode FORCE untuk menimpa.'
    };
  }

  if (force && lastRow > 1) {
    dbSheet.getRange(2, 1, lastRow - 1, dbSheet.getLastColumn()).clearContent();
  }

  const seed = getElectricSeedData_();
  if (!seed || !seed.length) {
    return { success: false, message: 'Seed data tidak ditemukan!' };
  }

  // Exact 10 columns: Rak, Item Code, Spesifikasi, Stok Awal, Masuk (In), Keluar (Out), Stok Akhir, Satuan, Minimal Stok, Remarks
  const rows = seed.map(item => [
    item[0] || '',       // Rak
    item[1] || '',       // Item Code
    item[2] || '',       // Spesifikasi
    Number(item[3]) || 0,// Stok Awal
    Number(item[4]) || 0,// Masuk (In)
    Number(item[5]) || 0,// Keluar (Out)
    Number(item[6]) || 0,// Stok Akhir
    item[7] || 'PCS',    // Satuan
    Number(item[8]) || 1,// Minimal Stok
    item[9] || ''        // Remarks
  ]);

  dbSheet.getRange(2, 1, rows.length, 10).setValues(rows);
  formatProfessionalSpreadsheet({ pin: getAdminPin_() });

  return {
    success: true,
    message: 'Berhasil mengimpor ' + rows.length + ' item katalog Gudang Elektrik ke database!'
  };
}

function initializeDatabase(payload) {
  if (payload && payload.pin && payload.pin !== getAdminPin_()) {
    return { success: false, message: 'Akses ditolak. PIN admin tidak valid.' };
  }

  try {
    ensureSheets();
    const force = payload ? !!payload.force : false;
    const res = seedInitialStock(force);
    return {
      success: true,
      message: 'Inisialisasi database Gudang Elektrik selesai. ' + (res.message || '')
    };
  } catch (e) {
    return { success: false, message: 'Gagal inisialisasi: ' + e.toString() };
  }
}

// ===================================================================
// QUERY DATA METADATA & STOK
// ===================================================================

function getInitialData() {
  try {
    ensureSheets();
    const ss = getSpreadsheet();
    const dbSheet = ss.getSheetByName(DB_SHEET);
    const data = dbSheet.getDataRange().getValues();
    const hMap = getHeaderMap(dbSheet);
    const keys = Object.keys(hMap);

    const rIdx = findColIdx_(keys, ['rak', 'lokasi_rak', 'lokasi', 'loc']);
    const cIdx = findColIdx_(keys, ['item_code', 'code_barang', 'kode']);

    const colRak  = rIdx !== -1 ? hMap[keys[rIdx]] : 0;
    const colCode = cIdx !== -1 ? hMap[keys[cIdx]] : 1;

    const rakSet = new Set();
    let hasItems = false;

    if (data.length > 1) {
      data.slice(1).forEach(r => {
        if (r[colCode]) {
          hasItems = true;
          const rVal = String(r[colRak] || '').trim();
          if (rVal) rakSet.add(rVal);
        }
      });
    }

    const props = PropertiesService.getScriptProperties();

    return {
      success: true,
      data: {
        raks: [...rakSet].sort(),
        sections: ELECTRICAL_SECTIONS,
        satuan: STANDARD_UNITS,
        adminPin: getAdminPin_(),
        isDatabaseEmpty: !hasItems,
        telegramEnabled: props.getProperty('TELEGRAM_ENABLED') === 'true',
        emailEnabled: props.getProperty('EMAIL_ENABLED') === 'true',
        emailRecipients: props.getProperty('EMAIL_RECIPIENTS') || '',
        emailSubjectPrefix: props.getProperty('EMAIL_SUBJECT_PREFIX') || '[GUDANG ELEKTRIK]'
      }
    };
  } catch (err) {
    return { success: false, message: 'Gagal inisialisasi: ' + err.toString() };
  }
}

function getStockData() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SHEET);
    if (!sheet) return { success: false, message: 'Sheet database belum dibuat.' };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        success: true,
        list: [],
        lastUpdated: formatDate_(new Date(), true).split(' ')[1]
      };
    }

    const hMap = getHeaderMap(sheet);
    const keys = Object.keys(hMap);

    const rIdx   = findColIdx_(keys, ['rak', 'lokasi_rak', 'lokasi', 'loc']);
    const cIdx   = findColIdx_(keys, ['item_code', 'code_barang', 'kode_barang', 'kode']);
    const nIdx   = findColIdx_(keys, ['spesifikasi', 'nama_barang', 'spacification', 'nama']);
    const saIdx  = findColIdx_(keys, ['stok_awal', 'initial_stock', 'awal']);
    const inIdx  = findColIdx_(keys, ['masuk_in', 'in', 'in_qty', 'masuk']);
    const outIdx = findColIdx_(keys, ['keluar_out', 'out', 'out_qty', 'keluar']);
    const fsIdx  = findColIdx_(keys, ['stok_akhir', 'final_stock', 'stok_berjalan', 'stok', 'akhir']);
    const satIdx = findColIdx_(keys, ['satuan', 'satuan_barang', 'unit']);
    const minIdx = findColIdx_(keys, ['minimal_stok', 'min_stok', 'min_stock', 'minimum_stok', 'min_qty', 'minimum']);
    const remIdx = findColIdx_(keys, ['remarks', 'remark', 'keterangan_tambahan', 'keterangan']);

    const colRak   = rIdx !== -1 ? hMap[keys[rIdx]] : 0;
    const colCode  = cIdx !== -1 ? hMap[keys[cIdx]] : 1;
    const colNama  = nIdx !== -1 ? hMap[keys[nIdx]] : 2;
    const colAwal  = saIdx !== -1 ? hMap[keys[saIdx]] : 3;
    const colIn    = inIdx !== -1 ? hMap[keys[inIdx]] : 4;
    const colOut   = outIdx !== -1 ? hMap[keys[outIdx]] : 5;
    const colAkhir = fsIdx !== -1 ? hMap[keys[fsIdx]] : 6;
    const colSat   = satIdx !== -1 ? hMap[keys[satIdx]] : 7;
    const colMin   = minIdx !== -1 ? hMap[keys[minIdx]] : 8;
    const colRem   = remIdx !== -1 ? hMap[keys[remIdx]] : 9;

    const list = data.slice(1)
      .filter(r => colCode !== -1 && r[colCode] !== undefined && String(r[colCode]).trim() !== '')
      .map((r, i) => {
        const cCode  = String(r[colCode]).trim();
        const cRak   = colRak !== -1 ? String(r[colRak] || '').trim() : '';
        const cNama  = colNama !== -1 ? String(r[colNama] || '').trim() : cCode;
        const sAwal  = colAwal !== -1 ? Number(r[colAwal]) || 0 : 0;
        const sIn    = colIn !== -1 ? Number(r[colIn]) || 0 : 0;
        const sOut   = colOut !== -1 ? Number(r[colOut]) || 0 : 0;
        let   sAkhir = colAkhir !== -1 && r[colAkhir] !== undefined && r[colAkhir] !== '' ? Number(r[colAkhir]) : (sAwal + sIn - sOut);
        if (isNaN(sAkhir)) sAkhir = sAwal + sIn - sOut;

        const cSat = colSat !== -1 && r[colSat] ? String(r[colSat]).trim() : 'PCS';
        const cMin = colMin !== -1 && r[colMin] !== undefined && r[colMin] !== '' ? Number(r[colMin]) : 1;
        const cRem = colRem !== -1 ? String(r[colRem] || '').trim() : '';

        return {
          no: i + 1,
          rak: cRak,
          gudang: cRak,
          code_barang: cCode,
          nama_barang: cNama || cCode,
          stok_awal: sAwal,
          in_qty: sIn,
          out_qty: sOut,
          final_stok: sAkhir,
          stok_akhir: sAkhir,
          satuan_barang: cSat,
          min_stok: cMin,
          remarks: cRem
        };
      });

    return {
      success: true,
      list: list,
      lastUpdated: formatDate_(new Date(), true).split(' ')[1]
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat data stok: ' + err.toString() };
  }
}

// ===================================================================
// LOG STOK SNAPSHOT
// ===================================================================

function logStokSnapshot(code, rak, newStok, triggeredBy, notes) {
  try {
    const ss = getSpreadsheet();
    const snapSheet = ss.getSheetByName('stok_snapshot');
    if (!snapSheet) return;

    const dbSheet = ss.getSheetByName(DB_SHEET);
    const dbData  = dbSheet.getDataRange().getValues();
    const hMap    = getHeaderMap(dbSheet);
    const keys    = Object.keys(hMap);

    const cIdx   = findColIdx_(keys, ['item_code', 'code_barang', 'kode']);
    const nIdx   = findColIdx_(keys, ['spesifikasi', 'nama_barang', 'nama']);
    const satIdx = findColIdx_(keys, ['satuan', 'satuan_barang', 'unit']);
    const minIdx = findColIdx_(keys, ['minimal_stok', 'min_stok', 'min_stock', 'min_qty']);
    const remIdx = findColIdx_(keys, ['remarks', 'remark', 'keterangan']);

    const colCode = cIdx !== -1 ? hMap[keys[cIdx]] : 1;
    const colNama = nIdx !== -1 ? hMap[keys[nIdx]] : 2;
    const colSat  = satIdx !== -1 ? hMap[keys[satIdx]] : 7;
    const colMin  = minIdx !== -1 ? hMap[keys[minIdx]] : 8;
    const colRem  = remIdx !== -1 ? hMap[keys[remIdx]] : 9;

    const row = dbData.find(r => String(r[colCode]).trim().toLowerCase() === String(code).trim().toLowerCase());

    const nama    = row ? row[colNama] : '';
    const satuan  = row ? row[colSat]  : 'PCS';
    const minStok = row ? (Number(row[colMin]) || 1) : 1;
    const remarks = row ? row[colRem]  : '';

    const snapId = 'SNP' + Date.now().toString().slice(-8);

    snapSheet.appendRow([
      snapId,
      new Date(),
      rak || '',
      code,
      nama,
      newStok,
      minStok,
      satuan,
      remarks,
      triggeredBy || 'SYSTEM',
      notes || ''
    ]);
  } catch (err) {
    Logger.log('[Warning] logStokSnapshot gagal: ' + err.toString());
  }
}

function createStockSnapshotAll(payload) {
  if (payload && payload.pin && payload.pin !== getAdminPin_()) {
    return { success: false, message: 'Akses ditolak. PIN admin tidak valid.' };
  }

  const stockRes = getStockData();
  if (!stockRes.success || !stockRes.list.length) {
    return { success: false, message: 'Tidak ada data stok untuk di-snapshot.' };
  }

  const ss = getSpreadsheet();
  const snapSheet = ss.getSheetByName('stok_snapshot');
  if (!snapSheet) return { success: false, message: 'Sheet stok_snapshot tidak ditemukan.' };

  const ts = new Date();
  const snapId = 'SNP-ALL-' + Date.now().toString().slice(-6);

  const rows = stockRes.list.map(r => [
    snapId, ts, r.rak || '', r.code_barang, r.nama_barang,
    r.final_stok, r.min_stok, r.satuan_barang, r.remarks || '',
    'MANUAL_BACKUP', 'Cadangan master stok penuh'
  ]);

  snapSheet.getRange(snapSheet.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
  return { success: true, message: 'Berhasil membuat snapshot untuk ' + rows.length + ' item!' };
}

// ===================================================================
// SUBMIT TRANSAKSI MUTASI (IN / OUT) — AUTO-RAK DARI MASTER
// ===================================================================

function submitTransaction(payload) {
  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(30000);
    ensureSheets();

    const ss = getSpreadsheet();
    const dbSheet = ss.getSheetByName(DB_SHEET);
    const mutasiSheet = ss.getSheetByName('mutasi');
    const dbData = dbSheet.getDataRange().getValues();
    const hMap = getHeaderMap(dbSheet);
    const keys = Object.keys(hMap);

    const rIdx   = findColIdx_(keys, ['rak', 'lokasi_rak', 'lokasi', 'loc']);
    const cIdx   = findColIdx_(keys, ['item_code', 'code_barang', 'kode']);
    const nIdx   = findColIdx_(keys, ['spesifikasi', 'nama_barang', 'nama']);
    const inIdx  = findColIdx_(keys, ['masuk_in', 'in', 'in_qty', 'masuk']);
    const outIdx = findColIdx_(keys, ['keluar_out', 'out', 'out_qty', 'keluar']);
    const fsIdx  = findColIdx_(keys, ['stok_akhir', 'final_stock', 'stok']);
    const satIdx = findColIdx_(keys, ['satuan', 'satuan_barang', 'unit']);
    const minIdx = findColIdx_(keys, ['minimal_stok', 'min_stok', 'min_stock', 'min_qty']);

    const colRak   = rIdx !== -1 ? hMap[keys[rIdx]] : 0;
    const colCode  = cIdx !== -1 ? hMap[keys[cIdx]] : 1;
    const colNama  = nIdx !== -1 ? hMap[keys[nIdx]] : 2;
    const colIn    = inIdx !== -1 ? hMap[keys[inIdx]] : 4;
    const colOut   = outIdx !== -1 ? hMap[keys[outIdx]] : 5;
    const colAkhir = fsIdx !== -1 ? hMap[keys[fsIdx]] : 6;
    const colSat   = satIdx !== -1 ? hMap[keys[satIdx]] : 7;
    const colMin   = minIdx !== -1 ? hMap[keys[minIdx]] : 8;

    let tglAmbil = '';
    let tglTerima = '';
    if (payload.type === 'OUT') {
      tglAmbil = parseHtmlDate_(payload.tanggal_pengambilan_barang);
    } else {
      tglTerima = parseHtmlDate_(payload.tanggal_penerimaan_barang);
    }

    const txId = 'TX' + Date.now().toString().slice(-6);
    const ts = new Date();

    const lookupCode = {};
    for (let i = 1; i < dbData.length; i++) {
      const code = String(dbData[i][colCode] || '').trim().toLowerCase();
      if (code && !lookupCode[code]) {
        lookupCode[code] = { rowIndex: i + 1, data: dbData[i] };
      }
    }

    const validatedItems = [];
    for (const item of payload.items) {
      const codeKey = String(item.code || '').trim().toLowerCase();
      const target = lookupCode[codeKey];

      if (!target) {
        throw new Error('Barang [' + item.code + '] tidak ditemukan di database master.');
      }

      const currentStok = Number(target.data[colAkhir]) || 0;
      const qty = Number(item.qty);
      if (qty <= 0) {
        throw new Error('Qty untuk [' + item.code + '] harus lebih besar dari 0.');
      }

      const newStok = payload.type === 'IN' ? (currentStok + qty) : (currentStok - qty);
      if (payload.type === 'OUT' && currentStok < qty) {
        throw new Error('Stok tidak mencukupi untuk ' + (target.data[colNama] || item.code) + '.\nStok tersedia: ' + currentStok + ' | Diminta: ' + qty);
      }

      validatedItems.push({ item, target, currentStok, qty, newStok });
    }

    const kritisAlert = [];
    const itemLines = [];
    const mutasiRows = [];

    validatedItems.forEach(({ item, target, currentStok, qty, newStok }) => {
      const rowIdx       = target.rowIndex;
      const itemCode     = target.data[colCode] || item.code;
      const itemNama     = target.data[colNama] || itemCode;
      const itemRak      = target.data[colRak]  || '';
      const itemSatuan   = target.data[colSat]  || 'PCS';
      const itemMinStok  = Number(target.data[colMin]) || 1;

      const currentIn  = Number(target.data[colIn]) || 0;
      const currentOut = Number(target.data[colOut]) || 0;

      if (payload.type === 'IN') {
        if (colIn !== -1) dbSheet.getRange(rowIdx, colIn + 1).setValue(currentIn + qty);
      } else {
        if (colOut !== -1) dbSheet.getRange(rowIdx, colOut + 1).setValue(currentOut + qty);
      }
      if (colAkhir !== -1) dbSheet.getRange(rowIdx, colAkhir + 1).setValue(newStok);

      logStokSnapshot(
        itemCode, itemRak, newStok,
        'MUTASI_' + payload.type,
        'TX: ' + txId + ' | NIK: ' + payload.nik + ' | Qty: ' + qty
      );

      mutasiRows.push([
        txId, ts, payload.type, payload.nik, payload.section, itemRak,
        itemCode, itemNama, qty, itemSatuan,
        payload.note || '', currentStok, newStok, 'WEBAPP',
        tglAmbil || '', tglTerima || ''
      ]);

      const arah = payload.type === 'IN' ? '▲' : '▼';
      const stokInfo = fmtNum(currentStok) + ' → *' + fmtNum(newStok) + '* ' + itemSatuan;
      itemLines.push(
        arah + ' *' + escapeTg(itemNama) + '*\n' +
        '   Kode: `' + escapeTg(itemCode) + '` | Qty: ' + (arah === '▲' ? '+' : '-') + fmtNum(qty) + ' ' + itemSatuan + '\n' +
        '   Lokasi Rak: ' + escapeTg(itemRak || '-') + ' | Stok: ' + stokInfo
      );

      if (payload.type === 'OUT' && newStok <= itemMinStok) {
        kritisAlert.push({ namaBarang: itemNama, satuanBarang: itemSatuan, newStok, minStok: itemMinStok, code: itemCode, rak: itemRak });
      }
    });

    if (mutasiRows.length > 0) {
      const startRow = mutasiSheet.getLastRow() + 1;
      mutasiSheet.getRange(startRow, 1, mutasiRows.length, 16).setValues(mutasiRows);
    }

    // Notifikasi Telegram
    try {
      const icon = payload.type === 'IN' ? '🟢 MASUK (IN)' : '🔴 KELUAR (OUT)';
      let tgMsg =
        '⚡ *MUTASI GUDANG ELEKTRIK — ' + icon + '*\n' +
        '━━━━━━━━━━━━━━━━━━━━━\n' +
        '📋 No. Bukti : `' + txId + '`\n' +
        '📅 Waktu     : ' + formatDate_(ts, true) + '\n' +
        '👤 Petugas   : ' + escapeTg(payload.nik) + '\n' +
        '🏭 Seksi     : ' + escapeTg(payload.section) + '\n' +
        (payload.note ? ('📝 Catatan   : ' + escapeTg(payload.note) + '\n') : '') +
        '━━━━━━━━━━━━━━━━━━━━━\n' +
        '*DAFTAR BARANG:*\n' +
        itemLines.join('\n') + '\n' +
        '━━━━━━━━━━━━━━━━━━━━━';

      if (kritisAlert.length > 0) {
        tgMsg += '\n⚠️ *PERINGATAN STOK MENIPIS / KRITIS:*\n';
        kritisAlert.forEach(k => {
          tgMsg += '• `' + escapeTg(k.code) + '` ' + escapeTg(k.namaBarang) +
                   ' (Rak ' + escapeTg(k.rak || '-') + ') Sisa: *' + fmtNum(k.newStok) + '* (Min: ' + fmtNum(k.minStok) + ')\n';
        });
      }

      sendTelegramMessage(tgMsg);
    } catch (eTg) {
      Logger.log('[Warning] Notifikasi telegram transaksi gagal: ' + eTg.toString());
    }

    return {
      success: true,
      message: 'Transaksi ' + (payload.type === 'IN' ? 'Barang Masuk' : 'Barang Keluar') + ' berhasil dicatat!',
      txId: txId,
      items: payload.items
    };
  } catch (err) {
    return { success: false, message: 'Gagal transaksi: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// ===================================================================
// ADMIN DATABASE CRUD (TAMBAH, EDIT, HAPUS, QUICK ADJUST)
// ===================================================================

function dbCRUD(payload) {
  if (!payload || payload.pin !== getAdminPin_()) {
    return { success: false, message: 'Akses ditolak. PIN admin tidak valid.' };
  }

  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(15000);
    ensureSheets();

    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SHEET);
    const data  = sheet.getDataRange().getValues();
    const hMap  = getHeaderMap(sheet);
    const keys  = Object.keys(hMap);

    const rIdx   = findColIdx_(keys, ['rak', 'lokasi_rak', 'lokasi', 'loc']);
    const cIdx   = findColIdx_(keys, ['item_code', 'code_barang', 'kode']);
    const nIdx   = findColIdx_(keys, ['spesifikasi', 'nama_barang', 'nama']);
    const saIdx  = findColIdx_(keys, ['stok_awal', 'initial_stock', 'awal']);
    const inIdx  = findColIdx_(keys, ['masuk_in', 'in', 'in_qty', 'masuk']);
    const outIdx = findColIdx_(keys, ['keluar_out', 'out', 'out_qty', 'keluar']);
    const fsIdx  = findColIdx_(keys, ['stok_akhir', 'final_stock', 'stok']);
    const satIdx = findColIdx_(keys, ['satuan', 'satuan_barang', 'unit']);
    const minIdx = findColIdx_(keys, ['minimal_stok', 'min_stok', 'min_stock', 'min_qty']);
    const remIdx = findColIdx_(keys, ['remarks', 'remark', 'keterangan']);

    const colRak   = rIdx !== -1 ? hMap[keys[rIdx]] : 0;
    const colCode  = cIdx !== -1 ? hMap[keys[cIdx]] : 1;
    const colNama  = nIdx !== -1 ? hMap[keys[nIdx]] : 2;
    const colAwal  = saIdx !== -1 ? hMap[keys[saIdx]] : 3;
    const colIn    = inIdx !== -1 ? hMap[keys[inIdx]] : 4;
    const colOut   = outIdx !== -1 ? hMap[keys[outIdx]] : 5;
    const colAkhir = fsIdx !== -1 ? hMap[keys[fsIdx]] : 6;
    const colSat   = satIdx !== -1 ? hMap[keys[satIdx]] : 7;
    const colMin   = minIdx !== -1 ? hMap[keys[minIdx]] : 8;
    const colRem   = remIdx !== -1 ? hMap[keys[remIdx]] : 9;

    // ── CREATE ──
    if (payload.op === 'CREATE') {
      const reqCode = String(payload.code || '').trim();
      const reqRak  = String(payload.rak  || '').trim();

      const isDup = data.slice(1).some(r =>
        String(r[colCode]).trim().toLowerCase() === reqCode.toLowerCase()
      );
      if (isDup) {
        return { success: false, message: 'Kode barang [' + reqCode + '] sudah terdaftar di database.' };
      }

      const initStk = Number(payload.stok_awal !== undefined ? payload.stok_awal : payload.stok) || 0;
      const minStk  = Number(payload.min) || 1;

      sheet.appendRow([
        reqRak,
        reqCode,
        payload.nama || reqCode,
        initStk,
        0,
        0,
        initStk,
        payload.satuan || 'PCS',
        minStk,
        payload.remarks || ''
      ]);

      logStokSnapshot(reqCode, reqRak, initStk, 'MASTER_CREATE', 'Item baru ditambahkan oleh admin');
      return { success: true, message: 'Barang [' + reqCode + '] berhasil ditambahkan ke database.' };
    }

    // ── UPDATE ──
    else if (payload.op === 'UPDATE') {
      const targetCode = String(payload.code || '').trim();
      let rowIndex = -1;

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][colCode]).trim().toLowerCase() === targetCode.toLowerCase()) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex === -1) {
        return { success: false, message: 'Barang [' + targetCode + '] tidak ditemukan.' };
      }

      if (payload.rak !== undefined && colRak !== -1) sheet.getRange(rowIndex, colRak + 1).setValue(payload.rak);
      if (payload.nama !== undefined && colNama !== -1) sheet.getRange(rowIndex, colNama + 1).setValue(payload.nama);
      if (payload.stok_awal !== undefined && colAwal !== -1) sheet.getRange(rowIndex, colAwal + 1).setValue(Number(payload.stok_awal) || 0);
      if (payload.satuan !== undefined && colSat !== -1) sheet.getRange(rowIndex, colSat + 1).setValue(payload.satuan);
      if (payload.min !== undefined && colMin !== -1) sheet.getRange(rowIndex, colMin + 1).setValue(Number(payload.min) || 1);
      if (payload.remarks !== undefined && colRem !== -1) sheet.getRange(rowIndex, colRem + 1).setValue(payload.remarks);
      if (payload.stok !== undefined && colAkhir !== -1) {
        sheet.getRange(rowIndex, colAkhir + 1).setValue(Number(payload.stok) || 0);
        logStokSnapshot(targetCode, payload.rak || '', Number(payload.stok) || 0, 'MASTER_UPDATE', 'Penyesuaian stok via Edit Barang');
      }

      return { success: true, message: 'Data barang [' + targetCode + '] berhasil diperbarui.' };
    }

    // ── DELETE ──
    else if (payload.op === 'DELETE') {
      const targetCode = String(payload.code || '').trim();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][colCode]).trim().toLowerCase() === targetCode.toLowerCase()) {
          sheet.deleteRow(i + 1);
          logStokSnapshot(targetCode, data[i][colRak] || '', 0, 'MASTER_DELETE', 'Item dihapus oleh admin');
          return { success: true, message: 'Barang [' + targetCode + '] berhasil dihapus dari database.' };
        }
      }
      return { success: false, message: 'Barang [' + targetCode + '] tidak ditemukan.' };
    }

    // ── QUICK ADJUST ──
    else if (payload.op === 'QUICK_ADJUST') {
      const targetCode = String(payload.code || '').trim();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][colCode]).trim().toLowerCase() === targetCode.toLowerCase()) {
          const oldStok = Number(data[i][colAkhir]) || 0;
          const newStok = Number(payload.newStok);

          sheet.getRange(i + 1, colAkhir + 1).setValue(newStok);

          const auditSheet = ss.getSheetByName('audit');
          if (auditSheet) {
            auditSheet.appendRow([
              'ADJ' + Date.now().toString().slice(-6),
              new Date(),
              'ADMIN',
              'QUICK_ADJUST',
              data[i][colRak] || '',
              targetCode,
              data[i][colNama] || targetCode,
              payload.reason || 'Penyesuaian stok cepat oleh admin',
              oldStok,
              newStok,
              newStok - oldStok,
              'WEBAPP'
            ]);
          }

          logStokSnapshot(targetCode, data[i][colRak] || '', newStok, 'QUICK_ADJUST', payload.reason);
          return { success: true, message: 'Stok [' + targetCode + '] berhasil disesuaikan menjadi ' + newStok + '.' };
        }
      }
      return { success: false, message: 'Barang [' + targetCode + '] tidak ditemukan.' };
    }

    return { success: false, message: 'Operasi CRUD tidak valid: ' + payload.op };
  } catch (err) {
    return { success: false, message: 'Gagal dbCRUD: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ===================================================================
// STOCK OPNAME / AUDIT FISIK (BULK AUDIT)
// ===================================================================

function bulkAudit(payload) {
  if (!payload || payload.pin !== getAdminPin_()) {
    return { success: false, message: 'Akses ditolak. PIN admin tidak valid.' };
  }

  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(25000);
    ensureSheets();

    const ss = getSpreadsheet();
    const dbSheet = ss.getSheetByName(DB_SHEET);
    const auditSheet = ss.getSheetByName('audit');
    const data = dbSheet.getDataRange().getValues();
    const hMap = getHeaderMap(dbSheet);
    const keys = Object.keys(hMap);

    const rIdx   = findColIdx_(keys, ['rak', 'lokasi_rak', 'lokasi', 'loc']);
    const cIdx   = findColIdx_(keys, ['item_code', 'code_barang', 'kode']);
    const nIdx   = findColIdx_(keys, ['spesifikasi', 'nama_barang', 'nama']);
    const fsIdx  = findColIdx_(keys, ['stok_akhir', 'final_stock', 'stok']);

    const colRak   = rIdx !== -1 ? hMap[keys[rIdx]] : 0;
    const colCode  = cIdx !== -1 ? hMap[keys[cIdx]] : 1;
    const colNama  = nIdx !== -1 ? hMap[keys[nIdx]] : 2;
    const colAkhir = fsIdx !== -1 ? hMap[keys[fsIdx]] : 6;

    const ts = new Date();
    const adjId = 'AUD-' + Date.now().toString().slice(-6);
    let changed = 0;
    const auditRows = [];

    payload.adjustments.forEach(adj => {
      const code = String(adj.code).trim().toLowerCase();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][colCode]).trim().toLowerCase() === code) {
          const oldVal = Number(data[i][colAkhir]) || 0;
          const newVal = Number(adj.newVal);

          if (oldVal !== newVal) {
            dbSheet.getRange(i + 1, colAkhir + 1).setValue(newVal);
            logStokSnapshot(adj.code, data[i][colRak] || '', newVal, 'STOCK_OPNAME', 'Audit: ' + adjId + ' | ' + payload.reason);
            auditRows.push([
              adjId, ts, 'ADMIN', 'BULK_AUDIT',
              data[i][colRak] || '', adj.code, data[i][colNama] || adj.code,
              payload.reason || 'Stock Opname Berkala',
              oldVal, newVal, newVal - oldVal, 'WEBAPP'
            ]);
            changed++;
          }
          break;
        }
      }
    });

    if (auditRows.length > 0 && auditSheet) {
      auditSheet.getRange(auditSheet.getLastRow() + 1, 1, auditRows.length, 12).setValues(auditRows);
    }

    return {
      success: true,
      message: 'Stock opname berhasil dieksekusi. ' + changed + ' item mengalami penyesuaian fisik.'
    };
  } catch (err) {
    return { success: false, message: 'Gagal eksekusi audit: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ===================================================================
// QUERY RIWAYAT MUTASI & LAPORAN
// ===================================================================

function getMutasiHistory(p) {
  try {
    ensureSheets();
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('mutasi');
    if (!sheet) return { success: true, list: [], total: 0, displayed: 0, hasMore: false };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, list: [], total: 0, displayed: 0, hasMore: false };

    const hMap = getHeaderMap(sheet);
    let rows = data.slice(1);

    const fromDate = p && p.from ? new Date(p.from + 'T00:00:00') : null;
    const toDate   = p && p.to   ? new Date(p.to   + 'T23:59:59') : null;

    rows = rows.filter(r => {
      const ts = r[hMap.timestamp];
      const d  = ts instanceof Date ? ts : new Date(ts);
      if (fromDate && d < fromDate) return false;
      if (toDate   && d > toDate)   return false;
      if (p && p.rak && String(r[hMap.rak] || '').trim().toLowerCase() !== p.rak.toLowerCase()) return false;
      if (p && p.type && p.type !== 'ALL' && r[hMap.type] !== p.type) return false;
      return true;
    });

    rows.reverse();

    const limit = (p && p.limit) || 200;
    const displayed = rows.slice(0, limit);

    const list = displayed.map(r => ({
      tx_id:      r[hMap.tx_id] || '',
      timestamp:  formatDate_(r[hMap.timestamp], true),
      type:       r[hMap.type] || '',
      nik:        r[hMap.nik] || '',
      section:    r[hMap.seksi] || r[hMap.section] || '',
      rak:        r[hMap.rak] || '',
      gudang:     r[hMap.rak] || '',
      code_barang:r[hMap.code_barang] || r[hMap.item_code] || '',
      nama_barang:r[hMap.nama_barang] || r[hMap.spesifikasi] || '',
      qty:        Number(r[hMap.qty]) || 0,
      satuan:     r[hMap.satuan] || 'PCS',
      keterangan: r[hMap.keterangan] || '',
      stok_before:Number(r[hMap.stok_before]) || 0,
      stok_after: Number(r[hMap.stok_after]) || 0
    }));

    return {
      success: true,
      list: list,
      total: rows.length,
      displayed: list.length,
      hasMore: rows.length > limit
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat history mutasi: ' + err.toString() };
  }
}

function getReportData(p) {
  try {
    ensureSheets();
    const ss = getSpreadsheet();
    const mutasiSheet = ss.getSheetByName('mutasi');
    if (!mutasiSheet) return { success: true, list: [] };

    const data = mutasiSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, list: [] };

    const hMap = getHeaderMap(mutasiSheet);
    const fromDate = p && p.from ? new Date(p.from + 'T00:00:00') : null;
    const toDate   = p && p.to   ? new Date(p.to   + 'T23:59:59') : null;

    const summaryMap = {};

    data.slice(1).forEach(r => {
      const ts = r[hMap.timestamp];
      const d  = ts instanceof Date ? ts : new Date(ts);
      if (fromDate && d < fromDate) return;
      if (toDate   && d > toDate)   return;

      const rak = String(r[hMap.rak] || '').trim();
      if (p && p.rak && rak.toLowerCase() !== p.rak.toLowerCase()) return;

      const code = String(r[hMap.code_barang] || r[hMap.item_code] || '').trim();
      const nama = String(r[hMap.nama_barang] || r[hMap.spesifikasi] || code).trim();
      const sat  = String(r[hMap.satuan] || 'PCS').trim();
      const qty  = Number(r[hMap.qty]) || 0;
      const type = r[hMap.type];

      const key = code + '|' + rak;
      if (!summaryMap[key]) {
        summaryMap[key] = {
          code, nama, rak, satuan: sat, in: 0, out: 0, lastStok: Number(r[hMap.stok_after]) || 0
        };
      }

      if (type === 'IN') summaryMap[key].in += qty;
      else if (type === 'OUT') summaryMap[key].out += qty;
      summaryMap[key].lastStok = Number(r[hMap.stok_after]) || 0;
    });

    const list = Object.values(summaryMap).sort((a, b) => (b.out + b.in) - (a.out + a.in));
    return { success: true, list: list };
  } catch (err) {
    return { success: false, message: 'Gagal memuat rekap laporan: ' + err.toString() };
  }
}

function getStokSnapshot() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('stok_snapshot');
    if (!sheet) return { success: true, list: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, list: [] };

    const hMap = getHeaderMap(sheet);
    const rows = data.slice(1).reverse().slice(0, 100);

    const list = rows.map(r => ({
      snapshot_id: r[hMap.snapshot_id] || '',
      timestamp:   formatDate_(r[hMap.timestamp], true),
      rak:         r[hMap.rak] || '',
      code_barang: r[hMap.code_barang] || r[hMap.item_code] || '',
      nama_barang: r[hMap.nama_barang] || r[hMap.spesifikasi] || '',
      stok_qty:    Number(r[hMap.stok_qty]) || 0,
      min_stok:    Number(r[hMap.minimal_stok] || r[hMap.min_stok]) || 0,
      satuan:      r[hMap.satuan] || 'PCS',
      remarks:     r[hMap.remarks] || '',
      triggered_by:r[hMap.triggered_by] || '',
      notes:       r[hMap.notes] || ''
    }));

    return { success: true, list: list };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===================================================================
// DOKUMEN GENERATOR (PDF, EXCEL, CSV) & MULTI-CHANNEL DISPATCH
// ===================================================================

function generateReportPdfBlob(title, meta, summaryText, rows, dateStr, timeStr) {
  const sanitize = s => String(s || '').replace(/[^a-zA-Z0-9_\-]/g, '_');
  const fileName = 'Laporan_' + sanitize(title) + '_' + dateStr + '.pdf';

  let tableHtml = '';
  if (rows && rows.length > 0) {
    const keys = Object.keys(rows[0]);
    tableHtml = `
      <table class="report-table">
        <thead>
          <tr>
            <th style="width:30px;text-align:center;">#</th>
            ` + keys.map(k => '<th>' + escapeHtml_(k.replace(/_/g, ' ').toUpperCase()) + '</th>').join('') + `
          </tr>
        </thead>
        <tbody>
          ` + rows.map((r, i) => `
            <tr class="` + (i % 2 === 0 ? 'even' : 'odd') + `">
              <td style="text-align:center;color:#64748B;font-weight:bold;">` + (i + 1) + `</td>
              ` + keys.map(k => {
                const val = r[k] !== undefined ? r[k] : '';
                const isNum = typeof val === 'number' || (val !== '' && !isNaN(val) && !isNaN(parseFloat(val)) && !String(val).includes('-') && !String(val).includes('/'));
                const align = isNum ? 'text-align:right;' : 'text-align:left;';
                return '<td style="' + align + '">' + escapeHtml_(val) + '</td>';
              }).join('') + `
            </tr>
          `).join('') + `
        </tbody>
      </table>
    `;
  } else {
    tableHtml = '<p style="text-align:center;color:#64748B;padding:24px;">Tidak ada baris data untuk ditampilkan.</p>';
  }

  const metaHtml = meta ? `
    <div class="meta-card">
      <table style="width:100%;font-size:9pt;border-collapse:collapse;">
        <tr>
          <td style="width:50%;vertical-align:top;">
            <b>Periode:</b> ` + escapeHtml_(meta.from || '-') + ` s/d ` + escapeHtml_(meta.to || '-') + `<br>
            <b>Lokasi Rak:</b> ` + escapeHtml_(meta.rak || 'Semua Rak') + `
          </td>
          <td style="width:50%;text-align:right;vertical-align:top;color:#475569;">
            <b>Dicetak:</b> ` + dateStr + ` ` + timeStr + ` WIB<br>
            <b>Sistem:</b> Gudang Elektrik Management
          </td>
        </tr>
      </table>
    </div>
  ` : '';

  const fullHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        @page { size: A4 landscape; margin: 10mm 12mm; }
        body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 8.5pt; color: #0F172A; margin: 0; padding: 0; }
        .header { border-bottom: 2px solid #0F172A; padding-bottom: 8px; margin-bottom: 12px; }
        .brand-title { font-size: 13pt; font-weight: 800; color: #0F172A; text-transform: uppercase; letter-spacing: 0.5px; }
        .brand-sub { font-size: 7.5pt; color: #64748B; margin-top: 2px; }
        .report-title { font-size: 10.5pt; font-weight: 700; color: #1E293B; margin-top: 6px; }
        .meta-card { background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 4px; padding: 6px 10px; margin-bottom: 12px; }
        .summary-box { background: #EFF6FF; border-left: 3px solid #2563EB; padding: 6px 10px; font-size: 8pt; color: #1E40AF; margin-bottom: 12px; }
        table.report-table { width: 100%; border-collapse: collapse; font-size: 7.5pt; }
        table.report-table th { background: #0F172A; color: #FFFFFF; font-weight: 700; text-align: left; padding: 5px 6px; border: 1px solid #0F172A; }
        table.report-table td { padding: 4px 6px; border: 1px solid #E2E8F0; }
        tr.even { background: #FFFFFF; }
        tr.odd { background: #F8FAFC; }
        .footer { margin-top: 16px; border-top: 1px solid #CBD5E1; padding-top: 6px; font-size: 7pt; color: #94A3B8; text-align: right; }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="brand-title">⚡ GUDANG ELEKTRIK — INVENTORY SYSTEM</div>
        <div class="brand-sub">Sistem Pengendalian Komponen & Material Elektrik</div>
        <div class="report-title">` + escapeHtml_(title) + `</div>
      </div>
      ` + metaHtml + `
      ` + (summaryText ? ('<div class="summary-box">' + escapeHtml_(summaryText) + '</div>') : '') + `
      ` + tableHtml + `
      <div class="footer">Dicetak otomatis dari Gudang Elektrik Web App pada ` + dateStr + ` ` + timeStr + ` WIB</div>
    </body>
    </html>
  `;

  return Utilities.newBlob(fullHtml, 'text/html', fileName).getAs('application/pdf').setName(fileName);
}

function generateReportExcelBlob(title, meta, summaryText, rows, dateStr, timeStr) {
  const sanitize = s => String(s || '').replace(/[^a-zA-Z0-9_\-]/g, '_');
  const fileName = 'Laporan_' + sanitize(title) + '_' + dateStr + '.xls';

  let tableXml = '';
  if (rows && rows.length > 0) {
    const keys = Object.keys(rows[0]);
    tableXml += '<Row ss:StyleID="HeaderStyle">';
    tableXml += '<Cell><Data ss:Type="String">#</Data></Cell>';
    keys.forEach(k => {
      tableXml += '<Cell><Data ss:Type="String">' + escapeHtml_(k.replace(/_/g, ' ').toUpperCase()) + '</Data></Cell>';
    });
    tableXml += '</Row>';

    rows.forEach((r, idx) => {
      tableXml += '<Row>';
      tableXml += '<Cell><Data ss:Type="Number">' + (idx + 1) + '</Data></Cell>';
      keys.forEach(k => {
        const val = r[k] !== undefined ? r[k] : '';
        const isNum = typeof val === 'number';
        if (isNum) {
          tableXml += '<Cell><Data ss:Type="Number">' + val + '</Data></Cell>';
        } else {
          tableXml += '<Cell><Data ss:Type="String">' + escapeHtml_(val) + '</Data></Cell>';
        }
      });
      tableXml += '</Row>';
    });
  }

  const xmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Styles>
  <Style ss:ID="HeaderStyle">
   <Font ss:Bold="1" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#0F172A" ss:Pattern="Solid"/>
   <Alignment ss:Vertical="Center"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Laporan">
  <Table>
   ` + tableXml + `
  </Table>
 </Worksheet>
</Workbook>`;

  return Utilities.newBlob(xmlContent, 'application/vnd.ms-excel', fileName);
}

function buildStockCSV(items, tanggal) {
  const BOM = '\uFEFF';
  const SEP = ',';

  const headers = [
    'No', 'Rak', 'Item Code', 'Spesifikasi', 'Stok Awal',
    'Masuk (In)', 'Keluar (Out)', 'Stok Akhir', 'Satuan',
    'Minimal Stok', 'Remarks', 'Status'
  ];

  let csv = BOM + headers.map(csvCell).join(SEP) + String.fromCharCode(13, 10);

  items.forEach((r, idx) => {
    const status = (+r.stok_akhir <= +r.min_stok) ? 'KRITIS' : 'AMAN';
    const row = [
      idx + 1,
      r.rak || '',
      r.code_barang || r.code,
      r.nama_barang || r.nama,
      r.stok_awal || 0,
      r.in_qty || 0,
      r.out_qty || 0,
      r.final_stok !== undefined ? r.final_stok : (r.stok_akhir || 0),
      r.satuan_barang || r.satuan || 'PCS',
      r.min_stok || 1,
      r.remarks || '',
      status
    ];
    csv += row.map(csvCell).join(SEP) + String.fromCharCode(13, 10);
  });

  return csv;
}

function csvCell(val) {
  if (val === null || val === undefined) return '""';
  const str = String(val).replace(/"/g, '""');
  return '"' + str + '"';
}

function buildMutasiSummary(mutasiSheet, tanggal) {
  try {
    const rows = mutasiSheet.getDataRange().getValues();
    if (rows.length <= 1) return null;

    const hMap = getHeaderMap(mutasiSheet);
    const fromDate = new Date(tanggal + 'T00:00:00');
    const toDate   = new Date(tanggal + 'T23:59:59');

    let totalIn = 0, totalOut = 0;
    const txIds = new Set();

    rows.slice(1).forEach(r => {
      const ts = r[hMap.timestamp];
      const date = ts instanceof Date ? ts : new Date(ts);
      if (isNaN(date.getTime())) return;
      if (date < fromDate || date > toDate) return;

      const qty = Number(r[hMap.qty]) || 0;
      if (r[hMap.type] === 'IN') { totalIn += qty; }
      else if (r[hMap.type] === 'OUT') { totalOut += qty; }
      if (r[hMap.tx_id]) txIds.add(r[hMap.tx_id]);
    });

    if (txIds.size === 0) return null;
    return 'Transaksi: ' + txIds.size + ' slip | Masuk: ' + fmtNum(totalIn) + ' | Keluar: ' + fmtNum(totalOut);
  } catch (e) {
    return null;
  }
}

function sendCustomReport(payload) {
  try {
    const props = PropertiesService.getScriptProperties();
    const tgEnabled = props.getProperty('TELEGRAM_ENABLED') === 'true';
    const emailEnabled = props.getProperty('EMAIL_ENABLED') === 'true' || Boolean(payload.customEmail);

    const now = new Date();
    const dateStr = formatDate_(now);
    const timeStr = Utilities.formatDate(now, TIMEZONE, 'HH:mm');

    const title = payload.title || 'LAPORAN GUDANG ELEKTRIK';
    const rows  = payload.rows || [];
    const meta  = payload.meta || { from: dateStr, to: dateStr, rak: 'Semua Rak' };

    const pdfBlob   = generateReportPdfBlob(title, meta, payload.summaryText || '', rows, dateStr, timeStr);
    const excelBlob = generateReportExcelBlob(title, meta, payload.summaryText || '', rows, dateStr, timeStr);

    const attachments = [];
    if (payload.format === 'pdf' || payload.format === 'both') attachments.push(pdfBlob);
    if (payload.format === 'excel' || payload.format === 'both') attachments.push(excelBlob);

    const results = [];

    // 1. Telegram
    if ((payload.channel === 'telegram' || payload.channel === 'both') && tgEnabled) {
      const summaryMsg =
        '⚡ *' + escapeTg(title) + '*\n' +
        '━━━━━━━━━━━━━━━━━━━━━\n' +
        '📅 Periode : ' + escapeTg(meta.from || dateStr) + ' s/d ' + escapeTg(meta.to || dateStr) + '\n' +
        '📍 Rak     : ' + escapeTg(meta.rak || 'Semua Rak') + '\n' +
        '📊 Total   : *' + rows.length + '* baris data\n' +
        (payload.summaryText ? ('📝 ' + escapeTg(payload.summaryText) + '\n') : '') +
        '━━━━━━━━━━━━━━━━━━━━━\n' +
        '📄 Dokumen laporan resmi terlampir di bawah.';

      sendTelegramMessage(summaryMsg);
      attachments.forEach(att => sendTelegramDocument(att, att.getName()));
      results.push('Telegram OK');
    }

    // 2. Email
    if ((payload.channel === 'email' || payload.channel === 'both') && emailEnabled) {
      const recipients = payload.customEmail || props.getProperty('EMAIL_RECIPIENTS') || '';
      if (recipients) {
        const prefix = props.getProperty('EMAIL_SUBJECT_PREFIX') || '[GUDANG ELEKTRIK]';
        const subject = prefix + ' ' + title + ' — ' + dateStr;
        const htmlBody = `
          <div style="font-family:sans-serif;max-width:680px;margin:auto;color:#0F172A;border:1px solid #E2E8F0;border-radius:8px;padding:24px;">
            <div style="background:#0F172A;color:white;padding:16px;border-radius:6px 6px 0 0;margin:-24px -24px 20px -24px;">
              <h2 style="margin:0;font-size:18px;">⚡ GUDANG ELEKTRIK — INVENTORY SYSTEM</h2>
              <p style="margin:4px 0 0 0;font-size:12px;color:#94A3B8;">Laporan Resmi Inventaris & Material Elektrik</p>
            </div>
            <h3 style="color:#1E293B;margin-top:0;">` + escapeHtml_(title) + `</h3>
            <p><b>Periode:</b> ` + escapeHtml_(meta.from || dateStr) + ` s/d ` + escapeHtml_(meta.to || dateStr) + `<br>
               <b>Lokasi Rak:</b> ` + escapeHtml_(meta.rak || 'Semua Rak') + `<br>
               <b>Total Baris:</b> ` + rows.length + ` item</p>
            ` + (payload.summaryText ? ('<div style="background:#F1F5F9;padding:10px 14px;border-radius:4px;margin:16px 0;font-size:13px;">' + escapeHtml_(payload.summaryText) + '</div>') : '') + `
            <p style="font-size:12px;color:#64748B;">Dokumen PDF dan Excel lengkap telah dilampirkan pada email ini.</p>
            <hr style="border:none;border-top:1px solid #E2E8F0;margin:20px 0;">
            <p style="font-size:11px;color:#94A3B8;">Dikirim otomatis oleh Sistem Web Gudang Elektrik.</p>
          </div>
        `;

        sendEmailReport(recipients, subject, htmlBody, attachments);
        results.push('Email OK (' + recipients + ')');
      }
    }

    logReportDispatch(payload.channel, payload.format, payload.customEmail || 'Default', 'SUCCESS', title + ' (' + rows.length + ' baris)', 'WEBAPP');

    return {
      success: true,
      message: 'Laporan berhasil dikirim! ' + (results.length ? '(' + results.join(', ') + ')' : '')
    };
  } catch (err) {
    logReportDispatch(payload.channel, payload.format, 'Error', 'FAILED', err.toString(), 'WEBAPP');
    return { success: false, message: 'Gagal mengirim laporan: ' + err.toString() };
  }
}

function sendDailyStockReport(channel = 'both', customRecipients = null) {
  const stockRes = getStockData();
  if (!stockRes.success) return { success: false, message: 'Gagal memuat data stok: ' + stockRes.message };

  const items = stockRes.list || [];
  const now = new Date();
  const dateStr = formatDate_(now);
  const timeStr = Utilities.formatDate(now, TIMEZONE, 'HH:mm');

  const payload = {
    title: 'LAPORAN REKAPITULASI STOK GUDANG ELEKTRIK',
    meta: { from: dateStr, to: dateStr, rak: 'Semua Rak' },
    rows: items,
    channel: channel,
    customEmail: customRecipients,
    format: 'both',
    summaryText: 'Total ' + items.length + ' item komponen elektrik tercatat pada ' + dateStr + ' pukul ' + timeStr + ' WIB.'
  };

  return sendCustomReport(payload);
}

function logReportDispatch(channel, format, recipients, status, details, user) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('log_laporan');
    if (!sheet) return;

    sheet.appendRow([
      new Date(),
      'LAPORAN_GUDANG_ELEKTRIK',
      channel || 'both',
      format || 'both',
      recipients || '-',
      status || 'SUCCESS',
      details || '-',
      user || 'SYSTEM'
    ]);
  } catch (e) {
    Logger.log('[Warning] logReportDispatch error: ' + e.toString());
  }
}

function getReportDispatchHistory() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('log_laporan');
    if (!sheet) return { success: true, list: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, list: [] };

    const hMap = getHeaderMap(sheet);
    const rows = data.slice(1).reverse().slice(0, 50);

    const list = rows.map(r => ({
      timestamp:        formatDate_(r[hMap.timestamp], true),
      tipe_laporan:     r[hMap.tipe_laporan] || 'LAPORAN',
      saluran:          r[hMap.saluran] || 'both',
      format_dokumen:   r[hMap.format_dokumen] || 'PDF & Excel',
      penerima:         r[hMap.penerima] || '-',
      status:           r[hMap.status] || 'SUCCESS',
      detail_ringkasan: r[hMap.detail_ringkasan] || '-',
      dieksekusi_oleh:  r[hMap.dieksekusi_oleh] || 'SYSTEM'
    }));

    return { success: true, list: list };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===================================================================
// TELEGRAM & EMAIL INTEGRATION
// ===================================================================

function sendTelegramMessage(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) return;

  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  const payload = {
    chat_id: chatId,
    text: text,
    parse_mode: 'Markdown'
  };

  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function sendTelegramDocument(blob, caption) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) return;

  const url = 'https://api.telegram.org/bot' + token + '/sendDocument';
  const payload = {
    chat_id: chatId,
    caption: caption || '',
    parse_mode: 'Markdown',
    document: blob
  };

  UrlFetchApp.fetch(url, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });
}

function sendEmailReport(to, subject, htmlBody, attachments) {
  if (!to) return;
  MailApp.sendEmail({
    to: to,
    subject: subject,
    htmlBody: htmlBody,
    attachments: attachments || []
  });
}

function testTelegram() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');

  if (!token || !chatId) {
    return { success: false, message: 'Bot Token atau Chat ID Telegram belum diatur!' };
  }

  const msg = '⚡ *TEST KONEKSI GUDANG ELEKTRIK*\nKoneksi bot Telegram berhasil terhubung!\nWaktu: ' + formatDate_(new Date(), true);
  try {
    sendTelegramMessage(msg);
    return { success: true, message: 'Pesan test berhasil dikirim ke Telegram!' };
  } catch (e) {
    return { success: false, message: 'Gagal mengirim test: ' + e.toString() };
  }
}

function testEmail() {
  const props = PropertiesService.getScriptProperties();
  const recipients = props.getProperty('EMAIL_RECIPIENTS');
  if (!recipients) {
    return { success: false, message: 'Email penerima belum diatur di konfigurasi!' };
  }

  try {
    const prefix = props.getProperty('EMAIL_SUBJECT_PREFIX') || '[GUDANG ELEKTRIK]';
    MailApp.sendEmail({
      to: recipients,
      subject: prefix + ' Test Koneksi Email Sistem Gudang Elektrik',
      htmlBody: '<p>Halo,<br><br>Ini adalah email uji coba dari <b>Sistem Gudang Elektrik</b>.<br>Koneksi notifikasi email berfungsi dengan baik.<br><br>Waktu: ' + formatDate_(new Date(), true) + '</p>'
    });
    return { success: true, message: 'Email test berhasil dikirim ke ' + recipients };
  } catch (e) {
    return { success: false, message: 'Gagal mengirim test email: ' + e.toString() };
  }
}

function getAdminConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    telegramEnabled:    props.getProperty('TELEGRAM_ENABLED') === 'true',
    telegramToken:      props.getProperty('TELEGRAM_BOT_TOKEN') || '',
    telegramChatId:     props.getProperty('TELEGRAM_CHAT_ID') || '',
    emailEnabled:       props.getProperty('EMAIL_ENABLED') === 'true',
    emailRecipients:    props.getProperty('EMAIL_RECIPIENTS') || '',
    emailSubjectPrefix: props.getProperty('EMAIL_SUBJECT_PREFIX') || '[GUDANG ELEKTRIK]'
  };
}

function saveAdminConfig(p) {
  if (!p || p.pin !== getAdminPin_()) {
    return { success: false, message: 'Akses ditolak. PIN admin tidak valid.' };
  }

  const props = PropertiesService.getScriptProperties();
  if (p.telegramEnabled !== undefined) props.setProperty('TELEGRAM_ENABLED', p.telegramEnabled ? 'true' : 'false');
  if (p.telegramToken !== undefined)   props.setProperty('TELEGRAM_BOT_TOKEN', p.telegramToken.trim());
  if (p.telegramChatId !== undefined)  props.setProperty('TELEGRAM_CHAT_ID', p.telegramChatId.trim());
  if (p.emailEnabled !== undefined)    props.setProperty('EMAIL_ENABLED', p.emailEnabled ? 'true' : 'false');
  if (p.emailRecipients !== undefined) props.setProperty('EMAIL_RECIPIENTS', p.emailRecipients.trim());
  if (p.emailSubjectPrefix !== undefined) props.setProperty('EMAIL_SUBJECT_PREFIX', p.emailSubjectPrefix.trim());

  return { success: true, message: 'Konfigurasi admin berhasil disimpan!' };
}

function saveTelegram(token, chatId) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('TELEGRAM_BOT_TOKEN', (token || '').trim());
  props.setProperty('TELEGRAM_CHAT_ID', (chatId || '').trim());
  props.setProperty('TELEGRAM_ENABLED', 'true');
  return { success: true, message: 'Pengaturan Telegram berhasil disimpan!' };
}

// ===================================================================
// SCHEDULED AUTOMATIC REPORT (TRIGGER)
// ===================================================================

function saveScheduleTrigger(p) {
  removeScheduleTrigger();

  const hour = (p && p.hour !== undefined) ? Number(p.hour) : 7;
  ScriptApp.newTrigger('sendScheduledReport')
    .timeBased()
    .atHour(hour)
    .everyDays(1)
    .inTimezone(TIMEZONE)
    .create();

  const props = PropertiesService.getScriptProperties();
  props.setProperty('SCHEDULE_ACTIVE', 'true');
  props.setProperty('SCHEDULE_HOUR', String(hour));
  props.setProperty('SCHEDULE_FREQ', (p && p.freq) || 'DAILY');
  props.setProperty('SCHEDULE_CHANNEL', (p && p.channel) || 'both');
  props.setProperty('SCHEDULE_FORMAT', (p && p.format) || 'both');

  return { success: true, message: 'Trigger laporan harian jam ' + String(hour).padStart(2, '0') + ':00 WIB berhasil dipasang!' };
}

function removeScheduleTrigger(p) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'sendScheduledReport') {
      ScriptApp.deleteTrigger(t);
    }
  });

  const props = PropertiesService.getScriptProperties();
  props.setProperty('SCHEDULE_ACTIVE', 'false');

  return { success: true, message: 'Trigger laporan terjadwal dinonaktifkan.' };
}

function getScheduleInfo() {
  const props = PropertiesService.getScriptProperties();
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some(t => t.getHandlerFunction() === 'sendScheduledReport');

  return {
    active: hasTrigger,
    count: hasTrigger ? 1 : 0,
    hour: Number(props.getProperty('SCHEDULE_HOUR') || 7),
    freq: props.getProperty('SCHEDULE_FREQ') || 'DAILY',
    channel: props.getProperty('SCHEDULE_CHANNEL') || 'both',
    format: props.getProperty('SCHEDULE_FORMAT') || 'both',
    lastRun: props.getProperty('SCHEDULE_LAST_RUN') || '-'
  };
}

function getTriggerInfo() {
  return getScheduleInfo();
}

function setupDailyTrigger() {
  return saveScheduleTrigger({ pin: getAdminPin_(), freq: 'DAILY', hour: 7, channel: 'both', format: 'both' });
}

function removeDailyTrigger() {
  return removeScheduleTrigger();
}

function sendScheduledReport() {
  const props = PropertiesService.getScriptProperties();
  const channel = props.getProperty('SCHEDULE_CHANNEL') || 'both';
  const format  = props.getProperty('SCHEDULE_FORMAT')  || 'both';
  props.setProperty('SCHEDULE_LAST_RUN', formatDate_(new Date(), true));

  const stockRes = getStockData();
  if (!stockRes.success) return;

  const items = stockRes.list || [];
  const now = new Date();
  const dateStr = formatDate_(now);
  const timeStr = Utilities.formatDate(now, TIMEZONE, 'HH:mm');

  const payload = {
    title: 'LAPORAN REKAPITULASI STOK ELEKTRIK',
    meta: { from: dateStr, to: dateStr, rak: 'Semua Rak' },
    rows: items,
    channel: channel,
    format: format,
    summaryText: 'Total ' + items.length + ' item komponen tercatat pada ' + dateStr + ' pukul ' + timeStr + ' WIB.'
  };

  sendCustomReport(payload);
}

// ===================================================================
// PROFESSIONAL SPREADSHEET FORMATTING (NAVY HEADER, ZEBRA, BORDERS)
// ===================================================================

function formatProfessionalSpreadsheet(p) {
  if (p && p.pin && p.pin !== getAdminPin_()) {
    return { success: false, message: 'Akses ditolak. PIN admin tidak valid.' };
  }

  const ss = getSpreadsheet();
  ensureSheets();

  const sheetConfigs = [
    {
      name: DB_SHEET,
      widths: [130, 160, 320, 90, 85, 85, 95, 80, 95, 160],
      alignments: ['center', 'center', 'left', 'right', 'right', 'right', 'right', 'center', 'right', 'left'],
      numberFormats: ['@', '@', '@', '#,##0', '#,##0', '#,##0', '#,##0', '@', '#,##0', '@'],
      hasRules: 'database'
    },
    {
      name: 'mutasi',
      widths: [110, 160, 90, 110, 130, 120, 150, 260, 90, 80, 220, 90, 90, 90, 110, 110],
      alignments: ['center', 'center', 'center', 'center', 'left', 'center', 'center', 'left', 'right', 'center', 'left', 'right', 'right', 'center', 'center', 'center'],
      numberFormats: ['@', 'yyyy-mm-dd hh:mm:ss', '@', '@', '@', '@', '@', '@', '#,##0', '@', '@', '#,##0', '#,##0', '@', 'yyyy-mm-dd', 'yyyy-mm-dd'],
      hasRules: 'mutasi'
    },
    {
      name: 'audit',
      widths: [110, 160, 90, 110, 120, 150, 260, 220, 90, 90, 90, 90],
      alignments: ['center', 'center', 'center', 'center', 'center', 'center', 'left', 'left', 'right', 'right', 'right', 'center'],
      numberFormats: ['@', 'yyyy-mm-dd hh:mm:ss', '@', '@', '@', '@', '@', '@', '#,##0', '#,##0', '#,##0', '@'],
      hasRules: 'audit'
    },
    {
      name: 'stok_snapshot',
      widths: [120, 160, 120, 150, 260, 90, 90, 80, 160, 110, 220],
      alignments: ['center', 'center', 'center', 'center', 'left', 'right', 'right', 'center', 'left', 'center', 'left'],
      numberFormats: ['@', 'yyyy-mm-dd hh:mm:ss', '@', '@', '@', '#,##0', '#,##0', '@', '@', '@', '@'],
      hasRules: 'snapshot'
    },
    {
      name: 'log_laporan',
      widths: [160, 160, 120, 140, 220, 100, 260, 140],
      alignments: ['center', 'left', 'center', 'center', 'left', 'center', 'left', 'center'],
      numberFormats: ['yyyy-mm-dd hh:mm:ss', '@', '@', '@', '@', '@', '@', '@'],
      hasRules: 'log'
    }
  ];

  sheetConfigs.forEach(conf => {
    const sh = ss.getSheetByName(conf.name);
    if (!sh) return;

    const lastRow = Math.max(sh.getLastRow(), 2);
    const lastCol = Math.max(sh.getLastColumn(), conf.widths.length);

    // 1. Header Format (Navy #0F172A)
    const hdr = sh.getRange(1, 1, 1, lastCol);
    hdr.setFontWeight('bold')
       .setFontFamily('Arial')
       .setFontSize(10)
       .setBackground('#0F172A')
       .setFontColor('#FFFFFF')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setWrap(false);
    sh.setRowHeight(1, 34);
    sh.setFrozenRows(1);

    // 2. Column widths
    conf.widths.forEach((w, i) => {
      if (i < lastCol) sh.setColumnWidth(i + 1, w);
    });

    // 3. Body formatting & zebra
    if (lastRow > 1) {
      for (let c = 0; c < conf.alignments.length; c++) {
        if (c < lastCol) {
          const bodyRange = sh.getRange(2, c + 1, lastRow - 1, 1);
          bodyRange.setFontFamily('Arial').setFontSize(9.5).setVerticalAlignment('middle');
          if (conf.alignments[c]) bodyRange.setHorizontalAlignment(conf.alignments[c]);
          if (conf.numberFormats[c]) bodyRange.setNumberFormat(conf.numberFormats[c]);
        }
      }

      for (let r = 2; r <= lastRow; r++) {
        const bg = (r % 2 === 0) ? '#FFFFFF' : '#F8FAFC';
        sh.getRange(r, 1, 1, lastCol).setBackground(bg);
      }

      sh.getRange(1, 1, lastRow, lastCol).setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
    }

    applySheetConditionalRules_(sh, conf.hasRules);
  });

  return { success: true, message: 'Spreadsheet Gudang Elektrik berhasil diformat rapi & berstandar korporat!' };
}

function applySheetConditionalRules_(sh, confName) {
  sh.clearConditionalFormatRules();
  const lastRow = Math.max(sh.getLastRow(), 2);
  const rules = [];

  if (confName === 'database') {
    // Col 7 (Stok Akhir) <= Col 9 (Minimal Stok) -> Merah muda soft #FEE2E2
    const stokRange = sh.getRange('G2:G' + lastRow);
    const ruleKritis = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2<=$I2')
      .setBackground('#FEE2E2')
      .setFontColor('#991B1B')
      .setBold(true)
      .setRanges([stokRange])
      .build();

    const ruleAman = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2>$I2')
      .setBackground('#DCFCE7')
      .setFontColor('#166534')
      .setBold(true)
      .setRanges([stokRange])
      .build();

    rules.push(ruleKritis, ruleAman);
  } else if (confName === 'mutasi') {
    const typeRange = sh.getRange('C2:C' + lastRow);
    const ruleIn = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('IN')
      .setBackground('#DCFCE7')
      .setFontColor('#166534')
      .setBold(true)
      .setRanges([typeRange])
      .build();

    const ruleOut = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OUT')
      .setBackground('#FEE2E2')
      .setFontColor('#991B1B')
      .setBold(true)
      .setRanges([typeRange])
      .build();

    rules.push(ruleIn, ruleOut);
  }

  sh.setConditionalFormatRules(rules);
}

// ===================================================================
// BINCARD GENERATOR (KARTU STOK A5 DUAL-SIDED)
// ===================================================================

function getBincardDataForPrint(filter) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SHEET);
    if (!sheet) throw new Error('Sheet database tidak ditemukan.');

    const data = sheet.getDataRange().getValues();
    const hMap = getHeaderMap(sheet);
    const keys = Object.keys(hMap);

    const rIdx   = findColIdx_(keys, ['rak', 'lokasi_rak', 'lokasi', 'loc']);
    const cIdx   = findColIdx_(keys, ['item_code', 'code_barang', 'kode']);
    const nIdx   = findColIdx_(keys, ['spesifikasi', 'nama_barang', 'nama']);
    const fsIdx  = findColIdx_(keys, ['stok_akhir', 'final_stock', 'stok']);
    const satIdx = findColIdx_(keys, ['satuan', 'satuan_barang', 'unit']);
    const minIdx = findColIdx_(keys, ['minimal_stok', 'min_stok', 'min_stock', 'min_qty']);
    const remIdx = findColIdx_(keys, ['remarks', 'remark', 'keterangan']);

    const colRak   = rIdx !== -1 ? hMap[keys[rIdx]] : 0;
    const colCode  = cIdx !== -1 ? hMap[keys[cIdx]] : 1;
    const colNama  = nIdx !== -1 ? hMap[keys[nIdx]] : 2;
    const colAkhir = fsIdx !== -1 ? hMap[keys[fsIdx]] : 6;
    const colSat   = satIdx !== -1 ? hMap[keys[satIdx]] : 7;
    const colMin   = minIdx !== -1 ? hMap[keys[minIdx]] : 8;
    const colRem   = remIdx !== -1 ? hMap[keys[remIdx]] : 9;

    const nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy');

    let items = data.slice(1)
      .filter(r => r[colCode])
      .map(r => ({
        rak:      String(r[colRak] || '').trim(),
        code:     String(r[colCode]).trim(),
        nama:     String(r[colNama] || '').trim(),
        satuan:   String(r[colSat] || 'PCS').trim(),
        stok:     Number(r[colAkhir]) || 0,
        min:      Number(r[colMin])   || 1,
        remarks:  String(r[colRem] || '').trim(),
        tglCetak: nowStr
      }));

    if (filter) {
      if (filter.rak) {
        items = items.filter(r => r.rak.toLowerCase() === filter.rak.toLowerCase());
      }
      if (filter.code) {
        items = items.filter(r => r.code.toLowerCase() === filter.code.toLowerCase());
      }
    }

    const pairedSheets = [];
    for (let i = 0; i < items.length; i += 2) {
      pairedSheets.push({
        itemLeft:  items[i],
        itemRight: items[i + 1] || null
      });
    }

    return { success: true, sheets: pairedSheets, totalItems: items.length };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function renderBincardPrintPage(param) {
  const code = param ? param.code : null;
  const rak  = param ? (param.rak || param.gudang) : null;
  const html = getBincardPrintHtml(code, rak);
  return HtmlService.createHtmlOutput(html)
    .setTitle('Cetak Bincard Stok A5 — Gudang Elektrik')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getBincardPrintHtml(code, rak) {
  const res = getBincardDataForPrint({ code: code, rak: rak });
  if (!res.success) {
    return '<h3>Error memuat data bincard: ' + res.message + '</h3>';
  }

  const sheets = res.sheets || [];
  let pagesHtml = '';

  sheets.forEach((sh, idx) => {
    pagesHtml += `
      <div class="print-sheet side-a">
        <div class="card-wrapper card-left">` + renderCardSideA_(sh.itemLeft) + `</div>
        <div class="cut-guide"></div>
        <div class="card-wrapper card-right">` + renderCardSideA_(sh.itemRight) + `</div>
      </div>
    `;

    pagesHtml += `
      <div class="print-sheet side-b">
        <div class="card-wrapper card-left">` + renderCardSideB_(sh.itemRight) + `</div>
        <div class="cut-guide"></div>
        <div class="card-wrapper card-right">` + renderCardSideB_(sh.itemLeft) + `</div>
      </div>
    `;
  });

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <title>Cetak Bincard Stok A5 — Gudang Elektrik</title>
      <style>
        @page { size: A4 landscape; margin: 0; }
        * { box-sizing: border-box; }
        body { margin: 0; padding: 0; font-family: 'Arial', sans-serif; background: #525659; font-size: 8pt; color: #000; }
        .no-print-bar { position: fixed; top: 0; left: 0; right: 0; background: #0F172A; color: white; padding: 10px 20px; display: flex; justify-content: space-between; align-items: center; z-index: 9999; box-shadow: 0 2px 8px rgba(0,0,0,0.3); }
        .no-print-bar button { background: #0284C7; color: white; border: none; padding: 8px 18px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 13px; }
        .no-print-bar button:hover { background: #0369A1; }
        .pages-container { margin-top: 55px; display: flex; flex-direction: column; align-items: center; gap: 20px; padding-bottom: 40px; }
        .print-sheet { width: 297mm; height: 210mm; background: white; position: relative; display: flex; page-break-after: always; box-shadow: 0 4px 12px rgba(0,0,0,0.15); overflow: hidden; }
        .card-wrapper { width: 148.5mm; height: 210mm; padding: 8mm 9mm; overflow: hidden; }
        .cut-guide { position: absolute; left: 148.5mm; top: 0; bottom: 0; border-left: 1px dashed #CBD5E1; }
        .bincard { width: 100%; height: 100%; border: 1.5px solid #000; padding: 6mm 7mm; display: flex; flex-direction: column; }
        .bc-head { width: 100%; border-collapse: collapse; margin-bottom: 6px; }
        .title-main { font-size: 13pt; font-weight: 900; letter-spacing: 0.5px; }
        .sub-title { font-size: 8pt; color: #475569; margin-top: 1px; font-weight: bold; }
        .side-badge { border: 1.5px solid #000; padding: 2px 8px; font-weight: 800; font-size: 8.5pt; text-transform: uppercase; }
        .meta-box { width: 100%; border-collapse: collapse; margin-bottom: 8px; border: 1.5px solid #000; }
        .meta-box td { border: 1px solid #000; padding: 4px 6px; font-size: 8pt; vertical-align: middle; }
        .meta-lbl { background: #F1F5F9; font-weight: bold; width: 18%; }
        .meta-val { width: 32%; }
        .table-data { width: 100%; border-collapse: collapse; flex: 1; border: 1.5px solid #000; }
        .table-data th { border: 1px solid #000; background: #E2E8F0; padding: 4px; font-size: 7.5pt; font-weight: bold; text-align: center; }
        .table-data td { border: 1px solid #000; padding: 3.5px 4px; font-size: 7.5pt; text-align: center; height: 6.8mm; }
        .text-left { text-align: left !important; }
        .font-bold { font-weight: bold; }
        .card-empty { width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; color: #94A3B8; font-style: italic; border: 1px dashed #CBD5E1; }
        @media print {
          body { background: transparent; }
          .no-print-bar { display: none !important; }
          .pages-container { margin: 0; padding: 0; gap: 0; }
          .print-sheet { box-shadow: none; margin: 0; }
        }
      </style>
    </head>
    <body>
      <div class="no-print-bar">
        <div>
          <b>Gudang Elektrik — Bincard Siap Cetak (A5 di A4)</b>
          <span style="font-size:11px;color:#94A3B8;margin-left:12px;">Total: ` + res.totalItems + ` item (` + sheets.length + ` lembar A4 bolak-balik)</span>
        </div>
        <button onclick="window.print()">🖨️ Cetak Dokumen (A4 Landscape)</button>
      </div>
      <div class="pages-container">
        ` + pagesHtml + `
      </div>
    </body>
    </html>
  `;
}

function renderCardSideA_(item) {
  if (!item) return '<div class="card-empty">(Halaman Kosong)</div>';

  let rows = `
    <tr>
      <td>` + item.tglCetak + `</td>
      <td>SALDO AWAL</td>
      <td class="text-left">Stok Terkini Sistem</td>
      <td>-</td>
      <td>-</td>
      <td class="font-bold">` + fmtNum(item.stok) + `</td>
      <td>SYS</td>
    </tr>
  `;
  for (let i = 0; i < 17; i++) {
    rows += '<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>';
  }

  return `
    <div class="bincard">
      <table class="bc-head">
        <tr>
          <td>
            <div class="title-main">KARTU STOK / BINCARD</div>
            <div class="sub-title">Gudang Elektrik — Inventory System</div>
          </td>
          <td style="text-align: right;">
            <span class="side-badge">SISI A (MUKA)</span>
          </td>
        </tr>
      </table>

      <table class="meta-box">
        <tr>
          <td class="meta-lbl">Lokasi Rak</td>
          <td class="meta-val font-bold" style="color:#0284C7;font-size:9.5pt;">` + escapeHtml_(item.rak || '-') + `</td>
          <td class="meta-lbl">Kode Barang</td>
          <td class="meta-val font-bold" style="font-size:9pt;font-family:monospace;">` + escapeHtml_(item.code) + `</td>
        </tr>
        <tr>
          <td class="meta-lbl">Spesifikasi</td>
          <td class="meta-val font-bold" colspan="3">` + escapeHtml_(item.nama) + `</td>
        </tr>
        <tr>
          <td class="meta-lbl">Remarks</td>
          <td class="meta-val">` + escapeHtml_(item.remarks || '-') + `</td>
          <td class="meta-lbl">Satuan / Min</td>
          <td class="meta-val font-bold">` + escapeHtml_(item.satuan) + ` / Min: ` + fmtNum(item.min) + `</td>
        </tr>
      </table>

      <table class="table-data">
        <thead>
          <tr>
            <th style="width: 14%;">Tanggal</th>
            <th style="width: 19%;">No. Bukti / WO</th>
            <th style="width: 25%;">Keterangan / Bagian</th>
            <th style="width: 11%;">Masuk</th>
            <th style="width: 11%;">Keluar</th>
            <th style="width: 11%;">Sisa</th>
            <th style="width: 9%;">Paraf</th>
          </tr>
        </thead>
        <tbody>
          ` + rows + `
        </tbody>
      </table>
    </div>
  `;
}

function renderCardSideB_(item) {
  if (!item) return '<div class="card-empty">(Halaman Kosong)</div>';

  let rows = '';
  for (let i = 0; i < 21; i++) {
    rows += '<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>';
  }

  return `
    <div class="bincard">
      <table class="bc-head">
        <tr>
          <td>
            <div class="title-main">KARTU STOK (LANJUTAN)</div>
            <div class="sub-title">Lokasi Rak: <b>` + escapeHtml_(item.rak || '-') + `</b> &bull; Kode: <b>` + escapeHtml_(item.code) + `</b></div>
          </td>
          <td style="text-align: right;">
            <span class="side-badge">SISI B (BELAKANG)</span>
          </td>
        </tr>
      </table>

      <table class="table-data" style="margin-top: 6px;">
        <thead>
          <tr>
            <th style="width: 14%;">Tanggal</th>
            <th style="width: 19%;">No. Bukti / WO</th>
            <th style="width: 25%;">Keterangan / Bagian</th>
            <th style="width: 11%;">Masuk</th>
            <th style="width: 11%;">Keluar</th>
            <th style="width: 11%;">Sisa</th>
            <th style="width: 9%;">Paraf</th>
          </tr>
        </thead>
        <tbody>
          ` + rows + `
        </tbody>
      </table>
    </div>
  `;
}

// ===================================================================
// SEED DATA 1.006 ITEM (DARI GUDANG ELEKTRIK.XLSX)
// Format 10 Kolom: [Rak, Item Code, Spesifikasi, Stok Awal, Masuk, Keluar, Stok Akhir, Satuan, Min Stok, Remarks]
// ===================================================================

function getElectricSeedData_() {
  return [["B-2.1 (Box 5)", "QCB NH2 400a", "Wohner Fuse Switch Disconnector 400A", 1, 0, 0, 1, "PCS", 1, "Ex AFR"], ["D-5.1", "Wipol", "Wipol Carbol 390ml", 6, 9, 15, 0, "PCS", 1, ""], ["D-5.1", "Wings Biru", "Wing Biru Sabun Krim Serbaguna (Sabun Colek)", -1, 20, 20, -1, "PCS", 1, ""], ["E-3.1", "RCM270024", "Weidmuller Relay 8pin; 24VDC", 6, 0, 6, 0, "PCS", 1, ""], ["E-3.1", "RCM 570730", "Weidmuller Relay 14pin; 220-230VAC", 0, 0, 0, 0, "PCS", 1, ""], ["B-4.1", "WM-G 1", "Water Mur Galvanis 1\"", 6, 0, 0, 6, "PCS", 1, ""], ["B-1.1", "PL-550-5-R3B", "Ueda Seisakusho Pressure Switch; 0.5-10 kgf/cm²", 6, 0, 0, 6, "KG", 1, ""], ["H-7.6", "P6PX", "Georgin Pressure Switch 0.5…10bar, Adj.≤6bar, SPDT, 5A@240VAC, G1/2B", 0, 0, 0, 0, "PCS", 1, ""], ["H-1.6", "MS24-112-R", "TURCK Rotational Speed Monitors; 20…250VAC/DC", 2, 0, 0, 2, "PCS", 1, ""], ["H-1.6", "Bi15U-CP40-FDZ30X2", "TURCK Proximity Switch; 20…250VAC; 10…300VDC", 1, 0, 0, 1, "PCS", 1, ""], ["H-1.6", "MP-10H-VP4X/S85", "TURCK Multi Prox; 10…65VDC; 200 mA", 1, 0, 0, 1, "PCS", 1, ""], ["H-4.6", "NI4-M12-AP6X", "Turck Inductive Sensor; 10…30VDC; Ø 12mm; NO", -1, 6, 5, 0, "PCS", 1, "3 Wire"], ["H-5.1", "Bi10-P30SR-FZ3X2", "Turck Sensor, Inductive ¦ Sd=10mm, M30, 20…250V, CO, 2w, IP67", 0, 3, 0, 3, "PCS", 1, ""], ["H-4.6", "Ni15U-M18M-VP44X-H1141", "Turck Proximity ¦ Sn=15mm, 10…55VDC, L=61.5mm", 0, 0, 0, 0, "PCS", 1, ""], ["Ikhrom", "HK-T01205", "Trkiro Torx Key Set (Kunci L Bintang)", 3, 0, 0, 3, "SET", 1, "Tools"], ["Rak WS", "59-001", "Traffic Light; Red (Cross) / Green(Arrow); 220V; IP65", 1, 1, 1, 1, "PCS", 1, ""], ["D-5.1", "TOD-02", "Tessa Toilet Tissues (Tisu Roll)", 2, 58, 57, 3, "ROLL", 1, ""], ["D-5.1", "TM-05", "Tessa Facial Tissues (Tisu box)", 29, 65, 95, -1, "BOX", 1, ""], ["Box 1", "LR2D1312", "Telemecanique Thermal Overload 5.5-8A", 1, 0, 0, 1, "CAN", 1, "Barang Lama"], ["Box 1", "LR2D1310", "Telemecanique Thermal Overload 4-6A", 7, 0, 0, 7, "CAN", 1, "Barang Lama"], ["Box 1", "LR2D1303-", "Telemecanique Thermal Overload 0.25-0.4A", 4, 0, 0, 4, "CAN", 1, "Barang Lama"], ["H-6.5", "ZB2-BZ105", "Telemecanique Push Button 1NC+1NO", 8, 0, 0, 8, "CAN", 1, ""], ["H-7.4", "XY2CE2A250", "Telemecanique Pullrope Switch ¦ 3A@240V~, 1NO+1NC, LH Side, Reset But.", 0, 2, 0, 2, "SET", 1, ""], ["H-4.1", "XS8D1A1MAU20", "Telemecanique Osiprox Inductive Proximity Switch; 24…240V; NO", 8, 0, 0, 8, "CAN", 1, ""], ["H-4.1", "XSDA600519", "Telemecanique Osiprox Inductive Proximity Switch; 24…240V; NO", 1, 0, 0, 1, "CAN", 1, ""], ["H-3.5", "XB2 MR42", "Telemecanique Mushroom Pushbutton", 2, 0, 0, 2, "CAN", 1, ""], ["H-2.6", "ZCKJ404+ZCKY13", "Telemecanique Limit Switch; 240V", 1, 2, 0, 3, "CAN", 1, ""], ["H-1.5", "XC1AC127", "Telemecanique Limit Switch Roll", 8, 0, 0, 8, "ROLL", 1, ""], ["H-8.7", "ZCKY11", "Telemecanique Limit Switch Lever (plastic roll)", 26, 0, 0, 26, "ROLL", 2, ""], ["H-8.7", "ZCKY13", "Telemecanique Limit Switch Lever (metal roll)", 9, 0, 0, 9, "ROLL", 1, ""], ["H-5.2", "ZCKE05", "Telemecanique Limit Switch Head (pegas lebih kencang)", 43, 0, 1, 42, "CAN", 2, ""], ["H-5.2", "ZCKE055", "Telemecanique Limit Switch Head", 10, 0, 0, 10, "CAN", 1, ""], ["H-6.1", "XCRT115", "Telemecanique Limit Switch", 2, 0, 1, 1, "CAN", 1, ""], ["H-7.1", "XCRA15", "Telemecanique Limit Switch", 8, 0, 0, 8, "CAN", 1, ""], ["H-8.1", "XCKM115", "Telemecanique Limit Switch", 6, 0, 0, 6, "CAN", 1, ""], ["E-3.2", "RHN 412E", "Telemecanique Inst-Relay; 48VDC", 3, 0, 0, 3, "CAN", 1, "Di belakang"], ["E-3.2", "RHN 412B", "Telemecanique Inst-Relay; 24VDC", 3, 0, 0, 3, "CAN", 1, "Di belakang"], ["E-3.2", "RHN 411B", "Telemecanique Inst-Relay; 24VAC", 3, 0, 0, 3, "CAN", 1, "Di belakang"], ["E-3.2", "RHN 411M", "Telemecanique Inst-Relay; 220VAC", 2, 0, 0, 2, "CAN", 1, "Di belakang"], ["H-3.1", "XSAV11801", "Telemecanique Inductive Proximity: Rotation monitoring; 24…240V; Ø 30mma", 8, 4, 4, 8, "CAN", 1, "2 Wire"], ["H-3.1", "XSAV11373", "Telemecanique Inductive Proximity: Rotation monitoring; 12…48VDC; Ø 30mm", 13, 5, 13, 5, "CAN", 1, "3 Wire"], ["H-6.2", "XS8C4A1PCG13", "Telemecanique Inductive Proximity; 12...48VDC; NC|NO (Kotak seperti Turck)", 1, 0, 0, 1, "CAN", 1, ""], ["H-3.2", "DWH-125SD", "Energy Meter ¦ 5(100)A, 3×380/220V~, 100×123×65mm, Rail Mount", 0, 1, 1, 0, "MTR", 1, ""], ["H-3.3", "XS630B1MAL2", "Telemecanique Inductive Proximity Switch; 24…240V; Ø 30mm; NO", -5, 11, 4, 2, "CAN", 1, "Kabel 2 core"], ["H-3.3", "XS630B1MBL2", "Telemecanique Inductive Proximity Switch; 24…240V; Ø 30mm; NC", 2, 0, 1, 1, "CAN", 1, "Kabel 2 core"], ["H-4.2", "XS4P30PA370L01B", "Telemecanique Inductive Proximity Switch; 12…48V; NO", 3, 0, 2, 1, "CAN", 1, ""], ["H-5.5", "XS618B1MAL2", "Telemecanique Inductive Prox Switch; 24..240VAC/DC; Ø 18mm; NO", -1, 4, 3, 0, "CAN", 1, "Kabel 2 core"], ["H-3.6", "XS4P30KP340", "Telemecanique Inductive Prox Switch; 12…24V; Ø 30mm; NO/NC", 2, 0, 0, 2, "CAN", 1, "Kabel 4 core"], ["Box 2", "CADN22M7", "Telemecanique Contactor 220VAC", 10, 0, 0, 10, "CAN", 1, "Barang Lama"], ["Box 2", "LAD8N11", "Telemecanique Contact Block (Aux Contact 1NC; 1NO)", 10, 10, 0, 20, "CAN", 2, "Barang Lama"], ["Box 2", "LX1FG220", "Telemecanique Coil for Contactor 220/230VAC", 2, 0, 0, 2, "CAN", 1, "Barang Lama"], ["Box 1", "LA7D1064", "Telemecanique Block Thermal", 33, 0, 0, 33, "CAN", 2, "Barang Lama"], ["Box 2", "LX1FJ220", "Telemecanique Coil for Contactor 220/230VAC", 2, 0, 0, 2, "CAN", 1, "Barang Lama"], ["C-4.1", "GT-PH1244", "Tekiro Double Plastic Hammer; Rubber Handle; 50mm", 1, 0, 1, 0, "PCS", 1, "Tools"], ["B-4.1", "TP-G 3/4", "Tee, Pipe ¦ Banded, NPS=¾″ (Ø20mm),Threaded, Galvanized", 10, 0, 0, 10, "PCS", 1, ""], ["B-4.1", "TP-G 1/2", "Tee, Pipe ¦ Banded, NPS=½\" (Ø mm),Threaded, Galvanized", 32, 0, 2, 30, "PCS", 2, ""], ["B-3.1", "TP-PVC 1/2", "Tee, Pipe NPS=½\" , PVC", 4, 0, 0, 4, "PCS", 1, ""], ["B-4.1", "TP-G 1", "Tee Pipe Galvanis 1\"", 9, 0, 0, 9, "PCS", 1, ""], ["Lemari Kaca", "Tape Masking 1", "Tape Masking; Self Adhesive; size 1\" (Isolasi Kertas)", 28, 0, 19, 9, "PCS", 1, ""], ["D-3.1", "TS-06N", "Taeha Solenoid Enclosure Set 6 (220VAC; 10 watt)", 0, 0, 0, 0, "SET", 1, ""], ["D-3.1", "TS-04N", "Taeha Solenoid Enclosur Set 4 ¼″ (220VAC)", 1, 0, 1, 0, "SET", 1, ""], ["F", "JTY-GD-2151EIS", "System Sensor Smoke Detector¦Sense:3±0.7%/ft.,2-Wire,Photoelect", 0, 3, 1, 2, "PCS", 1, "FSS"], ["D-5.1", "Swallow TCB", "Swallow Toilet Colour Ball (Kamper Toilet)", 10, 0, 10, 0, "PCS", 1, ""], ["D-5.1", "Swallow NDB", "Swallow Naphthalene Disk Ball; 1Kg; Kaper Putih", 3, 1, 1, 3, "KG", 1, ""], ["D-5.1", "Sunlight", "Sunlight 635ml", 1, 5, 4, 2, "PCS", 1, ""], ["Lemari Kaca", "Stabilo Y", "Stabilo Bos Yellow", 11, 0, 1, 10, "PCS", 1, ""], ["A-4.1", "Spanjaard 1010", "Spanjaard Electrical Contact Lubricant; 500g", 1, 0, 0, 1, "CAN", 1, ""], ["E-3.1", "PF083A", "Socket Timer 8pin", 2, 0, 2, 0, "PCS", 1, ""], ["E-3.1", "P2CF-08", "Socket Timer 8pin w/Lock", 0, 4, 1, 3, "PCS", 1, ""], ["E-3.1", "SCM 2CO ECO", "Socket Relay 8pin include R1M3", 11, 0, 5, 6, "PCS", 1, ""], ["E-3.1", "PYF14A", "Socket Relay 14pin", -27, 85, 23, 35, "PCS", 2, ""], ["B-4.1", "SB-G 3/4", "Socket Banded Galvanis ¾\" (female)", 2, 0, 8, -6, "PCS", 1, ""], ["B-4.1", "SB-G 1", "Socket Banded 1\", Galvanized (female)", 3, 0, 0, 3, "PCS", 1, ""], ["B-3.1", "SB-PVC 1", "Socket Banded 1\", PVC (female)", 50, 0, 0, 50, "PCS", 2, ""], ["B-3.1", "SR-PVC 2½ x 2", "Socket Banded 2½x2\", PVC (female)", 3, 0, 0, 3, "PCS", 1, ""], ["Lemari Kaca", "Snowman WBR", "Snowman white board marker; Red (Spidol)", 13, 0, 0, 13, "PCS", 2, ""], ["Lemari Kaca", "Snowman WBL", "Snowman white board marker; Blue (Spidol)", 21, 0, 1, 20, "PCS", 2, ""], ["Lemari Kaca", "Snowman WBB", "Snowman white board marker; Black (Spidol)", 68, 0, 14, 54, "PCS", 5, ""], ["Lemari Kaca", "Snowman PR", "Snowman Marke Pen Red (Spidol)", 0, 0, 0, 0, "PCS", 1, ""], ["Lemari Kaca", "Snowman PG", "Snowman Marke Pen Green (Spidol)", 0, 0, 0, 0, "PCS", 1, ""], ["Lemari Kaca", "Snowman PL", "Snowman Marke Pen Blue (Spidol)", 1, 0, 1, 0, "PCS", 1, ""], ["Lemari Kaca", "Snowman PB", "Snowman Marke Pen Black (Spidol)", 1, 0, 1, 0, "PCS", 1, ""], ["", "Omega Rel", "Slotted Rair; 35x7.5mm; 1M", 10, 10, 13, 7, "PCS", 1, ""], ["", "NU 316 ECP/C3", "SKF Cylindrical Roller Bearing 316", 1, 0, 1, 0, "ROLL", 1, ""], ["", "NU 314 ECP/C3", "SKF Cylindrical Roller Bearing 314", 2, 0, 0, 2, "ROLL", 1, ""], ["", "NU 221 ECP/C3", "SKF Cylindrical Roller Bearing 221", 1, 0, 0, 1, "ROLL", 1, ""], ["", "NU 220 ECP", "SKF Cylindrical Roller Bearing 220", 0, 0, 0, 0, "ROLL", 1, ""], ["", "305805-C-2Z", "SKF Ball Bearing Double; Metal Seal", 2, 0, 0, 2, "PCS", 1, "Untuk Hanger Hoist"], ["WS", "6322 M/C3", "SKF Ball Bearing 6322; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["", "6317-2Z/C3", "SKF Ball Bearing 6317; Metal Seal", 3, 1, 2, 2, "PCS", 1, ""], ["", "6316/C3VL0241", "SKF Ball Bearing 6316; Seal (N/A); INSOCOAT", -2, 2, 0, 0, "PCS", 1, ""], ["", "6316 M/C3", "SKF Ball Bearing 6316; Seal (N/A); Cage Tebal", 0, 2, 0, 2, "PCS", 1, ""], ["", "6314-2Z/C3", "SKF Ball Bearing 6314; Metal Seal", 0, 0, 0, 0, "PCS", 1, ""], ["", "6311-2Z/C3", "SKF Ball Bearing 6311; Metal Seal", 6, 0, 3, 3, "PCS", 1, ""], ["", "6310-2RS1/C3", "SKF Ball Bearing 6310; Rubber Seal", 8, 0, 5, 3, "PCS", 1, ""], ["", "6305-2RS1/C3", "SKF Ball Bearing 6305; Rubber Seal", 0, 4, 2, 2, "PCS", 1, ""], ["", "6305-C-2Z C3", "FAG Ball Bearing 6305; Metal Seal", 1, 0, 1, 0, "PCS", 1, ""], ["", "6304-2Z/C3", "SKF Ball Bearing 6304; Metal Seal", -1, 1, 0, 0, "PCS", 1, ""], ["", "6204-2Z/C3", "SKF Ball Bearing 6204; Metal Seal", -4, 8, 4, 0, "PCS", 1, ""], ["", "6302-2Z/C3", "SKF Ball Bearing 6302; Metal Seal", 10, 0, 0, 10, "PCS", 1, ""], ["", "625-2Z", "SKF Ball Bearing 625; Meatl Seal", 8, 0, 0, 8, "PCS", 1, ""], ["", "624-2Z", "SKF Ball Bearing 624; Metal Seal", 22, 0, 0, 22, "PCS", 2, ""], ["", "6226 MA/C3B20", "SKF Ball Bearing 6226; Seal (N/A); Cage Tebal", 1, 0, 0, 1, "PCS", 1, ""], ["", "6226/C3", "SKF Ball Bearing 6226; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["", "6224/C3", "SKF Ball Bearing 6224; Seal (N/A)", 4, 0, 0, 4, "PCS", 1, ""], ["", "6222/C3", "SKF Ball Bearing 6222; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["", "6215/C3", "SKF Ball Bearing 6215; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["", "6214-2Z/C3", "SKF Ball Bearing 6214; Metal Seal", 2, 4, 0, 6, "PCS", 1, ""], ["", "6207-2RS1/C3", "SKF Ball Bearing 6207; Rubber Seal", -9, 19, 8, 2, "PCS", 1, ""], ["", "6203-2RSH/C3", "SKF Ball Bearing 6203; Rubber Seal", 8, 0, 4, 4, "PCS", 1, ""], ["", "6200-2Z/C3", "SKF Ball Bearing 6200; Metal Seal", 6, 0, 0, 6, "PCS", 1, ""], ["", "6005-2Z", "SKF Ball Bearing 6005; Metal Seal", 6, 0, 0, 6, "PCS", 1, ""], ["", "60052RS", "Koyo Ball Bearing 6005; Ruber Seal", 0, 0, 0, 0, "PCS", 1, ""], ["", "608-2Z", "SKF Ball Bearing 608; Metal Seal", 3, 0, 0, 3, "PCS", 1, ""], ["", "6007-2Z-C3", "FAG Ball Bearing 6007; Metal Seal", 0, 8, 1, 7, "PCS", 1, ""], ["", "6008-2Z", "SKF Ball Bearing 6008; Metal Seal", 2, 0, 0, 2, "PCS", 1, ""], ["", "6003-2Z/C3", "SKF Ball Bearing 6003; Metal Seal", 5, 0, 0, 5, "PCS", 1, ""], ["WS", "6330 M/C3", "SKF Ball Bearing 6330; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["", "628-2Z", "SKF Ball Bearing 628; Metal Seal", 50, 0, 0, 50, "PCS", 2, ""], ["", "NU 320 ECP/C3", "SKF Cylindrical Roller Bearing Bearing 320", 1, 0, 0, 1, "ROLL", 1, ""], ["", "NU 320 ECM/C3", "SKF Cylindrical Roller Bearing Bearing 320", 1, 0, 1, 0, "ROLL", 1, ""], ["", "NU 224 ECP/C3", "SKF Cylindrical Roller Bearing Bearing 224", 2, 0, 0, 2, "ROLL", 1, ""], ["", "NJ 307 ECJ", "SKF Cylindrical Roller Bearing 307", 5, 0, 0, 5, "ROLL", 1, ""], ["", "NU 2208 ECP", "SKF Cylindrical Roller Bearing 2208", 2, 0, 0, 2, "ROLL", 1, ""], ["", "NJ 2208 ECP/C3", "SKF Cylindrical Roller Bearing 2208", 3, 0, 0, 3, "ROLL", 1, ""], ["", "NJ 206 ECP", "SKF Cylindrical Roller Bearing 206", 1, 0, 0, 1, "ROLL", 1, ""], ["", "6ES7331-1KF02-0AB0", "SIMATIC S7-300, Analog input SM 331, Isolated 8 AI, resolution 13 bits", 0, 1, 1, 0, "PCS", 1, ""], ["C-4.2", "Sikat Kawat Bukat", "Sikat Kawat Gurinda", 0, 0, 0, 0, "PCS", 1, "Tools"], ["D-4.1", "Sikat Kawat", "Sikat Kawat", -4, 20, 18, -2, "PCS", 1, "Tools"], ["D-5.1", "Sika B-AP", "Sika Boom AP; High Yield Expanding Foam", 20, 0, 11, 9, "PCS", 1, ""], ["E-5.1", "3RU2126-1JB0", "Siemens Thermal Overload on Contactor 7-10A", 5, 0, 0, 5, "PCS", 1, ""], ["E-5.1", "3RU2146-4JB0", "Siemens Thermal Overload on Contactor 45-63A", 1, 0, 0, 1, "PCS", 1, ""], ["E-5.1", "3RU2136-4HB0", "Siemens Thermal Overload on Contactor 40-50A", 2, 0, 0, 2, "PCS", 1, ""], ["E-5.1", "3RU2126-4AB0", "Siemens Thermal Overload on Contactor 11-16A", 20, 0, 5, 15, "PCS", 2, ""], ["A-2.1", "6EP1333-2BA20", "Siemens SITOP PSU100S ¦120/230V~,24VDC/5A,1f,IP20", 0, 2, 2, 0, "PCS", 1, ""], ["A-2.1", "6EP1334-2BA20", "Siemens SITOP PSU100S ¦120/230V~,24VDC/10A,1f,IP20", 1, 0, 0, 1, "PCS", 1, ""], ["B-2.1 (Box 4)", "6EP1336-3BA00", "Siemens SITOP Modular; input 120/230V; output 24VDC 20A", 1, 0, 0, 1, "PCS", 1, "Ex AFR"], ["H-2.4", "6ES7 321-1CH20-0AA0", "SIEMENS SIMATIC S7-300 Digital Input Module; DC 48V/125V", 5, 0, 0, 5, "PCS", 1, ""], ["H-2.4", "6ES7 332-5HF00-0AB0", "SIEMENS SIMATIC S7-300 Analog Output Module", 1, 0, 0, 1, "PCS", 1, ""], ["E-5.1", "3RB2153-4FC2", "Siemens Overload Relay 50...200A; 690V; 1NO+1NC", 0, 1, 0, 1, "PCS", 1, ""], ["D-3.2", "3RV1011-0EA10", "Siemens Motor Circuit Breaker 0.28-0.4A", 1, 0, 0, 1, "PCS", 1, ""], ["D-3.2", "3RV1011-1EA10", "Siemens Motor Circuit Breaker 2.8-4.0A", 3, 0, 3, 0, "PCS", 1, ""], ["D-2.1", "5SJ6306-7SC", "Siemens MCB 3P; 6A", 8, 0, 0, 8, "PCS", 1, ""], ["", "6GK5204-2BB10-2AA3", "Siemens Managed Switch Scalance X204-2, 2×BFOC, 4×RJ45", -1, 1, 0, 0, "PCS", 1, ""], ["D-4.2", "3VA9467-0FK21", "Siemens Extended Rotary Handle/Door Mounted Rotary (400/630)", 5, 0, 1, 4, "PCS", 1, ""], ["D-4.2", "3VA9157-0FK21", "Siemens Extended Rotary Handle/Door Mounted Rotary (100/160)", 4, 0, 0, 4, "PCS", 1, ""], ["D-4.2", "8UC7111-1BB10", "Siemens Extended Actuating Handle", 6, 0, 0, 6, "PCS", 1, ""], ["B-2.1 (Box 4)", "6ES7313-6CG04-0AB0", "Siemens CPU S7-300 (PLC)", 1, 0, 0, 1, "PCS", 1, "Ex AFR"], ["E-3.2", "3RH2131-1BB40", "Siemens Contactor Relay; 24VDC; 3NO+1NC", 2, 0, 0, 2, "PCS", 1, ""], ["E-2.2", "3TR2046-1KB40", "Siemens Contactor 45kW; 24VDC", 13, 0, 0, 13, "PCS", 2, ""], ["E-2.2", "3RT2027-BB40", "Siemens Contactor 32A; 15kW; 24VDC + 1NO; 1NC", 4, 7, 7, 4, "PCS", 1, ""], ["E-2.2", "3RT2037-1AL20", "Siemens Contactor 30kW; 65A; 230VAC", 1, 0, 1, 0, "PCS", 1, ""], ["E-2.2", "3RT2036-1AL20", "Siemens Contactor 22kW; 51A; 230VAC", 3, 0, 3, 0, "PCS", 1, ""], ["E-2.2", "3TF46 22-0XP0", "Siemens Contactor 22kW; 230VAC", 18, 0, 0, 18, "PCS", 2, ""], ["E-1", "3TF56 44-0XP0", "Siemens Contactor 200kW; 400A; 230VAC; 4NC + 4NO", 1, 0, 0, 1, "PCS", 1, ""], ["E-1", "3RT5075-6AP36", "Siemens Contactor 200kW; 400A; 230VAC; 2NC + 2NO", 1, 0, 0, 1, "PCS", 1, ""], ["E-3.2", "3RT2027-1AP00", "Siemens Contactor 1NC; 1NO; 32A; 230VAC", 0, 0, 0, 0, "PCS", 1, ""], ["E-3.2", "3RT2026-1BB40", "Siemens Contactor 1NC; 1NO; 25A; 24VDC", 2, 0, 2, 0, "PCS", 1, ""], ["E-3.2", "3RT2026-1AP00", "Siemens Contactor 1NC; 1NO; 25A; 230VAC", 0, 0, 0, 0, "PCS", 1, ""], ["E-3.2", "3RT6025-1AP00", "Siemens Contactor 1NC; 1NO; 17A; 230VAC", 3, 0, 2, 1, "PCS", 1, ""], ["E-3.2", "3RT2025-1AF00", "Siemens Contactor 1NC; 1NO; 17A; 230VAC", 2, 0, 0, 2, "PCS", 1, ""], ["E-2.2", "3TF44 22-1XB4", "Siemens Contactor 15kW; 24VDC", 0, 0, 0, 0, "PCS", 1, ""], ["D-3.2", "3RV2011-1KA10", "Siemens Circuit Breaker 9-12.5A (Rotary)", -3, 8, 5, 0, "PCS", 1, ""], ["D-3.2", "3RV2021-1JA10", "Siemens Circuit Breaker 7-10A (Rotary)", -5, 10, 5, 0, "PCS", 1, ""], ["D-3.2", "3RV2331-4JC10", "Siemens Circuit Breaker 65A (Rotary)", 0, 0, 0, 0, "PCS", 1, ""], ["D-3.2", "3RV2311-1FC10", "Siemens Circuit Breaker 5A (Rotary)", 1, 0, 1, 0, "PCS", 1, ""], ["D-3.2", "3RV2111-1GA10", "Siemens Circuit Breaker 4.5-6.3A (Rotary)", 5, 0, 5, 0, "PCS", 1, ""], ["D-3.2", "3RV1011-1GA10", "Siemens Circuit Breaker 4.5-6.3A", 2, 0, 2, 0, "PCS", 1, ""], ["D-2.1", "HK-11", "ABB Aux Contact MCB", 7, 0, 0, 7, "PCS", 1, ""], ["D-3.2", "3RV1021-1FA10", "Siemens Circuit Breaker 3.5-5A (Rotary)", 0, 0, 0, 0, "PCS", 1, ""], ["D-3.2", "3RV2031-4EA10", "Siemens Circuit Breaker 22-32A (Rotary)", 12, 0, 2, 10, "PCS", 1, ""], ["D-3.2", "3RV2021-4DA10", "Siemens Circuit Breaker 18-25A (Rotary)", 1, 0, 1, 0, "PCS", 1, ""], ["D-3.2", "3RV2031-4DA10", "Siemens Circuit Breaker 18-25A (Rotary)", 3, 0, 3, 0, "PCS", 1, ""], ["D-3.2", "3RV2311-1JC10", "Siemens Circuit Breaker 10A (Rotary)", 2, 0, 2, 0, "PCS", 1, ""], ["D-3.2", "3RV2021-4AA10", "Siemens Circuit Breaker 10-16A (Rotary)", 1, 0, 1, 0, "PCS", 1, ""], ["D-3.2", "3RV2021-1BA20", "Siemens Circuit Breaker 1.4-2.0A (Rotary)", 0, 0, 0, 0, "PCS", 1, ""], ["D-3.2", "3RV2111-1AA10", "Siemens Circuit Breaker 1.1-1.6A (Rotary)", 5, 0, 0, 5, "PCS", 1, ""], ["D-3.2", "3RV6011-0JA15", "Siemens Circuit Breaker 0.7-1.0A (Rotary)", 1, 0, 1, 0, "PCS", 1, ""], ["D-3.2", "3RV1011-0JA10", "Siemens Circuit Breaker 0.7-1.0A", 4, 0, 0, 4, "PCS", 1, ""], ["D-3.2", "3VU1340-1MF00", "Siemens Circuit Breaker 0.6-1A", 1, 0, 0, 1, "PCS", 1, ""], ["D-3.2", "3RV1121-0FA10", "Siemens Circuit Breaker 0.35-0.5A", 1, 0, 0, 1, "PCS", 1, ""], ["D-3.2", "3RV1011-0AA10", "Siemens Circuit Breaker 0.11-0.16", 3, 0, 0, 3, "PCS", 1, ""], ["B-2.1 (Box 4)", "6ES7323-1BL00-0AA0", "Siemens Chanel DI16/DO16xDC24V", 2, 0, 0, 2, "PCS", 1, "Ex AFR"], ["B-2.1 (Box 4)", "6ES7334-0CE01-0AA0", "Siemens Chanel AI4/AO2x8BIT", 2, 0, 0, 2, "PCS", 1, "Ex AFR"], ["D-4.2", "3VA9988-0AA12", "Siemens Auxiliary Switch for Breaker Siemens", 11, 0, 0, 11, "PCS", 2, ""], ["D-3.2", "PKZM0-0,4", "Moeller Motor Protective Breaker 0.25-0.4A", 1, 0, 0, 1, "PCS", 1, ""], ["D-2.1", "3RV2901-1A", "Siemens Aux Contact Breaker 1NC; 1NO", 8, 0, 5, 3, "PCS", 1, ""], ["E-3.1", "3RH2911-1FA22", "Siemens Aux Contact 2NO; 2NC", 13, 0, 2, 11, "PCS", 2, ""], ["E-3.1", "3RH2911-1HA22", "Siemens Aux Contact 2NO; 2NC", 10, 0, 0, 10, "PCS", 1, ""], ["F-4.1", "SC-L 35-10", "Shoe Cable L; 35mm; M10", 20, 0, 0, 20, "PCS", 2, ""], ["F-4.2", "SC 95-12", "Shoe Cable 95mm; M12", 60, 0, 3, 57, "PCS", 5, ""], ["F-4.1", "SC 70-10", "Shoe Cable 70mm; M10", 148, 0, 148, 0, "PCS", 1, ""], ["F-4.1", "SC 60-12", "Shoe Cable 60-95mm; M12", 120, 0, 6, 114, "PCS", 5, ""], ["F-4.1", "SC 60-10", "Shoe Cable 60-195mm; M10", 44, 0, 0, 44, "PCS", 2, ""], ["F-4.1", "SC 50-8", "Shoe Cable 50mm; M8", 60, 0, 0, 60, "PCS", 5, ""], ["F-4.1", "SC 50-10", "Shoe Cable 50mm; M10", 140, 0, 0, 140, "PCS", 5, ""], ["F-3.1", "Utilux 50-10", "Shoe Cable 50mm; M10", 3, 0, 0, 3, "PCS", 1, ""], ["F-4.2", "SC 400-16", "Shoe Cable 400mm; M16", 20, 0, 0, 20, "PCS", 2, ""], ["F-3.2", "RVS-5.5-4", "Shoe Cable 4-6mm; M4", 220, 0, 0, 220, "PCS", 5, ""], ["F-4.1", "SC 35-8", "Shoe Cable 35mm; M8", 35, 0, 0, 35, "PCS", 2, ""], ["F-4.1", "SC 35-10", "Shoe Cable 35mm; M10", 2, 70, 6, 66, "PCS", 5, ""], ["F-4.1", "SC 25-8", "Shoe Cable 25mm; M8", 104, 0, 0, 104, "PCS", 5, ""], ["F-4.2", "SC 240-16", "Shoe Cable 240mm; M16", 90, 0, 0, 90, "PCS", 5, ""], ["", "RVS-3.5-6", "Shoe Cable 2.5-4mm; M6", 92, 0, 50, 42, "PCS", 2, ""], ["F-4.2", "SC 185-16", "Shoe Cable 185mm; M16", 55, 0, 0, 55, "PCS", 5, ""], ["F-3.1", "SC 16-6", "Shoe Cable 16mm; M6", 10, 0, 6, 4, "PCS", 1, ""], ["F-4.2", "SC 150-16", "Shoe Cable 150mm; M16", 11, 0, 0, 11, "PCS", 2, ""], ["F-4.2", "SC 150-12", "Shoe Cable 150mm; M12", 30, 0, 3, 27, "PCS", 2, ""], ["F-4.2", "SC 120-12", "Shoe Cable 120mm; M12", 30, 0, 0, 30, "PCS", 2, ""], ["F-3.1", "SC 10-8", "Shoe Cable 10mm; M8", 13, 0, 10, 3, "PCS", 1, ""], ["F-3.1", "Utilux 10-6", "Shoe Cable 10mm; M6", 3, 0, 0, 3, "PCS", 1, ""], ["F-3.1", "SC 10-12", "Shoe Cable 10mm; M12", 74, 0, 10, 64, "PCS", 5, ""], ["F-4.2", "SC 100-16", "Shoe Cable 100-150mm; M16", 18, 0, 0, 18, "PCS", 2, ""], ["F-4.2", "SC 100-12", "Shoe Cable 100-150mm; M12", 86, 0, 0, 86, "PCS", 5, ""], ["Loker", "TWI-11-RH", "Shigematsu Respirator; Inhalation=70Pa,w; HeadBand (Masker Set)", 7, 19, 27, -1, "SET", 1, "APD"], ["C-3.1", "X/OV", "Shigematsu Gas Filter for Organic Vapors", 4, 0, 4, 0, "PCS", 1, "APD"], ["C-3.1", "T/OV", "Shigematsu Gas Filter for Organic Vapors", 8, 10, 6, 12, "PCS", 2, "APD"], ["C-3.1", "11061", "Shigematsu Filter Respirator; R2N; White (Filter Masker Putih)", 9, 226, 234, 1, "PCS", 1, "APD"], ["C-3.1", "CA-705/OV", "Shigematsu Chemical Filter for Organic Vapors", 1, 0, 2, -1, "PCS", 1, "APD"], ["D-4.1", "Seal Tape", "Seal Tape Rolls", 32, 20, 6, 46, "ROLL", 2, ""], ["D-2.2", "NSX160N TM160D", "Schneider/Merlin Gerin Breaker 3P; 160A", 0, 0, 0, 0, "PCS", 1, ""], ["H-4.5", "XB7EV05BP", "Schneider Yellow Pilot Lights with Integral LED; 24 VAC/DC", 23, 0, 17, 6, "PCS", 1, ""], ["E-4.1", "LRD16", "Schneider Thermal Overload 9-13A; Cl.10", 4, 0, 4, 0, "PCS", 1, ""], ["E-4.1", "LRD3365", "Schneider Thermal Overload 80-104A; Cl.10", 9, 0, 0, 9, "PCS", 1, ""], ["E-4.1", "LRD14L", "Schneider Thermal Overload 7-10A; Cl.20", 5, 0, 0, 5, "PCS", 1, ""], ["E-4.1", "LRD14", "Schneider Thermal Overload 7-10A; Cl.10", 12, 0, 1, 11, "PCS", 2, ""], ["E-5.1", "LR9F5367", "Schneider Thermal Overload 60-100A; Cl.10", 0, 0, 0, 0, "PCS", 1, ""], ["E-4.1", "LRD12L", "Schneider Thermal Overload 5.5-8A; Cl.20", 1, 0, 1, 0, "PCS", 1, ""], ["E-4.1", "LRD12", "Schneider Thermal Overload 5?5-8; Cl.10", 3, 0, 1, 2, "PCS", 1, ""], ["E-4.1", "LRD365", "Schneider Thermal Overload 48-65A; Cl.10", 3, 0, 0, 3, "PCS", 1, ""], ["E-4.1", "LRD10", "Schneider Thermal Overload 4-6A; Cl.10", 10, 0, 0, 10, "PCS", 1, ""], ["E-4.1", "LRD340", "Schneider Thermal Overload 30-40A; Cl.10; With Everlink", 8, 0, 0, 8, "PCS", 1, ""], ["E-5.1", "LR9F7375", "Schneider Thermal Overload 200-330A; Cl.10", 2, 0, 0, 2, "PCS", 1, ""], ["E-4.1", "LRD08", "Schneider Thermal Overload 2.5-4A; Cl.10", 12, 0, 0, 12, "PCS", 2, ""], ["E-4.1", "LRD22", "Schneider Thermal Overload 16-24A; Cl.10", 1, 0, 0, 1, "PCS", 1, ""], ["E-4.1", "LR97D25M7", "Schneider, Thermal Overload 5…25A, 200…240V~, 1CO, <3s, ƒ LC1D09…D38", 0, 3, 1, 2, "PCS", 1, ""], ["E-5.1", "LR9F5371", "Schneider Thermal Overload 132-220A; Cl.10", 1, 0, 0, 1, "PCS", 1, ""], ["E-4.1", "LRD21", "Schneider Thermal Overload 12-18A; Cl.10", 2, 0, 1, 1, "PCS", 1, ""], ["E-4.2", "LRD07", "Schneider Thermal Overload 1.6-2.5; Cl.10", 12, 0, 0, 12, "PCS", 2, ""], ["E-4.2", "LRD06L", "Schneider Thermal Overload 1-1.6A; Cl.20", 5, 0, 0, 5, "PCS", 1, ""], ["E-4.2", "LRD06", "Schneider Thermal Overload 1-1.6A; Cl.10", 5, 0, 1, 4, "PCS", 1, ""], ["E-4.2", "LRD05L", "Schneider Thermal Overload 0.63-1A; Cl.20", 1, 0, 1, 0, "PCS", 1, ""], ["E-4.2", "LRD05", "Schneider Thermal Overload 0.63-1A; Cl.10", 20, 0, 18, 2, "PCS", 1, ""], ["E-4.2", "LRD04", "Schneider Thermal Overload 0.4-0.63A Cl.10", 7, 2, 3, 6, "PCS", 1, ""], ["E-4.2", "LAD7B106", "Schneider Terminal Block for Thermal", 25, 0, 13, 12, "PCS", 2, ""], ["B-2.1 (Box 4)", "59604 S10MD", "Schneider Sepam", 1, 0, 0, 1, "PCS", 1, "Ex AFR"], ["H-5.6", "XB5AD41", "Schneider Selector; 2 Position; Spring Return", 1, 0, 1, 0, "PCS", 1, ""], ["H-5.7", "XB4BD21", "Schneider Selector Switch; 2 Position; Stay Put (support with ZBE-101/ZBE-102)", 1, 6, 7, 0, "PCS", 1, ""], ["H-5.6", "XB5AD33", "Schneider Selector Switch 3 Position; Stay Put (support with ZBE-101/ZBE-102)", -1, 6, 5, 0, "PCS", 1, ""], ["E-3.1", "RXM2AB1P7", "Schneider Relay 8 pin; 220VAC", 4, 0, 4, 0, "PCS", 1, ""], ["E-3.2", "CAD32ED", "Schneider Relay 3NO; 2NC; 48VDC", -6, 53, 38, 9, "PCS", 1, ""], ["E-3.2", "CAD32E7", "Schneider Relay 3NO; 2NC; 48VAC", 8, 0, 0, 8, "PCS", 1, ""], ["E-3.2", "CAD32Q7", "Schneider Relay 3NO; 2NC; 380VAC", 9, 0, 0, 9, "PCS", 1, ""], ["E-3.2", "CAD32BD", "Schneider Relay 3NO; 2NC; 24VDC", 1, 0, 1, 0, "PCS", 1, ""], ["E-3.2", "CAD32M7", "Schneider Relay 3NO; 2NC; 220VAC", 15, 24, 28, 11, "PCS", 2, ""], ["E-3.2", "CAD32FD", "Schneider Relay 3NO; 2NC; 110VDC", 1, 0, 0, 1, "PCS", 1, ""], ["H-4.5", "XB7EV04BP", "Schneider Red Pilot Lights with Integral LED; 24 VAC/DC", 22, 0, 5, 17, "PCS", 2, ""], ["H-5.7", "ZB4BP31", "Schneider Push Button NO (support w ZBE-102/ZBE-101)", 0, 0, 2, -2, "PCS", 1, ""], ["H-5.7", "ZB4BP4", "Schneider Push Button NC (Support w ZBE-102/ZBE-101", 4, 0, 2, 2, "PCS", 1, ""], ["E-3.1", "LADT4", "Schneider on delay 1NO; 1NC; 10-180sec", 2, 10, 1, 11, "PCS", 2, ""], ["E-3.1", "LADT2", "Schneider on delay 1NO; 1NC; 1-30sec", 12, 10, 10, 12, "PCS", 2, ""], ["E-3.1", "LADR2", "Schneider off delay 1NO; 1NO; 0.3-30sec", 56, 0, 2, 54, "PCS", 5, ""], ["H-5.6", "ZBV-G4", "Schneider Light Red; 110..120VAC", 0, 0, 0, 0, "PCS", 1, ""], ["H-6.6", "ZB4BV053", "Schneider Light Block Pilot Yelow; 220...230VAC; Ø 22mm; Schneider,", 9, 0, 5, 4, "PCS", 1, ""], ["H-6.6", "ZB4BV043", "Schneider Light Block Pilot Red; 220...230VAC; Ø 22mm; Schneider", 6, 0, 6, 0, "PCS", 1, ""], ["H-6.6", "ZB4BV033", "Schneider Light Block Pilot Green; 220...230VAC; Ø 22mm; Schneider", 16, 0, 6, 10, "PCS", 1, ""], ["H-5.6", "XB4BVG4", "Schneider Light Block Pilot Red; 110..120VAC; Ø 22mm; Schneider", 0, 0, 0, 0, "PCS", 1, ""], ["H-5.6", "XB4BVG5", "Schneider Light Block Pilot Orange; 110..230VAC; Ø 22mm; Harmony XB4", 6, 0, 0, 6, "PCS", 1, ""], ["H-5.6", "XB4BVG3", "Schneider Light Block Pilot Green; 110..120VAC; Ø 22mm; Harmony XB4", 6, 0, 0, 6, "PCS", 1, ""], ["D-2.1", "A9K14325", "Schneider iK60a MCB 3P; 25A", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.1", "A9F74202", "Schneider iC60N MCB 2P; 2A", -6, 38, 3, 29, "PCS", 2, ""], ["D-2.1", "A9F84206", "Schneider iC60N MCB 2P; 6A", 0, 4, 4, 0, "PCS", 1, ""], ["D-2.1", "A9F84204", "Schneider iC60N MCB 2P; 4A", 7, 0, 2, 5, "PCS", 1, ""], ["D-2.1", "A9F84225", "Schneider iC60N MCB 2P; 25A", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.1", "A9F84216", "Schneider iC60N MCB 2P; 16A", 7, 10, 1, 16, "PCS", 2, ""], ["D-2.1", "A9F84210", "Schneider iC60N MCB 2P; 10A", -2, 10, 8, 0, "PCS", 1, ""], ["D-2.1", "A9F74106", "Schneider iC60N MCB 1P; 6A", 27, 19, 8, 38, "PCS", 2, ""], ["D-2.1", "A9F84104", "Schneider iC60N MCB 1P; 4A", 8, 20, 8, 20, "PCS", 2, ""], ["D-2.1", "A9F74132", "Schneider iC60N MCB 1P; 32A", 2, 0, 0, 2, "PCS", 1, ""], ["D-2.1", "A9F74102", "Schneider iC60N MCB 1P; 2A", 25, 17, 22, 20, "PCS", 2, ""], ["D-2.1", "A9F74120", "Schneider iC60N MCB 1P; 20A", 3, 4, 0, 7, "PCS", 1, ""], ["D-2.1", "A9F84101", "Schneider iC60N MCB 1P; 1A", 10, 15, 10, 15, "PCS", 2, ""], ["D-2.1", "A9F74116", "Schneider iC60N MCB 1P; 16A", 9, 10, 9, 10, "PCS", 1, ""], ["D-2.1", "A9F74110", "Schneider iC60N MCB 1P; 10A", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.1", "20147", "Merlin Gerin MCB 3P+N; 25A", 1, 0, 1, 0, "PCS", 1, ""], ["D-2.1", "A9F74440", "Schneider iC60 MCB 4P; 40A", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.1", "A9F74306", "Schneider iC60 MCB 3P; 6A", 8, 1, 3, 6, "PCS", 1, ""], ["D-2.1", "A9F74316", "Schneider iC60 MCB 3P; 16A", 0, 0, 0, 0, "PCS", 1, ""], ["H-4.5", "XB7EV03BP", "Schneider Green Pilot Lights with Integral LED; 24VAC/DC", 22, 0, 22, 0, "PCS", 1, ""], ["D-4.2", "LV432598", "Schneider Extended Rotary Handle (Adapter Breaker) for NS 400/630", 4, 0, 0, 4, "PCS", 1, ""], ["D-4.2", "LV429338", "Schneider Extended Rotary Handle (Adapter Breaker) for NS 100/250", 1, 1, 0, 2, "PCS", 1, ""], ["D-4.2", "GV4APN01", "Schneider Extended Rotary Handle (Adapter Breaker) for Breaker GV Series", 7, 0, 0, 7, "PCS", 1, ""], ["Workshop", "NSYCRN325200P", "Schneider Enclosure Panel; W250xD200xH300; IP66", -5, 11, 6, 0, "BOX", 1, "Juntion Box"], ["H-4.7", "ZB5AS844", "Schneider Emergency Switch Yellow", 7, 0, 1, 6, "PCS", 1, ""], ["H-4.7", "ZB4BZ102+ZB4BS844", "Schneider Emergency Switch Silver (1 set dengan contact NC ZBE-102)", 3, 7, 1, 9, "SET", 1, ""], ["H-4.7", "ZB4BS844", "Schneider Emergency Switch Silver", 6, 3, 10, -1, "PCS", 1, ""], ["H-3.7", "XALK178", "Schneider Emergency Pushbutton Box", -7, 32, 19, 6, "BOX", 1, ""], ["E-5.1", "LR9G225", "Schneider Electronic Overload 57-225A; Push-in Terminal", 1, 0, 0, 1, "PCS", 1, ""], ["D-3.2", "DOMR01463", "Schneider Earth Leak Breaker; 4P; 63A; 400VAC; 300mA (RCCB)", 37, 2, 19, 20, "PCS", 2, ""], ["D-3.2", "A9R71225", "Schneider Earth Leak Breaker; 2P; 25A; 230VAC; 30mA (RCCB)", 4, 15, 11, 8, "PCS", 1, ""], ["D-4.2", "LV2429337", "Schneider Direct Rotary Handle (Adapter Breaker) for NSX100-250", 4, 0, 0, 4, "PCS", 1, ""], ["E-3.1", "LAD9R1V", "Schneider Contactor Reversing D09-38", 19, 0, 14, 5, "PCS", 1, ""], ["E-2.2", "LC2D80004M7", "Schneider Contactor Double 80A; 220VAC", 4, 0, 0, 4, "PCS", 1, ""], ["Belakang 32M7", "LC1D09M7", "Schneider Contactor 9A; 220VAC", 1, 0, 0, 1, "PCS", 1, ""], ["E-2.2", "LC1D95M7", "Schneider Contactor 95A; 220VAC", 0, 2, 1, 1, "PCS", 1, "Di atasnya LC1D80M7"], ["E-2.2", "LC1D8000D7", "Schneider Contactor 80A; 42VAC 2NC; 2NO; For Hoist", 9, 0, 1, 8, "PCS", 1, ""], ["E-2.1", "LC1D80P7", "Schneider Contactor 80A; 230VAC", 1, 0, 1, 0, "PCS", 1, ""], ["E-2.2", "LC1D80M7", "Schneider Contactor 80A; 220VAC", 1, 0, 1, 0, "PCS", 1, ""], ["E-2.1", "LC1D65AE7", "Schneider Contactor 65A; 48VAC", 2, 0, 0, 2, "PCS", 1, ""], ["E-2.1", "LC1D65AD7", "Schneider Contactor 65A; 42VAC", 5, 0, 0, 5, "PCS", 1, ""], ["E-2.1", "LC1D65008M7", "Schneider Contactor 65A; 220VAC; 2NC; 2NO; For Hoist", 0, 0, 0, 0, "PCS", 1, ""], ["E-2.1", "LC1D50AE7", "Schneider Contactor 50A; 48VAC", 5, 0, 1, 4, "PCS", 1, ""], ["E-2.1", "LC1D50AD7", "Schneider Contactor 50A; 42VAC", 4, 0, 0, 4, "PCS", 1, ""], ["E-2.1", "LC1D50M7", "Schneider Contactor 50A; 220VAC", 1, 0, 0, 1, "PCS", 1, ""], ["E1", "LC1F630M7", "Schneider Contactor 630A; 220VAC", 1, 0, 0, 1, "PCS", 1, "2nd"], ["E-1", "LC1F500M7", "Schneider Contactor 500A; 220VAC", 5, 0, 0, 5, "PCS", 1, ""], ["E-2.1", "LC1D40AD7", "Schneider Contactor 40A; 42VAC", 12, 0, 4, 8, "PCS", 1, ""], ["E-2.1", "LC1D40V7", "Schneider Contactor 40A; 400VAC", 3, 0, 0, 3, "PCS", 1, ""], ["E-2.1", "LC1D40P7C", "Schneider Contactor 40A; 230VAC", 1, 0, 1, 0, "PCS", 1, ""], ["E-2.1", "LC1D40AM7", "Schneider Contactor 40A; 220VAC", 2, 7, 4, 5, "PCS", 1, ""], ["E-1", "LC1F330M7", "Schneider Contactor 330; 220VAC", 5, 0, 0, 5, "PCS", 1, ""], ["E-2.1", "LC1D32E7", "Schneider Contactor 32A; 48VAC", 1, 6, 3, 4, "PCS", 1, ""], ["E-2.1", "LC1D18E7", "Schneider Contactor 18A; 48VAC", -1, 3, 2, 0, "PCS", 1, ""], ["E-3.2", "LC1D32Q7", "Schneider Contactor 32A; 380VAC", 9, 0, 0, 9, "PCS", 1, ""], ["E-2.1", "LC1D32M7", "Schneider Contactor 32A; 220VAC", 29, 40, 60, 9, "PCS", 1, ""], ["E-2.1", "LC1D25M7", "Schneider Contactor 25A; 220VAC", 1, 0, 0, 1, "PCS", 1, ""], ["E-2.1", "LC1D258E7", "Schneider Contactor 25A; 48VAC; 2NC; 2NO; For Hoist", 2, 6, 0, 8, "PCS", 1, ""], ["E-2.1", "LC1D258D7", "Schneider Contactor 25A; 42VAC; 2NC; 2NO; For Hoist", 4, 0, 4, 0, "PCS", 1, ""], ["E-2.1", "LC1D25D7", "Schneider Contactor 25A; 42VAC", 0, 0, 0, 0, "PCS", 1, ""], ["E-1", "LC1F265M7", "Schneider Contactor 256A; 220VAC", 8, 0, 0, 8, "PCS", 1, ""], ["E1", "LC1F185M7", "Schneider Contactor 185A; 220VAC", 7, 0, 2, 5, "PCS", 1, ""], ["E-1", "LC1D150M7", "Schneider Contactor 150A; 220VAC", 1, 0, 0, 1, "PCS", 1, ""], ["E1", "LC1F150M7", "Schneider Contactor 150A; 220VAC", 2, 0, 0, 2, "PCS", 1, ""], ["H-5.7", "ZB2-BE101", "Schneider Contact Point 1NO for Push Button Hoist", 7, 0, 0, 7, "PCS", 1, ""], ["H-5.7", "ZBE-101", "Schneider Contact Point 1NO (for Push Button and other)", 25, 0, 25, 0, "PCS", 1, ""], ["H-5.7", "ZBE-102", "Schneider Contact Point 1NC (for Push Button and other)", 52, 0, 24, 28, "PCS", 2, ""], ["H-6.5", "HENG1191", "Schneider Contact Block For Hoist Double Speed", 0, 0, 0, 0, "PCS", 1, ""], ["E-3.2", "LX1FF380", "Schneider Coil for Contactor 380VAC", 5, 0, 0, 5, "PCS", 1, "Di Belakang"], ["E-3.2", "LX1FF220", "Schneider Coil for Contactor 220VAC", 3, 0, 0, 3, "PCS", 1, "Di Belakang"], ["E-3.2", "LX1FH2202", "Schneider Coil for Contactor 220/230VAC", 4, 0, 0, 4, "PCS", 1, "Di Belakang"], ["E-3.2", "LX1FL220", "Schneider Coil 220/240VAC", 3, 0, 0, 3, "PCS", 1, "Di Belakang"], ["D-3.2", "GV2P32", "Schneider Breaker,Circuit¦3P,24…32A,35kA,400/415V~,ThermMagnetic(416A)", 0, 0, 0, 0, "PCS", 1, ""], ["D-3.2", "GV2P07", "Schneider Circuit Breaker 1.6-2.5A (Rotary)", 1, 0, 1, 0, "PCS", 1, ""], ["", "LC1DWK12M7", "Schneider Capacitor Switching 92A; 220-230VAC", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.1", "17871", "Schneider C65N C6 MCB 1P; 6A", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.1", "26924", "Schneider Aux Contact for MCB (Old)", 11, 0, 0, 11, "PCS", 2, ""], ["E-3.1", "LADN40", "Schneider Aux Contact 4NO", 16, 0, 6, 10, "PCS", 1, ""], ["E-3.1", "LADN31", "Schneider Aux Contact 3NO; 1NC", 10, 10, 17, 3, "PCS", 1, ""], ["E-3.1", "LADN22", "Schneider Aux Contact 2NO; 2NC", 8, 31, 24, 15, "PCS", 2, ""], ["D-2.1", "A9A26924", "Schneider Acti9 iOF Aux Contact for iC60", 26, 0, 4, 22, "PCS", 2, ""], ["D-2.1", "20146", "Merlin Gerin MCB 3P+N; 20A", 4, 0, 1, 3, "PCS", 1, ""], ["D-2.1", "A9N18524", "Schneider Acti 9 C120H MCB 4P; 100A", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.1", "17870", "Schneider C65N C4 MCB 1P; 4A", 2, 0, 2, 0, "PCS", 1, ""], ["D-2.1", "MGN22055", "Schneider C65H-DC C10 MCB 1P; 10A", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.1", "EZC100H2100", "Schneider Breaker Thermal Magnetic 2P; 100A", 1, 0, 0, 1, "PCS", 1, ""], ["D-2.2", "GV2ME14", "Schneider Circuit Breaker; 6-10A; 3P 690VAC", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.2", "3VA211010-7HL36-0AA0", "Siemens Circuit Breaker 3P; 100A", 1, 0, 0, 1, "PCS", 1, ""], ["D-2.2", "C63N32D630", "Schneider Breaker Micrologic 2.3 NSX630N ; 630A", 0, 0, 0, 0, "PCS", 1, "3 Shaft Modul"], ["D-2.2", "C40N32D400", "Schneider Breaker Micrologic 2.3 NSX400N ; 400A", 1, 0, 0, 1, "PCS", 1, ""], ["D-2.2", "NSX100N 100A", "Schneider Breaker 3P; 100A (Merlin Gerin NS100H)1", 1, 0, 1, 0, "PCS", 1, ""], ["D-2.2", "C10N3TM100", "Schneider Breaker 3P; 100A (Merlin Gerin NS100H)", 1, 0, 0, 1, "PCS", 1, "Adj.40-100A"], ["D-2.2", "C25N3TM250", "Schneider Breaker 3P; 250A (NSX250N)", 1, 3, 4, 0, "PCS", 1, ""], ["D-2.2", "NSX100N", "Schneider Breaker 3P; (Without Micro Logic)", 4, 0, 0, 4, "PCS", 1, "Di Belakang"], ["D-2.2", "GV4LE07", "Schneider Breaker 3P (EverLink) 7A", 3, 0, 0, 3, "PCS", 1, ""], ["D-2.2", "GV4LE50N", "Schneider Breaker 3P (EverLink) 50A", 0, 4, 3, 1, "PCS", 1, ""], ["D-2.2", "GV4LE25N", "Schneider Breaker 3P (EverLink) 25A", 0, 2, 2, 0, "PCS", 1, ""], ["D-2.2", "GV4LE12N", "Schneider Breaker 3P (EverLink) 12.5A", 4, 0, 0, 4, "PCS", 1, ""], ["D-4.2", "29450", "Schneider Auxiliary Switch for Breaker NS/X 100-630", 69, 0, 11, 58, "PCS", 5, ""], ["D-4.2", "GV4AE11", "Schneider Auxiliary Switch for Breaker GV Series", 3, 4, 0, 7, "PCS", 1, ""], ["E-3.2", "RM22TG20", "Schneider 3 Phase Control Relay", 12, 1, 8, 5, "PCS", 1, "Dibelakang Siemens"], ["Locker", "Sarung Tangan", "Sarung Tangan Bahan", 75, 245, 355, -35, "PCS", 1, "APD"], ["Ikhrom", "CD771", "Sanwa Digital Multimeter", 3, 0, 2, 1, "MTR", 1, "Tools"], ["Ikhrom", "Sands 15", "Sands Adjustable Wrench 15\"", 2, 0, 0, 2, "PCS", 1, "Tools"], ["", "Amplas G60", "Sandpaper, Waterproof ¦ Grit 60, 230×280mm, Silicon Carbide, One Face, Abrasive", 0, 15, 9, 6, "PCS", 1, ""], ["", "Amplas G120", "Sandpaper, Waterproof ¦ Grit 120, 230×280mm, Silicon Carbide, One Face, Abrasive", 0, 15, 3, 12, "PCS", 2, ""], ["", "SVC-500N", "Samoto Auto Regulator Voltage ; 140..240VAC/110;220VAC ±3% 500VA", 0, 5, 0, 5, "PCS", 1, ""], ["D-5.1", "BKC 20", "Sakdi Desinfektan-Antiseptik; 1L", 1, 0, 0, 1, "PCS", 1, ""], ["Lemari Kaca", "Safety Vest O", "Safety Vest Orange", 16, 0, 3, 13, "PCS", 2, "APD"], ["Atas Loker", "Virunga Cover", "Safety Jogger Safety Eyewear Virunga Cover Clear (Khusus yg berkaca mata)", -4, 15, 3, 8, "PCS", 1, "APD"], ["B-3.1", "Saddle Conduit G54", "Saddle Conduit Clamp; G54 (Clamp Oega for Pipe 2\")", -28, 50, 0, 22, "PCS", 2, ""], ["B-3.1", "Saddle Conduit G28", "Saddle Conduit Clamp; G28 (Clamp Oega for Pipe 1\")", 140, 50, 10, 180, "PCS", 5, ""], ["B-3.1", "Saddle Conduit G22", "Saddle Conduit Clamp; G22 (Clamp Oega for Pipe ¾\")", 200, 0, 66, 134, "PCS", 5, ""], ["B-3.1", "Saddle Conduit G16", "Saddle Conduit Clamp; G16 (Clamp Oega for Pipe ½\")", 100, 0, 0, 100, "PCS", 5, ""], ["B-3.1", "Clamp C PH-1", "Clamp, C ¦ PH-1, IØ17×OØ42×W34×L46mm, Bolt: M6×25mm, Jaw=33mm, ƒ Beam: 1…16mm", 110, 0, 22, 88, "PCS", 5, ""], ["B-3.2", "CC G-36", "Coupling, Threaded Conduit ¦ Size=G36(1¼\"), ?2?, Ø48.5×L56mm, Steel, Zinc Coated", 0, 100, 0, 100, "PCS", 5, ""], ["B-3.1", "CC G-42", "Coupling, Threaded Conduit ¦ Size=G42(1½″), ♀2♀, Ø54.5×L56mm, Steel, Zinc", 23, 0, 0, 23, "PCS", 2, ""], ["B-3.1", "CC G-54", "Coupling, Threaded Conduit ¦ Size=G54(2″), ♀2♀, Ø68×L64mm, Steel, Zinc", 10, 0, 0, 10, "PCS", 1, ""], ["B-3.1", "BC G-28", "Bushing, Conduit ¦ G-28, Flat Steel, Zinc Coated, Threaded", 14, 0, 4, 10, "PCS", 1, ""], ["B-3.1", "BC G-16", "Bushing, Conduit ¦ G-16, Flat Steel, Zinc Coated, Threaded", 50, 0, 0, 50, "PCS", 2, ""], ["H-1.3", "Z415 HL2", "RS Pro Metal Oxide Varistor; 240 - 242", 6, 0, 0, 6, "PCS", 1, "Barang Jarang Pakai"], ["Loker", "Roofing 70", "Roofing 70mm; Self Drilling; with rubber washer", 0, 500, 70, 430, "PCS", 5, ""], ["Loker", "Roofing 20", "Roofing 20mm; Self Drilling; with rubber washer", 0, 500, 542, -42, "PCS", 1, ""], ["D-5.1", "Rinso", "Rinso Detergent", 5, 5, 10, 0, "PCS", 1, ""], ["B-4.1", "RS-G 3/4 x 1/2", "Reducing Socket Galvanized ¾\" to ½\" (female)", 13, 0, 5, 8, "PCS", 1, ""], ["B-4.1", "RS-G 2x1", "Reducing Socket ; Threaded, Galvanized (female)", 5, 0, 0, 5, "PCS", 1, ""], ["C-4.2", "welding gloves", "REDRAM sarung tangan las/panas", 9, 0, 3, 6, "PCS", 1, "APD"], ["D-4.1", "Pressol 12633", "Pressol Grease Gun", 3, 0, 2, 1, "PCS", 1, "Tools"], ["B-4.1", "Plug-S 304-1", "Plug Stainless 1\",Threaded (Male)", 13, 0, 0, 13, "PCS", 2, ""], ["B-4.1", "Plug-G 1", "Plug Galvanized 1\", Threaded (Male)", 24, 0, 0, 24, "PCS", 2, ""], ["B-4.1", "Plug-G 1/2", "Plug Galvanized ½\" (Male)", 0, 0, 3, -3, "PCS", 1, ""], ["B-4.1", "Plug-G 3/4", "Plug Galvanis ¾\" Threaded (Male), Galvanized", 5, 0, 0, 5, "PCS", 1, ""], ["B-2.2", "Waste Plastic (R)", "Plastik Limbah B3 (Merah)", 7, 5, 7, 5, "PCS", 1, ""], ["B-2.1 (Box 4)", "PNOZ X4 24VDC", "Pillz Safety Relay; 24VDC; 3NO+1NC", 1, 0, 0, 1, "PCS", 1, "Ex AFR"], ["", "MACX MCR-SL-CAC-5-I-UP", "Phoenixcontact Current measuring transducer for 1 A and 5 A AC, the output signal 0...20 mA or 4...20 mA", 0, 4, 4, 0, "PCS", 1, ""], ["Laci", "MCR-SL-PT100-LP-NC", "Phoenix Contact Transmitter Temp; 12...30VDC; 4...20mA", 4, 5, 9, 0, "PCS", 1, ""], ["D-5.2", "UK 10 N", "Phoenix Contact Terminal block; 800 V; 57 A (Terminal UK)", 126, 0, 101, 25, "PCS", 2, "Satuan dalam pcs"], ["D-5.2", "UK 5 N", "Phoenix Contact Terminal block; 800 V; 32 A (Terminal UK)", 3475, 0, 745, 2730, "PCS", 5, "Satuan dalam pcs"], ["D-5.2", "USLKG 10 N", "Phoenix Contact Protective conductor/grounding terminal block", 68, 100, 18, 150, "PCS", 5, "Satuan dalam pcs"], ["D-5.2", "E-UK", "Phoenix Contact End Bracket mounting a DIN rail (Stopper terminal)", 91, 0, 20, 71, "PCS", 5, "Satuan dalam pcs"], [".", "Philips TCW069", "Philips TL Set ; 2xTL Led (Rumah TL)", 0, 17, 9, 8, "SET", 1, ""], ["D1", "Philips TL-D 36W", "Philips Neon TL; 36 watt; 220 V", 12, 0, 12, 0, "PCS", 1, ""], ["D1", "Philips TL-D 18W", "Philips Neon TL; 18 watt; 220 V", 90, 0, 65, 25, "PCS", 2, ""], ["D1", "Philips E40", "Philips LED TrueForce Core; 80 watt; Bulb", 26, 58, 76, 8, "PCS", 1, ""], ["D1", "Philips E27", "Philips LED TrueForce Core; 30 watt; Bulb", 47, 115, 140, 22, "PCS", 2, ""], ["D1", "Philips 929001184608", "Philips LED TL; 16 watt; 220-240V; 1200mm; Ø27.8mm; T8", -3, 110, 107, 0, "PCS", 1, ""], ["D1", "Philips 929003550867", "Philips, LED TL tube 1200mm 18W765 T8 APIG, 1800Lm, 6500K", 21, 20, 10, 31, "PCS", 2, ""], ["D1", "Philips BVP151", "Philips LED Floodinglight; 30 Watt; Spotlight", 16, 56, 72, 0, "PCS", 1, "(N/A)"], ["D1", "Philips BVP150", "Philips LED Floodinglight; 150 Watt; Spotlight", -3, 49, 46, 0, "PCS", 1, "(N/A)"], ["D1", "Philips BVP150 G3", "Philips LED Floodinglight; 30 Watt; Spotlight", 6, 0, 6, 0, "PCS", 1, ""], ["D1", "Philips BVP150 G4", "Philips LED Floodinglight; 150 Watt; Spotlight", 8, 0, 3, 5, "PCS", 1, ""], ["H-5.3", "NBB2-12GM50-E2-V1", "Pepperl+fuchs Inductive Sensor; Ø 12mm", 1, 0, 1, 0, "PCS", 1, ""], ["H-5.4", "NJ5-18GK-N-150", "Pepperl+fuchs Inductive Sensor; 8.2VDC; Ø 18mm; NC", 3, 1, 3, 1, "PCS", 1, "Kabel 2 core"], ["H-5.3", "NBB10-30GM50-WS", "Pepperl+fuchs Inductive Sensor; 20…253VAC; Ø 30mm; NO", 3, 0, 0, 3, "PCS", 1, "2 Wire + Gnd"], ["H-5.4", "NJ2-12GM40-E2", "Pepperl+fuchs Inductive Sensor; 10…60VDC; Ø 12mm; NO", 1, 0, 0, 1, "PCS", 1, "Kabel 3 core"], ["H-6.3", "NCB5-18GM40-40Z0", "Pepperl+fuchs Inductive Sensor; 10-30V; Ø 18mm; NO", 1, 0, 1, 0, "PCS", 1, "2 Wire"], ["H-5.4", "NJ1,5-8-N-Y18812", "Pepperl+Fuchs Inductive Sensor | NC", 1, 0, 1, 0, "PCS", 1, ""], ["C-4.1", "PCE 034", "PCE Plug Socket; 63A; 380-415V; 3P+E; IP67", 2, 0, 0, 2, "PCS", 1, "Welding Plug"], ["C-4.1", "PCE 0242", "PCE Plug Socket; 32A; 380-415V; 3P+E; IP67", 9, 0, 0, 9, "PCS", 1, "Welding Plug"], ["C-4.1", "PCE 434", "PCE Flanged Socket Sloping; 63A; 380-415V; 3P+E; IP67 (Fix)", 3, 0, 0, 3, "PCS", 1, "Welding Sockect"], ["D-3.1", "345B34", "Parker Lucifer; Pneumatic Distributors; 220-230V 50-60Hz", 3, 0, 0, 3, "PCS", 1, ""], ["C-2.2", "4114N/12MR", "Papst Fan; 24VDC; 145mA 35W (12-30VDC); 105x105mm", 6, 0, 0, 6, "PCS", 1, "Arah Rak B"], ["WS", "Panel Phone", "Panel Phone W400xD200xH600xT2mm", 2, -2, 0, 0, "PCS", 1, ""], ["WS", "Panel JB (S)", "Panel JB; W200xD100xH200xT2mm", 6, 2, 8, 0, "BOX", 1, "Juntion Box"], ["A-4.1", "Osbond 293", "Osbond Isolator Cleaner Compound; 500g", 3, 0, 0, 3, "PCS", 1, ""], ["A-4.1", "Osbond 44", "Osbond Epoxy Metal Repair; Resin + Hardener (Plastic Steel)", 7, 0, 1, 6, "PAIR", 1, ""], ["E-3.1", "H3CR-F8", "Omron Timer 1.2s to 300h; 100-240V (ON-OFF Control)", 1, 6, 5, 2, "PCS", 1, "(2 Jarum timer)"], ["E-3.1", "H3CR-A", "Omron Timer 1.2s to 300h; 100 -240V", 1, 0, 1, 0, "PCS", 1, ""], ["E-3.1", "RXM2AB2ED", "Schneider Relay 8pin; 48VDC / Test Button+LED", -2, 20, 7, 11, "PCS", 2, ""], ["E-3.1", "RXM4AB2BD", "Schneider, Relay 12pin; 24VDC, 4CO, 6A, w/ Test Button+LED", 30, 0, 0, 30, "PCS", 2, ""], ["E-3.1", "MY4N-GS 48D", "Omron Relay 14pin; 48VDC", 6, 15, 3, 18, "PCS", 2, ""], ["E-3.1", "MY4N-GS-R 24D", "Omron Relay 14pin; 24VDC", 6, 50, 45, 11, "PCS", 2, ""], ["E-3.1", "LY4N 24D", "Omron Relay 14pin; 24VDC", 4, 8, 0, 12, "PCS", 2, "Dibelakang"], ["E-3.1", "MY4N-GS-R 220/240A", "Omron Relay 14pin; 220/230VAC", 4, 15, 19, 0, "PCS", 1, ""], ["E-3.1", "MY4N-GS 100/110D", "Omron Relay 14pin; 100/110VDC", 40, 0, 0, 40, "PCS", 2, ""], ["E-3.1", "MY4N-GS 100/110A", "Omron Relay 14pin; 100/110VAC", 0, 0, 0, 0, "PCS", 1, ""], ["E-3.1", "MY4N 100/110A", "Omron Relay 14pin; 100/110VAC", 0, 0, 0, 0, "PCS", 1, ""], ["H-4.3", "E2E-X5Y1", "Omron Proximity Sensor; 24…240VAC; Ø 18mm; NO", 4, 6, 9, 1, "PCS", 1, "Kabel 2 core"], ["H-2.1", "WLNJ", "Omron Limit Switch; 48VDC; 250VAC (Buntut Tikus)", -5, 10, 5, 0, "PCS", 1, ""], ["H-2.1", "WLCA12-N", "Omron Limit Switch; 2A; 250VAC/DC; 1NC+1NO; Adjustable Roller Lever", 11, 0, 8, 3, "ROLL", 1, ""], ["H-2.3", "61F-GP-N", "Omron Floatless Level Switch; 220VAC", 1, 0, 0, 1, "PCS", 1, ""], ["H-2.2", "61-F-G-AP", "Omron Floatless Level Switch; 110/220VAC", 6, 0, 0, 6, "PCS", 1, ""], ["H-2.3", "61F-G2", "Omron Floatless Level Switch; 110/220VAC", 2, 0, 0, 2, "PCS", 1, ""], ["H-1.4", "PS-4S", "Omron Electrode Holder 4-Pin", 3, 0, 0, 3, "PCS", 1, ""], ["H-1.4", "PS-3S", "Omron Electrode Holder 3-Pin", 3, 0, 0, 3, "PCS", 1, ""], ["", "RSB6511529", "Oil Seal RS 65-115-29", 0, 0, 0, 0, "PCS", 1, ""], ["", "RSB458521", "Oil Seal RS 45-85-21", 0, 0, 0, 0, "PCS", 1, ""], ["", "TC 68-90-10", "Oil Seal 68-90-10", 1, 0, 0, 1, "PCS", 1, ""], ["", "TC 65-95-14", "Oil Seal 65-95-14", 2, 0, 0, 2, "PCS", 1, ""], ["", "TC 60-80-10", "Oil Seal 60-80-10", 0, 0, 0, 0, "PCS", 1, ""], ["", "TC 55-70-8", "Oil Seal 55-70-8", 12, 0, 0, 12, "PCS", 2, ""], ["", "TC 50-72-10", "Oil Seal 50-72-10", 7, 0, 0, 7, "PCS", 1, ""], ["", "TC 50-65-8", "Oil Seal 50-65-8", 6, 0, 0, 6, "PCS", 1, ""], ["", "TC 48-72-8", "Oil Seal 48-72-8", 1, 0, 0, 1, "PCS", 1, ""], ["", "TC 45-65-10", "Oil Seal 45-65-10", 5, 0, 0, 5, "PCS", 1, ""], ["", "TC 45-62-9", "Oil Seal 45-62-9", 10, 0, 0, 10, "PCS", 1, ""], ["", "TC 45-62-8", "Oil Seal 45-62-8", 7, 0, 0, 7, "PCS", 1, ""], ["", "TC 42-62-10", "Oil Seal 42-62-10", 7, 0, 0, 7, "PCS", 1, ""], ["", "TC 40-68-7", "Oil Seal 40-68-7", 2, 0, 0, 2, "PCS", 1, ""], ["", "TC 40-62-8", "Oil Seal 40-62-8", 10, 0, 2, 8, "PCS", 1, ""], ["", "TC 40-55-7", "Oil Seal 40-55-7", 7, 0, 0, 7, "PCS", 1, ""], ["", "TC 40-52-7", "Oil Seal 40-52-7", 5, 0, 0, 5, "PCS", 1, ""], ["", "TC 36-62-7", "Oil Seal 36-62-7", 2, 0, 0, 2, "PCS", 1, ""], ["", "TC 35-68-10", "Oil Seal 35-68-10", 2, 0, 0, 2, "PCS", 1, ""], ["", "TC 35-55-11", "Oil Seal 35-55-11", 5, 0, 0, 5, "PCS", 1, ""], ["", "TC 35-52-10", "Oil Seal 35-52-10", 6, 0, 0, 6, "PCS", 1, ""], ["", "TC 30-47-8", "Oil Seal 30-47-8", 2, 0, 2, 0, "PCS", 1, ""], ["", "TC 30-47-7", "Oil Seal 30-47-7", 2, 0, 1, 1, "PCS", 1, ""], ["", "TC 30-42-8", "Oil Seal 30-43-8", 7, 0, 0, 7, "PCS", 1, ""], ["", "TC 30-40-8", "Oil Seal 30-40-8", 14, 0, 0, 14, "PCS", 2, ""], ["", "TC 25-38-7", "Oil Seal 25-35-7", 9, 0, 0, 9, "PCS", 1, ""], ["", "TC 24-40-8", "Oil Seal 24-40-8", 3, 0, 0, 3, "PCS", 1, ""], ["", "TC 22-40-8", "Oil Seal 22-40-8", 2, 0, 0, 2, "PCS", 1, ""], ["", "TC 20-38-8", "Oil Seal 20-38-8", 6, 0, 0, 6, "PCS", 1, ""], ["", "TC 20-30-7", "Oil Seal 20-30-7", 1, 0, 0, 1, "PCS", 1, ""], ["", "TC 17-30-7", "Oil Seal 17-30-7", 2, 0, 0, 2, "PCS", 1, ""], ["", "TC 14-30-7", "Oil Seal 14-30-7", 4, 0, 0, 4, "PCS", 1, ""], ["H-7.5", "228-774A/219-218", "Numatics Solenoid for Aeration", 19, 0, 0, 19, "PCS", 2, ""], ["WS", "NU322C3", "NTN Cylindrical Bearing NU322", 1, 0, 0, 1, "PCS", 1, ""], ["", "6306ZZC3", "NTN Ball Bearing 6306; Metal Seal", 0, 0, 0, 0, "PCS", 1, ""], ["Laci", "BL-150P", "NT Cutter Spare Blade (Isi cutter)", 5, 2, 3, 4, "PCS", 1, "Tools"], ["Laci", "L-500", "NT Cutter (Cutter Set)", -1, 5, 6, -2, "SET", 1, ""], ["", "NU217WC3", "NSK Cylindrical Roller Bearing 217", 1, 0, 0, 1, "ROLL", 1, ""], ["", "6304DDUCM NS7S", "NSK Ball Bearing 6304; Rubber Seal", 0, 0, 0, 0, "PCS", 1, ""], ["", "6224C3", "NSK Ball Bearing 6224; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["F", "432288", "Notifier® Smoke Detector, FSP-851, 360µA@24VDC, 360°, w/ Base B710LP Ø15.5cm", 0, 3, 3, 0, "PCS", 1, "FSS"], ["F", "428100", "Notifier® Module Control ¦ FCM-1-REL 15…32VDC, 0.39/=5.1mA", 1, 0, 0, 1, "PCS", 1, ""], ["F", "428097", "Notifier, Monitor Modul FMM-1, 15…32VDC, 47kΩ, 0…49°C, 2│4wire, H4½×W4×D1¼″", 0, 3, 2, 1, "PCS", 1, "FSS"], ["Lemari Kaca", "Nitto Tape", "Nitto Plastic Tape; 0.2mm x 19 x 20mm ; 5kV; PVC Base", 91, 77, 169, -1, "PCS", 1, ""], ["C-4.2", "A24SBF 6", "Nippon Resibon Polish; 150x6x22.23 (6\")", 30, 0, 0, 30, "PCS", 2, "Mata Gerinda"], ["C-4.2", "A24SBF 5", "Nippon Resibon Polish; 125x6x22 (5\")", 28, 0, 2, 26, "PCS", 2, "Mata Gerinda"], ["C-4.2", "A24SBF 4", "Nippon Resibon Polish; 100x6x16 (4\")", 9, 10, 3, 16, "PCS", 2, "Mata Gerinda"], ["C-4.2", "A30PBF 14", "Nippon Resibon Cutting Wheel; 355x3x25.4 (14\")", 7, 20, 10, 17, "PCS", 2, "Mata Gerinda"], ["C-4.2", "A30PBF 4", "Nippon Resibon Cutting Wheel; 105x3x16 (4\")", -3, 31, 11, 17, "PCS", 2, "Mata Gerinda"], ["C-4.2", "RD-260 3.2", "Nikko Steel Welding Electrodes; RD-260; Ø3.2x350mm", 7, 0, 1, 6, "PCS", 1, "Kawat Las"], ["C-4.2", "RD-260 2.6", "Nikko Steel Welding Electrodes; RD-260; Ø2,6x350mm", 0, 0, 0, 0, "PCS", 1, "Kawat Las"], ["C-4.2", "NSN-312 3.2", "Nikko Steel Welding Electrodes; NSN-312 Ø3.2x350mm", 0, 0, 0, 0, "PCS", 1, "Kawat Las"], ["C-4.2", "RD-718 3.2", "Nikko Steel Welding Electrodes; RD-718 Ø3.2x350mm", 1, 0, 0, 1, "PCS", 1, ""], ["F", "PCH3016 7G24H", "Nihon Inter Thyristor VRRM=1600V, IOAV=30A@96°C, 26.7×93×30mm", 0, 0, 0, 0, "PCS", 1, ""], ["C-2.2", "D1751S24B6CZ", "Nidec Fan; 24VDC; 1.8A; Ø 170x170mm", 1, 0, 0, 1, "PCS", 1, "Arah Rak B"], ["", "6303-2NSEC3", "Nachi Ball Bearing 6303; Rubber Seal", 2, 0, 0, 2, "PCS", 1, ""], ["", "6205ZZECM", "Nachi Ball Bearing 6205; Rubber Seal", 0, 0, 0, 0, "PCS", 1, ""], ["C-4..1", "Steker", "N&L Plug Schuko Power (Stecker)", 12, 0, 12, 0, "PCS", 1, ""], ["E-4.2 (Back)", "U-PKZ2", "Moeller Undervoltage Trip Unit", 1, 0, 0, 1, "PCS", 1, "Barang Lama"], ["E-4.2 (Back)", "ZMR-32-PKZ2", "Moeller Motor Protective Trip Block (24-32A)", 1, 0, 0, 1, "PCS", 1, "Barang Lama"], ["E-4.2 (Back)", "ZM-25-PKZ2", "Moeller Motor Protective Trip Block (16-25A)", 3, 0, 0, 3, "PCS", 1, "Barang Lama"], ["E-4.2 (Back)", "PKZ2", "Moeller Breaker Basic Unit", 3, 0, 0, 3, "PCS", 1, "Barang Lama"], ["Box 2", "TH-N120KPTA 105", "Mitsubishi Thermal Overload 85-125A", 1, 0, 0, 1, "PCS", 1, "Barang Lama"], ["Box 2", "TH-N20KP 19", "Mitsubishi Thermal Overload 16-22A", 3, 0, 0, 3, "PCS", 1, "Barang Lama"], ["Box 2", "TH-N20KP 2.1", "Mitsubishi Thermal Overload 1.7-2.5A", 2, 0, 0, 2, "PCS", 1, "Barang Lama"], ["D-2.2", "NF32-SW 16A", "Mitsubishi No-Fuse Breaker 16A", 0, 0, 0, 0, "PCS", 1, ""], ["Gray Bucket 1", "CT75/5", "Mitsubishi Current Transformer Ratio 75:5 (Donut)", 0, 0, 0, 0, "PCS", 1, "2nd"], ["Gray Bucket 1", "CW-15L 250/5", "Mitsubishi Current Transformer Ratio 250:5 (Donut)", 1, 0, 0, 1, "PCS", 1, ""], ["E-1", "S-N80", "Mitsubishi Contactor ; 220VAC", 1, 0, 0, 1, "PCS", 1, "Barang Lama"], ["D-2.2", "NF400-SP 250A", "Mitsubishi Breaker 3P; 250A; Front Type (Bar)", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.2", "NF250-HV", "Mitsubishi Breaker 3P; 250A", 0, 0, 0, 0, "PCS", 1, "Barang Lama (Dibelakang)"], ["D-2.2", "NF250-CV", "Mitsubishi Breaker 3P; 250A", 1, 0, 0, 1, "PCS", 1, "Barang Lama (Dibelakang)"], ["Gudang", "MIT.S.B NYMHY 4x1.5", "MIT.S.B Cable Powe; NYMHY; 4C×1.5mm², 300/500V, Cu/PVC/PVC, Black", 0, 0, 0, 0, "MTR", 1, "Stock Dalam Meter"], ["Gudang", "MIT.S.B NYMHY 4x2.5", "MIT.S.B Cable Powe; NYMHY; 4C×2.5mm², 300/500V, Cu/PVC/PVC, Black", 200, 0, 100, 100, "MTR", 5, "Stock Dalam Meter"], ["Gudang", "MIT.S.B NYMHY 3x2.5", "MIT.S.B Cable Powe; NYMHY; 3C×2.5mm², 300/500V, Cu/PVC/PVC, Black", 200, 200, 300, 100, "MTR", 5, "Stock Dalam Meter"], ["Gudang", "MIT.S.B NYMHY 2x2.5", "MIT.S.B Cable Powe; NYMHY; 2C×2.5mm², 300/500V, Cu/PVC/PVC, Black", 100, 0, 100, 0, "MTR", 1, "Stock Dalam Meter"], ["Gudang", "MIT.S.B NYMHY 4x4", "MIT.S.B Cable Powe; NYMHY; 4C×4mm², 300/500V, Cu/PVC/PVC, Black", 100, 0, 0, 100, "MTR", 5, "Stock Dalam Meter"], ["Gudang", "MIT.S.B NYMHY 4x6", "MIT.S.B Cable Powe; NYMHY; 4C×6mm², 300/500V, Cu/PVC/PVC, Black", 0, 0, 0, 0, "MTR", 1, "Stock Dalam Meter"], ["D-2.2", "3VA1196-6EE36-0AA0", "Siemens Circuit Breaker 3P; 16A", 3, 0, 0, 3, "PCS", 1, ""], ["A-2.1", "4017105M", "Merlin Gerin Masterbloc Outside Rack Modul 3 Shaft ; 2x40x5 (S-T)", 15, 0, 0, 15, "PCS", 2, ""], ["A-2.1", "4017105L", "Merlin Gerin Masterbloc Outside Rack Modul 3 Shaft ; 1x40x5 (R)", 10, 0, 0, 10, "PCS", 1, ""], ["A-2.1", "4017105C", "Merlin Gerin Masterbloc Outside Rack Modul 1 Shaft ; 10x5 RRAW (RST)", 30, 0, 0, 30, "PCS", 2, ""], ["A-2.1", "4017025K", "Merlin Gerin Masterbloc Inside-Outsode Modul 3 Shaft ; 625A; 5 Finger", 23, 0, 0, 23, "PCS", 2, ""], ["A-2.1", "4017025 A", "Merlin Gerin Masterbloc Inside Modul 1 Shaft ; 125A", 28, 0, 0, 28, "PCS", 2, ""], ["D-4.2", "28052", "Merlin Gerin Extended Rotary Handle (Adapter Breaker) for NS 80", 0, 0, 0, 0, "PCS", 1, ""], ["D-2.2", "NS100H TM80D", "Merlin Gerin Breaker 3P; 80A", 1, 0, 1, 0, "PCS", 1, ""], ["D-2.2", "LV429841", "Schneider Breaker 3P; 80A w/ TM-D (NSX100 Series)", 1, 1, 1, 1, "PCS", 1, ""], ["D-2.2", "NS80H MA25", "Schneider Breaker 3P; MA 25", 1, 0, 0, 1, "PCS", 1, ""], ["D1", "C101L 25A", "Merlin Gerin Breaker 3P; 25A", 8, 0, 0, 8, "PCS", 1, "Barang Lama"], ["D-2.2", "NS80H MA12.5", "Merlin Gerin Breaker 3P; 12.5A", 1, 0, 0, 1, "PCS", 1, ""], ["C-4.1", "Mennekes 13106", "Mennekes Plug Socket; 63A; 380-415V; 3P+E; IP44", 3, 0, 3, 0, "PCS", 1, "Welding Plug"], ["C-4.1", "Mennekes 1443A", "Mennekes Plug Socket; 125A; 380-415V; 3P+E; IP67", 1, 0, 0, 1, "PCS", 1, "Welding Plug"], ["A-2.1", "LRS-100-24", "MeanWell PSU 108W, 24VDC/4.5A, η90%, 150mVp-p, SingleOut", 1, 0, 0, 1, "PCS", 1, ""], ["A-2.1", "LRS-150-24", "MeanWell, PSU 156W,230V~/24V ⎓│6.5A,η89%,200mVp-p,SingleOut", 0, 1, 1, 0, "PCS", 1, ""], ["E-4.2", "MBS 400/5", "MBS Current Tranformer 400/5 (Model Ring)", 3, 0, 0, 3, "PCS", 1, ""], ["E-4.2", "MBS 250/1", "MBS Current Tranformer 250/1 (Model Ring)", 4, 0, 0, 4, "PCS", 1, ""], ["D-3.1", "EV62-521MJ1", "Max-Air Solenoid Valve 5; ¼\"; 220VAC", 0, 0, 0, 0, "PCS", 1, ""], ["H-8.3", "AZ 6101", "Matsushita Limit Switch; 15A", 5, 0, 0, 5, "PCS", 1, ""], ["E-6.1", "ELAW-61P", "Matsushima Pull Switch; 15A; 250V; 2NC; IP67", 0, 2, 0, 2, "PCS", 1, ""], ["E-6.1", "ELAP-20", "Matsushima Belt Sway Switch; 15A; 250V; 2NC; IP67", 3, 6, 1, 8, "PCS", 1, ""], ["H-7.3", "56C-63-121JB", "MAC Solenoid Valve; 3Way; 220VAC", 5, 0, 0, 5, "PCS", 1, ""], ["Locker 6", "ALS-MCBL01", "Lototo Devic", 8, 0, 0, 8, "PCS", 1, "Lototo Device"], ["", "23412", "Loctite®609™ Compound Retaining, Gap≤0.2mm, GreenLiq., 50ml/tb", 0, 1, 1, 0, "PCS", 1, ""], ["H-1.2", "IT65-Y-10BND2ERQ", "Lika Electronic; Encoders", 3, 0, 0, 3, "PCS", 1, ""], ["", "90616-SDLS-165", "Leonovo SPM Transducer, Vibration ¦ ,M8, Side Entry, ƒ Leonova Analyzer", 0, 1, 1, 0, "PCS", 1, ""], ["", "90507-CA0B-103", "Leonovo SPM Measuring Cable; 8pin - 2pin; L=1,5m", 0, 0, 0, 0, "PCS", 1, ""], ["D-3.1", "31011021", "Legris Fitting Stud ; Drat ½; Tube 10mm (Push-in)", 18, 0, 3, 15, "PCS", 2, ""], ["D-3.1", "31991221", "Legris Fitting Push-In; Male; Elbow; Drat ½; Tube 12mm (Push-in)", 1, 0, 1, 0, "PCS", 1, ""], ["E-3.1", "412813", "Legrand Timer 24h; 230VAC", 2, 0, 2, 0, "PCS", 1, ""], ["C-4.1", "69733", "Legrand Outlet Socket; 2P+E; 250VAC; 16A; IP55; Surface Moun (Stopkontak Outbow)", 29, 7, 26, 10, "PCS", 1, ""], ["C-4.1", "7RAD27BK", "Tanxis, Socket 4 Outlet, 250V, 16A, ≤3500W, IP54, w/ Cover (Black)", 9, 0, 0, 9, "PCS", 1, ""], ["C-5.1", "ASTN-255", "KSS Stainless Steel Tie with coating", 3, 0, 1, 2, "PCS", 1, "Cable Tie"], ["C-5.1", "CV-380W", "KSS Nylon Cable Tie; UV Resistant; 380x7.6mm (Hitam)", 26, 15, 29, 12, "PCS", 2, ""], ["C-5.1", "CV-380", "KSS Nylon Cable Tie; 380x7.6mm", 128, 18, 137, 9, "PCS", 1, ""], ["C-5.1", "CV-350", "KSS Nylon Cable Tie; 350x4.6mm", 4, 12, 5, 11, "PCS", 2, ""], ["C-5.1", "CV-300", "KSS Nylon Cable Tie; 300x7.6mm", 30, 17, 23, 24, "PCS", 2, ""], ["C-5.1", "CV-250S", "KSS Nylon Cable Tie; 250×2.5mm", -4, 11, 8, -1, "PCS", 1, ""], ["C-5.1", "CV-250M", "KSS Nylon Cable Tie; 250×3.6mm", 9, 0, 2, 7, "PCS", 1, ""], ["C-5.1", "CV-150S", "KSS Nylon Cable Tie; 155x2.5mm", 7, 14, 18, 3, "PCS", 1, ""], ["Ikhrom", "PRWS2020", "Krisbow Webbing Sling; 2T; 2M", 1, 0, 0, 1, "PCS", 1, "Tools"], ["Atas Loker", "10051754", "Krisbow Spectacle Fit Face Model Clear (Safety Glasses Clear)", 9, 0, 6, 3, "PCS", 1, "APD"], ["Ikhrom", "KW0102648", "Krisbow Screwdriver Set (8pc)", 10, 0, 2, 8, "SET", 1, "Tools"], ["Ikhrom", "LPISS7", "Krisbow Pro Insulated Screwdriver Set (7pc)", 3, 0, 2, 1, "SET", 1, "Tools"], ["Ikhrom", "LPCWS1", "Krisbow Pro Combination Wrench Set 6-22MM", 2, 0, 0, 2, "SET", 1, "Tools"], ["C-4.1", "10158647", "Krisbow Plastic Hammer; Wood Handle; 45mm", 10, 0, 0, 10, "PCS", 1, "Tools"], ["Ikhrom", "KW0101637", "Krisbow Hexagon Key Set inch (Kunci L inch)", 9, 0, 1, 8, "SET", 1, "Tools"], ["Ikhrom", "KW0100332", "Krisbow Combination Wrench Set 6-22MM", 1, 0, 0, 1, "SET", 1, "Tools"], ["Ikhrom", "KW0100312", "Krisbow Combination Wrench 6 mm (Kunci Pas-Ring)", 7, 0, 1, 6, "PCS", 1, "Tools"], ["Ikhrom", "LPCW19", "Krisbow Combination Wrench 19 mm (Kunci Pas-Ring)", 1, 0, 0, 1, "PCS", 1, "Tools"], ["Ikhrom", "KW0100318", "Krisbow Combination Wrench 12MM (Ring-Pass)", 9, 0, 0, 9, "PCS", 1, "Tools"], ["Ikhrom", "KW0101354", "Krisbow Combination Plier 8'", 1, 0, 1, 0, "PCS", 1, "Tools"], ["Ikhrom", "LRAWS6", "Krisbow Adjustable Wrench Softgrip 6\" (Konci Inggris)", 13, 0, 4, 9, "PCS", 1, "Tools"], ["Ikhrom", "LRAWS10", "Krisbow Adjustable Wrench Softgrip 10\" (Konci Inggris)", 1, 0, 1, 0, "PCS", 1, "Tools"], ["Ikhrom", "LRAW8", "Krisbow Adjustable Wrench 8\" (Konci Inggris)", 3, 0, 2, 1, "PCS", 1, "Tools"], ["C-2.2", "CH10B", "Kraus&Naimer, Switch Box; Mounting+Handle (Local Switch Single)", 1, 6, 5, 2, "BOX", 1, ""], ["C-2.2", "CA10B", "Kraus&Naimer, Switch Box; Mounting+Handle (Local Switch Double)", 2, 0, 2, 0, "BOX", 1, ""], ["C-2.2", "CA10-A233FT2-V", "Kraus&Naimer Switch, Selector; 6Position, 48×48mm,", 0, 1, 1, 0, "PCS", 1, ""], ["C-2.2", "CA10 D-U813", "Kraus&Naimer Switch, Selector; 4Position, 48×48mm,", 0, 1, 0, 1, "PCS", 1, ""], ["", "N86KNE", "KNF Pump, Vacuum ¦ 230V/50Hz/IP00, Vacuum, 6Lpm, EPDM", -2, 2, 0, 0, "PCS", 1, ""], ["WS", "ERSC2010", "Kingmill Earth Rod; Solid Copper; Ø 20xL3000mm", 3, 0, 0, 3, "PCS", 1, ""], ["Laci", "CGUV70", "Kingmill Clamp Earth Rod Copper; Ø 70...185mm", 0, 0, 0, 0, "PCS", 1, ""], ["Atas Loker", "KY2222", "King's Safety Eyewear (Safety Glasses Smoke)", 8, 0, 6, 2, "PCS", 1, "APD"], ["Atas Loker", "KY1151F", "King's Safety Eyewear (Safety Glasses Clear)", 14, 60, 38, 36, "PCS", 2, "APD"], ["H-2.7", "EDO 93045508001", "Kiepe Elektrik Relay for Speed Sensor", 3, 0, 0, 3, "PCS", 1, ""], ["Ikhrom", "KEN-601-1060K", "Kennedy Hexagon Key Set Metric&Inch (Kunci L)", 3, 0, 1, 2, "SET", 1, "Tools"], ["Repair", "5930990K", "Kennedy Bag, Tool ¦ 320×160×170mm,8Pocket,601DurableFabrc,ShouldrStrp", 1, 16, 17, 0, "PCS", 1, "Tools"], ["D-5.2", "Silicon Gun", "Kenmaster Silicon Gun", 1, 0, 1, 0, "PCS", 1, ""], ["", "25TGQ2", "KDK Ceiling Exhaust Fan 10\"", 0, 2, 2, 0, "PCS", 1, ""], ["", "MV-300NEX", "Maspion Fan, Exhaust ¦ 45W, 220V, 47×23×45cm, Size 12″, In/Out", 1, 0, 0, 1, "PCS", 1, ""], ["B-2.2", "Karung", "Karung", 160, 96, 154, 102, "PCS", 5, ""], ["A-2.2", "FJ28082MAB", "JasonFan ; Extractor Fan; 2550rpm; 230VAC; 280X280X80,5mm", 9, 0, 1, 8, "PCS", 1, ""], ["H-4.4", "IGT203", "IFM Inductive Sensor; 10…36VDC; Ø 18mm; NO", -2, 7, 0, 5, "PCS", 1, ""], ["H-4.4", "IGT204", "IFM Inductive Sensor; 10…36VDC; Ø 18mm; NC", 7, 0, 0, 7, "PCS", 1, ""], ["H-4.4", "IG001", "IFM, Inductive Sensor; Sr=5mm, M18×L80mm, 20…250V, NO, 2w×Cbl.2m, IP67", 2, 0, 0, 2, "PCS", 1, ""], ["H-4.3", "EVC711", "IFM, Cable, Connection5P×M12×4C×1mm², =4A, L=2m, Brass, Angled, IP69K", 0, 5, 0, 5, "PCS", 1, ""], ["Ikhrom", "P-15-150", "Hozan Long Nose Pliers (Tang Long Nose)", 7, 0, 4, 3, "PCS", 1, "Tools"], ["Ikhrom", "N-12", "Hozan Cutting Pliers (Tang Pemotong)", 3, 0, 3, 0, "PCS", 1, "Tools"], ["D-4.1", "A14 GDME 250", "Hirschmann Socket A14; 250 AC/DC", 2, 0, 0, 2, "PCS", 1, ""], ["D-4.1", "GDM 12B 24", "Hirschmann Socket 12B; 24 AC/DC", 2, 0, 0, 2, "PCS", 1, ""], ["F-2.1", "Hilmax E40", "Hilmax Fitting Lamp E40 (Holder Lampu)", 10, 20, 18, 12, "PCS", 2, ""], ["F-2.1", "Hilmax E27", "Hilmax Fitting Lamp E27 (Holder Lampu)", 18, 40, 48, 10, "PCS", 1, ""], ["", "764L 12", "Hem-Let Fitting Union T ; Ø 12mm (Push-in)", 0, 0, 0, 0, "PCS", 1, ""], ["", "765L 12", "Hem-Let Fitting Union L ; Ø 12mm (Push-in)", 0, 0, 0, 0, "PCS", 1, ""], ["", "765L 10", "Hem-Let Fitting Union L ; Ø 10mm (Push-in)", 0, 0, 0, 0, "PCS", 1, ""], ["", "5099077014", "Haver MEC III (CPU Module)", 0, 1, 1, 0, "PCS", 1, ""], ["", "5099077215", "Haver MEC III (Control)", 0, 1, 1, 0, "PCS", 1, ""], ["A-4.1", "3333", "Guardian Laboratories; Clean-Tact (Contact Cleaner)", 17, 0, 6, 11, "PCS", 2, ""], ["D-5.1", "GE 2193", "Guardian Laboratories Insulating Spray (Red); 432g", 4, 22, 0, 26, "PCS", 2, ""], ["D-5.1", "GL 2233", "Guardian Laboratories Instant Penetrant and Release Agent; 350g", 3, 5, 10, -2, "PCS", 1, ""], ["D-3.1", "RCA 3-8V8010-336", "Goyen Controls Solenoid Valve Set 8 (RCA 3D2; 24VDC; 20 Watt)", 6, 0, 0, 6, "SET", 1, ""], ["A-4.2", "Glo-Tex Tnr", "Glo-Tex Thinner; 0.7 Liter", 3, 39, 27, 15, "LTR", 2, ""], ["A-4.2", "Glo-Tex Ylw", "Glo-Tex Super Synthetic Paint ; Yellow (Cat Katu & Besi)", 1, 20, 9, 12, "PCS", 2, ""], ["A-4.2", "Glo-Tex Gry", "Glo-Tex Super Synthetic Paint ; Grey (Cat Katu & Besi)", 4, 84, 72, 16, "PCS", 2, ""], ["A-4.2", "Glo-Tex Grn", "Glo-Tex Super Synthetic Paint ; Green (Cat Katu & Besi)", 20, 0, 8, 12, "PCS", 2, ""], ["A-4.2", "Glo-Tex Brn", "Glo-Tex Super Synthetic Paint ; Brown (Cat Katu & Besi)", 1, 0, 2, -1, "PCS", 1, ""], ["A-4.2", "Glo-Tex Blu", "Glo-Tex Super Synthetic Paint ; Blue (Cat Katu & Besi)", 26, 10, 7, 29, "PCS", 2, ""], ["A-4.2", "Glo-Tex Blk", "Glo-Tex Super Synthetic Paint ; Black (Cat Katu & Besi)", 1, 0, 1, 0, "PCS", 1, ""], ["D-5.1", "Glazing N10", "Glazing Neutral Cure Silicone Sealant", 59, 50, 56, 53, "PCS", 5, ""], ["E-5.2", "PG-40", "Gland Cable Ø 40", 180, 0, 11, 169, "PCS", 5, ""], ["E-5.2", "PG-29", "Gland Cable Ø 36", 125, 0, 0, 125, "PCS", 5, ""], ["E-5.2", "PG-32", "Gland Cable Ø 30", 45, 0, 0, 45, "PCS", 2, ""], ["E-5.2", "PG-21", "Gland Cable Ø 26", 92, 0, 11, 81, "PCS", 5, ""], ["E-5.2", "PG-16", "Gland Cable Ø 20", 170, 0, 38, 132, "PCS", 5, ""], ["E-3.1", "MS4SA", "Fuji Electric FA Timer 0.6s to 60h; 100-2240VAC", 1, 0, 0, 1, "PCS", 1, ""], ["A-4.2", "Fox CA", "Fox Polychloroprene Contact Adhesive; 600g (Aica Aibon)", 4, 0, 0, 4, "PCS", 1, ""], ["C-2.2", "Fort SA16-4-2", "Fort Rotary Selector Switch; 4Position, 2Pole, CO, 0│1│2│3", 4, 0, 1, 3, "PCS", 1, ""], ["D-3.1", "MOEH-3/2-1/8-B", "Festo Solenoid Valve; MOEH-3/2-1/8-B", 4, 0, 0, 4, "PCS", 1, ""], ["D-3.1", "MFH-5-1/4-B", "Festo Solenoid Valve 5; ¼\" (Exclude Coil)", 5, 4, 6, 3, "PCS", 1, ""], ["D-3.1", "MFH-5-1/2", "Festo Solenoid Valve 5 way; ½\"; internal pilot air supply (Exclude Coil)", 12, 0, 6, 6, "PCS", 1, ""], ["D-3.1", "3RV1901-1A", "Siemens Aux Contact Breaker 1NC; 1NO", 2, 0, 2, 0, "PCS", 1, ""], ["D-3.1", "MFH-3-1/4", "Festo Solenoid Valve 3; ¼\" (Exclude Coil)", 5, 0, 2, 3, "PCS", 1, ""], ["H-8.5", "MFHE-3-1/4B", "Festo Solenoid Valve 3 Way; ¼", 4, 0, 0, 4, "PCS", 1, ""], ["D-3.1", "MC-2-1/8", "Festo Solenoid Valve 2; ⅛\"(Exclude Coil)", 2, 14, 11, 5, "PCS", 1, ""], ["D-3.1", "QS-12", "Festo Push-in Connector; Tube 12mm", 8, 0, 0, 8, "PCS", 1, ""], ["D-3.1", "NPQH-D-G12-Q12", "Festo, Fitting, Push-in ¦ G½[M], Sz10mm, ƒTubing OØ12mm", 18, 0, 0, 18, "PCS", 2, ""], ["D-3.1", "QS-1/4-12", "Festo Fitting, Push-in ¦ ¼[M], N.Sz=8.5mm, ƒ Tubing Ø12mm", 4, 15, 0, 10, "PCS", 1, ""], ["D-3.1", "PEV-1/4B", "Festo Pressure Switch 1-12 Bara; ¼\"", 2, 0, 0, 2, "PCS", 1, ""], ["", "10352", "Festo Flow Control Valve ¦ , GRE-¼, G¼, 0…10bar, -10…70°C, 0…996lpm, 25g", 0, 11, 11, 0, "PCS", 1, ""], ["D-3.1", "QSL-1/2-16", "Festo Fitting Push-in; Drat 2 ; Tube 16mm (Push-in)", 9, 0, 0, 9, "PCS", 1, ""], ["D-3.1", "GRE-1/4", "Festo, GRE Valve, Flow Control ¦ 1/4, G¼, 0…10bar, -10…70°C, 0…996lpm, 25g", 5, 0, 5, 0, "PCS", 1, ""], ["D-3.1", "QSLF-1/8-6-B", "Festo Fitting Push-in L ; Drat ⅛ ; Tube 6mm (Push-in)", 10, 0, 0, 10, "PCS", 1, ""], ["D-3.1", "QS-1/4-8", "Festo Fitting Push-in L ; Drat 1/4; Tube 8mm (Push-in)", 41, 0, 5, 0, "PCS", 1, ""], ["D-3.1", "QS-3/8-10", "Festo Fitting Push-in L ; Drat 1/4; Tube 8mm (Push-in)", 19, 0, 0, 19, "PCS", 2, ""], ["D-4.1", "MSFW-230 4540", "Festo Coil (Small hole) for Solenoid + Socket", 24, 0, 6, 18, "PCS", 2, ""], ["D-4.1", "3592", "Festo Coil (Big hole) Solenoid", 10, 10, 10, 10, "PCS", 1, "Socket B12 (Ready 1)"], ["F-3.2", "E7508", "Ferrule Skun (Cable Pin; Press bagian depan) 0.75mm²", 400, 1200, 703, 897, "PCS", 5, ""], ["F-3.2", "E1508", "Ferrule Skun (Cable Pin; Press bagian depan) 1.5mm²", 400, 800, 300, 900, "PCS", 5, ""], ["F-3.2", "E2508", "Ferrule Skun (Cable Pin; Press bagian depan) 2.5mm²", 0, 600, 0, 600, "PCS", 5, ""], ["B-4.2", "Fan YX3-200L-4.6", "Fan Blade For 3 Phase Induction Motor 58mm x 240mm (9 sirip)", 1, 0, 0, 1, "PCS", 1, ""], ["B-4.2", "Fan 1LE0-180-2P", "Fan Blade For 3 Phase Induction Motor 49mm x 230mm (5 sirip)", 4, 0, 0, 4, "PCS", 1, ""], ["B-4.2", "Fan 1LE0-1DB CD", "Fan Blade For 3 Phase Induction Motor 42mm x 160mm (7 sirip)", 10, 0, 0, 10, "PCS", 1, "Fan Siemens"], ["B-4.2", "Fan Y2-100", "Fan Blade For 3 Phase Induction Motor 28mm x 146mm (5 sirip)", 10, 0, 0, 10, "PCS", 1, ""], ["B-4.2", "Fan Y2-90-2", "Fan Blade For 3 Phase Induction Motor 24mm x 130mm (5 sirip)", 9, 0, 0, 9, "PCS", 1, ""], ["B-4.2", "Fan 1LE0002-180-4P", "Fan Blade For 3 Phase Induction Motor 248mm x 260mm (7 sirip)", 4, 0, 0, 4, "PCS", 1, ""], ["", "NU218-E-XL-TVP2", "FAG Bearing, Cyl.Roller ¦ Ø90/160×W30mm, Cyl.Bore, Single Row", -2, 2, 0, 0, "ROLL", 1, ""], ["", "NU319-E-XL-M1-C3", "FAG Cylindrical Roller Bearing 319; Cage Kuning/Tebal", 1, 0, 0, 1, "ROLL", 1, ""], ["", "NU319-E-TVP2-C3", "FAG Cylindrical Roller Bearing 319", 1, 0, 0, 1, "ROLL", 1, ""], ["", "NU318-E-XL-TVP2-C3", "FAG Cylindrical Roller Bearing 318", 1, 0, 0, 1, "ROLL", 1, ""], ["", "22219-E1-XL", "FAG Bearing, Sph.Roller, Ø95/170×W43mm, Cyl.Bore, Double Row", 0, 1, 0, 1, "ROLL", 1, ""], ["", "NU316-E-M1-C3", "FAG Cylindrical Roller Bearing 316", 3, 0, 0, 3, "ROLL", 1, ""], ["", "NU315-EXL-M1-C3", "FAG Cylindrical Roller Bearing 315", 1, 1, 0, 2, "ROLL", 1, ""], ["", "NU312-E-Xl-TVP2-C3", "FAG Cylindrical Roller Bearing 312", 2, 0, 1, 1, "ROLL", 1, ""], ["WS", "NU222-E-XL-TVP2-C3", "FAG Cylindrical Roller Bearing 222", 1, 1, 1, 1, "ROLL", 1, ""], ["", "NU221-E-M1-C3", "FAG Cylindrical Roller Bearing 221", 2, 0, 0, 2, "ROLL", 1, ""], ["", "NU219-E-XL-M1-C3", "FAG Cylindrical Roller Bearing 219", 3, 0, 0, 3, "ROLL", 1, ""], ["WS", "NU324-E-M1-C3", "FAG Cylindrical Bearing NU324", 4, 0, 0, 4, "PCS", 1, "MD Coal Mill (DE)"], ["WS", "NU322-E-XL-M1-C3", "FAG Cylindrical Bearing NU322", 3, 0, 0, 3, "PCS", 1, ""], ["WS", "NU321-E-M1-C3", "FAG Cylindrical Bearing NU321", 1, 0, 0, 1, "PCS", 1, ""], ["WS", "NU317-E-M1-C3", "FAG Cylindrical Bearing NU317", 2, 0, 0, 2, "PCS", 1, ""], ["WS", "NU238-E-M1-C3", "FAG Cylindrical Bearing NU238", 0, 0, 0, 0, "PCS", 1, "MD KIln"], ["WS", "NU2230-E-M1-C3", "FAG Cylindrical Bearing NU2230", 1, 0, 0, 1, "PCS", 1, ""], ["6203", "NU203-E-XL-TVP2", "FAG Cylindrical Bearing NU203", 0, 0, 0, 0, "PCS", 1, ""], ["6308", "NJ2308-E-TVP2", "FAG Cylindrical Bearing NJ2308", 2, 0, 0, 2, "PCS", 1, ""], ["", "6319-C3", "FAG Ball Bearing 6319; Seal (N/A)", 3, 0, 2, 1, "PCS", 1, ""], ["", "6318-C3", "FAG Ball Bearing 6318; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["", "6316-2Z-C3", "FAG Ball Bearing 6316; Metal Seal", 1, 0, 1, 0, "PCS", 1, ""], ["", "6315-2Z-C3", "FAG Ball Bearing 6315; Metal Seal", 6, 0, 3, 3, "PCS", 1, ""], ["", "6314-2Z-C3", "FAG Ball Bearing 6314; Metal Seal", 9, 0, 7, 2, "PCS", 1, ""], ["", "6313-2Z-C3", "FAG Ball Bearing 6313; Metal Seal", 2, 4, 4, 2, "PCS", 1, ""], ["", "6312-C-2HRS-C3", "FAG Ball Bearing 6312; Rubber Seal", 0, 0, 0, 0, "PCS", 1, ""], ["", "6312-2Z-C3", "FAG Ball Bearing 6312; Metal Seal", 5, 5, 10, 0, "PCS", 1, ""], ["", "6311-C-2Z-C3", "FAG Ball Bearing 6311; Metal Seal", 2, 0, 2, 0, "PCS", 1, ""], ["", "6310-2Z-C3", "FAG Ball Bearing 6310; Metal Seal", 8, 0, 7, 1, "PCS", 1, ""], ["", "6309-2RSR-C3", "FAG Ball Bearing 6309; Rubber Seal", 0, 14, 5, 9, "PCS", 1, ""], ["", "6309-2Z-C3", "FAG Ball Bearing 6309; Metal Seal", 1, 17, 7, 11, "PCS", 2, ""], ["", "6308-2Z-C3", "FAG Ball Bearing 6308; Metal Seal", 13, 6, 11, 8, "PCS", 1, ""], ["", "6307-2RSR-C3", "FAG Ball Bearing 6307; Rubber Seal", 12, 0, 6, 6, "PCS", 1, ""], ["", "6307-2Z-C3", "FAG Ball Bearing 6307; Metal Seal", 4, 0, 4, 0, "PCS", 1, ""], ["", "6306-2RSR-C3", "FAG Ball Bearing 6306; Rubber Seal", -1, 6, 6, -1, "PCS", 1, ""], ["", "6306-2Z-C3", "FAG Ball Bearing 6306; Metal Seal", 5, 0, 5, 0, "PCS", 1, ""], ["", "6303-C-2Z-C3", "FAG Ball Bearing 6303; Metal Seal", 6, 6, 1, 11, "PCS", 2, ""], ["", "626-2RSR-HLN", "FAG Ball Bearing 626; Rubber Seal", -15, 20, 0, 5, "PCS", 1, ""], ["", "6224-C3", "FAG Ball Bearing 6224; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["", "6222-C3", "FAG Ball Bearing 6222; Seal (N/A)", 7, 0, 0, 7, "PCS", 1, ""], ["", "6220-C3", "FAG Ball Bearing 6220; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["", "6219-C3", "FAG Ball Bearing 6219; Seal (N/A)", 6, 0, 0, 6, "PCS", 1, ""], ["", "6217-2Z-C3", "FAG Ball Bearing 6217; Metal Seal", 3, 0, 0, 3, "PCS", 1, ""], ["", "6216-C3", "FAG Ball Bearing 6216; Seal (N/A)", 0, 0, 0, 0, "PCS", 1, ""], ["", "6216-2Z-C3", "FAG Ball Bearing 6216; Metal Seal", -2, 2, 0, 0, "PCS", 1, ""], ["", "6215-C3", "FAG Ball Bearing 6215; Seal (N/A)", 0, 0, 0, 0, "PCS", 1, ""], ["", "6215-2Z-C3", "FAG Ball Bearing 6215; Metal Seal", -2, 4, 0, 2, "PCS", 1, ""], ["", "6213-C3", "FAG Ball Bearing 6213; Seal (N/A)", 3, 0, 0, 3, "PCS", 1, ""], ["", "6213ZRN", "FAG Ball Bearing 6213; Metal Seal (one side)", 2, 0, 0, 2, "PCS", 1, ""], ["", "6213-2Z-C3", "FAG Ball Bearing 6213; Metal Seal", 3, 0, 3, 0, "PCS", 1, ""], ["", "6212-2RSR-C3", "FAG Ball Bearing 6212; Ruber Seal", 16, 0, 8, 8, "PCS", 1, ""], ["", "6212-2Z-C3", "FAG Ball Bearing 6212; Metal Seal", 5, 0, 2, 3, "PCS", 1, ""], ["6212", "NU 212 ECP/C3", "SKF Cylinder Bearing NU 212", 1, 0, 0, 1, "PCS", 1, ""], ["", "6211-2Z/C3", "FAG Ball Bearing 6211; Metal Seal", 4, 0, 0, 4, "PCS", 1, ""], ["", "6210-2Z-C3", "FAG Ball Bearing 6210; Metal Seal", 18, 0, 13, 5, "PCS", 1, ""], ["", "6209-2Z-C3", "FAG Ball Bearing 6209; Metal Seal", 8, 10, 2, 16, "PCS", 2, ""], ["", "6208-2RSR-C3", "FAG Ball Bearing 6208; Rubber Seal", 8, 34, 35, 7, "PCS", 1, ""], ["", "6208-2Z-C3", "FAG Ball Bearing 6208; Metal Seal", 0, 0, 0, 0, "PCS", 1, ""], ["", "6207-C3", "FAG Ball Bearing 6207; seal (N/A)", 4, 0, 0, 4, "PCS", 1, ""], ["", "6206-2RSR-C3", "FAG Ball Bearing 6206; Rubber Seal", 7, 0, 7, 0, "PCS", 1, ""], ["", "6206-C-2HRS-C3", "FAG Ball Bearing 6206; Rubber Seal", -3, 25, 22, 0, "PCS", 1, ""], ["", "6205-2RSR-C3", "FAG Ball Bearing 6205; Rubber Seal", 1, 0, 2, -1, "PCS", 1, ""], ["", "6205-C-2HRS-C3", "FAG Ball Bearing 6205; Rubber Seal", 6, 35, 40, 1, "PCS", 1, ""], ["", "6205-C-2Z-C3", "FAG Ball Bearing 6205; Metal Seal", -6, 10, 4, 0, "PCS", 1, ""], ["", "6204-C-2HRS", "FAG Ball Bearing 6204; Ruber Seal", 12, 16, 23, 5, "PCS", 1, ""], ["", "6204-2RSR-C3", "FAG Ball Bearing 6204; Rubber Seal", 7, 0, 7, 0, "PCS", 1, ""], ["", "6204-C-2Z", "FAG Ball Bearing 6204; Rubber Seal", 0, 0, 0, 0, "PCS", 1, ""], ["", "6203-2Z", "FAG Ball Bearing 6203; Rubber Seal", 1, 0, 1, 0, "PCS", 1, ""], ["", "6202-2RSR-C3", "FAG Ball Bearing 6202; Rubber Seal", 4, 10, 14, 0, "PCS", 1, ""], ["", "6202-C-HRS", "FAG Ball Bearing 6202; Rubber Seal", -4, 15, 3, 8, "PCS", 1, ""], ["", "6201-2Z-C3", "FAG Ball Bearing 6201; Metal Seal", 8, 0, 4, 4, "PCS", 1, ""], ["", "6200-C-2HRS-C3", "FAG Ball Bearing 6200 ; Rubber Seal", 13, 0, 2, 11, "PCS", 2, ""], ["", "6004-2ZR-C3", "FAG Ball Bearing 6004 ; Metal Seal", 1, 0, 1, 0, "PCS", 1, ""], ["", "6002-C-2Z-C3", "FAG Ball Bearing 6002; Metal Seal", 6, 0, 6, 0, "PCS", 1, ""], ["", "6001-C-2Z-C3", "FAG Ball Bearing 6001; Metal Seal", 4, 0, 0, 4, "PCS", 1, ""], ["", "22222-E1-XL-C3", "FAG Ball Bearing 22222; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["", "22222 E", "SKF Ball Bearing 22222; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["WS", "6330-M-C3", "FAG Ball Bearing 6330; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["WS", "6324-M-C3", "FAG Ball Bearing 6324; Seal (N/A)", 4, 0, 0, 4, "PCS", 1, "MD Coal Mill (NDE)"], ["WS", "6320-C3", "FAG Ball Bearing 6320; Seal (N/A)", 5, 0, 5, 0, "PCS", 1, ""], ["WS", "NU320 ECP/C3", "SKF Cylinder Bearing NU320; Seal (N/A)", 2, 0, 0, 2, "PCS", 1, ""], ["WS", "6326-M-C3", "FAG Ball Bearing 6326; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["WS", "6238-M-C3", "FAG Ball Bearing 6238; Seal (N/A)", 0, 0, 0, 0, "PCS", 1, "MD KIln"], ["WS", "6048-M-C3", "FAG Ball Bearing 6048; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["WS", "22230-E1-C3", "FAG Ball Bearing 22230; Seal (N/A)", 1, 0, 0, 1, "PCS", 1, ""], ["Ikhrom", "U.53K1", "Facom Tapped Beam 40-110 mm (Tracker Palang / Puller Palang)", 1, 0, 0, 1, "PCS", 1, "Tools"], ["Ikhrom", "U.23K1A", "Facom Tapped Beam 25-95 mm (Tracker Palang / Puller Palang)", 2, 0, 1, 1, "PCS", 1, "Tools"], ["Ikhrom", "83H.JL12", "Facom Hexagon Key Set (Kunci L)", 5, 0, 1, 4, "SET", 1, "Tools"], ["D-4.1", "378A", "Facom Grease Gun", -2, 2, 0, 0, "PCS", 1, "Tools"], ["Ikhrom", "187.18VE", "Facom Combination Plier", 1, 0, 1, 0, "PCS", 1, "Tools"], ["Ikhrom", "76.8", "Facom Angled open socket wrench 8 mm (Kunci shok L)", 3, 0, 3, 0, "PCS", 1, "Tools"], ["Ikhrom", "76.22", "Facom Angled open socket wrench 22 mm (Kunci shok L)", 6, 0, 0, 6, "PCS", 1, "Tools"], ["Ikhrom", "76.19", "Facom Angled open socket wrench 19 mm (Kunci shok L)", 4, 0, 1, 3, "PCS", 1, "Tools"], ["Ikhrom", "76.17", "Facom Angled open socket wrench 17 mm (Kunci shok L)", 1, 0, 1, 0, "PCS", 1, "Tools"], ["Ikhrom", "76.13", "Facom Angled open socket wrench 13 mm (Kunci shok L)", 3, 0, 2, 1, "PCS", 1, "Tools"], ["Ikhrom", "76.099999999999994", "Facom Angled open socket wrench 10 mm (Kunci shok L)", 3, 0, 3, 0, "PCS", 1, "Tools"], ["Lemari Kaca", "F-Co Tape", "F-Co Tape; Salf-Bonding Insulating Tape; 0.5mm x 20mm x 10m (L) (HTM)", 43, 0, 1, 42, "PCS", 2, ""], ["D-5.1", "Excel PO", "Excel Penetrating Oil; 380cc", 6, 20, 26, 0, "PCS", 1, ""], ["D-5.1", "OS-60", "SK Super Coroban ¦ Oil, Penetrating : FP=60?,Dens=0.775@15? g/cm³,450mL/Can", 6, 0, 0, 6, "CAN", 1, ""], ["E-3.1", "LAD96560", "Everlink size 3 power terminal block for Contactor", 10, 0, 0, 10, "PCS", 1, ""], ["D-4.1", "Eterna 4", "Eterna Paint Brushes 4\" (Kuas)", 51, 20, 41, 30, "PCS", 2, "Tools"], ["D-4.1", "Eterna 3", "Eterna Paint Brushes 3\" (Kuas)", 73, 23, 62, 34, "PCS", 2, "Tools"], ["D-4.1", "Eterna 2", "Eterna Paint Brushes 2\" (Kuas)", 68, 40, 68, 40, "PCS", 2, "Tools"], ["D-4.1", "Eterna 1", "Eterna Paint Brushes 1\" (Kuas)", 39, 20, 11, 48, "PCS", 2, "Tools"], ["E-4.2", "ETE 400/5", "ETE Current Tranformer 400/5 (Model Ring)", 2, 0, 0, 2, "PCS", 1, ""], ["Lemari Kaca", "9V1", "Energizer Max Battery 9V", 37, 7, 8, 36, "PCS", 2, ""], ["H-4.2", "FTC 968", "Endress+ Hauser Level Limit Switch", 1, 0, 0, 1, "PCS", 1, ""], ["B-4.1", "EP-G 3/4", "Elbow, Pipe ¦ 90°Banded, NPS=¾″ (Ø20mm),Threaded, Galvanized", 7, 0, 0, 7, "PCS", 1, ""], ["B-4.1", "EP-G 1/2", "Elbow, Pipe ¦ 90°Banded, NPS=½\" (Ø mm),Threaded, Galvanized", 33, 0, 1, 32, "PCS", 2, ""], ["B-4.1", "EP-G 1", "Elbow, Pipe ¦ 90°Banded, NPS=1\" (Ø mm),Threaded, Galvanized", 0, 0, 0, 0, "PCS", 1, ""], ["B-3.1", "EL-PVC 1", "Elbow, Pipe ¦ 90°Banded, NPS=1\" , PVC", 50, 0, 0, 50, "PCS", 2, ""], ["B-3.1", "EL-PVC 1/2", "Elbow, Pipe ¦ 90°Banded, NPS=½\" , PVC", 2, 0, 0, 2, "PCS", 1, ""], ["B-3.1", "EL-PVC 2", "Elbow, Pipe ¦ 90°Banded, 2\" , PVC", 10, 0, 0, 10, "PCS", 1, ""], ["B-3.1", "EL-ALY 1 3/8", "Elbow, Fitting ; 90° ; 1⅜\" ; Alloy", 60, 0, 0, 60, "PCS", 5, "Air Conditioner"], ["B-3.1", "EL-ALY 1", "Elbow, Fitting ; 90° ; 1\" ; Alloy", 13, 0, 0, 13, "PCS", 2, "Air Conditioner"], ["B-3.1", "EL-ALY 3/4", "Elbow, Fitting ; 90° ; ¾\" ; Alloy", 17, 0, 0, 17, "PCS", 2, "Air Conditioner"], ["B-3.1", "EL-ALY 5/8", "Elbow, Fitting ; 90° ; ⅝\" ; Alloy", 56, 0, 0, 56, "PCS", 5, "Air Conditioner"], ["B-3.1", "EL-ALY 1/2", "Elbow, Fitting ; 90° ; ½\" ; Alloy", 1, 0, 0, 1, "PCS", 1, "Air Conditioner"], ["B-3.1", "EL-ALY 3/8", "Elbow, Fitting ; 90° ; ⅜\" ; Alloy", 20, 0, 0, 20, "PCS", 2, "Air Conditioner"], ["B-3.1", "SB-ALY 1 3/8", "Socket, Fitting ; 1 ⅜\" ; Alloy", 10, 0, 0, 10, "PCS", 1, "Air Conditioner"], ["B-3.1", "SB-ALY 7/8", "Socket, Fitting ; ⅞\" ; Alloy", 23, 0, 0, 23, "PCS", 2, "Air Conditioner"], ["B-3.1", "SB-ALY 5/8", "Socket, Fitting ; ⅝\" ; Alloy", 67, 0, 0, 67, "PCS", 5, "Air Conditioner"], ["B-3.1", "SB-ALY 1/2", "Socket, Fitting ; ½\" ; Alloy", 70, 0, 0, 70, "PCS", 5, "Air Conditioner"], ["H-1.7", "FS229.55ACM9B", "EH Running Counter; 220VAC", 2, 0, 0, 2, "PCS", 1, "Barang Jarang Pakai"], ["C-2.2", "6314HR", "Ebmpapst Fan; 24VDC; 1.5 A 36W (12-28VDC)", 8, 0, 0, 8, "PCS", 1, "Arah Rak B"], ["C-2.2", "4650N", "Ebmpapst Fan 160m³/h, 2650rpm, 19W, 1f230V~50Hz, 119×119×38mm", 0, 3, 0, 3, "PCS", 1, "Arah Rak B"], ["C-2.2", "W2E142-BB01-01", "Ebmpapst Fan; 230VAC; 27/28W", 3, 1, 3, 1, "PCS", 1, "Arah Rak B"], ["B-2.2", "7114N/2H", "Ebmpapst Axial Fan; Ø 150x38mm; 24VDC;", 4, 0, 0, 4, "PCS", 1, "Arah Rak B"], ["F-2.2", "Dynabolt 6x65", "Dynabolt M6 x 65", 76, 0, 44, 32, "PCS", 2, ""], ["F-2.2", "Dynabolt 20x120", "Dynabolt M20 x 120", 100, 0, 0, 100, "PCS", 5, ""], ["F-2.2", "Dynabolt 16x123", "Dynabolt M16 x 123", 100, 0, 0, 100, "PCS", 5, ""], ["F-2.2", "Dynabolt 12x60", "Dynabolt M12 x 60", 55, 0, 0, 55, "PCS", 5, ""], ["Ikhrom", "Mop Dust", "Duster Mop ; wiper type COTTON; Handle WOOD (sapu debu)", 7, 3, 5, 5, "PCS", 1, "Tools"], ["B-4.1", "DNH-G 1/2", "Double Nipple Hexagon NPS= ½\", Glvanized", 9, 0, 9, 0, "PCS", 1, ""], ["B-4.1", "DNH-G 1 1/2", "Double Nipple Hexagon NPS= 1 ½\", Glvanized", 5, 0, 5, 0, "PCS", 1, ""], ["B-4.1", "DNH-G 3/4", "Double Nipple Hexagon Galvanis ¾\"", 4, 0, 2, 2, "PCS", 1, ""], ["B-4.1", "DNH-G 1", "Double Nipple Hexagon Galvanis 1\"", 7, 0, 0, 7, "PCS", 1, ""], ["A-4.2", "Dextone PS 48", "Dextone Plastic Steel; Epoxy Adhesives; 48g", 5, 0, 3, 2, "PCS", 1, ""], ["A-4.2", "Super ND TNR", "Delta Super N.D Thinner;", 6, 7, 5, 8, "PCS", 1, ""], ["Rak Luar", "(H)07VVH6-F 4x6", "Delta Cable Flat; PVC 4Gx6mm²; 450/750V", 150, 0, 20, 130, "MTR", 5, "Stock Dalam Meter"], ["Rak Luar", "(H)07VVH6-F 12x1.5", "Delta Cable Flat; PVC 12Gx1.5mm²; 450/750V", 150, 0, 0, 150, "MTR", 5, "Stock Dalam Meter"], ["Gray Bucket 1", "S13-75", "Current Transformer for Modul Ratio 75:1", 1, 1, 0, 2, "PCS", 1, "2nd"], ["Gray Bucket 1", "CT75/5", "Current Transformer for Modul Ratio 75/5", 1, 0, 0, 1, "PCS", 1, "2nd"], ["Gray Bucket 1", "S13-5", "Current Transformer for Modul Ratio 5:1", 9, 0, 0, 9, "PCS", 1, "2nd"], ["Gray Bucket 1", "S13-40", "Current Transformer for Modul Ratio 40:1", 2, 0, 0, 2, "PCS", 1, "2nd"], ["Gray Bucket 1", "S13-30", "Current Transformer for Modul Ratio 30:1", 4, 0, 0, 4, "PCS", 1, ""], ["Gray Bucket 1", "S13-3", "Current Transformer for Modul Ratio 3:1", 1, 0, 0, 1, "PCS", 1, "2nd"], ["Gray Bucket 1", "S14-240", "Current Transformer for Modul Ratio 240:1 (Long)", 2, 0, 0, 2, "PCS", 1, "2nd"], ["Gray Bucket 1", "S14-50", "Current Transformer for Modul Ratio 50:1 (Long)", 1, 0, 0, 1, "PCS", 1, "2nd"], ["Gray Bucket 1", "S13-240", "Current Transformer for Modul Ratio 240:1", 2, 0, 0, 2, "PCS", 1, "2nd"], ["Gray Bucket 1", "S13-200", "Current Transformer for Modul Ratio 200:1", 1, 0, 0, 1, "PCS", 1, "2nd"], ["Gray Bucket 1", "S13-2.5", "Current Transformer for Modul Ratio 2.5:1", 4, 0, 0, 4, "PCS", 1, "2nd"], ["Gray Bucket 1", "CT180/1", "Current Transformer for Modul Ratio 180:1", 0, 0, 0, 0, "PCS", 1, "2nd"], ["Gray Bucket 1", "S13-16", "Current Transformer for Modul Ratio 16:1", 5, 0, 0, 5, "PCS", 1, "2nd"], ["Gray Bucket 1", "S13-125", "Current Transformer for Modul Ratio 125:1", 5, 0, 0, 5, "PCS", 1, "2nd"], ["Gray Bucket 1", "CT100/1", "Current Transformer for Modul Ratio 100:1", 1, 0, 0, 1, "PCS", 1, "2nd"], ["Gray Bucket 1", "S13-100", "Current Transformer for Modul Ratio 100:1", 4, 0, 4, 0, "PCS", 1, "2nd"], ["Gray Bucket 1", "S13-10", "Current Transformer for Modul Ratio 10:1", 3, 0, 0, 3, "PCS", 1, ""], ["Gray Bucket 1", "S13-12.5", "Current Transformer for Modul Ratio 10:1", 5, 0, 0, 5, "PCS", 1, ""], ["Gray Bucket 1", "S13-20", "Current Transformer for Modul Ratio 20:1", 1, 0, 0, 1, "PCS", 1, ""], ["Gray Bucket 1", "S13-400", "Current Transformer for Modul Ratio 400:1", 2, 0, 0, 2, "PCS", 1, ""], ["Laci", "Clamp Hose ¾", "Clamp, Hose ¦ Size=¾″, Ø(12…19)×W8mm, 80% Stainless Steel 304", 14, 0, 2, 12, "PCS", 2, ""], ["Laci", "Clamp Hose ½", "Clamp, Hose ¦ Size=½″, Ø(8…12)×W8mm, 80% Stainless Steel 304", 6, 0, 0, 6, "PCS", 1, ""], ["Laci", "Clamp Hose 1", "Clamp, Hose ¦ Size=1″, Ø(16…25)×W8mm, 80% Stainless Steel 304", 18, 0, 0, 18, "PCS", 2, ""], ["", "Clamp M4", "Clamp Wire Rope M4 (U-Clamp)", 100, 53, 0, 153, "PCS", 5, ""], ["D-3.1", "ADK11-25", "CKD Solenoid Valve 2 way; 1\"; 220VAC", -2, 5, 2, 1, "PCS", 1, ""], ["D-3.1", "APK11-8", "CKD Solenoid Valve 2 way; 1/4\"; 220VAC", 1, 0, 1, 0, "PCS", 1, ""], ["Locker", "RC73 204060", "CE® Carbon Brush RC73; 20×40×60mm", 0, 5, 0, 5, "PCS", 1, ""], ["Loker", "Dwg.EL-611-03-14", "CE® Carbon Brush RC73; 16×32×50mm, Metal Graphite", 0, 0, 0, 0, "PCS", 1, "For MD Indosin"], ["Loker", "Dwg.IDCMLe21521600", "CE Holder, Stang Grounding Ø16×L152mm, M10, ƒGround M.", 0, 0, 0, 0, "PCS", 1, "For CF and other"], ["Loker", "Dwg.2178/BFF/2532", "CE Carbon Brush RC87; 25×32×65mm; M.Graphite", 6, 0, 6, 0, "PCS", 1, "BF Fan FM7-8A"], ["Loker", "Dwg.IDCMEa1025120", "CE Carbon Brush Grounding 12.5×25×40mm", 0, 0, 0, 0, "PCS", 1, "For CF and other"], ["Loker", "Dwg.2178/A2/GS", "CE Carbon Brush 4041; 20x32x50mm", 0, 10, 0, 10, "PCS", 1, "For Kiln Drive"], ["Loker", "Dwg.2107/RPM/204000", "CE Carbom Brush M76; 65×40×20mm", 0, 10, 0, 10, "PCS", 1, "For MD RP"], ["Loker", "Dwg.2178/RMF/BH3232", "CE Brush Holder; Box(32×32)×L106mm (Ring Motor)", 0, 1, 0, 1, "BOX", 1, "For MD FM7-8A"], ["Loker", "Dwg.2178/BHRM/2532", "CE Brush Holder; Box(32×25)×L45mm; Const.Force", 0, 0, 0, 0, "BOX", 1, "Raw Mill Drive (Cegelec)"], ["Loker", "Dwg.IDCMLa12321008", "CE Brush Holder; Box_2a(32×10mm), 2_Clamping", 0, 0, 0, 0, "BOX", 1, "For Packer"], ["Loker", "Dwg.IDCMLa12251208", "CE Brush Holder Grounding; Box_2a(25×12.5mm), 2_Clamp", 0, 0, 0, 0, "BOX", 1, "For CF and other"], ["Loker", "Dwg.2178/BHEP/2040", "CE Brush Holder Box=20×40mm", 3, 0, 3, 0, "BOX", 1, "Coal Mill Drive/Turbo/EP-RM7/Crusher"], ["D-4.1", "Cape 3\"", "Cape Scraper Stainless 3\"", 22, 29, 5, 46, "PCS", 2, "Tools"], ["D-4.1", "Cape 2\"", "Cape Scraper Stainless 2\"", 13, 15, 6, 22, "PCS", 2, "Tools"], ["C-5.2", "Microfarad 8", "Capacitor 8 μF; 400~ 50/60Hz", 4, 0, 1, 3, "PCS", 1, ""], ["C-5.2", "Microfarad 6", "Capacitor 6 μF; 400/450~ 50/60Hz", 1, 0, 0, 1, "PCS", 1, ""], ["C-5.2", "Microfarad 50", "Capacitor 50 μF; 450~ 50/60Hz", 10, 0, 0, 10, "PCS", 1, ""], ["C-5.2", "Microfarad 5", "Capacitor 5 μF; 400~ 50/60Hz", 5, 0, 0, 5, "PCS", 1, ""], ["C-5.2", "Microfarad 40", "Capacitor 40 μF; 450~ 50/60Hz", 9, 0, 0, 9, "PCS", 1, ""], ["Box3", "FX86P0305J", "Capacitor 3µF ±5%; 1300VDC; 60Arms (Bentuk bulat seperti CT)", 20, 0, 0, 20, "PCS", 2, "Rarely used"], ["Box3", "PHG 491", "Capacitor 2µF ±5%; 1600-2000VDC; 40A (Bentuk bulat seperti CT)", 3, 0, 0, 3, "PCS", 1, "Rarely used"], ["Box3", "FPX669179B", "Capacitor 2µF ±5%; 1300VDC; 40A (Bentuk bulat seperti CT)", 3, 0, 0, 3, "PCS", 1, "Rarely used"], ["H-2.5", "1221", "Camsco Snap Switch; 15A; 250V", 9, 0, 0, 9, "PCS", 1, ""], ["H-1.1", "CA-22010N", "Camsco Auto Lighter; AC220V; 10A (Photo Cell)", 3, 2, 4, 1, "PCS", 1, ""], ["B-3.1", "RW-Clamp G16", "Conduit Clamp; Raceway; for pipe Ø½; Nut+Bolt", 8, 0, 0, 8, "PCS", 1, "C-Rail Clamp"], ["B-3.1", "RW-Clamp G22", "Conduit Clamp; Raceway; for pipe Ø¾; Nut+Bolt", 125, 0, 20, 105, "PCS", 5, "C-Rail Clamp"], ["B-3.1", "RW-Clamp G36", "Conduit Clamp; Raceway; for pipe Ø1¼; Nut+Bolt", 100, 0, 0, 100, "PCS", 5, "C-Rail Clamp"], ["B-3.1", "RW-Clamp G54", "Conduit Clamp; Raceway; for pipe Ø2; Nut+Bolt", 50, 0, 0, 50, "PCS", 2, "C-Rail Clamp"], ["F-3.2", "CS 95", "Cable Sleeve; 95mm", 98, 0, 0, 98, "PCS", 5, ""], ["F-3.2", "CS 70", "Cable Sleeve; 70mm", 67, 0, 0, 67, "PCS", 5, ""], ["F-3.1", "CS 6", "Cable Sleeve; 6mm", 20, 0, 0, 20, "PCS", 2, ""], ["F-3.2", "CS 50", "Cable Sleeve; 50mm", 125, 0, 0, 125, "PCS", 5, ""], ["F-3.1", "CS 4", "Cable Sleeve; 4mm", 25, 0, 0, 25, "PCS", 2, ""], ["F-3.2", "CS 400", "Cable Sleeve; 400mm", 12, 0, 0, 12, "PCS", 2, ""], ["F-3.2", "CS 35", "Cable Sleeve; 35mm", 87, 0, 0, 87, "PCS", 5, ""], ["F-3.2", "CS 25", "Cable Sleeve; 25mm", 16, 0, 0, 16, "PCS", 2, ""], ["F-3.2", "CS 185", "Cable Sleeve; 185mm", 35, 0, 0, 35, "PCS", 2, ""], ["F-3.1", "CS 16", "Cable Sleeve; 16mm", 76, 0, 15, 61, "PCS", 5, ""], ["F-3.2", "CS 120", "Cable Sleeve; 120mm", 55, 0, 0, 55, "PCS", 5, ""], ["F-3.1", "CS 10", "Cable Sleeve; 10mm", 60, 0, 0, 60, "PCS", 5, ""], ["F-3.2", "CS 150", "Cable Sleeve; 10mm", 40, 0, 0, 40, "PCS", 2, ""], ["Gudang", "NYAF 1x1.5 Rd", "Cable Control single core 1.5mm² (Red)", 51, 250, 1, 300, "MTR", 5, "Stock Dalam Meter"], ["Gudang", "NYAF 1x1.5 Blk", "Cable Control single core 1.5mm² (Black)", -200, 300, 100, 0, "MTR", 1, "Stock Dalam Meter"], ["Gudang", "NYAF 1x0.75 Rd", "Cable Control single core 0.75mm² (Red)", -150, 200, 50, 0, "MTR", 1, "Stock Dalam Meter"], ["Gudang", "NYAF 1x0.75 Brw", "Cable Control single core 0.75mm² (Brown)", 100, 300, 200, 200, "MTR", 5, "Stock Dalam Meter"], ["Gudang", "NYAF 1x0.75 Blk", "Cable Control single core 0.75mm² (Black)", 25, 425, 450, 0, "MTR", 1, "Stock Dalam Meter"], ["Gudang", "NYAF 1x10 Blk", "Cable single core 10mm² (Black)", 200, 0, 0, 200, "MTR", 5, "Stock Dalam Meter"], ["D-4.1", "14A GDM 250", "Brucken Socket 14 A GDM; 250 AC/DC; for Solenoid (Big Hole)", 46, 0, 0, 46, "PCS", 2, ""], ["", "TMS005TS", "Brioschi Step Down Trafo; 50VA; 0...230...400/0...24...48V", 1, 0, 0, 1, "PCS", 1, ""], ["Locker 6", "90848", "Brady MCB Lock PIN (Blue)", 12, 0, 0, 12, "PCS", 2, "Lototo Device"], ["F-5.1", "HH Bolt 8x60 4.6", "Bolt M8; L:60mm; Grade 4.6 (Hexagonal Head)", 8, 0, 0, 8, "PCS", 1, "Baut Putih"], ["F-5.1", "HH Bolt 8x50 8.8", "Bolt M8; L:50mm; Grade 8.8; LD (Hexagonal Head)", 250, 0, 10, 240, "PCS", 5, "Baut Hita"], ["F-5.1", "HH Bolt 8x50 4.6", "Bolt M8; L:50mm; Grade 4.6 (Hexagonal Head)", 88, 0, 6, 82, "PCS", 5, "Baut Putih"], ["F-5.1", "HH Bolt 8x40 4.6", "Bolt M8; L:40mm; Grade 4.6 (Hexagonal Head)", 144, 0, 18, 126, "PCS", 5, "Baut Putih"], ["F-5.1", "HH Bolt 8x30 4.6", "Bolt M8; L:30mm; Grade 4.6 (Hexagonal Head)", 355, 0, 73, 282, "PCS", 5, "Baut Putih"], ["F-5.1", "HH Bolt 8x15 4.6", "Bolt M8; L:15mm; Grade 4.6 (Hexagonal Head)", 300, 0, 91, 209, "PCS", 5, "Baut Putih"], ["F-5.2", "HH Bolt 6x50 4.6", "Bolt M6; L:50mm; Grade 4.6 (Hexagonal Head)", 10, 0, 0, 10, "PCS", 1, "Baut Putih"], ["F-5.2", "HH Bolt 6x40 4.6", "Bolt M6; L:40mm; Grade 4.6 (Hexagonal Head)", 9, 0, 6, 3, "PCS", 1, "Baut Putih"], ["F-5.2", "HH Bolt 6x30 4.6", "Bolt M6; L:30mm; Grade 4.6 (Hexagonal Head)", 22, 0, 34, -12, "PCS", 1, "Baut Putih"], ["F-5.2", "HH Bolt 6x20 4.6", "Bolt M6; L:20mm; Grade 4.6 (Hexagonal Head)", 186, 0, 96, 90, "PCS", 5, "Baut Putih"], ["F-5.2", "HH Bolt 5x30 4.6", "Bolt M5; L:30mm; Grade 4.6 (Hexagonal Head)", 288, 0, 13, 275, "PCS", 5, "Baut Putih"], ["F-5.2", "HH Bolt 4x30 4.6", "Bolt M4; L:30mm; Grade 4.6 (Hexagonal Head)", 420, 0, 55, 365, "PCS", 5, "Baut Putih"], ["F-5.2", "HH Bolt 4x25 4.6", "Bolt M4; L:25mm; Grade 4.6 (Hexagonal Head)", 165, 0, 43, 122, "PCS", 5, "Baut Putih"], ["", "HH Bolt 16x60 8.8 LD", "Bolt M16; L:60mm; Grade 8.8; LD (Hexagonal Head)", 40, 0, 0, 40, "PCS", 2, "Baut Hitam"], ["", "HH Bolt 16x50 4.6", "Bolt M16; L:50mm; Grade 4.6 (Hexagonal Head)", 100, 0, 0, 100, "PCS", 5, "Baut Putih"], ["", "HH Bolt 16x40 8.8", "Bolt M16; L:40mm; Grade 8.8 (Hexagonal Head)", 50, 0, 0, 50, "PCS", 2, "Baut Hitam"], ["", "HH Bolt 16x40 4.6", "Bolt M16; L:40mm; Grade 4.6 (Hexagonal Head)", 250, 0, 0, 250, "PCS", 5, "Baut Putih"], ["", "HH Bolt 16x30 4.6", "Bolt M16; L:30mm; Grade 4.6 (Hexagonal Head)", 30, 0, 0, 30, "PCS", 2, "Baut Putih"], ["", "HH Bolt 16x130 8.8 LD", "Bolt M16; L:130mm; Grade 8.8; LD (Hexagonal Head)", 20, 0, 0, 20, "PCS", 2, "Baut Hitam"], ["", "HH Bolt 16x100 8.8 LD", "Bolt M16; L:100mm; Grade 8.8; LD (Hexagonal Head)", 75, 0, 40, 35, "PCS", 2, "Baut Hitam"], ["", "HH Bolt 14x50 4.6", "Bolt M14; L:50mm; Grade 4.6 (Hexagonal Head)", 65, 0, 0, 65, "PCS", 5, "Baut Putih"], ["", "HH Bolt 14x40 8.8 LD", "Bolt M14; L:40mm; Grade 8.8; LD (Hexagonal Head)", 7, 0, 0, 7, "PCS", 1, "Baut Hitam"], ["", "HH Bolt 12x70 8.8 LD", "Bolt M12; L:70mm; Grade 8.8; LD (Hexagonal Head)", 175, 0, 0, 175, "PCS", 5, "Baut Hitam"], ["", "HH Bolt 12x60 4.6", "Bolt M12; L:60mm; Grade 4.6 (Hexagonal Head)", 45, 0, 0, 45, "PCS", 2, "Baut Putih"], ["", "HH Bolt 12x50 4.6", "Bolt M12; L:50mm; Grade 4.6 (Hexagonal Head)", 50, 0, 0, 50, "PCS", 2, "Baut Putih"], ["", "HH Bolt 12x40 4.6", "Bolt M12; L:40mm; Grade 4.6 (Hexagonal Head)", 8, 0, 0, 8, "PCS", 1, "Baut Putih"], ["", "HH Bolt 12x30 8.8 LD", "Bolt M12; L:30mm; Grade 8.8; LD (Hexagonal Head)", 80, 0, 0, 80, "PCS", 5, "Baut Hitam"], ["", "HH Bolt 12x30 4.6", "Bolt M12; L:30mm; Grade 4.6 (Hexagonal Head)", 40, 0, 0, 40, "PCS", 2, "Baut Putih"], ["", "SH Bolt 12x100 12.9", "Bolt M12; L:100mm; Grade 12.9; LD (SocketHead)", 0, 175, 0, 175, "PCS", 5, "Baut Hitam"], ["", "HH Bolt 10x60 8.8 LD", "Bolt M10; L:60mm; Grade 8.8; LD (Hexagonal Head)", 50, 0, 0, 50, "PCS", 2, "Baut Hitam"], ["", "HH Bolt 10x50 8.8 LD", "Bolt M10; L:50mm; Grade 8.8; LD (Hexagonal Head)", 40, 0, 0, 40, "PCS", 2, "Baut Hitam"], ["", "HH Bolt 10x50 4.6", "Bolt M10; L:50mm; Grade 4.6 (Hexagonal Head)", 50, 0, 29, 21, "PCS", 2, "Baut Putih"], ["", "HH Bolt 10x45 4.6", "Bolt M10; L:45mm; Grade 4.6 (Hexagonal Head)", 74, 0, 10, 64, "PCS", 5, "Baut Putih"], ["", "HH Bolt 10x25 4.6", "Bolt M10; L:25mm; Grade 4.6 (Hexagonal Head)", 297, 0, 0, 297, "PCS", 5, "Baut Putih"], ["", "SH Bolt 10x100 8.8 LD", "Bolt M10; L:100mm; Grade 8.8; LD (SocketHead)", 5, 0, 0, 5, "PCS", 1, "Baut Hitam"], ["", "HH Bolt 10x100 8.8 LD", "Bolt M10; L:100mm; Grade 8.8; LD (Hexagonal Head)", 15, 0, 0, 15, "PCS", 2, "Baut Hitam"], ["A-4.2", "NP Org", "Bodelac Primer and Top Coat Paint; Orange (Cat Anti Karat)", 1, 0, 0, 1, "PCS", 1, ""], ["Ikhrom", "85126", "Benz Webbing Sling; 2T; 6M", 1, 0, 0, 1, "PCS", 1, "Tools"], ["C-3.1", "IRDH275", "Bender Insulation Resistance A-Isometer", 1, 0, 1, 0, "MTR", 1, "Ring Motor"], ["F", "428101", "AutoPulse Module Control FCM1 ; 1×(1…159)Addr, 15…32VDC, 0.39/=5.1mA", 0, 2, 2, 0, "PCS", 1, "FSS"], ["D-3.1", "ULFBB262C002", "Asco Solenoid Valve 2; ⅛\"; 24VDC", 2, 0, 0, 2, "PCS", 1, ""], ["D-3.1", "8262G002", "Asco Solenoid Valve 2; ⅛\"; 220VAC", 3, 0, 0, 3, "PCS", 1, ""], ["D-3.1", "UG1120283", "Asco Solenoid Valve 2; 1/4\"; 220VAC", 6, 4, 9, 1, "PCS", 1, ""], ["D-3.1", "SCG353A044", "Asco Solenoid Valve 2 way; 1\"; 220VAC", 9, 0, 0, 9, "PCS", 1, ""], ["Lemari Kaca", "Artline 70 Blk", "Artline 70 High Performance; Permanen Marker; 1.5m/m; Blue (Spidol)", 0, 24, 18, 6, "PCS", 1, ""], ["Lemari Kaca", "Artline 70 Blu", "Artline 70 High Performance; Permanen Marker; 1.5m/m; Blue (Spidol)", 0, 0, 0, 0, "PCS", 1, ""], ["Laci", "SCT011-020", "APCS Transducer Current ; 0...20A / 4...20mA; 2 wire", 0, 0, 0, 0, "PCS", 1, ""], ["F", "73327", "Ansul Elect.Actuator; 12VDC/0,57A; ½\" (Solenoid for FSS)", 0, 1, 0, 1, "PCS", 1, ""], ["C-4.2", "Alphatec 37-175", "Ansell Alphatec 37-175 (Sarung tangan kimia)", 15, 0, 4, 11, "PCS", 2, "APD"], ["Lemari Kaca", "Alphatec 1500", "Ansell Alphatec 1500 (Baju EP)", 13, 60, 19, 54, "PCS", 5, "APD"], ["Rak Besi", "Angle Steel 30", "Angle Steel 30x30x3x6000mm (Besi Siku)", 0, 10, 3, 7, "PCS", 1, "(Luar Workshop)"], ["Locker 6", "ALS-PL01YLW", "Allsafe Personal PadLock (Ylw)", 39, 0, 0, 39, "PCS", 2, "Lototo Device"], ["Locker 6", "ALS-MCBL03", "Allsafe MCB Lockout POW (Red)", 9, 0, 0, 9, "PCS", 1, "Lototo Device"], ["Locker 6", "ALS-MCBL02", "Allsafe MCB Lockout PIN (Blue)", 10, 0, 0, 10, "PCS", 1, "Lototo Device"], ["Lemari Kaca", "ALS-NH02", "Allsafe Lockout Hasps; 4 Padlocks; Yellow (For rotary braker and other)", 16, 80, 0, 96, "PCS", 5, "Lototo Device"], ["Locker 6", "ALS-PL01KAYLW (20)", "Allsafe Group PadLock; 20 PadLock; 2 Key (Ylw)", 1, 0, 0, 1, "PCS", 1, "Lototo Device"], ["Locker 6", "ALS-PL01KAYLW (10)", "Allsafe Group PadLock; 10 PadLock; 2 Key (Ylw)", 4, 0, 0, 4, "PCS", 1, "Lototo Device"], ["Locker", "ALS-CBL01", "Allsafe Cirecuit Breaker Lockout (MCB/NS/GV)", -10, 10, 0, 0, "PCS", 1, "Lototo Device"], ["H-3.4", "#SDAV-10111", "Agrichema Solenoid Valve DN12; 3/2-Way; 0.5…16bar; 110…240V", 1, 0, 0, 1, "PCS", 1, ""], ["Lemari Kaca", "AA -DC", "ABC Dry Cell Battery 1.5V (AA)", 27, 0, 25, 2, "PCS", 1, ""], ["Lemari Kaca", "AAA-A", "ABC Alkaline Battery 1.5V (AAA)", 119, 30, 29, 120, "PCS", 5, ""], ["Lemari Kaca", "AA-A", "ABC Alkaline Battery 1.5V (AA)", 28, 57, 57, 28, "PCS", 2, ""], ["E-1", "AX370-30-11-80", "ABB Contactor 370A; 220VAC", 1, 0, 0, 1, "PCS", 1, ""], ["E-1", "A300-30-11", "ABB Contactor 300A; 220VAC", 0, 0, 0, 0, "PCS", 1, ""], ["Hub.Ikhrom", "6502", "3M Respirator, HalfMask, Multi Acid Gas (Khusus AFR)", 2, 4, 3, 3, "PCS", 1, "APD"], ["Atas Locker", "1990", "3M Chin Strap", 51, 20, 27, 44, "PCS", 2, "APD"], ["Atas Loker", "340-4002", "3M \"UltraFit\" Earplug", 16, 50, 43, 23, "LTR", 2, "APD"], ["D-5.1", "YP", "Yuri Porstex; 1000ml", 3, 5, 7, 1, "PCS", 1, ""], ["D-5.1", "YHS", "Yuri Hand Soap; 3.7L", 12, 5, 13, 4, "PCS", 1, ""], ["B-2.1 (Box 4)", "DIN1*1 ISO-L-A3-P1-08", "Yuan Isolator Amplifier Current; 24VDC; in 0-20mA; Out 0-100mA", 2, 0, 0, 2, "PCS", 1, "Ex AFR"], ["", "S-293", "Yahagi Motor Siren", 0, 1, 0, 1, "PCS", 1, ""], ["B-3.1", "Sankyo Valve 3/4", "Sankyo; Ball Valve ¾\"", 9, 0, 0, 9, "PCS", 1, ""], ["B-3.1", "Sankyo Valve 1", "Sankyo; Ball Valve 1\"", 11, 0, 0, 11, "PCS", 2, ""], ["B-3.1", "KITZ Valve 1/4", "KITZ; Ball Valve ¼\"", 1, 0, 0, 1, "PCS", 1, ""], ["Gudang", "Wire Rope 6", "Rope, Wire ¦ Ø6mm×7×19 IWRC, RHRL, SUS 304, MBL:2460kgf", 100, 0, 0, 100, "MTR", 5, "Satuan dalam meter"], ["C-3.2", "Dexta JB", "Dexta Box, Junction ¦ 8.5×8.5×4.5cm, Plastic, Black", 23, 0, 1, 22, "BOX", 2, ""], ["C-4.1", "Hensel KF0404", "Hensel Box, Terminal ¦ KF0404B, 10P, 32A, W104×H104×D70mm, IP69, w/ Terminal", 10, 12, 13, 9, "BOX", 1, ""], ["Rak Oil (WS)", "Synthgard 9797", "Synthgard, Grease 9797 Munex, NLGI=2, DP>280°C, 4BT=500kgf, 15kg/pl", 5, 0, 0, 5, "KG", 1, ""], ["Rak Oil (WS)", "LE 6402", "Lubrication Engineers 6402 (Comp/Turbin Oil)", 2, 0, 0, 2, "PCS", 1, ""], ["Rak Oil (WS)", "LE 4622", "Lubrication Engineers 4622 (Multiplex Oil)", 1, 0, 0, 1, "PCS", 1, ""], ["Rak Oil (WS)", "LE 6401", "Lubrication Engineers 6401 (Comp/Turbin Oil)", 3, 0, 0, 3, "PCS", 1, ""], ["Rak Oil (WS)", "Shield 4758", "Shield 4758 Sulphonex (Grease)", 2, 0, 0, 2, "PCS", 1, ""], ["Rak Oil (WS)", "Shield 7502", "Shield Aeropress XR 7502 (Oil)", 2, 0, 0, 2, "PCS", 1, ""], ["Rak Oil (WS)", "2173", "Solvent, Renovating Safety ¦ Guardian, Elecsolv, 25L/Pail", 2, 0, 0, 2, "PCS", 1, ""], ["Rak Oil (WS)", "GL-530", "Global® Cleaner, Scale & Rust , GL-530, Liquid, Black, 20kg/pai", 0, 0, 0, 0, "KG", 1, ""], ["Rak Oil (WS)", "GL-151", "Global® Cleaner, Solvent Degreaser ¦ Speed Clean, GL-151, Liquid, Clear, 20L/pl", 0, 1, 0, 1, "PCS", 1, ""], ["Rak (WS)", "PTN-16", "Conduit, Flexible ¦ PTN-16 PVC, Size=½″, Ø16mm, Steel Core, PVC Jacket, -10…+70℃", 0, 100, 0, 100, "MTR", 5, "Satuan dalam Meter"], ["Rak (WS)", "PTN-21", "Conduit, Flexible ¦ PTN-16 PVC, Size=3/4″, Ø21mm, Steel Core, PVC Jacket, -10…+70℃", 100, 0, 0, 100, "MTR", 5, "Satuan dalam Meter"], ["Rak (WS)", "Terpal 10x20", "Tarpaulin, Plastic ¦ A5, Size: 10×20m, Blue, w/ Eyelets", 0, 0, 0, 0, "PCS", 1, ""], ["Rak (WS)", "Terpal 4x6", "Tarpaulin, Plastic ¦ A5, Size: 4×6m, Blue, w/ Eyelets", 7, 0, 0, 7, "PCS", 1, ""], ["Rak (WS)", "Puma AZ", "Puma®,Accu-Zuur Electrolyte Acid, Liquid, Colorless, SpGr≥1.260, 1L/bt (Brand: Puma)", 24, 0, 0, 24, "PCS", 2, ""], ["Rak (WS)", "Pumma WD", "Puma®, Water, Demineralize pH:6…7, SG=1, <1µS/cm, Liquid, Colorless, 1L/bottle", 60, 0, 0, 60, "PCS", 5, ""], ["Rak (WS)", "Enggart E-20", "Enggart, Fan ¦ Type E-20, 6000CFM, 1440rpm, 0.75kW, 3φ380V50Hz, ≤1.95A, Sz.20″", 2, 0, 0, 2, "PCS", 1, ""], ["Rak (WS)", "PIACS DC FM.AAA.MS1 V02", "FL-Smidtn, PIACS DC; FM.AAA.MS1.V02", 6, 0, 0, 6, "PCS", 1, "2nd (EP Cool 8)"], ["Rak (WS)", "PSTX470-600-70", "ABB SoftStarter¦250kW,3f400V~,470A,Class10,IP20¦PSTX470-600-70, 1SFA898116R7000", 1, 0, 0, 1, "PCS", 1, ""], ["Rak (WS)", "CM5-5PT", "Grundfos, WaterPump; 220VAC; 900waat", 1, 0, 0, 1, "PCS", 1, ""], ["Rak (WS)", "Daikin 2PK", "Daikin AC, Split ¦ , FTC/RC50NV14, 2PK, 17100 Btu/h, 1524W, 220V,R32, 48/35¦52dBA", 1, 0, 0, 1, "PCS", 1, "Outdoor Only"], ["Rak (WS)", "Daikin Fan 10", "Daikin Motor Fan Outdoor for 10 PK", 1, 0, 0, 1, "PCS", 1, ""], ["Laci", "JOD-4801-020", "Joden, AC-DC Adaptor; 230VAC/7.5VDC; 1A", 1, 0, 0, 1, "PCS", 1, ""], ["B-2.1", "V1-FA50", "Eletta, Monitor, Flow ¦ Ø2″, 60…120Lpm, 1A@30VDC, SPDT, Brass, Flange", 0, 1, 0, 1, "PCS", 1, ""], ["E-3.1", "XSZBD10", "Telemecanique Bracket,Adjustment¦flat form 80×80×26mm,mount&fixing accs.", 0, 4, 0, 4, "CAN", 1, ""], ["D-3.1", "82402019101", "Buschjost Valve, Solenoid¦Ø12mm,2/2way,0.1…16bar¦ \\ (Brand: Norgren)", 0, 1, 0, 1, "PCS", 1, ""], ["Gudang", "Shade 12", "SLM Shade, Lamp ¦, Ø12mm, Stainless Steel, Shiny Reflector, Hanging, ƒ Base=E27", 12, 0, 0, 12, "PCS", 2, ""], ["Gudang", "Shade 16", "SLM Shade, Lamp ¦ , Ø16mm, Stainless Steel, Shiny Reflector, Hanging, ƒ Base=E40", 13, 0, 0, 13, "PCS", 2, ""], ["Gudang", "MT55-0GAF31", "MaxAir Acromatic Actuator, Pneumatic ¦ , MT55-0GAF31, 676Nm, 90°, F10/12, 27sqdr, 2×Act.", 1, 0, 0, 1, "PCS", 1, ""], ["Gudang", "99.700.6905.5", "Wieland Connector, Female ¦ SLIBUCXY2425, #99.700.6905.5, 24Pin×1.5mm², 10A", 0, 69, 0, 69, "PCS", 5, ""], ["Gudang", "99.701.6905.5", "Wieland Connector, Male ¦ SLISTCXY2425, 24Pin×1.5mm², 10A", 0, 69, 0, 69, "PCS", 5, ""], ["Rak WS", "Rivetti 450", "Rivetti, Paku Rivet 47x24.5x19.5", 0, 0, 0, 0, "PCS", 1, ""], ["Rak WS", "Rivetti 550", "Rivetti, Paku Rivet 55.5x27.5x20", 0, 0, 0, 0, "PCS", 1, ""], ["H-4.5", "AD22-22DS Y", "Shemsco, Lamp, Pilot ¦ LED,12- 48VDC, Yellow, Ø22mm, Plastic", -1, 30, 13, 16, "PCS", 2, ""], ["H-4.5", "AD22-22DS G", "Shemsco, Lamp, Pilot ¦ LED,12- 48VDC, Green, Ø22mm, Plastic", -1, 30, 13, 16, "PCS", 2, ""], ["H-4.5", "AD22-22DS R", "Shemsco, Lamp, Pilot ¦ LED,12- 48VDC, Red, Ø22mm, Plastic", 29, 0, 15, 14, "PCS", 2, ""], ["Rak WS", "1025U1204", "Legris Tubing, Polyurethane ¦, Ø12/8mm, Flex.Polyester, Blue, 25m/rl", 0, 100, 0, 100, "PCS", 5, ""], ["D-3.1", "QS-1/2-12", "Festo, Fitting, Push-in ¦ QS-G½-12, G½[M], N.Size=11mm, ƒ Tubing OØ12mm", -1, 5, 4, 0, "PCS", 1, ""], ["D-3.1", "768L", "Ham-Let, Connector, Male ¦ , Let-Lok, 12×½, NPT ½\", Ø12×L49.0×N38.9×I22.8mm", -11, 40, 9, 20, "PCS", 2, ""], ["Locker 6", "ALS-AH01", "Allsafe Lockout Hasps; Red Sz.1\", Mat: Alumunium", 0, 30, 0, 30, "PCS", 2, "Lototo Device"], ["Locker 6", "ALS-MCBL06", "Allsafe Lockout, Circuit Breaker ¦ Ø7×W17×H28mm, Yellow, ABS", 0, 20, 0, 20, "PCS", 2, "Lototo Device"], ["F", "K8411", "Kidde, Station, Manual ¦ 2×Act., ≤10kΩ, 580µA, Die-Cast Metal, Push│Pull", 0, 1, 0, 1, "PCS", 1, "FSS"], ["Rak WS", "SPW-500 BIT", "Shimizu, Pump, Submersible ¦ 210Lpm, 12m, 0.5kW, 2″pipe, Cable_10m", 0, 2, 0, 2, "PCS", 1, ""], ["", "ALS-CBL05", "Allsafe, Lockout, Circuit Breaker ¦ Allsafe,192×102×27|Shck.Ø7mm,Yellow, ƒElct", 10, 0, 0, 10, "PCS", 1, "Lototo Device (Stick)"], ["", "ALS-ACL01", "Allsafe, Adjustable Lockout (wire)", 10, 0, 0, 10, "PCS", 1, "Lototo Device (Wire)"]];
}
