/**
 * TRADELOG v3 - Google Apps Script
 * Backend: Google Sheet sebagai database, pengaturan, jurnal emosional, harga pasar.
 */

// GANTI tulisan di bawah dengan kode rahasia buatanmu sendiri (bebas, minimal 8 karakter).
// Kode yang sama nanti diketik di aplikasi TradeLog saat pertama kali dibuka.
const API_TOKEN = 'GANTI_KODE_INI';

const SHEET_TRADES = 'Trades';
const SHEET_SETTINGS = 'Settings';
const SHEET_DAILY = 'Harian';

const HEADERS = ['ID', 'Tanggal', 'Instrumen', 'Arah', 'Entry', 'SL', 'TP', 'Exit',
  'Lot', 'Fee', 'P&L', 'R', 'Setup', 'Emosi', 'Catatan', 'Hasil',
  'Emosi sesudah', 'Disiplin', 'Ikut rencana', 'Kesalahan', 'Pelajaran', 'Link chart'];

const DAILY_HEADERS = ['Tanggal', 'Mood', 'Energi', 'Tidur (jam)', 'Rencana',
  'Review', 'Catatan emosional', 'Pelajaran'];

const ALLOWED_TZ = ['Asia/Jakarta', 'Asia/Makassar', 'Asia/Jayapura', 'Asia/Singapore', 'UTC'];

const DEFAULT_CONFIG = {
  appName: 'TradeLog',
  userName: 'Trader',
  accent: '#2F6BFF',
  currency: '$',
  tz: 'Asia/Jakarta',
  dailyLossLimit: 0,
  setups: ['Breakout', 'Pullback', 'Reversal', 'Trend following', 'Support/Resistance', 'News'],
  emotions: ['Tenang', 'Percaya diri', 'Ragu', 'FOMO', 'Takut', 'Serakah', 'Balas dendam', 'Bosan', 'Lelah'],
  mistakes: ['Tidak ada', 'Entry terlalu cepat', 'Entry telat', 'Geser/hapus SL', 'Lot terlalu besar',
    'Overtrading', 'Tidak sesuai rencana', 'Cut profit terlalu cepat', 'Revenge trade'],
  checklist: ['Sudah cek kalender berita', 'Tren timeframe besar searah', 'Entry sesuai setup',
    'SL sudah ditentukan', 'Risiko maksimal 1-2% dari modal', 'Kondisi emosi stabil'],
  market: [
    { name: 'XAUUSD (Emas)', symbol: 'GC=F', color: '#E8B84A' },
    { name: 'IHSG', symbol: '^JKSE', color: '#35D08A' },
    { name: 'USD/IDR', symbol: 'IDR=X', color: '#7FB6FF' }
  ]
};

/* ---------- MENU & TAMPILAN ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📈 Jurnal Trade')
    .addItem('Setup awal (jalankan sekali)', 'setup')
    .addToUi();
}

// Web App ini hanya jadi "pintu data" untuk aplikasi TradeLog. Tidak menampilkan data apa pun lewat link.
function doGet() {
  return ContentService.createTextOutput('TradeLog API aktif.');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    if (!API_TOKEN || API_TOKEN === 'GANTI_KODE_INI' || String(API_TOKEN).length < 8) {
      throw new Error('Kode akses belum diatur di Code.gs (baris API_TOKEN).');
    }
    const req = JSON.parse(e.postData.contents);
    if (!req.token || req.token !== API_TOKEN) throw new Error('Kode akses salah.');
    const map = {
      getAppData: getAppData, getMarket: getMarket, addTrade: addTrade, updateTrade: updateTrade,
      deleteTrade: deleteTrade, saveDaily: saveDaily, saveAll: saveAll
    };
    const fn = map[req.fn];
    if (!fn) throw new Error('Fungsi tidak dikenal.');
    return json_({ ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/* ---------- SETUP & SKEMA ---------- */

function setup() {
  ensureSheets_();
  SpreadsheetApp.getUi().alert('Setup selesai. Sekarang buka aplikasi TradeLog kamu.');
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let t = ss.getSheetByName(SHEET_TRADES);
  if (!t) t = ss.insertSheet(SHEET_TRADES);
  const cols = Math.max(t.getLastColumn(), 1);
  const cur = t.getRange(1, 1, 1, cols).getValues()[0];
  if (cur[0] !== 'ID' || cols < HEADERS.length) {
    t.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#0F1426').setFontColor('#FFFFFF');
    t.setFrozenRows(1);
    t.getRange('K:K').setNumberFormat('#,##0.00');
  }

  let s = ss.getSheetByName(SHEET_SETTINGS);
  if (!s) {
    s = ss.insertSheet(SHEET_SETTINGS);
    s.getRange('A1:B1').setValues([['Modal awal', 10000]]);
    s.getRange('A2:B2').setValues([['Multiplier default', 1]]);
    s.getRange('A3:B3').setValues([['Target profit bulanan', 500]]);
    s.getRange('A4:B4').setValues([['Instrumen', 'Multiplier (nilai per 1 lot per 1 poin harga)']])
      .setFontWeight('bold');
    s.getRange('A5:B11').setValues([
      ['XAUUSD', 100], ['EURUSD', 100000], ['GBPUSD', 100000], ['BTCUSD', 1],
      ['BBCA', 100], ['BBRI', 100], ['TLKM', 100]
    ]);
    s.setColumnWidth(1, 180); s.setColumnWidth(2, 340);
  }

  let d = ss.getSheetByName(SHEET_DAILY);
  if (!d) {
    d = ss.insertSheet(SHEET_DAILY);
    d.getRange(1, 1, 1, DAILY_HEADERS.length).setValues([DAILY_HEADERS])
      .setFontWeight('bold').setBackground('#0F1426').setFontColor('#FFFFFF');
    d.setFrozenRows(1);
  }
}

/* ---------- KONFIGURASI ---------- */

function getConfig_() {
  let c = {};
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty('config');
    if (raw) c = JSON.parse(raw);
  } catch (e) { c = {}; }
  return Object.assign({}, DEFAULT_CONFIG, c);
}

function cleanList_(arr, max) {
  if (!Array.isArray(arr)) return [];
  const out = [];
  arr.forEach(x => {
    const v = String(x == null ? '' : x).trim().slice(0, 80);
    if (v && out.indexOf(v) === -1) out.push(v);
  });
  return out.slice(0, max || 50);
}

function cleanConfig_(c) {
  const d = DEFAULT_CONFIG;
  const out = {};
  out.appName = String(c.appName || d.appName).trim().slice(0, 30) || d.appName;
  out.userName = String(c.userName || d.userName).trim().slice(0, 30) || d.userName;
  out.accent = /^#[0-9a-fA-F]{6}$/.test(c.accent) ? c.accent : d.accent;
  out.currency = String(c.currency == null ? d.currency : c.currency).trim().slice(0, 4);
  out.tz = ALLOWED_TZ.indexOf(c.tz) >= 0 ? c.tz : d.tz;
  out.dailyLossLimit = Math.max(0, Number(c.dailyLossLimit) || 0);
  out.setups = cleanList_(c.setups, 40);
  out.emotions = cleanList_(c.emotions, 30);
  out.mistakes = cleanList_(c.mistakes, 30);
  out.checklist = cleanList_(c.checklist, 20);
  if (!out.emotions.length) out.emotions = d.emotions;
  const mk = [];
  (Array.isArray(c.market) ? c.market : []).forEach(m => {
    const symbol = String(m.symbol || '').trim();
    if (!/^[A-Za-z0-9^=.\-]{1,20}$/.test(symbol)) return;
    mk.push({
      name: String(m.name || symbol).trim().slice(0, 30),
      symbol: symbol,
      color: /^#[0-9a-fA-F]{6}$/.test(m.color) ? m.color : '#7FB6FF'
    });
  });
  out.market = mk.slice(0, 6);
  return out;
}

// Simpan semua pengaturan sekaligus dari halaman Pengaturan.
function saveAll(p) {
  ensureSheets_();
  const cfg = cleanConfig_(p.config || {});
  PropertiesService.getDocumentProperties().setProperty('config', JSON.stringify(cfg));

  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  s.getRange('A1:B1').setValues([['Modal awal', Number(p.balance) || 0]]);
  s.getRange('A3:B3').setValues([['Target profit bulanan', Number(p.target) || 0]]);

  const list = [];
  const seen = {};
  (Array.isArray(p.instruments) ? p.instruments : []).forEach(i => {
    const name = String(i.name || '').trim().toUpperCase().slice(0, 20);
    const mult = Number(i.mult);
    if (name && mult > 0 && !seen[name]) { seen[name] = 1; list.push([name, mult]); }
  });
  const last = s.getLastRow();
  if (last >= 5) s.getRange(5, 1, last - 4, 2).clearContent();
  if (list.length) s.getRange(5, 1, list.length, 2).setValues(list);
  return true;
}

function getSettings_() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  const balance = Number(s.getRange('B1').getValue()) || 0;
  const defMult = Number(s.getRange('B2').getValue()) || 1;
  const target = Number(s.getRange('B3').getValue()) || 0;
  const last = s.getLastRow();
  const mult = {};
  if (last >= 5) {
    s.getRange(5, 1, last - 4, 2).getValues().forEach(r => {
      if (r[0]) mult[String(r[0]).trim().toUpperCase()] = Number(r[1]) || defMult;
    });
  }
  return { balance: balance, defMult: defMult, target: target, mult: mult };
}

/* ---------- DATA ---------- */

function toDateStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v == null ? '' : v);
}

function str_(v) { return v == null ? '' : String(v); }
function numOrNull_(v) { return v === '' || v == null || isNaN(Number(v)) ? null : Number(v); }

function getAppData() {
  ensureSheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = getSettings_();

  const trades = [];
  const sh = ss.getSheetByName(SHEET_TRADES);
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues().forEach(r => {
      if (!r[0]) return;
      trades.push({
        id: str_(r[0]), date: toDateStr_(r[1]), inst: str_(r[2]), dir: str_(r[3]),
        entry: numOrNull_(r[4]), sl: numOrNull_(r[5]), tp: numOrNull_(r[6]), exit: numOrNull_(r[7]),
        lot: numOrNull_(r[8]), fee: Number(r[9]) || 0,
        pnl: Number(r[10]) || 0, r: numOrNull_(r[11]),
        setup: str_(r[12]), emotion: str_(r[13]), note: str_(r[14]), result: str_(r[15]),
        afterEmotion: str_(r[16]), discipline: numOrNull_(r[17]), followed: str_(r[18]),
        mistake: str_(r[19]), lesson: str_(r[20]), link: str_(r[21])
      });
    });
  }
  trades.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0);

  const daily = [];
  const dsh = ss.getSheetByName(SHEET_DAILY);
  if (dsh.getLastRow() > 1) {
    dsh.getRange(2, 1, dsh.getLastRow() - 1, DAILY_HEADERS.length).getValues().forEach(r => {
      if (!r[0]) return;
      daily.push({
        date: toDateStr_(r[0]), mood: numOrNull_(r[1]), energy: numOrNull_(r[2]), sleep: numOrNull_(r[3]),
        plan: str_(r[4]), review: str_(r[5]), emotion: str_(r[6]), lesson: str_(r[7])
      });
    });
  }
  daily.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0);

  const instruments = Object.keys(st.mult).map(k => ({ name: k, mult: st.mult[k] }));
  return { trades: trades, daily: daily, balance: st.balance, target: st.target,
           instruments: instruments, config: getConfig_() };
}

function addTrade(t) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TRADES);
  if (!sh) throw new Error('Sheet Trades belum ada. Jalankan Setup awal dulu.');

  const inst = String(t.inst || '').trim().toUpperCase();
  const entry = Number(t.entry), exit = Number(t.exit), lot = Number(t.lot);
  const sl = t.sl === '' || t.sl == null ? '' : Number(t.sl);
  const tp = t.tp === '' || t.tp == null ? '' : Number(t.tp);
  const fee = Number(t.fee) || 0;
  if (!inst || !t.date || isNaN(entry) || isNaN(exit) || isNaN(lot)) {
    throw new Error('Tanggal, instrumen, entry, exit, dan lot wajib diisi.');
  }

  const st = getSettings_();
  const mult = st.mult[inst] || st.defMult;
  const dir = t.dir === 'Sell' ? -1 : 1;
  const gross = (exit - entry) * dir * lot * mult;
  const pnl = Math.round((gross - fee) * 100) / 100;

  let r = '';
  if (sl !== '' && !isNaN(sl) && sl !== entry) {
    const risk = Math.abs(entry - sl) * lot * mult;
    r = Math.round((gross / risk) * 100) / 100;
  }
  const result = pnl > 0 ? 'Win' : pnl < 0 ? 'Loss' : 'BE';
  const id = Utilities.getUuid().slice(0, 8);
  const disc = Number(t.discipline);
  const link = /^https?:\/\//i.test(String(t.link || '')) ? String(t.link).slice(0, 500) : '';

  sh.appendRow([id, '', inst, t.dir === 'Sell' ? 'Sell' : 'Buy', entry, sl, tp, exit, lot, fee, pnl, r,
    str_(t.setup).slice(0, 60), str_(t.emotion).slice(0, 40), str_(t.note).slice(0, 1000), result,
    str_(t.afterEmotion).slice(0, 40), disc >= 1 && disc <= 5 ? disc : '',
    t.followed === 'Ya' || t.followed === 'Tidak' ? t.followed : '',
    str_(t.mistake).slice(0, 60), str_(t.lesson).slice(0, 1000), link]);

  // simpan tanggal sebagai teks supaya tidak bergeser zona waktu
  sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(String(t.date).slice(0, 10));
  return { id: id, pnl: pnl, r: r };
}

// Edit kolom ulasan/emosional setelah trade selesai.
function updateTrade(id, patch) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TRADES);
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  const col = { setup: 13, emotion: 14, note: 15, afterEmotion: 17, discipline: 18,
                followed: 19, mistake: 20, lesson: 21, link: 22 };
  for (let i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) !== String(id)) continue;
    const row = i + 1;
    Object.keys(patch || {}).forEach(k => {
      if (!col[k]) return;
      let v = patch[k];
      if (k === 'discipline') { v = Number(v); v = v >= 1 && v <= 5 ? v : ''; }
      else if (k === 'followed') { v = v === 'Ya' || v === 'Tidak' ? v : ''; }
      else if (k === 'link') { v = /^https?:\/\//i.test(String(v || '')) ? String(v).slice(0, 500) : ''; }
      else { v = str_(v).slice(0, 1000); }
      sh.getRange(row, col[k]).setValue(v);
    });
    return true;
  }
  throw new Error('Trade tidak ditemukan.');
}

function deleteTrade(id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TRADES);
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = ids.length - 1; i >= 1; i--) {
    if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 1); return true; }
  }
  return false;
}

/* ---------- JURNAL HARIAN (EMOSIONAL) ---------- */

function saveDaily(d) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DAILY);
  const date = String(d.date || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal tidak valid.');
  const mood = Number(d.mood);
  if (!(mood >= 1 && mood <= 5)) throw new Error('Pilih mood dulu.');
  const energy = Number(d.energy);
  const row = [date, mood, energy >= 1 && energy <= 5 ? energy : '', numOrNull_(d.sleep) == null ? '' : Number(d.sleep),
    str_(d.plan).slice(0, 2000), str_(d.review).slice(0, 2000),
    str_(d.emotion).slice(0, 2000), str_(d.lesson).slice(0, 2000)];

  const last = sh.getLastRow();
  let target = -1;
  if (last > 1) {
    const dates = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (toDateStr_(dates[i][0]) === date) { target = i + 2; break; }
    }
  }
  if (target < 0) target = last + 1;
  sh.getRange(target, 1).setNumberFormat('@');
  sh.getRange(target, 1, 1, row.length).setValues([row]);
  return true;
}

/* ---------- PASAR (simbol bisa diatur di Pengaturan) ---------- */

function getMarket() {
  const list = getConfig_().market;
  const keyRaw = JSON.stringify(list);
  const key = 'mk_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, keyRaw));
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const out = list.map(m => {
    try {
      const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' +
        encodeURIComponent(m.symbol) + '?range=1mo&interval=1d';
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
      const j = JSON.parse(res.getContentText());
      const r = j.chart.result[0];
      const closes = r.indicators.quote[0].close;
      const tz = Session.getScriptTimeZone();
      const series = [], lab = [];
      closes.forEach((c, i) => {
        if (typeof c === 'number' && isFinite(c)) {
          series.push(c);
          lab.push(Utilities.formatDate(new Date(r.timestamp[i] * 1000), tz, 'dd MMM'));
        }
      });
      if (!series.length) throw new Error('Data kosong');
      const price = typeof r.meta.regularMarketPrice === 'number' ? r.meta.regularMarketPrice : series[series.length - 1];
      const prev = series.length > 1 ? series[series.length - 2] : price;
      const change = price - prev;
      return { name: m.name, symbol: m.symbol, color: m.color, price: price, change: change,
               pct: prev ? (change / prev) * 100 : 0, hi: Math.max.apply(null, series),
               lo: Math.min.apply(null, series), labels: lab, series: series };
    } catch (e) {
      return { name: m.name, symbol: m.symbol, color: m.color, error: String(e) };
    }
  });
  cache.put(key, JSON.stringify(out), 300);
  return out;
}
