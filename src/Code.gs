/**
 * WMS Warehouse — Kesesuaian Rak + Kontrol Muat (1 aplikasi)
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me | Who has access: Anyone with the link (atau domain Anda)
 * Data disimpan di Google Sheet "WMS_DB" (dibuat otomatis / memakai sheet yang terikat script).
 */
const SHEET_NAME = 'DB';
const CHUNK = 40000; // batas sel Sheet 50.000 karakter


/* ============ KEAMANAN ============
 * Script properties (Project Settings > Script properties):
 *  (PIN kini juga bisa diatur di tab "PIN" pada spreadsheet database; lihat bagian TABEL PIN)
 *  APP_PIN        : (opsional) PIN tim gudang (baca + tulis). Diminta sekali di browser, lalu diingat di perangkat.
 *  VIEW_PIN       : (opsional) PIN manager, hanya bisa MEMBACA. Dipakai di link dashboard: <URL web app>?v=dash
 *  ALLOWED_EMAILS : (opsional) email yang boleh memakai app, pisahkan koma. Hanya andal bila semua
 *                   pengguna satu domain Google Workspace; akun Gmail biasa terbaca kosong.
 *  SRC_ALLOW      : (opsional) ID spreadsheet sumber yang boleh dibaca (pisahkan koma).
 *                   Cara yang lebih mudah: daftarkan di tab "Config" (otomatis dibuat), kolom Izinkan = YA.
 */
function props_(k) { return (PropertiesService.getScriptProperties().getProperty(k) || '').split(/[,\s]+/).map(s => s.trim()).filter(Boolean); }

const INTERNAL_ = {};   // penanda panggilan dari editor Apps Script (tidak bisa dipalsukan dari browser)

function authEmail_() {
  const allow = props_('ALLOWED_EMAILS').map(s => s.toLowerCase());
  if (!allow.length) return;
  let me = '';
  try { me = (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  if (!me || allow.indexOf(me) < 0) throw new Error('Akses ditolak untuk akun ini');
}

/* ============ TABEL PIN (tab "PIN" di spreadsheet database) ============
 * Dibuat OTOMATIS saat aplikasi dibuka pertama kali (atau jalankan siapkanPIN dari editor) dengan 2 baris default:
 *   Tim Gudang = TULIS (baca + tulis)  |  Manager = BACA (hanya dashboard, link <URL>?v=dash)
 * PIN default diacak 8 digit. Ubah langsung di tab "PIN", tambah baris untuk peran/orang lain.
 * Kolom Aktif = TIDAK untuk mencabut PIN. Perubahan berlaku maksimal ~1 menit (cache).
 * Script properties APP_PIN / VIEW_PIN (bila diisi) TETAP berlaku, berdampingan dengan tabel ini.
 * CATATAN: siapa pun yang bisa membuka spreadsheet database bisa melihat PIN. Batasi akses share spreadsheet-nya. */
const PIN_SHEET = 'PIN';
const PIN_HDR = ['Peran', 'PIN', 'Akses (TULIS / BACA)', 'Aktif (YA / TIDAK)', 'Catatan'];
function pinRnd_() { let s = String(1 + Math.floor(Math.random() * 9)); for (let i = 0; i < 7; i++) s += Math.floor(Math.random() * 10); return s; }

function pinSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(PIN_SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(PIN_SHEET);
  sh.getRange(1, 1, 1, PIN_HDR.length).setValues([PIN_HDR]).setFontWeight('bold').setBackground('#0a6c74').setFontColor('#ffffff').setWrap(true);
  sh.setFrozenRows(1);
  sh.getRange(2, 1, 200, 2).setNumberFormat('@');          // PIN disimpan sebagai teks (angka 0 di depan tidak hilang)
  sh.getRange(2, 3, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TULIS', 'BACA'], true).build());
  sh.getRange(2, 4, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['YA', 'TIDAK'], true).build());
  sh.getRange(2, 1, 2, PIN_HDR.length).setValues([
    ['Tim Gudang', pinRnd_(), 'TULIS', 'YA', 'Default: baca + tulis. Ganti PIN-nya sebelum dibagikan.'],
    ['Manager', pinRnd_(), 'BACA', 'YA', 'Default: hanya baca (link dashboard ?v=dash).']
  ]);
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 150); sh.setColumnWidth(4, 130); sh.setColumnWidth(5, 380);
  return sh;
}

/** Daftar PIN aktif dari tabel: [{pin, w}] (cache 60 detik agar tidak membaca sheet di setiap panggilan) */
function pinList_() {
  const C = CacheService.getScriptCache();
  try { const c = C.get('PINTBL'); if (c) return JSON.parse(c); } catch (e) {}
  let out = [];
  try {
    const sh = pinSheet_(), n = sh.getLastRow();
    if (n >= 2) sh.getRange(2, 1, n - 1, 4).getValues().forEach(r => {
      const pin = String(r[1] == null ? '' : r[1]).trim();
      if (pin && String(r[3]).trim().toUpperCase() === 'YA') out.push({ pin: pin, w: String(r[2]).trim().toUpperCase() === 'TULIS' });
    });
  } catch (e) { out = []; }
  try { C.put('PINTBL', JSON.stringify(out), 60); } catch (e) {}
  return out;
}

/** Semua PIN yang berlaku: Script properties (APP_PIN tulis, VIEW_PIN baca) + tabel PIN */
function pinAll_() {
  const P = PropertiesService.getScriptProperties();
  const pin = (P.getProperty('APP_PIN') || '').trim(), vpin = (P.getProperty('VIEW_PIN') || '').trim();
  const all = pinList_().slice();
  if (pin) all.push({ pin: pin, w: true });
  if (vpin) all.push({ pin: vpin, w: false });
  return all;
}

/** Jalankan SEKALI dari editor (pilih "siapkanPIN" lalu Run): membuat tab PIN + PIN default, hasilnya tampil di Execution log. */
function siapkanPIN() {
  CacheService.getScriptCache().remove('PINTBL');
  const sh = pinSheet_();
  Logger.log('Tab PIN siap di: ' + ss_().getUrl() + '  (lihat kolom PIN)');
  return sh.getLastRow() - 1;
}

/** Semua fungsi yang dipanggil dari browser menerima PIN sebagai argumen pertama (tok).
 *  PIN dibaca dari tab "PIN" (dan/atau Script properties APP_PIN / VIEW_PIN). Tanpa satu pun PIN TULIS, aplikasi tertutup. */
function auth_(tok, write) {
  if (tok === INTERNAL_) return;
  authEmail_();
  const all = pinAll_();
  if (!all.some(e => e.w)) throw new Error('PIN belum diisi: isi tab "PIN" di spreadsheet database');   // go-live: aplikasi tidak boleh terbuka tanpa PIN
  const t = String(tok == null ? '' : tok).trim();
  let ok = false;
  all.forEach(e => { if (same_(t, e.pin) && (e.w || !write)) ok = true; });   // tanpa berhenti di kecocokan pertama
  if (ok) return;
  Utilities.sleep(2000); throw new Error('PIN_REQUIRED');
}
/** perbandingan string tanpa berhenti di karakter pertama yang beda */
function same_(a, b) { a = String(a); b = String(b); let d = a.length ^ b.length; for (let i = 0; i < Math.max(a.length, b.length); i++) d |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0); return d === 0; }
function authCheck(tok) { auth_(tok); return true; }

function srcId_(u) { const m = String(u).match(/\/d\/([a-zA-Z0-9_-]+)/); return m ? m[1] : String(u).trim(); }

function doGet(e) {
  try { authEmail_(); } catch (err) { return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:2rem">Akses ditolak. Hubungi admin.</p>'); }
  if (!pinAll_().some(e => e.w)) return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:2rem">Aplikasi belum dikonfigurasi. Admin: isi PIN di tab <b>PIN</b> pada spreadsheet database (atau Script property APP_PIN), lalu muat ulang.</p>');
  const view = !!(e && e.parameter && e.parameter.v === 'dash');   // link manager: <URL>?v=dash (tampilan baca-saja)
  const out = HtmlService.createHtmlOutputFromFile('Index');
  if (view) out.setContent(out.getContent().replace('window.__VIEW=false/*__MODE__*/', 'window.__VIEW=true'));
  return out
    .setTitle(view ? 'Dashboard KPI WMS' : 'WMS Warehouse')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function ss_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SSID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  ss = SpreadsheetApp.create('WMS_DB');
  props.setProperty('SSID', ss.getId());
  return ss;
}

function sheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['key', 'idx', 'data', 'updated']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Ambil semua data (rak & muat) sekali jalan */
function rev_(k) { return Number(PropertiesService.getScriptProperties().getProperty('REV_' + k) || 0); }

/** Cek versi saja (murah): browser hanya menarik seluruh data bila versi berubah */
function dbRev(tok) { auth_(tok); return { rak: rev_('rak'), muat: rev_('muat') }; }

function dbGetAll(tok) {
  auth_(tok);
  const sh = sheet_();
  const last = sh.getLastRow();
  const out = { rak: '', muat: '', rev: { rak: rev_('rak'), muat: rev_('muat') } };
  if (last < 2) return out;
  const vals = sh.getRange(2, 1, last - 1, 3).getValues();
  const parts = { rak: [], muat: [] };
  vals.forEach(r => { if (parts[r[0]]) parts[r[0]].push([Number(r[1]), String(r[2])]); });
  Object.keys(parts).forEach(k => {
    out[k] = parts[k].sort((a, b) => a[0] - b[0]).map(x => x[1]).join('');
  });
  return out;
}

/** Simpan 1 key (rak / muat) sebagai JSON string terpotong-potong.
 *  baseRev = versi data yang terakhir dilihat browser. Bila server sudah lebih baru (perangkat lain
 *  menyimpan lebih dulu), data TIDAK ditimpa: server mengembalikan {conflict:true, data} agar browser menggabungkan.
 *  Versi lama disalin ke DB_Backup (hingga 20 salinan, jarak min. 30 menit, atau langsung bila ukuran turun drastis). */
function dbSet(tok, key, json, baseRev) {
  auth_(tok, true);
  if (key !== 'rak' && key !== 'muat') throw new Error('key tidak valid');
  if (typeof json !== 'string' || !json) throw new Error('data kosong');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cur = rev_(key);
    const sh = sheet_();
    const last = sh.getLastRow();
    let keep = [], old = [];
    if (last >= 2) {
      const vals = sh.getRange(2, 1, last - 1, 4).getValues();
      keep = vals.filter(r => r[0] !== key);
      old = vals.filter(r => r[0] === key);
    }
    const br = (baseRev == null || baseRev === '') ? null : Number(baseRev);
    if ((br === null && cur > 0) || (br !== null && br !== cur)) {   // klien tanpa versi tidak boleh menimpa data yang sudah ada
      const data = old.slice().sort((a, b) => Number(a[1]) - Number(b[1])).map(r => String(r[2])).join('');
      log_('Konflik', key, json.length, 'rev klien ' + (br === null ? '-' : br) + ' ≠ server ' + cur);
      return { conflict: true, rev: cur, data: data };
    }
    const oldLen = old.reduce((a, r) => a + String(r[2]).length, 0);
    const drop = oldLen > 2000 && json.length < oldLen * 0.5;
    backup_(key, old, drop);
    const now = new Date();
    const rows = [];
    for (let i = 0, n = 0; i < json.length || n === 0; i += CHUNK, n++) {
      rows.push([key, n, json.substr(i, CHUNK), now]);
    }
    const all = keep.concat(rows);
    const prev = last >= 2 ? last - 1 : 0;
    sh.getRange(2, 3, all.length, 1).setNumberFormat('@');   // cegah potongan JSON terbaca sebagai angka/tanggal
    sh.getRange(2, 1, all.length, 4).setValues(all);          // TULIS DULU: bila gagal di sini, data lama tetap utuh
    if (prev > all.length) sh.getRange(all.length + 2, 1, prev - all.length, 4).clearContent();   // baru buang sisa baris lama
    PropertiesService.getScriptProperties().setProperty('REV_' + key, String(cur + 1));
    log_('Simpan', key, json.length, drop ? 'PERINGATAN: ukuran data turun drastis (' + oldLen + ' → ' + json.length + '). Versi lama ada di DB_Backup.' : '');
    return { ok: true, rev: cur + 1 };
  } finally {
    lock.releaseLock();
  }
}

const BK_KEEP = 20, BK_GAP = 30 * 60000;
function backup_(key, oldRows, force) {
  if (!oldRows.length) return;
  const ss = ss_();
  let b = ss.getSheetByName('DB_Backup');
  if (!b) { b = ss.insertSheet('DB_Backup'); b.appendRow(['key', 'idx', 'data', 'updated']); b.setFrozenRows(1); }
  const n = b.getLastRow();
  let rows = n >= 2 ? b.getRange(2, 1, n - 1, 4).getValues() : [];
  rows = rows.filter(r => r[0] !== key);                       // buang format backup lama (1 salinan)
  const ids = [];
  rows.forEach(r => { const id = String(r[0]); if (id.indexOf(key + '@') === 0 && ids.indexOf(id) < 0) ids.push(id); });
  const latest = ids.reduce((m, id) => Math.max(m, Number(id.split('@')[1]) || 0), 0);
  const now = Date.now();
  if (!force && now - latest < BK_GAP) return;
  const id = key + '@' + now;
  ids.push(id);
  let all = rows.concat(old_(oldRows, id));
  const dead = ids.slice(0, Math.max(ids.length - BK_KEEP, 0));
  if (dead.length) all = all.filter(r => dead.indexOf(String(r[0])) < 0);
  if (all.length) {
    b.getRange(2, 1, all.length, 3).setNumberFormat('@');
    b.getRange(2, 1, all.length, 4).setValues(all);
  }
  if (n - 1 > all.length) b.getRange(all.length + 2, 1, n - 1 - all.length, 4).clearContent();
}
function old_(rows, id) { return rows.map(r => [id, r[1], r[2], r[3]]); }

/** Jejak audit: siapa/kapan menyimpan apa (maks 500 baris terakhir) */
function log_(aksi, key, size, note) {
  try {
    const ss = ss_();
    let l = ss.getSheetByName('Log');
    if (!l) {
      l = ss.insertSheet('Log');
      l.getRange(1, 1, 1, 6).setValues([['Waktu', 'Aksi', 'Data', 'Ukuran (char)', 'Pengguna', 'Catatan']]).setFontWeight('bold').setBackground('#0a6c74').setFontColor('#ffffff');
      l.setFrozenRows(1);
    }
    let user = '';
    try { user = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail(); } catch (e) {}
    const lr = l.getLastRow() + 1;
    txt_(l.getRange(lr, 1, 1, 6)).setValues([[Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss'), aksi, key, String(size), user || '-', String(note || '')]]);
    const n = l.getLastRow();
    if (n > 501) l.deleteRows(2, n - 501);
  } catch (e) {}
}


/* ============ REKAP ke Google Spreadsheet (Harian / Mingguan / Bulanan) ============
 * Dipanggil dari aplikasi. Seluruh tab direbuat dari data terbaru, jadi tanggal yang
 * dihapus di aplikasi ikut hilang dari rekap. Persentase memakai RUMUS sheet.
 * Minggu (W1..W5) ditentukan dari tanggal SENIN: tgl 1-7 = W1, 8-14 = W2, 15-21 = W3,
 * 22-28 = W4, 29-31 = W5. Bulanan = jumlah data HARIAN pada bulan kalender yang sama (sama dengan dashboard).
 */
const BLN_F = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const BLN_S = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
const HARI_F = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
const REKAP_HDR = {
  Harian: ['Tanggal','Hari','Minggu','Total Mobil','Mobil WMS','Mobil Tidak WMS','% Mobil WMS','Total SPM','SPM WMS','SPM Copy Rekom.','SPM Full OK','Total Item','Item Copy Rekom.','Item OK (Aktual)','% Copy Rekom.','% Item OK','SPM Sudah Dicek','SPM Belum Dicek','Status Data','Diupdate','Bulan','Mobil Masuk'],
  Mingguan: ['Mulai (Senin)','Bulan (tgl Senin)','Minggu','Periode','Total Mobil','Mobil WMS','Mobil Tidak WMS','% Mobil WMS','Total SPM','SPM WMS','SPM Copy Rekom.','SPM Full OK','Total Item','Item Copy Rekom.','Item OK (Aktual)','% Copy Rekom.','% Item OK','SPM Sudah Dicek','Hari Data','Diupdate'],
  Bulanan: ['Bulan','Hari Data','Total Mobil','Mobil WMS','Mobil Tidak WMS','% Mobil WMS','Total SPM','SPM WMS','SPM Copy Rekom.','SPM Full OK','Total Item','Item Copy Rekom.','Item OK (Aktual)','% Copy Rekom.','% Item OK','SPM Sudah Dicek'],
  /* KPI SKU selisih (qty muat ≠ qty SPM): kolom teks dahulu, angka sesudahnya */
  Selisih_SKU: ['Tanggal','Bulan','No SPM','SKU','Deskripsi','Ship To','Nopol','PIC Muat','PIC Check','Rak','Jenis','Tindak Lanjut','Catatan Tindak Lanjut','Diupdate','Qty Checker Out (SPM)','Qty Muat','Selisih'],
  Selisih_PIC: ['Bulan','PIC Muat','SKU Dimuat','SKU Selisih','% SKU Selisih','Sudah Ditindak','Belum Ditindak','SKU Bersama PIC Lain'],
  Selisih_Bulanan: ['Bulan','Hari Data','SKU Dimuat','SKU Selisih','% SKU Selisih','Ditangkap Checker','Belum Dicek Checker','Sudah Ditindak','Belum Ditindak','PIC Tertinggi','SKU PIC Tertinggi']
};

function rekapInfo(tok) {
  auth_(tok, true); return ss_().getUrl(); }

function bulan_(iso) { return BLN_F[Number(iso.slice(5, 7)) - 1] + ' ' + iso.slice(0, 4); }
function wLabel_(mon) { return 'W' + Math.ceil(Number(mon.slice(8, 10)) / 7); }
function periode_(mon) {
  const a = new Date(mon + 'T00:00:00Z'), b = new Date(a.getTime() + 6 * 864e5);
  const f = d => d.getUTCDate() + ' ' + BLN_S[d.getUTCMonth()];
  return f(a) + ' – ' + f(b) + ' ' + b.getUTCFullYear();
}

/** Rumus persen TANPA koma/titik-koma: aman di semua locale spreadsheet.
 *  (Locale Indonesia memakai ';' sebagai pemisah argumen; rumus ber-koma seperti IF(a,b,c) bisa jadi #ERROR!) */
function pct_(num, den) { return '=' + num + '/(' + den + '+(' + den + '=0))*(' + den + '>0)'; }

function dayRow_(x, r, now) {
  const hari = HARI_F[new Date(x.date + 'T00:00:00Z').getUTCDay()];
  return [x.date, hari, bulan_(x.mon) + ' ' + wLabel_(x.mon),
    x.tMob, (x.tMob ? x.mobWms : 0), x.notMob, pct_('E' + r, 'D' + r),
    x.tSpm, x.spmWms, x.spmCopy, x.spmWms - x.spmCopy,
    x.lines, x.copy, x.ok, pct_('M' + r, 'L' + r), pct_('N' + r, 'L' + r),
    x.checked, Math.max(x.spmWms - x.checked, 0), (x.tMob ? (x.spmWms > x.tMob ? 'Total mobil < SPM WMS, periksa' : 'Lengkap') : 'Total mobil belum diisi'), now, bulan_(x.date), Number(x.mMasuk) || 0];
}

function weekRow_(x, r, now) {
  return [x.mon, bulan_(x.mon), wLabel_(x.mon), periode_(x.mon),
    x.tMob, x.mobWms, x.notMob, pct_('F' + r, 'E' + r),
    x.tSpm, x.spmWms, x.spmCopy, x.spmWms - x.spmCopy,
    x.lines, x.copy, x.ok, pct_('N' + r, 'M' + r), pct_('O' + r, 'M' + r),
    x.checked, x.hari, now];
}

function monthRow_(bulan, r, H) {
  // jumlahkan dari tab Harian (kolom U = Bulan menurut TANGGAL kalender); dihitung di script agar bebas rumus ber-koma
  const rows = (H || []).filter(x => String(x[20]) === String(bulan));
  const sum = i => rows.reduce((s, x) => s + lapNum_(x[i]), 0);
  return [bulan, rows.length,
    sum(3), sum(4), sum(5), pct_('D' + r, 'C' + r),
    sum(7), sum(8), sum(9), sum(10),
    sum(11), sum(12), sum(13), pct_('L' + r, 'K' + r), pct_('M' + r, 'K' + r),
    sum(16)];
}

function writeTable_(ss, name, rows, textCols, pctCols, extraText) {
  const hdr = REKAP_HDR[name];
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#0a6c74').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  const old = sh.getLastRow();
  if (old > 1) sh.getRange(2, 1, old - 1, sh.getMaxColumns()).clear();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, textCols).setNumberFormat('@');   // cegah "Oktober 2026" / "2026-10-05" diubah jadi tanggal
    (extraText || []).forEach(c => sh.getRange(2, c, rows.length, 1).setNumberFormat('@'));
    sh.getRange(2, 1, rows.length, hdr.length).setValues(rows);
    pctCols.forEach(c => sh.getRange(2, c + 1, rows.length, 1).setNumberFormat('0.0%'));
  }
  sh.autoResizeColumns(1, hdr.length);
}

/** KPI SKU selisih: detail per SKU, ringkasan per PIC per bulan, ringkasan bulanan (seluruhnya dibangun ulang dari data aplikasi).
 *  Selisih = Qty Muat − Qty SPM per SKU per No SPM. % memakai RUMUS sheet. */
function saveSelisih_(ss, sel, now) {
  sel = sel || {};
  const rows = (sel.rows || []).slice().sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.spm).localeCompare(String(b.spm)) || String(a.item).localeCompare(String(b.item)));
  writeTable_(ss, 'Selisih_SKU', rows.map((x, i) => {
    const r = i + 2;
    return [x.date, bulan_(x.date), x.spm, x.item, x.desc, x.ship, x.nopol, x.pic, x.chk || 'BELUM DICEK', x.rak,
      x.jenis === 'kurang' ? 'KURANG' : 'LEBIH', x.fu, x.cat, now, Number(x.qs) || 0, Number(x.qm) || 0, '=P' + r + '-O' + r];
  }), 14, []);
  const pic = (sel.pic || []).slice().sort((a, b) => String(a.b).localeCompare(String(b.b)) || (Number(b.n) - Number(a.n)) || String(a.pic).localeCompare(String(b.pic)));
  writeTable_(ss, 'Selisih_PIC', pic.map((x, i) => {
    const r = i + 2, n = Number(x.n) || 0, fu = Number(x.fu) || 0;
    return [bulan_(x.b + '-01'), x.pic, Number(x.lines) || 0, n,
      pct_('D' + r, 'C' + r), fu, Math.max(n - fu, 0), Number(x.sh) || 0];
  }), 2, [4]);
  const bln = (sel.bln || []).slice().sort((a, b) => String(a.b).localeCompare(String(b.b)));
  writeTable_(ss, 'Selisih_Bulanan', bln.map((x, i) => {
    const r = i + 2, n = Number(x.n) || 0, fu = Number(x.fu) || 0;
    return [bulan_(x.b + '-01'), Number(x.hari) || 0, Number(x.lines) || 0, n, pct_('D' + r, 'C' + r),
      Number(x.caught) || 0, Number(x.belum) || 0, fu, Math.max(n - fu, 0), x.top || '-', Number(x.topN) || 0];
  }), 1, [4], [10]);
  log_('Rekap Selisih', 'Selisih_SKU/PIC/Bulanan', rows.length, rows.length + ' SKU selisih, ' + pic.length + ' baris PIC-bulan, ' + bln.length + ' bulan');
  return rows.length;
}

function saveRekap(tok, p) {
  auth_(tok, true);
  if (!p) return { skipped: true };
  if ((!p.days || !p.days.length) && !p.clear) return { skipped: true };
  p.days = p.days || [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
    const days = p.days.slice().sort((a, b) => a.date.localeCompare(b.date));
    const existed = {};   // tanggal yang sudah ada di tab Harian = akan DITIMPA (kunci = tanggal)
    const hs = ss.getSheetByName('Harian');
    if (hs && hs.getLastRow() > 1) hs.getRange(2, 1, hs.getLastRow() - 1, 1).getValues().forEach(r => { const k = dkey_(r[0]); if (k) existed[k] = 1; });
    let upd = 0, ins = 0;
    days.forEach(x => { if (existed[x.date]) upd++; else ins++; });
    const weeks = (p.weeks || []).slice().sort((a, b) => a.mon.localeCompare(b.mon));
    writeTable_(ss, 'Harian', days.map((x, i) => dayRow_(x, i + 2, now)), 3, [6, 14, 15], [21]);
    writeTable_(ss, 'Mingguan', weeks.map((x, i) => weekRow_(x, i + 2, now)), 4, [7, 15, 16]);
    const months = [];
    days.forEach(x => { const b = bulan_(x.date); if (months.indexOf(b) < 0) months.push(b); });
    writeTable_(ss, 'Bulanan', (function (H) { return months.map((b, i) => monthRow_(b, i + 2, H)); })(lapVals_('Harian')), 1, [5, 13, 14]);
    log_('Rekap', 'Harian/Mingguan/Bulanan', days.length, days.length + ' hari, ' + weeks.length + ' minggu, ' + months.length + ' bulan');
    let selN = 0;
    if (p.sel || p.clear) selN = saveSelisih_(ss, p.sel, now);
    lapAuto_();
    return { url: ss.getUrl(), days: days.length, weeks: weeks.length, months: months.length, upd: upd, ins: ins, sel: selN };
  } finally {
    lock.releaseLock();
  }
}


/* ============ REKAP CEK RAK (1 baris per TANGGAL, diperbarui bila dikirim ulang; kolom Petugas = petugas terakhir) ============ */
const RAK_HDR = ['Tanggal','Petugas','Total Rak','Rak Dicek (Item)','Rak Sesuai (Item)','% Kesesuaian Rak','Qty Dicek','Qty Sama','% Qty per Bin','Perlu Adjustment','Belum Dicek (Rak)','Belum Dicek (Qty)','Diupdate','Item Stok Dicek','Item Stok Sama','% Stock Item','Total Stok WMS','Total Stok Fisik'];
/* KPI urut: 1 Kesesuaian Rak = item fisik vs WMS sesuai/tidak (kolom D-F) · 2 Stock Item = stok per item semua bin vs WMS (kolom N-R) · 3 Qty per Bin = qty tiap bin (kolom G-I) */

/** Perbaiki kolom persen di tab Rak untuk SEMUA baris lama (rumus lama ber-koma bisa #ERROR! di locale Indonesia) */
function rakFixPct_() {
  const sh = ss_().getSheetByName('Rak');
  if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1;
  const f = (a, b) => Array.from({ length: n }, (_, i) => [pct_(a + (i + 2), b + (i + 2))]);
  [[6, 'E', 'D'], [9, 'H', 'G'], [16, 'O', 'N']].forEach(c => {
    sh.getRange(2, c[0], n, 1).setValues(f(c[1], c[2]));
    sh.getRange(2, c[0], n, 1).setNumberFormat('0.0%');
  });
}

/** Hitung ulang metrik rak untuk 1 tanggal dari Rak_Detail (null bila belum ada detail) */
function rakAgg_(tgl) {
  const seen = {};   // 1 baris per rak per tanggal (duplikat lama: yang terakhir dipakai), agar tidak dihitung dobel
  lapVals_('Rak_Detail').forEach(r => { if (dkey_(r[0]) === tgl && rkey_(r[1])) seen[rkey_(r[1])] = r; });
  const rows = Object.keys(seen).map(k => seen[k]);
  if (!rows.length) return null;
  const o = { ic: 0, io: 0, qc: 0, qo: 0, adj: 0 };
  rows.forEach(r => {
    const it = String(r[5]).trim(), q = String(r[9]).trim();
    if (it === 'Sesuai' || it === 'Beda') { o.ic++; if (it === 'Sesuai') o.io++; }
    if (q === 'Sama' || q === 'Selisih') { o.qc++; if (q === 'Sama') o.qo++; }
    if (String(r[10]).trim().toUpperCase() === 'YA') o.adj++;
  });
  return o;
}

function saveRak(tok, p) {
  auth_(tok, true);
  if (!p || !p.tgl) return { skipped: true };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    let sh = ss.getSheetByName('Rak');
    if (!sh) {
      sh = ss.insertSheet('Rak');
      sh.getRange(1, 1, 1, RAK_HDR.length).setValues([RAK_HDR]).setFontWeight('bold').setBackground('#0a6c74').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
      sh.setFrozenRows(1);
    }
    const pet = p.petugas || '-';
    const n = sh.getLastRow();
    const hits = [];   // kunci ringkasan = TANGGAL (satu baris per tanggal, simpan ulang = menimpa)
    if (n > 1) {
      const v = sh.getRange(2, 1, n - 1, 1).getValues();
      for (let i = 0; i < v.length; i++) if (dkey_(v[i][0]) === p.tgl) hits.push(i + 2);
    }
    const row = hits.length ? hits[0] : Math.max(n, 1) + 1;
    for (let k = hits.length - 1; k >= 1; k--) sh.deleteRow(hits[k]);   // buang duplikat lama tanggal yang sama
    const r = row;
    /* Ringkasan = akumulasi SEMUA sesi di tanggal ini (dari Rak_Detail), bukan hanya sesi terakhir. Stok item dipertahankan bila sesi ini tidak mengecek stok. */
    const A = rakAgg_(p.tgl);
    const oldS = hits.length ? sh.getRange(r, 14, 1, 5).getValues()[0] : null;
    const keepS = !(Number(p.sc) > 0) && oldS && Number(oldS[0]) > 0;
    p = Object.assign({}, p, A ? { ic: A.ic, io: A.io, qc: A.qc, qo: A.qo, adj: A.adj } : {},
      keepS ? { sc: Number(oldS[0]), sk: Number(oldS[1]), tw: Number(oldS[3]), tp: Number(oldS[4]) } : {});
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
    sh.getRange(r, 1, 1, 2).setNumberFormat('@');
    sh.getRange(r, 1, 1, RAK_HDR.length).setValues([[p.tgl, pet, p.total, p.ic, p.io,
      pct_('E' + r, 'D' + r), p.qc, p.qo, pct_('H' + r, 'G' + r),
      p.adj, Math.max(p.total - p.ic, 0), Math.max(p.total - p.qc, 0), now,
      p.sc || 0, p.sk || 0, pct_('O' + r, 'N' + r), p.tw || 0, p.tp || 0]]);
    [6, 9, 16].forEach(c => sh.getRange(r, c).setNumberFormat('0.0%'));
    sh.getRange(1, 1, 1, RAK_HDR.length).setValues([RAK_HDR]).setFontWeight('bold').setBackground('#0a6c74').setFontColor('#ffffff').setWrap(true);   // pastikan header terbaru
    rakFixPct_();   // baris lama ikut diperbaiki
    sh.autoResizeColumns(1, RAK_HDR.length);
    log_('Rekap Rak', p.tgl, p.total, pet + ': stock item ' + (p.sk || 0) + '/' + (p.sc || 0) + ', rak ' + p.io + '/' + p.ic + ', qty ' + p.qo + '/' + p.qc + ', adj ' + p.adj);
    lapAuto_();
    return { url: ss.getUrl(), row: r, mode: hits.length ? 'update' : 'insert' };
  } finally {
    lock.releaseLock();
  }
}


/* ============ DETAIL SO / KESESUAIAN RAK (kunci = Tanggal + Rak, simpan ulang = menimpa) ============ */
const RD_HDR = ['Tanggal','Rak','Kode Item WMS','Nama Item','Qty WMS','Status Item','Item Aktual / Catatan','Qty Aktual','Selisih Qty','Status Qty','Perlu Adjustment','Petugas','Diupdate'];

/** Teks yang dikirim pengguna: paksa jadi teks agar "=..." / "+..." tidak dijalankan sebagai rumus Sheet */
function txt_(rng) { return rng.setNumberFormat('@'); }

/** kunci rak seragam (abaikan spasi ganda & huruf besar/kecil) agar "a1-01-a" dan "A1-01-A" dianggap rak yang sama */
function rkey_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().toUpperCase(); }

function dkey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}

/** Simpan ulang rak yang sama di tanggal yang sama: hasil yang SUDAH ada (kesesuaian atau qty) tidak boleh hilang
 *  hanya karena sesi berikutnya cuma mengisi salah satunya (mis. pagi cek kesesuaian, siang hitung qty). Yang baru menimpa yang lama bila diisi. */
function rdMerge_(old, row) {
  const done = (v, a, b) => String(v).trim() === a || String(v).trim() === b;
  const o = old.slice(), n = row.slice();
  if (!done(n[5], 'Sesuai', 'Beda') && done(o[5], 'Sesuai', 'Beda')) { n[5] = o[5]; n[6] = o[6]; }
  if (!done(n[9], 'Sama', 'Selisih') && done(o[9], 'Sama', 'Selisih')) { n[7] = o[7]; n[8] = o[8]; n[9] = o[9]; n[10] = o[10]; }
  if (String(n[5]).trim() === 'Beda') n[10] = '-';   // item beda = pindah bin, bukan adjustment
  return n;
}

function saveRakDetail(tok, p) {
  auth_(tok, true);
  if (!p || !p.tgl) return { skipped: true };
  p.rows = p.rows || [];
  const drop = (p.drop || []).map(x => String(x).trim()).filter(Boolean);
  if (!p.rows.length && !drop.length) return { skipped: true };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    let sh = ss.getSheetByName('Rak_Detail');
    if (!sh) {
      sh = ss.insertSheet('Rak_Detail');
      sh.getRange(1, 1, 1, RD_HDR.length).setValues([RD_HDR]).setFontWeight('bold').setBackground('#0a6c74').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
      sh.setFrozenRows(1);
    }
    const n = sh.getLastRow(), W = RD_HDR.length;
    let data = n > 1 ? sh.getRange(2, 1, n - 1, W).getValues() : [];
    let del = 0;
    if (drop.length) {   // rak yang dikeluarkan dari SO: buang baris tanggal ini agar tidak ikut laporan / adjustment
      const before = data.length;
      data = data.filter(r => !(dkey_(r[0]) === p.tgl && drop.map(rkey_).indexOf(rkey_(r[1])) >= 0));
      del = before - data.length;
    }
    const idx = {};
    data.forEach((r, i) => { idx[dkey_(r[0]) + '|' + rkey_(r[1])] = i; });   // bila sudah ada duplikat lama, yang terakhir dipakai
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
    const pet = p.petugas || '-';
    let upd = 0, ins = 0;
    p.rows.forEach(x => {
      let row = [p.tgl, x.rak, x.kode, x.nama, x.wms, x.itemSt, x.itemAkt, x.qtyAkt, x.selisih, x.qtySt, x.adj ? 'YA' : '-', pet, now];
      const k = p.tgl + '|' + rkey_(x.rak);
      if (k in idx) { row = rdMerge_(data[idx[k]], row); data[idx[k]] = row; upd++; }          // tanggal + rak sama: timpa
      else { data.push(row); idx[k] = data.length - 1; ins++; }
    });
    data.sort((a, b) => dkey_(b[0]).localeCompare(dkey_(a[0])) || String(a[1]).localeCompare(String(b[1])));
    [1, 2, 3, 4, 6, 7, 10, 11, 12, 13].forEach(c => txt_(sh.getRange(2, c, data.length, 1)));   // kolom teks (catatan, petugas, dll.) tidak boleh jadi rumus
    if (data.length) sh.getRange(2, 1, data.length, W).setValues(data);
    if (n - 1 > data.length) sh.getRange(data.length + 2, 1, n - 1 - data.length, W).clearContent();   // sisa baris setelah penghapusan
    sh.autoResizeColumns(1, W);
    log_('Rekap Rak Detail', p.tgl, p.rows.length, pet + ': ' + upd + ' ditimpa, ' + ins + ' baru, ' + del + ' dihapus');
    return { url: ss.getUrl(), upd: upd, ins: ins, del: del, total: p.rows.length };
  } finally {
    lock.releaseLock();
  }
}

/** Status penyimpanan per tanggal (untuk penanda "tersimpan" di aplikasi dan dashboard) */
function sheetStatus(tok) {
  auth_(tok);
  const ss = ss_(), out = { muat: {}, rak: {} };
  const rd = (name, cd, cu, tgt) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const n = sh.getLastRow();
    if (n < 2) return;
    const d = sh.getRange(2, cd, n - 1, 1).getValues(), u = sh.getRange(2, cu, n - 1, 1).getValues();
    d.forEach((r, i) => {
      const k = dkey_(r[0]);
      if (k) tgt[k] = u[i][0] instanceof Date ? Utilities.formatDate(u[i][0], Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm') : String(u[i][0]);
    });
  };
  rd('Harian', 1, 20, out.muat);
  rd('Rak_Detail', 1, 13, out.rak);
  return out;
}

/** URL web app (untuk link dashboard manager: URL + ?v=dash) */
function appUrl(tok) { auth_(tok, true); return ScriptApp.getService().getUrl(); }


/* ============ SUMBER TOTAL MOBIL dari spreadsheet lain ============
 * DUA MODE:
 *
 * (A) MODE SATU SEL — angka total mobil ada di 1 sel (mis. DASHBOARD NEW!N43):
 *     cfg = {url, sheet, cell, cellTgl?, cellSpm?}
 *       cell    : alamat sel total mobil, mis. 'N43'
 *       cellTgl : (opsional) alamat sel yang berisi TANGGAL angka tsb, mis. 'N40'.
 *                 Jika kosong, angka dipetakan ke tanggal yang diminta (harus tepat 1 tanggal).
 *       cellSpm : (opsional) alamat sel total SPM.
 *
 * (B) MODE TABEL — 1 baris per tanggal:
 *     cfg = {url, sheet, headRow, colTgl, colMob, modeMob, colSpm, modeSpm}
 *     Kolom boleh huruf (A, C, AA) atau nama judul kolom. Mode: sum | count | unique.
 *
 * Dijalankan dengan akun pemilik script, jadi akun itu harus punya akses ke sheet sumber. */
/* Daftar spreadsheet sumber yang boleh dibaca ada di tab "Config" (dibuat otomatis di spreadsheet database).
 * Kolom C "Izinkan" harus YA. Bila sumber baru diminta tapi belum terdaftar, barisnya otomatis ditambahkan
 * dengan status TIDAK; admin tinggal mengubahnya jadi YA. Script property SRC_ALLOW tetap berlaku (opsional). */
const CFG_SHEET = 'Config';
const CFG_HDR = ['ID / Link spreadsheet sumber', 'Catatan', 'Izinkan (YA / TIDAK)', 'Terakhir diminta'];

function cfgSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(CFG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CFG_SHEET);
    sh.getRange(1, 1, 1, CFG_HDR.length).setValues([CFG_HDR]).setFontWeight('bold').setBackground('#0a6c74').setFontColor('#ffffff').setWrap(true);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 380); sh.setColumnWidth(2, 220); sh.setColumnWidth(3, 150); sh.setColumnWidth(4, 140);
    sh.getRange(2, 3, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['YA', 'TIDAK'], true).build());
  }
  return sh;
}

function srcAllow_(id) {
  if (props_('SRC_ALLOW').map(srcId_).indexOf(id) >= 0) return;
  try {
    const sh = cfgSheet_(), n = sh.getLastRow();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
    let found = false;
    if (n >= 2) {
      const v = sh.getRange(2, 1, n - 1, 4).getValues();
      for (let i = 0; i < v.length; i++) {
        if (!String(v[i][0]).trim() || srcId_(v[i][0]) !== id) continue;
        if (String(v[i][2]).trim().toUpperCase() === 'YA') return;   // sudah diizinkan
        sh.getRange(i + 2, 4).setValue(now); found = true; break;
      }
    }
    if (!found && n < 200) sh.appendRow([id, '', 'TIDAK', now]);   // batas 200 baris agar Config tidak dipenuhi ID sembarang
  } catch (e) {}
  throw new Error('Spreadsheet sumber belum diizinkan. Buka tab "Config" di spreadsheet database WMS, ubah kolom Izinkan menjadi YA untuk ID ' + id + ', lalu coba lagi.');
}

function srcOpen_(cfg) {
  if (!cfg || !cfg.url) throw new Error('Link spreadsheet sumber belum diisi');
  const u = String(cfg.url).trim();
  srcAllow_(srcId_(u));
  return /^https?:/i.test(u) ? SpreadsheetApp.openByUrl(u) : SpreadsheetApp.openById(u);
}

/** Baca 1 sel. Mengembalikan nilai mentah (angka / tanggal / teks). */
function srcCellRaw_(sh, addr, label) {
  const a1 = String(addr || '').trim().toUpperCase();
  if (!/^[A-Z]{1,2}[0-9]{1,6}$/.test(a1)) throw new Error('Alamat sel ' + label + ' tidak valid: "' + addr + '" (contoh: N43)');
  return sh.getRange(a1).getValue();
}

function srcNum_(v, label) {
  if (typeof v === 'number') return v;
  const t = String(v).trim();
  if (t === '') return 0;
  const n = parseFloat(/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(t) ? t.replace(/\./g, '').replace(',', '.') : t.replace(',', '.'));
  if (isNaN(n)) throw new Error('Isi sel ' + label + ' bukan angka: "' + t + '"');
  return n;
}

/** MODE SATU SEL */
function srcCell_(cfg, dates) {
  const ss = srcOpen_(cfg);
  const sh = ss.getSheetByName(cfg.sheet);
  if (!sh) throw new Error('Tab "' + cfg.sheet + '" tidak ada');
  let d = '';
  if (cfg.cellTgl) {
    d = srcDate_(srcCellRaw_(sh, cfg.cellTgl, 'tanggal'));
    if (!d) throw new Error('Sel tanggal ' + cfg.cellTgl + ' tidak terbaca sebagai tanggal');
  } else if (dates && dates.length === 1) {
    d = dates[0];
  } else {
    throw new Error('Mode satu sel hanya bisa untuk 1 tanggal. Pilih 1 tanggal, atau isi cellTgl (sel yang berisi tanggal).');
  }
  if (dates && dates.length && dates.indexOf(d) < 0) return {};   // tanggal di sheet sumber bukan yang diminta
  const mob = srcNum_(srcCellRaw_(sh, cfg.cell, 'total mobil'), cfg.cell);
  const spm = cfg.cellSpm ? srcNum_(srcCellRaw_(sh, cfg.cellSpm, 'total SPM'), cfg.cellSpm) : null;
  const out = {};
  const masuk = cfg.cellMasuk ? srcNum_(srcCellRaw_(sh, cfg.cellMasuk, 'mobil masuk'), cfg.cellMasuk) : null;
  out[d] = { mob: mob, spm: spm, masuk: masuk };
  return out;
}

function srcInfo(tok, cfg) {
  auth_(tok, true);
  const ss = srcOpen_(cfg);
  const names = ss.getSheets().map(s => s.getName());
  const sh = cfg.sheet ? ss.getSheetByName(cfg.sheet) : null;
  if (cfg.cell) {   // mode satu sel: tampilkan isi sel agar bisa dicek
    let val = '';
    if (sh) val = String(srcCellRaw_(sh, cfg.cell, 'total mobil'));
    return { title: ss.getName(), names: names, headers: [], found: !!sh, cellValue: val };
  }
  let headers = [];
  if (sh && sh.getLastColumn() > 0) headers = sh.getRange(Math.max(1, Number(cfg.headRow) || 1), 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  return { title: ss.getName(), names: names, headers: headers, found: !!sh };
}

function srcCol_(spec, headers) {
  const s = String(spec || '').trim();
  if (!s) return -1;
  const byName = headers.findIndex(h => String(h).trim().toLowerCase() === s.toLowerCase());
  if (byName >= 0) return byName;
  if (/^[A-Za-z]{1,2}$/.test(s)) { let n = 0; s.toUpperCase().split('').forEach(c => n = n * 26 + c.charCodeAt(0) - 64); return n - 1; }
  throw new Error('Kolom "' + s + '" tidak ditemukan');
}

function srcDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

/** dates: array 'yyyy-MM-dd'. Hasil: { 'yyyy-MM-dd': {mob, spm|null} } hanya untuk tanggal yang ditemukan */
function srcTotals(tok, cfg, dates) {
  auth_(tok, true);
  if (cfg && cfg.cell) return srcCell_(cfg, dates);   // MODE SATU SEL
  const ss = srcOpen_(cfg);
  const sh = ss.getSheetByName(cfg.sheet);
  if (!sh) throw new Error('Tab "' + cfg.sheet + '" tidak ada');
  const hr = Math.max(1, Number(cfg.headRow) || 1), last = sh.getLastRow(), lc = sh.getLastColumn();
  const out = {};
  if (last <= hr) return out;
  const headers = sh.getRange(hr, 1, 1, lc).getDisplayValues()[0];
  const ci = { tgl: srcCol_(cfg.colTgl, headers), mob: srcCol_(cfg.colMob, headers), spm: srcCol_(cfg.colSpm, headers), masuk: srcCol_(cfg.colMasuk, headers) };
  if (ci.tgl < 0 || ci.mob < 0) throw new Error('Kolom tanggal / mobil belum diisi');
  const modes = { mob: cfg.modeMob || 'sum', spm: cfg.modeSpm || 'sum', masuk: cfg.modeMasuk || 'sum' };
  const want = {}; (dates || []).forEach(d => want[d] = 1);
  const acc = {};
  sh.getRange(hr + 1, 1, last - hr, lc).getValues().forEach(r => {
    const d = srcDate_(r[ci.tgl]);
    if (!want[d]) return;
    const a = acc[d] = acc[d] || { mob: 0, mobSet: {}, spm: 0, spmSet: {}, masuk: 0, masukSet: {} };
    ['mob', 'spm', 'masuk'].forEach(k => {
      if (ci[k] < 0) return;
      const v = r[ci[k]], t = String(v).trim();
      if (modes[k] === 'sum') a[k] += parseFloat(t.replace(',', '.')) || 0;
      else if (modes[k] === 'unique') { if (t) a[k + 'Set'][t.toUpperCase().replace(/\s+/g, '')] = 1; }
      else if (t !== '') a[k] += 1;
    });
  });
  Object.keys(acc).forEach(d => {
    const a = acc[d], f = k => modes[k] === 'unique' ? Object.keys(a[k + 'Set']).length : a[k];
    out[d] = { mob: f('mob'), spm: ci.spm < 0 ? null : f('spm'), masuk: ci.masuk < 0 ? null : f('masuk') };
  });
  return out;
}

/** TES dari editor Apps Script: pilih fungsi testSrc lalu Run, hasil di Execution log.
 *  Ganti tanggal di bawah dengan tanggal yang ingin diuji. */
function srcDefault_() { const id = (PropertiesService.getScriptProperties().getProperty('SRC_ID') || '').trim(); if (!id) throw new Error('Isi Script property SRC_ID (ID spreadsheet sumber) dulu'); return srcId_(id); }
function testSrc() {
  const cfg = {
    url: srcDefault_(),
    sheet: 'DASHBOARD NEW',
    cell: 'N43'
    // cellTgl: 'N40',   // aktifkan bila ada sel yang berisi tanggal
    // cellSpm: 'N44'    // aktifkan bila total SPM ada di 1 sel
  };
  Logger.log(JSON.stringify(srcInfo(INTERNAL_, cfg)));                    // cellValue = isi N43
  Logger.log(JSON.stringify(srcTotals(INTERNAL_, cfg, ['2026-10-12'])));  // hasil untuk 1 tanggal
}

/** Jalankan SEKALI dari editor (pilih "daftarSumber" lalu Run): membuat tab Config dan mengizinkan sheet sumber DASHBOARD NEW. */
function daftarSumber() {
  const id = srcDefault_();
  const sh = cfgSheet_(), n = sh.getLastRow();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
  let row = 0;
  if (n >= 2) {
    const v = sh.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() && srcId_(v[i][0]) === id) { row = i + 2; break; }
  }
  if (row) { sh.getRange(row, 3).setValue('YA'); sh.getRange(row, 4).setValue(now); }
  else sh.appendRow([id, 'DASHBOARD NEW (total mobil N43)', 'YA', now]);
  Logger.log('Sumber diizinkan: ' + id);
}


/* =====================================================================
 * LAPORAN MINGGUAN & BULANAN OTOMATIS  (tab: Laporan_Mingguan, Laporan_Bulanan, Adjustment_List)
 * ---------------------------------------------------------------------
 * Dibangun ulang OTOMATIS setiap kali aplikasi menyimpan (Rekap Muat / Simpan SO Rak).
 * Tidak perlu input atau kumpul data manual: tinggal buka tab Laporan lalu setor.
 * Sumber: tab Harian (muat), Selisih_SKU + Selisih_Bulanan (selisih SKU),
 *         Rak (kesesuaian rak per tanggal), Rak_Detail (adjustment per rak).
 * Adjustment MUAT = Qty Muat − Qty Checker Out per SKU (acuan = hasil validasi checker, kolom "Qty SPM" di laporan WMS): muat 50 vs checker 25 = +25 (kelebihan muat, di-adjust), kebalikannya minus. Digabung dengan adjustment rak = TOTAL NET ADJUSTMENT.
 * Aturan periode: Mingguan = Senin–Minggu (label W1..W5 mengikuti tgl Senin, sama dengan tab Mingguan).
 *                 Bulanan  = bulan kalender dari TANGGAL data.
 * Warna angka %: hijau >= 95%, kuning 85–<95%, merah < 85% (ubah di LAP_HI / LAP_MID).
 * Manual: menu "Laporan WMS" > "Bangun ulang laporan" (atau jalankan buildLaporan dari editor).
 * ===================================================================== */
const LAP_HI = 0.95, LAP_MID = 0.85;
const LAP_GRP = {
  id:   { t: 'PERIODE',                         dark: '#334155', light: '#e2e8f0' },
  muat: { t: 'MUAT (PROSES WMS)',               dark: '#0a6c74', light: '#d5eef0' },
  sel:  { t: 'SELISIH SKU (QTY MUAT ≠ CHECKER OUT)',    dark: '#b45309', light: '#fdebd3' },
  rak:  { t: 'SO RAK: KESESUAIAN RAK · STOCK ITEM · QTY PER BIN',         dark: '#1d4ed8', light: '#dbe7fe' },
  adj:  { t: 'TOTAL ADJUSTMENT',                dark: '#b91c1c', light: '#fde0e0' }
};
/* t: n = angka dijumlah dari data | f = rumus (fm) | fmt: i = angka, p = persen | cf = warnai persen | m = hanya bulanan */
const LAP_COLS = [
  { g: 'id', k: 'days', h: 'Hari Muat' }, { g: 'id', k: 'rakDays', h: 'Hari Cek Rak' },

  { g: 'muat', k: 'tMob', h: 'Total Mobil' }, { g: 'muat', k: 'mMasuk', h: 'Mobil Masuk' }, { g: 'muat', k: 'mobWms', h: 'Mobil WMS' }, { g: 'muat', k: 'notMob', h: 'Mobil Tidak WMS' },
  { g: 'muat', k: 'pMob', h: '% Mobil WMS', fm: '={mobWms}/({tMob}+({tMob}=0))*({tMob}>0)', fmt: 'p', cf: 1 },
  { g: 'muat', k: 'tSpm', h: 'Total SPM' }, { g: 'muat', k: 'spmWms', h: 'SPM WMS' }, { g: 'muat', k: 'spmCopy', h: 'SPM Copy Rekom.' },
  { g: 'muat', k: 'spmOk', h: 'SPM Full OK', fm: '={spmWms}-{spmCopy}' },
  { g: 'muat', k: 'lines', h: 'Total Item' }, { g: 'muat', k: 'copy', h: 'Item Copy Rekom.' }, { g: 'muat', k: 'ok', h: 'Item OK (Aktual)' },
  { g: 'muat', k: 'pCopy', h: '% Copy Rekom.', fm: '={copy}/({lines}+({lines}=0))*({lines}>0)', fmt: 'p' },
  { g: 'muat', k: 'pOk', h: '% Item OK', fm: '={ok}/({lines}+({lines}=0))*({lines}>0)', fmt: 'p', cf: 1 },
  { g: 'muat', k: 'checked', h: 'SPM Sudah Dicek' }, { g: 'muat', k: 'belum', h: 'SPM Belum Dicek' },

  { g: 'sel', k: 'selLines', h: 'SKU Dimuat', m: 1 },
  { g: 'sel', k: 'selN', h: 'SKU Selisih' },
  { g: 'sel', k: 'pSel', h: '% SKU Selisih', fm: '={selN}/({selLines}+({selLines}=0))*({selLines}>0)', fmt: 'p', m: 1 },
  { g: 'sel', k: 'selK', h: 'Muat Kurang' }, { g: 'sel', k: 'selL', h: 'Muat Lebih' }, { g: 'sel', k: 'selQ', h: 'Total Selisih Qty (abs)' },
  { g: 'sel', k: 'selFu', h: 'Sudah Ditindak' }, { g: 'sel', k: 'selBlm', h: 'Belum Ditindak', fm: '={selN}-{selFu}' },
  { g: 'sel', k: 'pFu', h: '% Ditindak', fm: '={selFu}/({selN}+({selN}=0))*({selN}>0)', fmt: 'p', cf: 1 },

  /* urutan KPI rak: 1 Kesesuaian Rak (item fisik vs WMS) · 2 Stock Item (stok per item) · 3 Qty per Bin */
  { g: 'rak', k: 'ic', h: 'Rak Dicek (Item)' }, { g: 'rak', k: 'io', h: 'Rak Sesuai (Item)' },
  { g: 'rak', k: 'pRak', h: '% Kesesuaian Rak', fm: '={io}/({ic}+({ic}=0))*({ic}>0)', fmt: 'p', cf: 1 },
  { g: 'rak', k: 'sc', h: 'Item Stok Dicek' }, { g: 'rak', k: 'sk', h: 'Item Stok Sama' },
  { g: 'rak', k: 'pSt', h: '% Stock Item', fm: '={sk}/({sc}+({sc}=0))*({sc}>0)', fmt: 'p', cf: 1 },
  { g: 'rak', k: 'tw', h: 'Total Stok WMS' }, { g: 'rak', k: 'tp', h: 'Total Stok Fisik' },
  { g: 'rak', k: 'dSt', h: 'Selisih Stok (Fisik − WMS)', fm: '={tp}-{tw}' },
  { g: 'rak', k: 'qc', h: 'Bin Qty Dicek' }, { g: 'rak', k: 'qo', h: 'Bin Qty Sama' },
  { g: 'rak', k: 'pQty', h: '% Qty per Bin', fm: '={qo}/({qc}+({qc}=0))*({qc}>0)', fmt: 'p', cf: 1 },

  { g: 'adj', k: 'adjRak', h: 'Item Perlu Adjustment' },
  { g: 'adj', k: 'pAdj', h: '% Perlu Adjustment', fm: '={adjRak}/({ic}+({ic}=0))*({ic}>0)', fmt: 'p' },
  { g: 'adj', k: 'adjD', h: 'Baris Adj (Detail Rak)' }, { g: 'adj', k: 'qSel', h: 'Baris Qty Selisih' },
  { g: 'adj', k: 'adjNeg', h: 'Adj Rak (−)' }, { g: 'adj', k: 'adjPos', h: 'Adj Rak (+)' },
  { g: 'adj', k: 'adjNet', h: 'Net Adj Rak', fm: '={adjNeg}+{adjPos}' },
  /* Adjustment dari selisih MUAT: Qty Muat − Qty Checker Out (acuan = validasi checker). Muat > Checker = (+), Muat < Checker = (−). Dilakukan agar stok balance. */
  { g: 'adj', k: 'mPos', h: 'Adj Muat (+) Muat > Checker' }, { g: 'adj', k: 'mNeg', h: 'Adj Muat (−) Muat < Checker' },
  { g: 'adj', k: 'mNet', h: 'Net Adj Muat', fm: '={mPos}+{mNeg}' },
  { g: 'adj', k: 'tNet', h: 'TOTAL NET ADJUSTMENT (Rak + Muat)', fm: '={adjNet}+{mNet}' }
];

function lapColL_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function lapNum_(v) { const n = Number(String(v == null ? '' : v).replace(',', '.')); return isNaN(n) ? 0 : n; }
function lapVals_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function lapMon_(iso) { const d = new Date(iso + 'T00:00:00Z'); return new Date(d.getTime() - ((d.getUTCDay() + 6) % 7) * 864e5).toISOString().slice(0, 10); }
function lapBlank_() {
  const o = {}; LAP_COLS.forEach(c => { if (!c.fm) o[c.k] = 0; }); return o;
}
const LAP_ISO = /^\d{4}-\d{2}-\d{2}$/;
/* AREA SO YANG DIPAKAI SEKARANG (fokus dulu): rak di luar daftar ini diabaikan di report Kesesuaian & Belum_Dicek. Tambah area di sini bila sudah siap; kosongkan [] = semua area. Samakan dengan SCOPE di Index.html. */
const RAK_SCOPE = ['A1', 'A2', 'B1', 'B2', 'C1', 'C2', 'D1', 'D2'];
function inScope_(rak) { return !RAK_SCOPE.length || RAK_SCOPE.indexOf(String(rak).split('-')[0].trim().toUpperCase()) >= 0; }

/** Kumpulkan semua tab data menjadi agregat per minggu (W) dan per bulan (B) */
function lapAgg_() {
  const W = {}, B = {};
  const at = iso => {
    const mk = lapMon_(iso), bk = iso.slice(0, 7);
    return [W[mk] || (W[mk] = lapBlank_()), B[bk] || (B[bk] = lapBlank_())];
  };
  lapVals_('Harian').forEach(r => {
    const d = dkey_(r[0]); if (!LAP_ISO.test(d)) return;
    at(d).forEach(o => {
      o.days++; o.tMob += lapNum_(r[3]); o.mobWms += lapNum_(r[4]); o.notMob += lapNum_(r[5]);
      o.tSpm += lapNum_(r[7]); o.spmWms += lapNum_(r[8]); o.spmCopy += lapNum_(r[9]);
      o.lines += lapNum_(r[11]); o.copy += lapNum_(r[12]); o.ok += lapNum_(r[13]);
      o.checked += lapNum_(r[16]); o.belum += lapNum_(r[17]); o.mMasuk += lapNum_(r[21]);
    });
  });
  lapVals_('Rak').forEach(r => {
    const d = dkey_(r[0]); if (!LAP_ISO.test(d)) return;
    at(d).forEach(o => {
      o.rakDays++; o.ic += lapNum_(r[3]); o.io += lapNum_(r[4]); o.qc += lapNum_(r[6]); o.qo += lapNum_(r[7]);
      o.adjRak += lapNum_(r[9]); o.sc += lapNum_(r[13]); o.sk += lapNum_(r[14]); o.tw += lapNum_(r[16]); o.tp += lapNum_(r[17]);
    });
  });
  lapVals_('Rak_Detail').forEach(r => {
    const d = dkey_(r[0]); if (!LAP_ISO.test(d)) return;
    const sel = lapNum_(r[8]), isAdj = String(r[10]).trim().toUpperCase() === 'YA', isSel = String(r[9]).trim() === 'Selisih';
    at(d).forEach(o => {
      if (isSel) o.qSel++;
      if (isAdj) { o.adjD++; if (sel < 0) o.adjNeg += sel; else o.adjPos += sel; }
    });
  });
  lapVals_('Selisih_SKU').forEach(r => {
    const d = dkey_(r[0]); if (!LAP_ISO.test(d)) return;
    const qs = lapNum_(r[14]), qm = lapNum_(r[15]);
    at(d).forEach(o => {
      o.selN++; if (String(r[10]).toUpperCase() === 'KURANG') o.selK++; else o.selL++;
      o.selQ += Math.abs(qm - qs); if (qm - qs > 0) o.mPos += qm - qs; else o.mNeg += qm - qs;
      if (String(r[11]).toUpperCase() === 'SUDAH') o.selFu++;
    });
  });
  const lines = {};   // SKU Dimuat per bulan (dari Selisih_Bulanan; hanya tersedia per bulan)
  lapVals_('Selisih_Bulanan').forEach(r => { lines[String(r[0]).trim()] = lapNum_(r[2]); });
  Object.keys(B).forEach(k => { B[k].selLines = lines[bulan_(k + '-01')] || 0; });
  return { W: W, B: B };
}

/** Tulis 1 tab laporan. rows = [{ids:[...], o:{...}}] sudah urut tampil (terbaru di atas) */
function lapWrite_(name, idHdr, rows, monthly) {
  const ss = ss_();
  let sh = ss.getSheetByName(name), created = false;
  if (!sh) { sh = ss.insertSheet(name); created = true; }
  const old = sh.getFilter(); if (old) old.remove();
  try { sh.getRange(1, 1, 2, sh.getMaxColumns()).breakApart(); } catch (e) {}   // bersihkan merge dari versi lama
  sh.clear(); sh.setConditionalFormatRules([]);
  const cols = LAP_COLS.filter(c => !c.m || monthly);
  const nId = idHdr.length, nC = nId + cols.length, nR = rows.length;
  const L = {}; cols.forEach((c, i) => { L[c.k] = lapColL_(nId + i + 1); });
  const fill = (fm, r) => fm.replace(/\{(\w+)\}/g, (m, k) => L[k] + r);
  const first = 4, lastR = first + nR - 1;

  const hdr = idHdr.concat(cols.map(c => c.h));
  const grp = idHdr.map(() => 'id').concat(cols.map(c => c.g));
  const grow = grp.map((g, i) => (i === 0 || grp[i - 1] !== g) ? LAP_GRP[g].t : '');
  const total = ['TOTAL'].concat(idHdr.slice(1).map(() => '')).concat(cols.map(c => {
    if (c.fm) return fill(c.fm, 3);
    const l = L[c.k]; return nR ? '=SUM(' + l + first + ':' + l + lastR + ')' : 0;
  }));
  const body = rows.map((x, i) => x.ids.concat(cols.map(c => c.fm ? fill(c.fm, first + i) : x.o[c.k])));

  const need = Math.max(nR + 3, 3); if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (sh.getMaxColumns() < nC) sh.insertColumnsAfter(sh.getMaxColumns(), nC - sh.getMaxColumns());
  sh.getRange(1, 1, 3 + nR, nId).setNumberFormat('@');
  sh.getRange(1, 1, 1, nC).setValues([grow]);
  sh.getRange(2, 1, 1, nC).setValues([hdr]);
  sh.getRange(3, 1, 1, nC).setValues([total]);
  if (nR) sh.getRange(first, 1, nR, nC).setValues(body);

  // format angka
  cols.forEach((c, i) => { sh.getRange(3, nId + i + 1, nR + 1, 1).setNumberFormat(c.fmt === 'p' ? '0.0%' : '#,##0'); });
  // pita grup + header
  let a = 0;
  for (let i = 1; i <= nC; i++) {
    if (i === nC || grp[i] !== grp[a]) {
      const g = LAP_GRP[grp[a]], w = i - a;
      sh.getRange(1, a + 1, 1, w).setBackground(g.dark).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('left').setWrap(false).setFontSize(11);   // TANPA merge: merge yang melintasi kolom beku ditolak Sheets
      sh.getRange(2, a + 1, 1, w).setBackground(g.light).setFontColor('#0f172a').setFontWeight('bold');
      sh.getRange(1, a + 1, nR + 3, 1).setBorder(null, true, null, null, null, null, g.dark, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      a = i;
    }
  }
  sh.getRange(2, 1, 1, nC).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(2, 52);
  sh.getRange(3, 1, 1, nC).setFontWeight('bold').setBackground('#f1f5f9').setBorder(true, null, true, null, null, null, '#334155', SpreadsheetApp.BorderStyle.SOLID);
  if (nR) {
    sh.getRange(first, 1, nR, nC).setBorder(null, null, true, null, null, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(first, nId + 1, nR, nC - nId).setHorizontalAlignment('right');
  }
  sh.getRange(3, nId + 1, nR + 1, nC - nId).setHorizontalAlignment('right');
  sh.setColumnWidths(1, nId, monthly ? 130 : 150); sh.setColumnWidths(nId + 1, nC - nId, 92);
  sh.setFrozenRows(3); sh.setFrozenColumns(Math.min(nId, 2));

  // warna % (hijau/kuning/merah); 0 = belum ada data, tidak diwarnai
  const rules = [];
  cols.forEach((c, i) => {
    if (!c.cf) return;
    const rg = sh.getRange(3, nId + i + 1, nR + 1, 1), x = lapColL_(nId + i + 1) + '3', pm = Math.round(LAP_MID * 100), ph = Math.round(LAP_HI * 100), x100 = x + '*100';
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=(' + x + '>0)*(' + x100 + '<' + pm + ')').setBackground('#fecaca').setFontColor('#991b1b').setRanges([rg]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=(' + x100 + '>=' + pm + ')*(' + x100 + '<' + ph + ')').setBackground('#fef08a').setFontColor('#854d0e').setRanges([rg]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=' + x100 + '>=' + ph).setBackground('#bbf7d0').setFontColor('#166534').setRanges([rg]).build());
  });
  sh.setConditionalFormatRules(rules);
  if (created) { try { ss.setActiveSheet(sh); ss.moveActiveSheet(1); } catch (e) {} }
  return sh;
}

/** Daftar adjustment (rak bertanda YA + semua selisih muat) dengan label minggu/bulan, siap difilter untuk disetor.
 *  Kolom Adjustment: (+) / (−). Muat = Qty Muat − Qty SPM. Rak = kolom Selisih Qty di Rak_Detail. */
function lapAdjList_() {
  const ss = ss_(), name = 'Adjustment_List';
  const hdr = ['Tanggal', 'Minggu', 'Bulan', 'Sumber', 'Rak (asal ambil)', 'No SPM', 'Kode Item / SKU', 'Nama Item', 'Qty Acuan (WMS / Checker Out)', 'Qty Aktual (Fisik / Muat)', 'Adjustment (+/−)', 'Petugas / PIC Muat', 'Status Tindak Lanjut'];
  const rows = [];
  const lab = d => { const mk = lapMon_(d); return [wLabel_(mk) + ' · ' + periode_(mk), bulan_(d)]; };
  lapVals_('Rak_Detail').forEach(r => {
    const d = dkey_(r[0]);
    if (!LAP_ISO.test(d) || String(r[10]).trim().toUpperCase() !== 'YA') return;
    rows.push([d].concat(lab(d), ['RAK', String(r[1]), '', String(r[2]), String(r[3]), lapNum_(r[4]), r[7] === '' ? '' : lapNum_(r[7]), lapNum_(r[8]), String(r[11]), '']));
  });
  lapVals_('Selisih_SKU').forEach(r => {
    const d = dkey_(r[0]); if (!LAP_ISO.test(d)) return;
    const qs = lapNum_(r[14]), qm = lapNum_(r[15]); if (qm === qs) return;
    rows.push([d].concat(lab(d), ['MUAT', String(r[9]), String(r[2]), String(r[3]), String(r[4]), qs, qm, qm - qs, String(r[7]), String(r[11])]));
  });
  rows.sort((a, b) => b[0].localeCompare(a[0]) || a[3].localeCompare(b[3]) || a[4].localeCompare(b[4]));
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const f = sh.getFilter(); if (f) f.remove();
  sh.clear();
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#b91c1c').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 8).setNumberFormat('@');
    sh.getRange(2, 12, rows.length, 2).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, hdr.length).setValues(rows);
    sh.getRange(2, 11, rows.length, 1).setNumberFormat('+#,##0;-#,##0;0').setFontWeight('bold');
    sh.getRange(1, 1, rows.length + 1, hdr.length).createFilter();
  }
  sh.autoResizeColumns(1, hdr.length);
  lapAdjRak_(rows);
  return rows.length;
}

/** Adjustment per RAK per bulan: gabungan selisih SO rak + selisih muat (muat membawa kolom Rak = asal ambil barang).
 *  Catatan: bila 1 baris muat diambil dari beberapa rak (mis. "A1, B2"), qty ditampilkan pada gabungan rak tsb apa adanya. */
function lapAdjRak_(rows) {
  const ss = ss_(), name = 'Adjustment_per_Rak';
  const hdr = ['Bulan', 'Rak', 'Adj SO Rak (fisik − WMS)', 'Adj Muat (+) Muat > Checker', 'Adj Muat (−) Muat < Checker', 'Net Adj Muat', 'Net Total (SO + Muat)', 'Kejadian SO', 'Kejadian Muat'];
  const m = {};
  rows.forEach(r => {
    const k = r[2] + '|' + (r[4] || '(tanpa rak)');
    const o = m[k] || (m[k] = { b: r[2], d: r[0], rak: r[4] || '(tanpa rak)', so: 0, mp: 0, mn: 0, ns: 0, nm: 0 });
    const v = lapNum_(r[10]);
    if (r[3] === 'RAK') { o.so += v; o.ns++; } else { if (v > 0) o.mp += v; else o.mn += v; o.nm++; }
  });
  const list = Object.keys(m).map(k => m[k]).sort((a, b) => b.d.slice(0, 7).localeCompare(a.d.slice(0, 7)) || Math.abs(b.so + b.mp + b.mn) - Math.abs(a.so + a.mp + a.mn) || a.rak.localeCompare(b.rak));
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const f = sh.getFilter(); if (f) f.remove();
  sh.clear();
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#b91c1c').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  if (list.length) {
    const vals = list.map((o, i) => [o.b, o.rak, o.so, o.mp, o.mn, '=D' + (i + 2) + '+E' + (i + 2), '=C' + (i + 2) + '+F' + (i + 2), o.ns, o.nm]);
    sh.getRange(2, 1, list.length, 2).setNumberFormat('@');
    sh.getRange(2, 1, list.length, hdr.length).setValues(vals);
    sh.getRange(2, 3, list.length, 5).setNumberFormat('+#,##0;-#,##0;0');
    sh.getRange(2, 7, list.length, 1).setFontWeight('bold');
    sh.getRange(1, 1, list.length + 1, hdr.length).createFilter();
  }
  sh.autoResizeColumns(1, hdr.length);
}

/* =====================================================================
 * REPORT KESESUAIAN RAK  (tab: Kesesuaian_Mingguan, Kesesuaian_Bulanan)
 * ---------------------------------------------------------------------
 * Sumber: Rak_Detail (hasil SO rak per TANGGAL SO). Dibangun ulang otomatis bersama laporan lain.
 * Mingguan = Senin–Minggu menurut TANGGAL SO (label W1..W5 mengikuti tgl Senin). Bulanan = bulan kalender tgl SO.
 * Per periode: baris per BLOK (bagian rak sebelum "-", mis. A1 dari A1-01-A), subtotal per ZONA (huruf pertama, mis. A)
 * bila zona punya >1 blok, lalu TOTAL PERIODE.
 *  Total Rak     = jumlah rak (lokasi unik) yang dicek di periode itu (rak yang dicek >1x di periode sama dihitung 1x, hasil SO TERAKHIR yang dipakai)
 *  Sesuai        = rak dengan Status Item "Sesuai"          Tidak Sesuai = Total − Sesuai
 *  % Kesesuaian  = Sesuai / Total Rak (rumus sheet)         Perlu Adjustment = rak bertanda YA     Net Adj Qty = jumlah Selisih Qty rak tsb
 * ===================================================================== */
function lapKesAgg_(rows) {
  const list = [];
  rows.forEach(r => {
    const d = dkey_(r[0]), rak = rkey_(r[1]), it = String(r[5]).trim();
    if (!LAP_ISO.test(d) || !rak || !inScope_(rak) || (it !== 'Sesuai' && it !== 'Beda')) return;   // hanya rak yang item-nya benar-benar dicek
    list.push({ d: d, rak: rak, ok: it === 'Sesuai', adj: String(r[10]).trim().toUpperCase() === 'YA', sel: lapNum_(r[8]) });
  });
  list.sort((a, b) => a.d.localeCompare(b.d));   // lama → baru: SO terakhir menimpa
  const W = {}, B = {};
  list.forEach(x => {
    const mk = lapMon_(x.d), bk = x.d.slice(0, 7);
    (W[mk] || (W[mk] = {}))[x.rak] = x;
    (B[bk] || (B[bk] = {}))[x.rak] = x;
  });
  return { W: W, B: B };
}
function lapRakBlok_(rak) { const b = String(rak).split('-')[0].trim(); return b || '(tanpa rak)'; }
function lapRakZona_(blok) { const c = String(blok).charAt(0).toUpperCase(); return /[A-Z0-9]/.test(c) ? c : '-'; }
function lapKesSum_(items) {
  const o = { tot: 0, ok: 0, adj: 0, net: 0 };
  items.forEach(x => { o.tot++; if (x.ok) o.ok++; if (x.adj) { o.adj++; o.net += x.sel; } });
  return o;
}
/** map {rak: x} → baris tampil [{kind:'b'|'z'|'t', zona, blok, o}] */
function lapKesRows_(map) {
  const byZ = {};
  Object.keys(map).forEach(rak => {
    const b = lapRakBlok_(rak), z = lapRakZona_(b);
    ((byZ[z] || (byZ[z] = {}))[b] || (byZ[z][b] = [])).push(map[rak]);
  });
  const out = [], all = [];
  Object.keys(byZ).sort().forEach(z => {
    const bl = Object.keys(byZ[z]).sort(), zItems = [];
    bl.forEach(b => { out.push({ kind: 'b', zona: z, blok: b, o: lapKesSum_(byZ[z][b]) }); byZ[z][b].forEach(x => zItems.push(x)); });
    if (bl.length > 1) out.push({ kind: 'z', zona: z, blok: 'Subtotal Zona ' + z, o: lapKesSum_(zItems) });
    zItems.forEach(x => all.push(x));
  });
  out.push({ kind: 't', zona: 'SEMUA', blok: 'TOTAL PERIODE', o: lapKesSum_(all) });
  return out;
}

function lapKesWrite_(name, idHdr, periods) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const old = sh.getFilter(); if (old) old.remove();
  sh.clear(); sh.setConditionalFormatRules([]);
  const hdr = idHdr.concat(['Zona', 'Blok', 'Total Rak', 'Sesuai', 'Tidak Sesuai', '% Kesesuaian', 'Perlu Adjustment', 'Net Adj Qty (+/−)']);
  const nId = idHdr.length, nC = hdr.length;
  const cT = lapColL_(nId + 3), cS = lapColL_(nId + 4), cP = lapColL_(nId + 6);
  const body = [], kinds = [], starts = [];
  periods.forEach(p => {
    starts.push(body.length);
    lapKesRows_(p.map).forEach(x => {
      const r = body.length + 2;
      body.push(p.ids.concat([x.zona, x.blok, x.o.tot, x.o.ok, x.o.tot - x.o.ok, pct_(cS + r, cT + r), x.o.adj, x.o.net]));
      kinds.push(x.kind);
    });
  });
  const need = Math.max(body.length + 1, 2); if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (sh.getMaxColumns() < nC) sh.insertColumnsAfter(sh.getMaxColumns(), nC - sh.getMaxColumns());
  sh.getRange(1, 1, 1, nC).setValues([hdr]).setFontWeight('bold').setBackground('#0a6c74').setFontColor('#ffffff').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40); sh.setFrozenRows(1);
  const n = body.length;
  if (n) {
    sh.getRange(2, 1, n, nId + 2).setNumberFormat('@');
    sh.getRange(2, 1, n, nC).setValues(body);
    sh.getRange(2, nId + 3, n, 3).setNumberFormat('#,##0');
    sh.getRange(2, nId + 6, n, 1).setNumberFormat('0.0%');
    sh.getRange(2, nId + 7, n, 1).setNumberFormat('#,##0');
    sh.getRange(2, nId + 8, n, 1).setNumberFormat('+#,##0;-#,##0;0');
    sh.getRange(2, nId + 3, n, 6).setHorizontalAlignment('right');
    sh.getRange(2, 1, n, nC).setBorder(null, null, true, null, null, null, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
    const rl = (k, a, c) => { const L = []; kinds.forEach((x, i) => { if (x === k) L.push(a + (i + 2) + ':' + c + (i + 2)); }); return L; };
    const z = rl('z', 'A', lapColL_(nC)), t = rl('t', 'A', lapColL_(nC));
    if (z.length) sh.getRangeList(z).setBackground('#e2e8f0').setFontWeight('bold');
    if (t.length) sh.getRangeList(t).setBackground('#d5eef0').setFontWeight('bold');
    const sep = starts.filter(i => i > 0).map(i => 'A' + (i + 2) + ':' + lapColL_(nC) + (i + 2));
    if (sep.length) sh.getRangeList(sep).setBorder(true, null, null, null, null, null, '#334155', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    // warna % kesesuaian: hijau >= 95%, kuning 85–<95%, merah < 85% (hanya bila Total Rak > 0)
    const rg = sh.getRange(2, nId + 6, n, 1), x = cP + '2', tt = cT + '2', x100 = x + '*100';
    const pm = Math.round(LAP_MID * 100), ph = Math.round(LAP_HI * 100);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=(' + tt + '>0)*(' + x100 + '<' + pm + ')').setBackground('#fecaca').setFontColor('#991b1b').setRanges([rg]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=(' + tt + '>0)*(' + x100 + '>=' + pm + ')*(' + x100 + '<' + ph + ')').setBackground('#fef08a').setFontColor('#854d0e').setRanges([rg]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=(' + tt + '>0)*(' + x100 + '>=' + ph + ')').setBackground('#bbf7d0').setFontColor('#166534').setRanges([rg]).build()
    ]);
  }
  sh.setColumnWidths(1, nId, 140); sh.setColumnWidth(nId + 1, 70); sh.setColumnWidth(nId + 2, 150); sh.setColumnWidths(nId + 3, 6, 100);
  return n;
}

function lapKesesuaian_() {
  const A = lapKesAgg_(lapVals_('Rak_Detail'));
  const wk = Object.keys(A.W).sort().reverse().map(mk => ({ ids: [bulan_(mk), wLabel_(mk), periode_(mk)], map: A.W[mk] }));
  const bl = Object.keys(A.B).sort().reverse().map(bk => ({ ids: [bulan_(bk + '-01')], map: A.B[bk] }));
  const a = lapKesWrite_('Kesesuaian_Mingguan', ['Bulan (tgl Senin)', 'Minggu', 'Periode (Senin–Minggu)'], wk);
  const b = lapKesWrite_('Kesesuaian_Bulanan', ['Bulan'], bl);
  return { wk: wk.length, bl: bl.length, rows: a + b };
}

/* =====================================================================
 * COVERAGE SO RAK PER BULAN  (rak yang SUDAH / BELUM dicek di bulan tertentu)
 * Dua pengecekan terpisah: KESESUAIAN (status item Sesuai/Beda) dan SO QTY (status qty Sama/Selisih).
 * Sebuah rak dianggap "sudah" bila minimal 1x dicek di bulan itu (tanggal SO, bukan tanggal simpan).
 * KESESUAIAN wajib semua rak; SO QTY hanya sampel acak per bulan (target & sampel diatur di aplikasi: R.cfg / R.samples).
 * ===================================================================== */
function rakCoverMap_(bln) {
  const m = {}, dates = {};
  const rows = lapVals_('Rak_Detail').map(r => ({ d: dkey_(r[0]), rak: rkey_(r[1]), it: String(r[5]).trim(), q: String(r[9]).trim() }))
    .filter(x => LAP_ISO.test(x.d) && x.d.slice(0, 7) === bln && x.rak).sort((a, b) => a.d.localeCompare(b.d));
  rows.forEach(x => {
    const o = m[x.rak] || (m[x.rak] = { it: '', tIt: '', q: '', tQ: '' });
    if (x.it === 'Sesuai' || x.it === 'Beda') { o.it = x.it; o.tIt = x.d; }
    if (x.q === 'Sama' || x.q === 'Selisih') { o.q = x.q; o.tQ = x.d; }
    dates[x.d] = 1;
  });
  return { map: m, dates: Object.keys(dates).sort() };
}

/** Dipakai aplikasi: peta rak → status cek di bulan 'YYYY-MM' (kosong = bulan ini) */
function rakCover(tok, bln) {
  auth_(tok);
  const b = /^\d{4}-\d{2}$/.test(String(bln)) ? String(bln) : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const c = rakCoverMap_(b);
  return { bln: b, map: c.map, dates: c.dates };
}

/** Tab "Belum_Dicek": semua rak masuk SO + status KESESUAIAN (semua rak) dan SO QTY (hanya rak SAMPEL acak, diatur di aplikasi) bulan berjalan; yang belum di atas */
function lapBelum_() {
  const ss = ss_(), name = 'Belum_Dicek';
  const bln = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  let wms = [], cfg = {}, samp = [];
  try {
    const all = dbGetAll(INTERNAL_);
    if (all.rak) {
      const R = JSON.parse(all.rak);
      wms = (R.wms || []).filter(r => r && r.so !== false && String(r.rak || '').trim() && inScope_(r.rak));
      cfg = Object.assign({ qMode: 'pct', qVal: 10 }, R.cfg || {});
      samp = (((R.samples || {})[bln]) || {}).raks || [];
    }
  } catch (e) {}
  const c = rakCoverMap_(bln).map, S = {};
  samp.forEach(k => { S[rkey_(k)] = 1; });
  const dI = d => d ? d.slice(8) + '-' + d.slice(5, 7) + '-' + d.slice(0, 4) : '';
  const seen = {}, L = [];
  wms.forEach(r => {
    const k = rkey_(r.rak); if (seen[k]) return; seen[k] = 1;
    const x = c[k] || {}, a = !!x.it, b = !!x.q, sp = !!S[k], blok = lapRakBlok_(k), qBelum = sp && !b;
    const st = !a && qBelum ? 'BELUM KESESUAIAN + SAMPEL QTY' : (!a ? 'BELUM KESESUAIAN' : (qBelum ? 'BELUM SO QTY (SAMPEL)' : 'LENGKAP'));
    L.push({ ord: !a && qBelum ? 0 : (!a ? 1 : (qBelum ? 2 : 3)), row: [bulan_(bln + '-01'), lapRakZona_(blok), blok, String(r.rak).trim(), r.kode ? String(r.nama || r.kode) : 'kosong', Number(r.onhand) || 0,
      a ? 'SUDAH (' + x.it + ')' : 'BELUM', dI(x.tIt), sp ? 'SAMPEL' : '-', b ? 'SUDAH (' + x.q + ')' : (sp ? 'BELUM' : '-'), dI(x.tQ), st] });
  });
  L.sort((p, q) => p.ord - q.ord || p.row[3].localeCompare(q.row[3], undefined, { numeric: true }));
  const n = L.length, ka = L.filter(x => x.row[6] !== 'BELUM').length, kq = L.filter(x => x.row[9].indexOf('SUDAH') === 0).length;
  const tq = Math.min(n, Math.ceil(cfg.qMode === 'n' ? Number(cfg.qVal) || 0 : n * (Number(cfg.qVal) || 0) / 100));
  const pc = (v, t) => t ? Math.round(v / t * 1000) / 10 + '%' : '-';
  let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
  const old = sh.getFilter(); if (old) old.remove();
  sh.clear();
  const hdr = ['Bulan', 'Zona', 'Blok', 'Rak', 'Item (WMS)', 'Qty WMS', 'Kesesuaian', 'Tgl Kesesuaian', 'Sampel SO Qty', 'SO Qty', 'Tgl SO Qty', 'Status'];
  if (sh.getMaxRows() < n + 2) sh.insertRowsAfter(sh.getMaxRows(), n + 2 - sh.getMaxRows());
  if (sh.getMaxColumns() < hdr.length) sh.insertColumnsAfter(sh.getMaxColumns(), hdr.length - sh.getMaxColumns());
  sh.getRange(1, 1).setValue('Bulan ' + bulan_(bln + '-01') + ' · rak masuk SO: ' + n + ' · KESESUAIAN (semua rak) ' + ka + '/' + n + ' (' + pc(ka, n) + ') · SO QTY (sampel acak) ' + kq + '/' + tq + ' (' + pc(kq, tq) + ')' + (samp.length ? '' : ' · sampel qty belum diacak di aplikasi') + ' · diperbarui ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm')).setFontWeight('bold');
  sh.getRange(2, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#0a6c74').setFontColor('#ffffff').setWrap(true);
  sh.setFrozenRows(2);
  if (n) {
    sh.getRange(3, 1, n, hdr.length).setNumberFormat('@');
    sh.getRange(3, 1, n, hdr.length).setValues(L.map(x => x.row));
    sh.getRange(3, 6, n, 1).setNumberFormat('#,##0');
    const col = { 0: '#fecaca', 1: '#fed7aa', 2: '#fef08a', 3: '#bbf7d0' }, grp = { 0: [], 1: [], 2: [], 3: [] };
    L.forEach((x, i) => grp[x.ord].push('L' + (i + 3)));
    Object.keys(grp).forEach(g => { if (grp[g].length) sh.getRangeList(grp[g]).setBackground(col[g]).setFontWeight('bold'); });
    sh.getRange(2, 1, n + 1, hdr.length).createFilter();
  }
  sh.autoResizeColumns(1, hdr.length);
  return n;
}

/** Bangun ulang semua tab laporan. Aman dijalankan berulang kali. */
function buildLaporan() {
  try { rakFixPct_(); } catch (e) {}   // menu 'Bangun ulang laporan' juga memperbaiki % lama di tab Rak
  const A = lapAgg_();
  const wk = Object.keys(A.W).sort().reverse().map(mk => ({ ids: [periode_(mk), bulan_(mk), wLabel_(mk)], o: A.W[mk] }));
  const bl = Object.keys(A.B).sort().reverse().map(bk => ({ ids: [bulan_(bk + '-01')], o: A.B[bk] }));
  lapWrite_('Laporan_Mingguan', ['Periode (Senin–Minggu)', 'Bulan (tgl Senin)', 'Minggu'], wk, false);
  lapWrite_('Laporan_Bulanan', ['Bulan'], bl, true);
  const n = lapAdjList_();
  const kes = lapKesesuaian_();
  try { lapBelum_(); } catch (e) { log_('Belum_Dicek GAGAL', 'Laporan', 0, String(e && e.message || e)); }
  log_('Laporan', 'Laporan_Mingguan/Bulanan', wk.length + bl.length, wk.length + ' minggu, ' + bl.length + ' bulan, ' + n + ' baris adjustment');
  return { weeks: wk.length, months: bl.length, adj: n, kesesuaian: kes };
}
/** Dipanggil otomatis setelah simpan dari aplikasi; kegagalan laporan tidak boleh menggagalkan penyimpanan data */
function lapAuto_() { try { buildLaporan(); } catch (e) { log_('Laporan GAGAL', 'Laporan', 0, String(e && e.message || e)); } }

/* onOpen() sengaja TIDAK ada di file ini: menu "Laporan WMS" didefinisikan di Riset_Lengkap.gs
 * (Apps Script hanya boleh punya satu onOpen). */
