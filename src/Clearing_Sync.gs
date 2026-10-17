/* =====================================================================
 * SALIN SHEET "REKAP" (kolom A–N) dari spreadsheet lain, hampir realtime
 * Tempel sebagai file baru di Apps Script (Clearing_Sync.gs). Memakai fungsi Code.gs: srcId_, srcAllow_, cfgSheet_, ss_.
 *
 * Langkah (sekali saja, dari editor Apps Script):
 *   1) Run izinkanClearing      -> mendaftarkan spreadsheet sumber di tab Config (Izinkan = YA)
 *   2) Run pasangRealtimeClearing -> memasang trigger + menyalin pertama kali (setujui izin akses)
 * Hasil: tab "CLEARING" di spreadsheet ini (isinya salinan tab REKAP sumber), selalu mengikuti sumber.
 *
 * Cara kerja "realtime":
 *   - trigger onChange di spreadsheet sumber (jalan saat ada yang mengedit di sumber)
 *   - trigger tiap 1 menit sebagai jaring pengaman (batas tercepat Apps Script)
 *   - bila isi tidak berubah (dicek lewat sidik MD5), tab tujuan tidak ditulis ulang
 * Catatan: tab tujuan ditimpa penuh tiap ada perubahan. Jangan menulis manual di tab CLEARING ini.
 * ===================================================================== */
const CLR = {
  src: '1HUO0bYa-ClVROP8aAsl58tnRvL9mHCGENUQD4sV3yD4',   // ID spreadsheet sumber
  tab: 'REKAP',                                           // nama tab sumber (huruf besar/kecil bebas)
  from: 'A', to: 'N',                                     // rentang kolom yang disalin
  dest: 'CLEARING'                                        // nama tab tujuan di spreadsheet ini
};

function clrCol_(s) { let n = 0; String(s).toUpperCase().split('').forEach(c => { n = n * 26 + c.charCodeAt(0) - 64; }); return n; }

function salinClearing_(force) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return 'sibuk, dilewati';
  try {
    const id = srcId_(CLR.src);
    srcAllow_(id);                                        // wajib YA di tab Config
    const src = SpreadsheetApp.openById(id);
    const want = String(CLR.tab).trim().toLowerCase();
    const sh = src.getSheets().filter(s => s.getName().trim().toLowerCase() === want)[0];
    if (!sh) throw new Error('Tab "' + CLR.tab + '" tidak ditemukan di spreadsheet sumber. Tab yang ada: ' + src.getSheets().map(s => s.getName()).join(', '));

    const c1 = clrCol_(CLR.from), c2 = clrCol_(CLR.to), w = c2 - c1 + 1;
    const lr = Math.max(sh.getLastRow(), 1);
    const rng = sh.getRange(1, c1, lr, w);
    const vals = rng.getValues(), fmts = rng.getNumberFormats();

    const sig = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(vals) + '|' + JSON.stringify(fmts)));
    const P = PropertiesService.getScriptProperties();
    if (!force && P.getProperty('CLR_SIG') === sig) return 'tidak berubah';

    const ss = ss_();
    let d = ss.getSheetByName(CLR.dest);
    if (!d) d = ss.insertSheet(CLR.dest);
    if (d.getMaxRows() < lr) d.insertRowsAfter(d.getMaxRows(), lr - d.getMaxRows());
    if (d.getMaxColumns() < w) d.insertColumnsAfter(d.getMaxColumns(), w - d.getMaxColumns());
    const oldLast = d.getLastRow();
    const safe = vals.map(r => r.map(v => (typeof v === 'string' && v.charAt(0) === '=') ? "'" + v : v));   // teks "=..." tidak dijalankan sebagai rumus
    d.getRange(1, 1, lr, w).setNumberFormats(fmts);
    d.getRange(1, 1, lr, w).setValues(safe);
    if (oldLast > lr) d.getRange(lr + 1, 1, oldLast - lr, d.getMaxColumns()).clearContent();   // sisa baris lama
    d.setFrozenRows(1);
    d.getRange(1, 1, 1, w).setFontWeight('bold');
    P.setProperty('CLR_SIG', sig);
    P.setProperty('CLR_AT', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss'));
    return 'disalin ' + lr + ' baris x ' + w + ' kolom';
  } finally {
    lock.releaseLock();
  }
}

/** Dipakai trigger (tidak memaksa tulis bila isi sama) */
function clearingSync() { try { Logger.log(salinClearing_(false)); } catch (e) { log_('Sync CLEARING GAGAL', CLR.dest, 0, String(e && e.message || e)); } }

/** Salin sekarang juga (manual / menu) */
function salinClearing() { const r = salinClearing_(true); Logger.log(r); return r; }

/** Daftarkan spreadsheet sumber di tab Config dengan Izinkan = YA */
function izinkanClearing() {
  const id = srcId_(CLR.src), sh = cfgSheet_(), n = sh.getLastRow();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
  let row = 0;
  if (n >= 2) {
    const v = sh.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() && srcId_(v[i][0]) === id) { row = i + 2; break; }
  }
  if (row) { sh.getRange(row, 3).setValue('YA'); sh.getRange(row, 4).setValue(now); }
  else sh.appendRow([id, 'CLEARING (salin otomatis)', 'YA', now]);
  Logger.log('Sumber diizinkan: ' + id);
}

/** Pasang trigger realtime + salin pertama kali. Aman dijalankan ulang (trigger lama diganti). */
function pasangRealtimeClearing() {
  izinkanClearing();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'clearingSync').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('clearingSync').timeBased().everyMinutes(1).create();
  let onch = 'ya';
  try { ScriptApp.newTrigger('clearingSync').forSpreadsheet(srcId_(CLR.src)).onChange().create(); }
  catch (e) { onch = 'GAGAL (' + String(e && e.message || e) + ') - tetap jalan tiap 1 menit'; }
  Logger.log('Trigger tiap 1 menit: ya | onChange sumber: ' + onch);
  Logger.log(salinClearing_(true));
}

/** Matikan sinkronisasi */
function hapusRealtimeClearing() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'clearingSync').forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('Trigger sinkronisasi CLEARING dihapus');
}
