/* =====================================================================
 * SERVER untuk tab "Clearing" di aplikasi web
 * Tempel sebagai file baru (Clearing_UI.gs). Butuh Clearing_Sync.gs (konstanta CLR, salinClearing_).
 * clearingData(tok, refresh):
 *   refresh = false -> baca tab CLEARING yang sudah ada di spreadsheet ini (cepat)
 *   refresh = true  -> salin dulu dari spreadsheet sumber (butuh PIN TULIS), lalu baca
 * ===================================================================== */
function clearingData(tok, refresh) {
  auth_(tok, !!refresh);
  let msg = '';
  if (refresh) msg = salinClearing_(false);
  const sh = ss_().getSheetByName(CLR.dest);
  const at = PropertiesService.getScriptProperties().getProperty('CLR_AT') || '';
  if (!sh || sh.getLastRow() < 1) return { headers: [], rows: [], at: at, msg: msg || 'Tab ' + CLR.dest + ' belum ada. Tekan Perbarui atau jalankan pasangRealtimeClearing.' };
  const v = sh.getRange(1, 1, sh.getLastRow(), Math.max(sh.getLastColumn(), 1)).getDisplayValues();
  return { headers: v[0], rows: v.slice(1).filter(r => r.some(c => String(c).trim() !== '')), at: at, msg: msg };
}
