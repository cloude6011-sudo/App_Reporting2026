/* =====================================================================
 * API JSON untuk frontend yang di-hosting di luar Apps Script (mis. GitHub Pages)
 * Tempel sebagai file baru (Api.gs), lalu Deploy > Manage deployments > Edit > Version: New version.
 * Deploy sebagai Web app: Execute as = Me, Who has access = Anyone (bukan "Anyone with Google account").
 *
 * Permintaan : POST <URL /exec>, body teks JSON {"fn":"dbGetAll","args":["PIN", ...]}
 *              (args[0] = PIN, sama seperti google.script.run). Content-Type text/plain agar tanpa preflight CORS.
 * Jawaban    : {"ok":true,"data":...}  atau  {"ok":false,"error":"pesan"}
 * Keamanan   : hanya fungsi di daftar putih; setiap fungsi tetap memeriksa PIN lewat auth_().
 * ===================================================================== */
function apiFn_(n) {
  switch (n) {
    case 'authCheck': return authCheck;
    case 'dbRev': return dbRev;
    case 'dbGetAll': return dbGetAll;
    case 'dbSet': return dbSet;
    case 'saveRekap': return saveRekap;
    case 'saveRak': return saveRak;
    case 'saveRakDetail': return saveRakDetail;
    case 'sheetStatus': return sheetStatus;
    case 'rekapInfo': return rekapInfo;
    case 'appUrl': return appUrl;
    case 'srcInfo': return srcInfo;
    case 'srcTotals': return srcTotals;
    case 'rakCover': return rakCover;
    case 'clearingData': return clearingData;
    case 'simpanClearing': return simpanClearing;
    case 'kirimTelegram': return kirimTelegram;
    case 'kirimRiset': return kirimRiset;
    case 'tgStatus': return tgStatus;
    case 'tgAppGet': return tgAppGet;
    case 'tgAppSave': return tgAppSave;
    case 'tgAppTest': return tgAppTest;
    case 'tgAppFind': return tgAppFind;
  }
  return null;
}

function doPost(e) {
  let out;
  try {
    authEmail_();
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const f = apiFn_(String(req.fn || ''));
    if (!f) throw new Error('Fungsi tidak tersedia');
    const args = Array.isArray(req.args) ? req.args : [];
    out = { ok: true, data: f.apply(null, args) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out === undefined ? null : out)).setMimeType(ContentService.MimeType.JSON);
}
