/* =====================================================================
 * KIRIM LAPORAN APA PUN KE TELEGRAM (tanpa salin-tempel)
 * Tempel sebagai file baru (Telegram_Laporan.gs). Butuh Telegram_Clearing.gs (tgCfg_, tgSend_, ...).
 * Pengaturan token & chat ID tetap lewat Menu > Pengaturan Telegram di aplikasi.
 *
 * kirimTelegram(tok, {text, force})  -> dipakai tombol "✈️ Telegram" di semua laporan di aplikasi.
 * kirimRiset(tok)                    -> membangun Riset Data bulan ini lalu mengirimnya.
 * risetKeTelegram()                  -> menu Sheet "Laporan WMS" > "✈️ Kirim Riset ke Telegram".
 * Hasil: {state:'sent'|'dup'|'off'|'nocfg'|'fail', n, err}
 * Teks gaya WhatsApp (*tebal*, _miring_, ```tabel```) otomatis diubah jadi format Telegram dan
 * dipecah per bagian bila panjang, jadi tabel/kode tidak terpotong.
 * ===================================================================== */

/** rapikan: buang CR, spasi di ujung baris, baris kosong berlebih; batasi panjang */
function tgNeat_(t) {
  return String(t == null ? '' : t).replace(/\r/g, '').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim().slice(0, 20000);
}

/** kirim teks ke semua chat; cegah kirim ganda teks yang sama dalam 90 detik kecuali force */
function kirimTeksTelegram_(text, force, label) {
  text = tgNeat_(text);
  if (text.length < 20) throw new Error('Tidak ada isi laporan untuk dikirim');
  const c = tgCfg_();
  if (!c.token || !c.chats.length) return { state: 'nocfg' };
  if (!c.on) return { state: 'off' };
  const sig = 'TGL_' + Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text.replace(/_Update: [^\n]*_/, '')));
  const C = CacheService.getScriptCache();
  if (!force && C.get(sig)) return { state: 'dup' };
  const r = tgSend_(text);
  if (r.n) { try { C.put(sig, '1', 90); } catch (e) {} }
  try { log_('Kirim Telegram', label || 'Laporan', text.length, r.n ? 'terkirim ke ' + r.n + ' chat' + (r.fail.length ? ' (gagal: ' + r.fail.join(' | ') + ')' : '') : 'GAGAL: ' + r.fail.join(' | ')); } catch (e) {}
  return r.n ? { state: 'sent', n: r.n, err: r.fail.join(' | ') } : { state: 'fail', err: r.fail.join(' | ') || 'tidak terkirim' };
}

/** dipanggil aplikasi (PIN tulis). p = {text, judul?, force?} */
function kirimTelegram(tok, p) {
  auth_(tok, true);
  p = p || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return kirimTeksTelegram_(p.text, !!p.force, String(p.judul || '').slice(0, 60)); }
  finally { lock.releaseLock(); }
}

/** teks Telegram untuk hasil risetData() */
function risetTeks_(r) {
  const ico = { BAIK: '🟢', WASPADA: '🟡', KRITIS: '🔴' };
  const pad = (s, n) => { s = String(s); return s.length >= n ? s.slice(0, n) : s + ' '.repeat(n - s.length); };
  const val = k => k.v == null ? '-' : (k.u === 'p' ? riPct_(k.v) : (k.v > 0 ? '+' : '') + k.v);
  const L = ['*📊 RISET DATA WMS · ' + r.bln + '*', '_Diperbarui ' + r.updated + '_', '━━━━━━━━━━━━━━━━━━━━', '', '*A. KPI BULAN INI*', '```'];
  r.kpi.forEach(k => L.push((ico[k.status] || '⚪') + ' ' + pad(String(k.n).replace(/ \(makin kecil makin baik\)/, ''), 24) + pad(val(k), 8) + (k.prev == null ? '' : '(lalu ' + (k.u === 'p' ? riPct_(k.prev) : k.prev) + ')')));
  L.push('```', '_🟢 baik · 🟡 waspada · 🔴 kritis_', '', '*B. TEMUAN OTOMATIS*');
  r.temuan.forEach(t => L.push('• ' + t));
  if (r.topPic && r.topPic.length) { L.push('', '*C. TOP PIC SELISIH SKU*'); r.topPic.forEach(x => L.push('• ' + x.pic + ' — ' + x.n + ' SKU' + (x.belum ? ' (' + x.belum + ' belum ditindak)' : ''))); }
  if (r.topRak && r.topRak.length) { L.push('', '*D. RAK ADJUSTMENT TERBESAR*'); r.topRak.forEach(x => L.push('• ' + x.rak + ' — net ' + (x.net > 0 ? '+' : '') + x.net)); }
  return L.join('\n');
}

/** dipanggil aplikasi (PIN tulis): bangun riset bulan ini lalu kirim. Butuh ±10–30 detik. */
function kirimRiset(tok, force) {
  auth_(tok, true);
  const r = risetData(INTERNAL_);
  if (r.empty) throw new Error('Belum ada data untuk dianalisis');
  const res = kirimTeksTelegram_(risetTeks_(r), !!force, 'Riset Data ' + r.bln);
  res.bln = r.bln;
  return res;
}

/** menu Sheet "Laporan WMS" */
function risetKeTelegram() {
  const ui = SpreadsheetApp.getUi();
  try {
    ss_().toast('Menyusun riset lalu mengirim ke Telegram… ±10–30 detik', 'Telegram', 30);
    const r = risetData(INTERNAL_);
    if (r.empty) { ui.alert('Riset Data', 'Belum ada data untuk dianalisis.', ui.ButtonSet.OK); return; }
    const res = kirimTeksTelegram_(risetTeks_(r), true, 'Riset Data ' + r.bln);
    const msg = { sent: 'Terkirim ke ' + res.n + ' chat Telegram.' + (res.err ? '\nGagal: ' + res.err : ''), off: 'Telegram nonaktif. Aktifkan di Pengaturan Telegram.', nocfg: 'Telegram belum diatur (token / chat ID).', fail: 'GAGAL: ' + res.err }[res.state] || res.state;
    ui.alert('Kirim Riset ke Telegram', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Gagal', String(e && e.message || e), ui.ButtonSet.OK);
  }
}
