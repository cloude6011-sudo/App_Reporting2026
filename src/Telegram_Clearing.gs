/* =====================================================================
 * TELEGRAM BOT untuk laporan CLEARING
 * Tempel sebagai file baru (Telegram_Clearing.gs). Butuh Code.gs (auth_, ss_, log_).
 *
 * Alur: di tab Clearing aplikasi tekan "💾 Simpan"
 *       -> hasil dicatat di tab "Clearing_Hasil"
 *       -> laporan otomatis dikirim ke Telegram (jika aktif).
 *
 * Pengaturan: menu Laporan WMS > Pengaturan Telegram (token bot, chat ID, aktif/nonaktif, tes kirim).
 * Token & chat ID disimpan di Script Properties (BUKAN di sheet), jadi tidak terlihat peserta sheet.
 * Izin baru yang diminta saat pertama dipakai: "Terhubung ke layanan eksternal" (kirim ke api.telegram.org).
 * Bila appsscript.json punya daftar oauthScopes sendiri, tambahkan:
 *   https://www.googleapis.com/auth/script.external_request
 * ===================================================================== */
const TG_SHEET = 'Clearing_Hasil';
const TG_MAX = 3800;                                   // batas aman per pesan (Telegram maks 4096)

/* ---------- konfigurasi ---------- */
function tgCfg_() {
  const P = PropertiesService.getScriptProperties();
  return { token: P.getProperty('TG_TOKEN') || '', chats: (P.getProperty('TG_CHAT') || '').split(',').map(s => s.trim()).filter(Boolean), on: P.getProperty('TG_ON') === '1' };
}
function tgMask_(t) { t = String(t || ''); const i = t.indexOf(':'); return t ? t.slice(0, i > 0 ? i : 4) + ':••••' + t.slice(-3) : ''; }
function tgClean_(msg, token) { let m = String(msg || ''); if (token) m = m.split(token).join('•••'); return m.slice(0, 300); }
/** Hanya editor/pemilik spreadsheet (dipanggil dari dialog Sheet), bukan pengguna web app */
function tgAdmin_() {
  let me = '';
  try { me = (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  const ss = ss_(), ok = ss.getEditors().map(u => u.getEmail().toLowerCase());
  try { const o = ss.getOwner(); if (o) ok.push(o.getEmail().toLowerCase()); } catch (e) {}
  if (!me || ok.indexOf(me) < 0) throw new Error('Hanya editor spreadsheet yang boleh mengubah pengaturan Telegram');
}

/* ---------- Telegram API ---------- */
function tgApi_(token, method, payload) {
  try {
    const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload || {}), muteHttpExceptions: true });
    let j = {}; try { j = JSON.parse(r.getContentText()); } catch (e) {}
    return { ok: !!j.ok, desc: j.description || ('HTTP ' + r.getResponseCode()), result: j.result };
  } catch (e) { return { ok: false, desc: tgClean_(e && e.message || e, token) }; }
}
/** teks gaya WhatsApp (*tebal*, _miring_, ```kode```) -> HTML Telegram */
function tgHtml_(t) {
  t = String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  t = t.replace(/```\n?([\s\S]*?)```/g, (m, c) => '<pre>' + c.replace(/\n$/, '') + '</pre>');
  t = t.replace(/\*([^*\n]+)\*/g, '<b>$1</b>');
  t = t.replace(/^_(.+)_$/gm, '<i>$1</i>');
  return t;
}
/** pecah per bagian (batas baris kosong) agar blok ``` tidak terpotong */
function tgChunks_(t) {
  const out = []; let cur = '';
  String(t).split('\n\n').forEach(sec => {
    if (sec.length > TG_MAX) {                          // bagian sangat panjang: pecah per baris
      if (cur) { out.push(cur); cur = ''; }
      let b = ''; sec.split('\n').forEach(l => { if ((b + '\n' + l).length > TG_MAX) { out.push(b); b = l; } else b = b ? b + '\n' + l : l; });
      if (b) cur = b; return;
    }
    if (cur && (cur + '\n\n' + sec).length > TG_MAX) { out.push(cur); cur = sec; } else cur = cur ? cur + '\n\n' + sec : sec;
  });
  if (cur) out.push(cur);
  return out;
}
/** kirim ke semua chat; hasil {n: chat berhasil, fail: [..]} */
function tgSend_(text) {
  const c = tgCfg_(), res = { n: 0, fail: [] };
  const parts = tgChunks_(text).map(tgHtml_);
  c.chats.forEach(chat => {
    let ok = true, why = '';
    for (let i = 0; i < parts.length && ok; i++) {
      let r = tgApi_(c.token, 'sendMessage', { chat_id: chat, text: parts[i], parse_mode: 'HTML', disable_web_page_preview: true });
      if (!r.ok && /parse entities/i.test(r.desc)) r = tgApi_(c.token, 'sendMessage', { chat_id: chat, text: parts[i].replace(/<[^>]+>/g, ''), disable_web_page_preview: true });   // cadangan: teks polos
      if (!r.ok) { ok = false; why = r.desc; }
    }
    if (ok) res.n++; else res.fail.push(chat + ': ' + tgClean_(why, c.token));
  });
  return res;
}

/* ---------- dipanggil dialog Pengaturan (hanya editor sheet) ---------- */
function tgGetConfig() { tgAdmin_(); return tgGetConfig_(); }
function tgGetConfig_() { const c = tgCfg_(); return { tokenMask: tgMask_(c.token), chats: c.chats.join(', '), on: c.on, hasToken: !!c.token }; }
function tgSaveConfig(o) { tgAdmin_(); return tgSaveConfig_(o); }
function tgSaveConfig_(o) {
  const P = PropertiesService.getScriptProperties(), old = tgCfg_();
  const token = String(o && o.token || '').trim() || old.token;
  const chats = String(o && o.chats || '').split(/[,\s;]+/).map(s => s.trim()).filter(Boolean);
  if (!/^\d{5,}:[\w-]{20,}$/.test(token)) throw new Error('Format token bot tidak valid (contoh: 123456789:AAH...)');
  const bad = chats.filter(c => !/^-?\d+$/.test(c) && !/^@\w{4,}$/.test(c));
  if (bad.length) throw new Error('Chat ID tidak valid: ' + bad.join(', ') + ' (angka, atau @namachannel)');
  if (o && o.on && !chats.length) throw new Error('Isi minimal satu chat ID untuk mengaktifkan');
  P.setProperty('TG_TOKEN', token); P.setProperty('TG_CHAT', chats.join(',')); P.setProperty('TG_ON', o && o.on ? '1' : '0');
  const me = tgApi_(token, 'getMe', {});
  return me.ok ? 'Tersimpan. Bot: @' + me.result.username + (o && o.on ? ' · pengiriman AKTIF' : ' · pengiriman nonaktif') : 'Tersimpan, tapi token ditolak Telegram: ' + me.desc;
}
function tgTest() { tgAdmin_(); return tgTest_(); }
function tgTest_() {
  const c = tgCfg_();
  if (!c.token || !c.chats.length) throw new Error('Simpan token dan chat ID dulu');
  const r = tgSend_('✅ *Tes bot WMS Warehouse*\nBila pesan ini terbaca, laporan Clearing akan dikirim ke chat ini.');
  return 'Terkirim ke ' + r.n + ' dari ' + c.chats.length + ' chat' + (r.fail.length ? '\nGagal: ' + r.fail.join(' | ') : '');
}
/** bantu cari chat ID: kirim pesan ke bot (atau tambahkan bot ke grup & kirim pesan), lalu tekan Cari */
function tgFindChats() { tgAdmin_(); return tgFindChats_(); }
function tgFindChats_() {
  const c = tgCfg_();
  if (!c.token) throw new Error('Simpan token bot dulu');
  const r = tgApi_(c.token, 'getUpdates', { limit: 100 });
  if (!r.ok) throw new Error(r.desc);
  const seen = {};
  (r.result || []).forEach(u => { const m = u.message || u.edited_message || u.channel_post || (u.my_chat_member && { chat: u.my_chat_member.chat }); if (m && m.chat) seen[m.chat.id] = { id: String(m.chat.id), type: m.chat.type, title: m.chat.title || [m.chat.first_name, m.chat.last_name].filter(Boolean).join(' ') || m.chat.username || '' }; });
  return Object.keys(seen).map(k => seen[k]);
}
function tgSetup() {
  tgAdmin_();
  const html = `<style>body{font:14px Arial,sans-serif;margin:14px;color:#0f172a}label{display:block;font-weight:bold;margin:10px 0 3px}input[type=text],input[type=password]{width:100%;box-sizing:border-box;padding:7px}
small{color:#64748b;display:block;margin-top:2px}button{padding:8px 12px;border:1px solid #94a3b8;border-radius:6px;background:#fff;cursor:pointer;margin-right:6px}button.p{background:#1d4ed8;color:#fff;border-color:#1d4ed8}
#st{margin-top:12px;white-space:pre-wrap;font-size:13px}.c{display:inline-block;margin:3px 4px 0 0;padding:3px 8px;border:1px solid #cbd5e1;border-radius:99px;cursor:pointer;font-size:12px}</style>
<label>Token bot</label><input type="password" id="tk" autocomplete="off" placeholder="123456789:AAH..."><small id="tkm">Dapat dari @BotFather. Kosongkan bila tidak diganti.</small>
<label>Chat ID <span style="font-weight:normal">(pisahkan koma untuk beberapa chat)</span></label><input type="text" id="ch" placeholder="-1001234567890, 987654321">
<small>Kirim satu pesan ke bot (atau tambahkan bot ke grup lalu kirim pesan), kemudian tekan "Cari Chat ID".</small><div id="found"></div>
<label><input type="checkbox" id="on"> Kirim laporan otomatis saat hasil Clearing disimpan</label>
<p><button class="p" onclick="sv()">Simpan</button><button onclick="ts()">Kirim tes</button><button onclick="fd()">Cari Chat ID</button></p><div id="st"></div>
<script>
function $(i){return document.getElementById(i)}
function msg(t){$('st').textContent=t}
google.script.run.withSuccessHandler(function(c){$('ch').value=c.chats;$('on').checked=c.on;if(c.hasToken)$('tkm').textContent='Token tersimpan: '+c.tokenMask+' · kosongkan kolom bila tidak diganti.'}).withFailureHandler(function(e){msg(e.message)}).tgGetConfig();
function sv(){msg('Menyimpan…');google.script.run.withSuccessHandler(function(r){msg(r);$('tk').value=''}).withFailureHandler(function(e){msg('Gagal: '+e.message)}).tgSaveConfig({token:$('tk').value,chats:$('ch').value,on:$('on').checked})}
function ts(){msg('Mengirim tes…');google.script.run.withSuccessHandler(msg).withFailureHandler(function(e){msg('Gagal: '+e.message)}).tgTest()}
function fd(){msg('Mencari…');google.script.run.withSuccessHandler(function(a){if(!a.length){msg('Belum ada chat. Kirim pesan ke bot dulu, lalu coba lagi.');return}msg('Klik chat untuk menambahkan:');var f=$('found');f.innerHTML='';a.forEach(function(x){var s=document.createElement('span');s.className='c';s.textContent=x.id+' · '+x.title+' ('+x.type+')';s.onclick=function(){var v=$('ch').value.split(/[,\\s]+/).filter(Boolean);if(v.indexOf(x.id)<0)v.push(x.id);$('ch').value=v.join(', ')};f.appendChild(s)})}).withFailureHandler(function(e){msg('Gagal: '+e.message)}).tgFindChats()}
</script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(520), 'Pengaturan Telegram');
}

/* ---------- dipanggil aplikasi web ---------- */
/** status singkat untuk tab Clearing (PIN baca cukup) */
function tgStatus(tok) { auth_(tok, false); const c = tgCfg_(); return { on: c.on && !!c.token && c.chats.length > 0, n: c.chats.length, set: !!c.token }; }

/** "Simpan" di tab Clearing: catat hasil, lalu kirim laporan ke Telegram bila aktif (PIN tulis) */
function simpanClearing(tok, p, force) {
  auth_(tok, true);
  p = p || {};
  const text = String(p.text || '').slice(0, 20000);
  if (!text || text.length < 20) throw new Error('Tidak ada hasil untuk disimpan');
  const num = v => Math.max(0, Math.round(Number(v) || 0));
  const tot = num(p.tot), ok = num(p.ok), no = num(p.no), of = num(p.of), qd = num(p.qd);
  const qdList = (Array.isArray(p.qdList) ? p.qdList : []).slice(0, 200).map(v => String(v).slice(0, 40)).join(', ').slice(0, 2000);
  const noList = (Array.isArray(p.noList) ? p.noList : []).slice(0, 500).map(v => String(v).replace(/\D/g, '').slice(0, 7)).filter(Boolean).join(', ');
  const sig = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text.replace(/_Update: [^\n]*_/, '')));
  const P = PropertiesService.getScriptProperties();
  if (!force && P.getProperty('TG_CLR_SIG') === sig) return { dup: true };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_(), tz = Session.getScriptTimeZone(), at = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy HH:mm:ss');
    const c = tgCfg_();
    let tg;
    if (!c.token || !c.chats.length) tg = { state: 'nocfg' };
    else if (!c.on) tg = { state: 'off' };
    else { const r = tgSend_(text); tg = r.n ? { state: 'sent', n: r.n, err: r.fail.join(' | ') } : { state: 'fail', err: r.fail.join(' | ') || 'tidak terkirim' }; }
    let sh = ss.getSheetByName(TG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(TG_SHEET);
      sh.getRange(1, 1, 1, 11).setValues([['Waktu simpan', 'File laporan', 'Voucher', 'Cocok', 'Belum di CLEARING', 'Hanya di CLEARING', '% Cocok', 'Voucher belum di CLEARING', 'Telegram', 'Qty beda', 'Voucher qty beda (laporan→CLEARING)']])
        .setFontWeight('bold').setBackground('#d5eef0');
      sh.setFrozenRows(1); sh.setColumnWidths(1, 2, 150); sh.setColumnWidth(8, 360); sh.setColumnWidth(9, 220);
    }
    if (!String(sh.getRange(1, 10).getValue())) sh.getRange(1, 10, 1, 2).setValues([['Qty beda', 'Voucher qty beda (laporan→CLEARING)']]).setFontWeight('bold').setBackground('#d5eef0');   // sheet lama (9 kolom)
    const st = { sent: 'Terkirim ke ' + (tg.n || 0) + ' chat' + (tg.err ? ' (gagal: ' + tg.err + ')' : ''), fail: 'GAGAL: ' + tg.err, off: 'Nonaktif', nocfg: 'Belum diatur' }[tg.state];
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 11).setNumberFormats([['@', '@', '0', '0', '0', '0', '0.0%', '@', '@', '0', '@']])
      .setValues([[at, String(p.file || '').slice(0, 150), tot, ok, no, of, tot ? ok / tot : 0, noList, st, qd, qdList]]);
    P.setProperty('TG_CLR_SIG', sig);
    try { log_('Simpan Clearing', TG_SHEET, tot, ok + ' cocok, ' + no + ' belum, ' + qd + ' qty beda · Telegram: ' + tg.state); } catch (e) {}
    return { saved: true, at: at, tg: tg };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- pengaturan dari menu aplikasi web (butuh PIN TULIS, bukan PIN baca/manager) ---------- */
function tgAppGet(tok) { auth_(tok, true); return tgGetConfig_(); }
function tgAppSave(tok, o) { auth_(tok, true); return tgSaveConfig_(o); }
function tgAppTest(tok) { auth_(tok, true); return tgTest_(); }
function tgAppFind(tok) { auth_(tok, true); return tgFindChats_(); }
