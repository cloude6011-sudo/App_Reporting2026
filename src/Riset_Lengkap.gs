/* =====================================================================
 * RISET DATA + RAPIKAN LAPORAN
 * Tempel sebagai FILE BARU di Apps Script (nama bebas, mis. Riset_Tambahan.gs).
 * Memakai fungsi yang sudah ada di Code.gs (lapAgg_, buildLaporan, auth_, dll.).
 *
 * 1) risetData(tok)      -> dipanggil tombol "Riset Data" di aplikasi web.
 *                           Membangun ulang laporan, lalu membuat tab "Riset_Data":
 *                           KPI bulan ini vs bulan lalu, temuan otomatis, tren 8 minggu,
 *                           Top PIC selisih, Top rak adjustment. Hasil ringkas dikirim ke aplikasi.
 * 2) rapikanLaporan_()   -> dipanggil otomatis di akhir buildLaporan (lihat langkah di bawah):
 *                           tab "Daftar_Isi" (link ke semua tab + penjelasan), urutan tab,
 *                           warna tab per kelompok, kolom Laporan_* bisa dilipat per kelompok.
 * ===================================================================== */

const RI_SHEET = 'Riset_Data', TOC_SHEET = 'Daftar_Isi';
const RI_COL = {
  'BAIK': ['#bbf7d0', '#166534'], 'WASPADA': ['#fef08a', '#854d0e'],
  'KRITIS': ['#fecaca', '#991b1b'], 'BELUM ADA DATA': ['#e2e8f0', '#475569']
};

function riDiv_(a, b) { return b > 0 ? a / b : null; }
function riStat_(v, low) {
  if (v == null) return 'BELUM ADA DATA';
  if (low) return v <= 0.05 ? 'BAIK' : (v <= 0.15 ? 'WASPADA' : 'KRITIS');
  return v >= LAP_HI ? 'BAIK' : (v >= LAP_MID ? 'WASPADA' : 'KRITIS');
}
function riPct_(v) { return v == null ? '-' : (Math.round(v * 1000) / 10) + '%'; }

function riKpi_(o, selLines) {
  o = o || lapBlank_();
  return [
    { n: '% Mobil WMS', v: riDiv_(o.mobWms, o.tMob), u: 'p', t: '≥ 95%' },
    { n: '% Item OK (Aktual)', v: riDiv_(o.ok, o.lines), u: 'p', t: '≥ 95%' },
    { n: '% Kesesuaian Rak', v: riDiv_(o.io, o.ic), u: 'p', t: '≥ 95%' },
    { n: '% Stock Item', v: riDiv_(o.sk, o.sc), u: 'p', t: '≥ 95%' },
    { n: '% Qty per Bin', v: riDiv_(o.qo, o.qc), u: 'p', t: '≥ 95%' },
    { n: '% SKU Selisih (makin kecil makin baik)', v: riDiv_(o.selN, selLines), u: 'p', low: true, t: '≤ 5%' },
    { n: '% Selisih Ditindak', v: riDiv_(o.selFu, o.selN), u: 'p', t: '≥ 95%' },
    { n: 'Net Adjustment (Rak + Muat)', v: o.adjNeg + o.adjPos + o.mPos + o.mNeg, u: 'n', plain: true, t: 'mendekati 0' }
  ];
}

/** Tombol "Riset Data" di aplikasi web memanggil fungsi ini. */
function risetData(tok) {
  auth_(tok, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    buildLaporan();                 // pastikan semua tab laporan memakai data terbaru
    SpreadsheetApp.flush();
    const ss = ss_(), A = lapAgg_();
    const bKeys = Object.keys(A.B).sort(), wKeys = Object.keys(A.W).sort();
    if (!bKeys.length) return { empty: true };
    const cur = bKeys[bKeys.length - 1], prv = bKeys.length > 1 ? bKeys[bKeys.length - 2] : null;
    const label = bulan_(cur + '-01');
    const K = riKpi_(A.B[cur], A.B[cur].selLines);
    const P = prv ? riKpi_(A.B[prv], A.B[prv].selLines) : null;
    K.forEach((k, i) => { k.prev = P ? P[i].v : null; k.status = k.plain ? (k.v == null ? 'BELUM ADA DATA' : 'BAIK') : riStat_(k.v, k.low); });

    // ---- data pendukung dari tab yang sudah dibangun ----
    const picAll = lapVals_('Selisih_PIC').filter(r => String(r[0]).trim() === label)
      .map(r => ({ pic: String(r[1]), lines: lapNum_(r[2]), n: lapNum_(r[3]), p: lapNum_(r[4]), belum: lapNum_(r[6]), sh: lapNum_(r[7]) }))
      .sort((a, b) => b.n - a.n || b.p - a.p).slice(0, 10);
    const rakAll = lapVals_('Adjustment_per_Rak').filter(r => String(r[0]).trim() === label)
      .map(r => ({ rak: String(r[1]), so: lapNum_(r[2]), nm: lapNum_(r[5]), net: lapNum_(r[6]), ks: lapNum_(r[7]), km: lapNum_(r[8]) }))
      .sort((a, b) => Math.abs(b.net) - Math.abs(a.net)).slice(0, 10);
    let belumRak = 0, belumInfo = '';
    const bd = ss.getSheetByName('Belum_Dicek');
    if (bd && bd.getLastRow() >= 3) {
      belumInfo = String(bd.getRange(1, 1).getValue());
      belumRak = bd.getRange(3, 12, bd.getLastRow() - 2, 1).getValues().filter(r => String(r[0]).trim() !== 'LENGKAP').length;
    }
    const mobKosong = lapVals_('Harian').filter(r => String(r[20]).trim() === label && String(r[18]).trim() === 'Total mobil belum diisi').length;

    // ---- temuan otomatis ----
    const T = [], o = A.B[cur];
    K.forEach(k => {
      if (k.status === 'KRITIS') T.push('KRITIS: ' + k.n + ' = ' + riPct_(k.v) + ' (target ' + k.t + ')');
      if (k.u === 'p' && k.v != null && k.prev != null && (k.low ? k.v - k.prev > 0.05 : k.prev - k.v > 0.05))
        T.push('MEMBURUK: ' + k.n + ' ' + riPct_(k.prev) + ' → ' + riPct_(k.v) + ' dibanding bulan lalu');
    });
    if (o.selN - o.selFu > 0) T.push(String(o.selN - o.selFu) + ' SKU selisih muat belum ditindak lanjuti');
    if (picAll.length && picAll[0].n > 0) T.push('PIC dengan SKU selisih terbanyak: ' + picAll[0].pic + ' (' + picAll[0].n + ' SKU)');
    if (belumRak > 0) T.push(belumRak + ' rak belum lengkap dicek bulan ini (lihat tab Belum_Dicek)');
    if (mobKosong > 0) T.push(mobKosong + ' hari muat belum diisi Total Mobil');
    if (!T.length) T.push('Tidak ada temuan kritis bulan ini.');

    // ---- susun tab Riset_Data (6 kolom) ----
    const W6 = 6, rows = [], sty = [], pf = [], kr = [];
    const add = (arr, type) => { const a = arr.slice(); while (a.length < W6) a.push(''); rows.push(a); sty.push(type || ''); return rows.length; };
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
    add(['RISET DATA WMS · ' + label], 'title');
    add(['Diperbarui ' + now + ' · hijau ≥ ' + Math.round(LAP_HI * 100) + '%, kuning ≥ ' + Math.round(LAP_MID * 100) + '%, merah di bawahnya'], 'sub');
    add([], 'blank');
    add(['1. KPI BULAN INI vs BULAN LALU'], 'sec');
    add(['Indikator', 'Bulan ini', 'Bulan lalu', 'Perubahan (poin)', 'Status', 'Target'], 'hdr');
    K.forEach(k => {
      const r = add([k.n, k.v == null ? '' : (k.u === 'p' ? k.v : k.v), k.prev == null ? '' : k.prev,
        (k.v == null || k.prev == null) ? '' : (k.u === 'p' ? (k.v - k.prev) * 100 : k.v - k.prev), k.status, k.t], 'row');
      kr.push({ r: r, u: k.u, s: k.status });
    });
    add([], 'blank');
    add(['2. TEMUAN OTOMATIS'], 'sec');
    T.forEach(t => add(['• ' + t], 'note'));
    add([], 'blank');
    add(['3. TREN 8 MINGGU TERAKHIR'], 'sec');
    add(['Minggu', '% Item OK', '% Kesesuaian Rak', '% Qty per Bin', '% Ditindak', 'SKU Selisih'], 'hdr');
    const wr = rows.length + 1;
    wKeys.slice(-8).reverse().forEach(mk => {
      const w = A.W[mk];
      add([wLabel_(mk) + ' · ' + periode_(mk), riDiv_(w.ok, w.lines) == null ? '' : riDiv_(w.ok, w.lines),
        riDiv_(w.io, w.ic) == null ? '' : riDiv_(w.io, w.ic), riDiv_(w.qo, w.qc) == null ? '' : riDiv_(w.qo, w.qc),
        riDiv_(w.selFu, w.selN) == null ? '' : riDiv_(w.selFu, w.selN), w.selN], 'row');
    });
    pf.push([wr, 2, rows.length - wr + 1, 4, '0.0%']);
    add([], 'blank');
    add(['4. TOP PIC SELISIH SKU · ' + label], 'sec');
    add(['PIC Muat', 'SKU Dimuat', 'SKU Selisih', '% SKU Selisih', 'Belum Ditindak', 'SKU Bersama PIC Lain'], 'hdr');
    const pr = rows.length + 1;
    if (picAll.length) picAll.forEach(x => add([x.pic, x.lines, x.n, x.p, x.belum, x.sh], 'row')); else add(['(belum ada data selisih)'], 'note');
    if (picAll.length) pf.push([pr, 4, picAll.length, 1, '0.0%']);
    add([], 'blank');
    add(['5. TOP RAK BERMASALAH (ADJUSTMENT TERBESAR) · ' + label], 'sec');
    add(['Rak', 'Adj SO Rak', 'Net Adj Muat', 'Net Total', 'Kejadian SO', 'Kejadian Muat'], 'hdr');
    const rr = rows.length + 1;
    if (rakAll.length) rakAll.forEach(x => add([x.rak, x.so, x.nm, x.net, x.ks, x.km], 'row')); else add(['(belum ada adjustment)'], 'note');
    if (rakAll.length) pf.push([rr, 2, rakAll.length, 3, '+#,##0;-#,##0;0']);

    let sh = ss.getSheetByName(RI_SHEET);
    if (!sh) sh = ss.insertSheet(RI_SHEET);
    sh.clear(); sh.setConditionalFormatRules([]);
    const n = rows.length;
    if (sh.getMaxRows() < n) sh.insertRowsAfter(sh.getMaxRows(), n - sh.getMaxRows());
    if (sh.getMaxColumns() < W6) sh.insertColumnsAfter(sh.getMaxColumns(), W6 - sh.getMaxColumns());
    sh.getRange(1, 1, n, W6).setNumberFormat('General');
    sh.getRange(1, 1, n, 1).setNumberFormat('@');
    sh.getRange(1, 1, n, W6).setValues(rows);
    sty.forEach((t, i) => {
      const r = sh.getRange(i + 1, 1, 1, W6);
      if (t === 'title') r.setFontSize(16).setFontWeight('bold').setFontColor('#0a6c74');
      else if (t === 'sub') r.setFontColor('#64748b').setFontStyle('italic');
      else if (t === 'sec') r.setBackground('#0a6c74').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
      else if (t === 'hdr') r.setBackground('#d5eef0').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
      else if (t === 'row') r.setBorder(null, null, true, null, null, null, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
    });
    kr.forEach(k => {
      sh.getRange(k.r, 2, 1, 2).setNumberFormat(k.u === 'p' ? '0.0%' : '+#,##0;-#,##0;0');
      sh.getRange(k.r, 4).setNumberFormat('+0.0;-0.0;0.0');
      const c = RI_COL[k.s] || RI_COL['BELUM ADA DATA'];
      sh.getRange(k.r, 5).setBackground(c[0]).setFontColor(c[1]).setFontWeight('bold').setHorizontalAlignment('center');
      sh.getRange(k.r, 2, 1, 5).setHorizontalAlignment('right');
      sh.getRange(k.r, 5).setHorizontalAlignment('center');
    });
    pf.forEach(f => { sh.getRange(f[0], f[1], f[2], f[3]).setNumberFormat(f[4]).setHorizontalAlignment('right'); });
    sh.setColumnWidth(1, 330); sh.setColumnWidths(2, 5, 130);
    sh.setFrozenRows(2);
    log_('Riset Data', RI_SHEET, n, T.length + ' temuan, bulan ' + label);
    try { rapikanLaporan_(); } catch (e) {}

    return {
      url: ss.getUrl() + '#gid=' + sh.getSheetId(), bln: label, updated: now,
      kpi: K.map(k => ({ n: k.n, u: k.u, v: k.v, prev: k.prev, status: k.status, t: k.t })),
      temuan: T,
      topPic: picAll.slice(0, 5).map(x => ({ pic: x.pic, n: x.n, belum: x.belum })),
      topRak: rakAll.slice(0, 5).map(x => ({ rak: x.rak, net: x.net }))
    };
  } finally {
    lock.releaseLock();
  }
}

/* =====================================================================
 * RAPIKAN LAPORAN: Daftar_Isi, urutan tab, warna tab, kolom bisa dilipat
 * ===================================================================== */
const RI_GRP = {
  'LAPORAN': '#0a6c74', 'TINDAK LANJUT': '#b91c1c', 'REKAP DATA': '#1d4ed8', 'SISTEM': '#64748b'
};
const RI_TABS = [
  ['Riset_Data', 'LAPORAN', 'Ringkasan analisis: KPI vs bulan lalu, temuan otomatis, tren, Top PIC & rak', 'Dibuat tombol "Riset Data" di aplikasi'],
  ['Laporan_Mingguan', 'LAPORAN', '1 baris per minggu (Senin–Minggu): muat, selisih SKU, SO rak, adjustment. Kolom bisa dilipat per kelompok', 'Setoran laporan mingguan'],
  ['Laporan_Bulanan', 'LAPORAN', '1 baris per bulan, kolom sama dengan mingguan + SKU dimuat', 'Setoran laporan bulanan'],
  ['Kesesuaian_Mingguan', 'LAPORAN', 'Kesesuaian rak per zona / blok per minggu', 'Evaluasi akurasi lokasi rak'],
  ['Kesesuaian_Bulanan', 'LAPORAN', 'Kesesuaian rak per zona / blok per bulan', 'Evaluasi akurasi lokasi rak'],
  ['Adjustment_List', 'TINDAK LANJUT', 'Daftar adjustment (rak + muat) per baris, siap difilter minggu / bulan', 'Disetor ke admin untuk adjust stok'],
  ['Adjustment_per_Rak', 'TINDAK LANJUT', 'Adjustment per rak per bulan (SO + muat)', 'Mencari rak yang sering selisih'],
  ['Belum_Dicek', 'TINDAK LANJUT', 'Rak yang belum dicek bulan berjalan (yang belum di atas)', 'Pembagian tugas cek rak'],
  ['Harian', 'REKAP DATA', 'Rekap muat per tanggal', 'Sumber laporan muat'],
  ['Mingguan', 'REKAP DATA', 'Rekap muat per minggu', 'Sumber laporan muat'],
  ['Bulanan', 'REKAP DATA', 'Rekap muat per bulan', 'Sumber laporan muat'],
  ['Selisih_Bulanan', 'REKAP DATA', 'KPI SKU selisih per bulan', 'Sumber laporan selisih'],
  ['Selisih_PIC', 'REKAP DATA', 'SKU selisih per PIC muat per bulan', 'Evaluasi PIC'],
  ['Selisih_SKU', 'REKAP DATA', 'Detail tiap SKU selisih (qty muat ≠ checker out)', 'Telusur & tindak lanjut'],
  ['Rak', 'REKAP DATA', 'Ringkasan cek rak per tanggal', 'Sumber laporan rak'],
  ['Rak_Detail', 'REKAP DATA', 'Detail SO per rak per tanggal', 'Sumber laporan rak & adjustment'],
  ['DB', 'SISTEM', 'Penyimpanan data aplikasi · JANGAN DIUBAH MANUAL', 'Dipakai aplikasi'],
  ['DB_Backup', 'SISTEM', 'Salinan cadangan data aplikasi (maks 20)', 'Pemulihan bila data hilang'],
  ['Log', 'SISTEM', 'Jejak audit simpan data (500 baris terakhir)', 'Menelusuri siapa menyimpan apa'],
  ['PIN', 'SISTEM', 'Daftar PIN akses · batasi share spreadsheet ini', 'Admin'],
  ['Config', 'SISTEM', 'Izin spreadsheet sumber total mobil', 'Admin']
];

function riGroupCols_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastColumn() < 2) return;
  const lc = sh.getLastColumn(), top = sh.getRange(1, 1, 1, lc).getValues()[0];
  const starts = [];
  top.forEach((v, i) => { if (String(v).trim()) starts.push(i + 1); });
  starts.forEach((s, i) => {
    if (i === 0) return;                                   // kelompok PERIODE tetap terlihat
    const w = (i + 1 < starts.length ? starts[i + 1] : lc + 1) - s;
    try {
      if (sh.getColumnGroupDepth(s) === 0) sh.getRange(1, s, 1, w).shiftColumnGroupDepth(1);
    } catch (e) {}
  });
}

function rapikanLaporan_() {
  const ss = ss_();
  let toc = ss.getSheetByName(TOC_SHEET);
  if (!toc) toc = ss.insertSheet(TOC_SHEET);
  toc.clear();
  const exist = RI_TABS.filter(t => ss.getSheetByName(t[0]));
  const hdr = ['Kelompok', 'Tab (klik untuk pindah)', 'Isi', 'Dipakai untuk'];
  const n = exist.length;
  if (toc.getMaxRows() < n + 4) toc.insertRowsAfter(toc.getMaxRows(), n + 4 - toc.getMaxRows());
  if (toc.getMaxColumns() < 4) toc.insertColumnsAfter(toc.getMaxColumns(), 4 - toc.getMaxColumns());
  toc.getRange(1, 1).setValue('DAFTAR ISI · WMS Warehouse').setFontSize(16).setFontWeight('bold').setFontColor('#0a6c74');
  toc.getRange(2, 1).setValue('Mulai dari kelompok LAPORAN (hijau-biru) untuk setoran, TINDAK LANJUT (merah) untuk pekerjaan, REKAP DATA (biru) sebagai sumber, SISTEM (abu) jangan diubah.').setFontColor('#64748b').setFontStyle('italic');
  toc.getRange(3, 1, 1, 4).setValues([hdr]).setFontWeight('bold').setBackground('#0a6c74').setFontColor('#ffffff');
  if (n) {
    toc.getRange(4, 1, n, 4).setNumberFormat('@');
    toc.getRange(4, 1, n, 4).setValues(exist.map(t => [t[1], t[0], t[2], t[3]]));
    exist.forEach((t, i) => {
      const s = ss.getSheetByName(t[0]), r = i + 4, c = RI_GRP[t[1]];
      toc.getRange(r, 2).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(t[0]).setLinkUrl('#gid=' + s.getSheetId()).build());
      toc.getRange(r, 1).setBackground(c).setFontColor('#ffffff').setFontWeight('bold');
      s.setTabColor(c);
    });
    toc.getRange(4, 1, n, 4).setBorder(null, null, true, null, null, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
  }
  toc.setColumnWidth(1, 120); toc.setColumnWidth(2, 190); toc.setColumnWidth(3, 560); toc.setColumnWidth(4, 260);
  toc.setFrozenRows(3);
  toc.setTabColor('#0a6c74');

  ['Laporan_Mingguan', 'Laporan_Bulanan'].forEach(riGroupCols_);

  try {                                                    // urutan tab: Daftar_Isi, laporan, tindak lanjut, rekap, sistem
    [TOC_SHEET].concat(exist.map(t => t[0])).forEach((nm, i) => {
      const s = ss.getSheetByName(nm);
      if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); }
    });
    ss.setActiveSheet(toc);
  } catch (e) {}
}

/** Bisa dijalankan manual dari editor / menu: merapikan tab tanpa membangun ulang data */
function rapikanTab() { rapikanLaporan_(); }


/* =====================================================================
 * MENU "Laporan WMS" di Google Sheet: Riset Data, Bangun ulang laporan, Rapikan tab
 * Tempel sebagai file baru (Riset_Menu.gs), lalu HAPUS fungsi onOpen lama di Code.gs
 * (hanya boleh ada satu onOpen). Syarat: Riset_Tambahan.gs sudah terpasang.
 * Setelah ditempel, muat ulang spreadsheet agar menu muncul.
 * ===================================================================== */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Laporan WMS')
      .addItem('📊 Riset Data', 'risetDariMenu')
      .addSeparator()
      .addItem('Bangun ulang laporan', 'buildLaporan')
      .addItem('Rapikan tab (Daftar Isi, warna, urutan)', 'rapikanTab')
      .addSeparator()
      .addItem('⚠ Kosongkan semua data (kecuali Config & PIN)', 'kosongkanSemuaData')
      .addToUi();
  } catch (e) {}
}

function risetDariMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    ss_().toast('Menganalisis data… ±10–30 detik', 'Riset Data', 30);
    const r = risetData(INTERNAL_);
    if (r.empty) { ui.alert('Riset Data', 'Belum ada data untuk dianalisis.', ui.ButtonSet.OK); return; }
    const pct = v => v == null ? '-' : (Math.round(v * 1000) / 10) + '%';
    const lines = r.kpi.map(k => '• ' + k.n + ': ' + (k.u === 'p' ? pct(k.v) : (k.v == null ? '-' : k.v)) + '  [' + k.status + ']');
    const sh = ss_().getSheetByName(RI_SHEET);
    if (sh) ss_().setActiveSheet(sh);
    ui.alert('Riset Data · ' + r.bln,
      lines.join('\n') + '\n\nTEMUAN:\n' + r.temuan.map(t => '• ' + t).join('\n') + '\n\nDetail lengkap ada di tab Riset_Data.',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Riset Data gagal', String(e && e.message || e), ui.ButtonSet.OK);
  }
}

/* =====================================================================
 * KOSONGKAN SEMUA DATA (kecuali tab Config dan PIN)
 * - Tab laporan/rekap dihapus (dibuat ulang otomatis saat aplikasi menyimpan data baru).
 * - DB (data rak & muat aplikasi), DB_Backup, Log dikosongkan, header dipertahankan.
 * - Versi data dinaikkan agar perangkat yang terbuka menarik data kosong.
 * - Dua konfirmasi + salinan spreadsheet otomatis ke Google Drive sebelum menghapus.
 * CATATAN: perangkat yang masih menyimpan data lama di browser bisa mengirimnya lagi
 * ke server saat menyimpan. Setelah reset, bersihkan data lokal di tiap perangkat / buka aplikasi dan cek dulu.
 * ===================================================================== */
const RESET_DEL = ['Riset_Data', 'Daftar_Isi', 'Laporan_Mingguan', 'Laporan_Bulanan', 'Kesesuaian_Mingguan', 'Kesesuaian_Bulanan',
  'Adjustment_List', 'Adjustment_per_Rak', 'Belum_Dicek', 'Harian', 'Mingguan', 'Bulanan',
  'Selisih_Bulanan', 'Selisih_PIC', 'Selisih_SKU', 'Rak', 'Rak_Detail'];
const RESET_CLR = ['DB', 'DB_Backup', 'Log'];

function kosongkanSemuaData() {
  const ui = SpreadsheetApp.getUi(), ss = ss_();
  const a = ui.alert('KOSONGKAN SEMUA DATA',
    'Yang akan dihapus:\n• Data rak & muat di aplikasi (tab DB), DB_Backup, dan Log\n• Semua tab laporan/rekap: ' + RESET_DEL.join(', ') +
    '\n\nTab Config dan PIN TIDAK diubah.\nSalinan spreadsheet dibuat otomatis di Google Drive sebelum menghapus.\n\nLanjutkan?', ui.ButtonSet.YES_NO);
  if (a !== ui.Button.YES) return;
  const p = ui.prompt('Konfirmasi terakhir', 'Ketik KOSONGKAN (huruf besar semua) untuk melanjutkan:', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK || p.getResponseText().trim() !== 'KOSONGKAN') {
    ui.alert('Dibatalkan', 'Tidak ada data yang diubah.', ui.ButtonSet.OK); return;
  }
  let bk = '';
  try {
    bk = ss.copy(ss.getName() + ' - cadangan sebelum kosongkan ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm')).getUrl();
  } catch (e) {
    const c = ui.alert('Salinan otomatis GAGAL', String(e && e.message || e) + '\n\nLanjut TANPA cadangan? (Pilih TIDAK, lalu buat salinan manual lewat File > Buat salinan.)', ui.ButtonSet.YES_NO);
    if (c !== ui.Button.YES) return;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let dl = 0, cl = 0;
  try {
    RESET_DEL.forEach(n => { const s = ss.getSheetByName(n); if (s) { try { ss.deleteSheet(s); dl++; } catch (e) {} } });
    RESET_CLR.forEach(n => {
      const s = ss.getSheetByName(n);
      if (s && s.getLastRow() > 1) { s.getRange(2, 1, s.getLastRow() - 1, s.getMaxColumns()).clearContent(); cl++; }
    });
    const P = PropertiesService.getScriptProperties();
    ['rak', 'muat'].forEach(k => P.setProperty('REV_' + k, String(rev_(k) + 1)));
    SpreadsheetApp.flush();
    log_('RESET', 'SEMUA DATA', 0, dl + ' tab dihapus, ' + cl + ' tab dikosongkan, cadangan: ' + (bk || 'TIDAK ADA'));
  } finally {
    lock.releaseLock();
  }
  ui.alert('Selesai', dl + ' tab laporan dihapus, ' + cl + ' tab data dikosongkan. Config dan PIN tidak diubah.' + (bk ? '\n\nCadangan:\n' + bk : '\n\nTidak ada cadangan.'), ui.ButtonSet.OK);
}
