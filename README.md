# App Reporting 2026 (WMS)

Aplikasi pelaporan WMS berbasis **Google Apps Script** + web app (HTML).

## Hosting GitHub Pages
Frontend: **https://cloude6011-sudo.github.io/App_Reporting2026/** (dashboard manager: tambahkan `?v=dash`).

Frontend memanggil backend Apps Script lewat `fetch` (endpoint `doPost` di `src/Api.gs`). Satu file `Index.html` bisa dipakai di dua tempat: di Apps Script (`google.script.run`) dan di GitHub Pages (`fetch`).

**Setup backend (sekali):**
1. Tempel `src/Api.gs` sebagai file baru di project Apps Script.
2. Deploy > Manage deployments > Edit > Version: **New version**. Pengaturan: Execute as **Me**, Who has access **Anyone** (bukan "Anyone with Google account"). Link `/exec` tetap sama.
3. Bila link `/exec` berubah, ubah `API_URL` di `index.html` dan `src/Index.html`.

**Diagnosa:** bila aplikasi gagal terhubung, buka `cek.html` (https://cloude6011-sudo.github.io/App_Reporting2026/cek.html). Halaman itu menguji versi deploy, endpoint API, dan PIN, lalu menjelaskan langkah yang kurang.

**Catatan:** `index.html` (root, untuk Pages) dan `src/Index.html` (untuk Apps Script) harus identik; salin ulang bila salah satu diubah.

## Link aplikasi
Web app (Apps Script): https://script.google.com/macros/s/AKfycbwgZmFrt5XCgGk8_hIoeNCwRc0xiCnICbeDhY8HofgS5kbFLHWRk13zuoA9hnsj31sjuw/exec

Akses memakai PIN (tab `PIN` di spreadsheet database). Dashboard manager (baca saja): tambahkan `?v=dash` di akhir link.

## Struktur
| File | Fungsi |
|---|---|
| `index.html` | Salinan `src/Index.html` untuk GitHub Pages |
| `src/Api.gs` | Endpoint `doPost` (JSON) untuk frontend di luar Apps Script |
| `src/Index.html` | Antarmuka web app (frontend), di Apps Script dinamai `Index` |
| `src/Code.gs` | Server utama: PIN/auth, penyimpanan DB, rekap, laporan mingguan/bulanan, kesesuaian rak |
| `src/Riset_Lengkap.gs` | Riset Data, Daftar Isi/rapikan tab, menu `Laporan WMS` (`onOpen`), kosongkan data |
| `src/Telegram_Laporan.gs` | Kirim laporan apa pun ke Telegram (`kirimTelegram`, `kirimRiset`) + menu Sheet |
| `src/Clearing_UI.gs` | Server untuk tab Clearing di aplikasi web (`clearingData`) |
| `src/Clearing_Sync.gs` | Sinkronisasi data Clearing dari spreadsheet sumber |
| `src/Telegram_Clearing.gs` | Pengiriman laporan Clearing ke Telegram |

## Deploy ke Apps Script
Salin isi file di `src/` ke project Apps Script (`Index.html` sebagai file HTML bernama `Index`), lalu deploy sebagai Web App.

## Kirim laporan ke Telegram
Tekan **✈️ Telegram** (tanpa salin-tempel) di: Laporan (Cek Rak, Muat Harian, Muat Mingguan, Clearing), tab Clearing, Tarikan Adjustment, dan Cek rak bulan ini (belum dicek). Kartu **Riset Data** di tab Laporan dan menu Sheet `Laporan WMS > Kirim Riset ke Telegram` mengirim ringkasan KPI bulan ini.

- Atur sekali di Menu > Pengaturan Telegram (token bot + chat ID). Hanya PIN **tulis** yang bisa mengirim; mode manager (`?v=dash`) tidak menampilkan tombolnya.
- Pesan otomatis dirapikan (judul tebal, tabel monospace, dipecah bila panjang). Laporan yang sama dalam 90 detik ditahan agar tidak terkirim ganda (bisa dipaksa).
- Setelah memperbarui kode: tempel `Telegram_Laporan.gs`, ganti `Api.gs` dan `Riset_Lengkap.gs`, lalu Deploy > New version.

### Isi laporan yang dikirim
- **Muat Harian/Mingguan:** bagian *Belum ada hasil counting checker* (no lori, SPM, customer, PIC muat, item + qty untuk dicek) dan *Rincian & analisa kemungkinan* untuk tiap selisih (item tertukar, rak ter-scan dobel, salah ketik 10x, hitungan antar-PIC, dll).
- **Clearing:** voucher yang belum ada di CLEARING (dengan item/qty/rak/posisi dan deteksi salah ketik nomor), *no voucher sama tetapi qty beda* (qty laporan, CLEARING terbaru, CLEARING sebelumnya, selisih, analisa), kode item beda, dan voucher yang hanya ada di CLEARING. Pilih kolom qty/kode item CLEARING di tab Clearing (terdeteksi otomatis). Hasil qty beda ikut dicatat di tab Sheet Clearing (kolom Qty beda).

## Laporan Clearing (voucher vs hasil counting)
Upload laporan voucher di tab Clearing, lalu dicocokkan dengan tab CLEARING. Laporan (Telegram/Salin/WhatsApp) berisi:
- **Belum ada hasil counting**: no voucher, kode/nama item, qty, **no lori dan PIC**, serta analisa (mis. sudah N hari tanpa counting → prioritas).
- **No voucher sama, qty beda**: qty voucher (laporan) vs qty terbaru di CLEARING, riwayat qty sebelum → sesudah, dan analisa kemungkinan (salah ketik nol, 2×/½, revisi belum terbawa, input dobel, kode item beda).
- **Ada di CLEARING tapi tidak di laporan**.

No lori dan PIC diambil otomatis dari kolom file upload atau tab CLEARING bila headernya bernama mis. `No Lori`/`Nopol`/`Plat`/`Kendaraan` dan `PIC`/`Petugas`/`Checker`/`Operator`. Bila tidak ada, aplikasi memperkirakannya dari data Muat (kode item, qty, dan tanggal yang berdekatan) dan memberi label "perkiraan dari data Muat".

### Presisi pencocokan voucher
- Nomor voucher dibaca tahan format: apostrof, spasi, awalan huruf/`No.`, sufiks `-01`, banyak nomor dalam satu sel, nomor terselip (`12-1000009`), dan angka 0 di depan yang hilang.
- Kolom voucher dipilih otomatis berdasarkan jumlah yang benar-benar cocok (bukan hanya judul kolom).
- Voucher dengan beberapa baris item dicocokkan per kode item.
- Sebelum menyatakan "belum ada", seluruh kolom CLEARING dicari. Bila nomor ada di kolom lain, laporan menyebut kolom dan barisnya.
- Saran **salah ketik** hanya muncul bila nomor mirip itu belum dipakai di laporan dan kode item tidak bertentangan; tingkat keyakinannya ditampilkan (`item sama` / `qty sama` / `belum terverifikasi`).
- Tombol **✔ Validasi: benar salah ketik** mengonfirmasi saran itu. Voucher lalu dihitung cocok (persentase bisa menjadi 100%), tersimpan di perangkat, bisa dibatalkan, dan dicatat di bagian F laporan.

## Catatan
- Hanya boleh ada satu `onOpen`; letaknya di `Riset_Lengkap.gs`.
- Script properties yang dipakai: `APP_PIN`, `VIEW_PIN`, `ALLOWED_EMAILS`, `SRC_ALLOW`, `SRC_ID`, `TG_TOKEN`, `TG_CHAT`, `TG_ON` (semua diisi di Project Settings, bukan di kode).

## Keamanan
- Token bot Telegram & chat ID disimpan di **Script Properties** (tidak ada di repo).
- Jangan commit token, PIN, atau file kredensial (lihat `.gitignore`).
