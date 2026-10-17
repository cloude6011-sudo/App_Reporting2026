# App Reporting 2026 (WMS)

Aplikasi pelaporan WMS berbasis **Google Apps Script** + web app (HTML).

## Link aplikasi
Web app (Apps Script): https://script.google.com/macros/s/AKfycbzVpOOlTmf1hnm7Mro5NEmwKixg2TcBvTjzOMVxbUn4tH3spxq3nuYs4QCmEHUrjrDIaQ/exec

Akses memakai PIN (tab `PIN` di spreadsheet database). Dashboard manager (baca saja): tambahkan `?v=dash` di akhir link.

## Struktur
| File | Fungsi |
|---|---|
| `src/Index.html` | Antarmuka web app (frontend), di Apps Script dinamai `Index` |
| `src/Code.gs` | Server utama: PIN/auth, penyimpanan DB, rekap, laporan mingguan/bulanan, kesesuaian rak |
| `src/Riset_Lengkap.gs` | Riset Data, Daftar Isi/rapikan tab, menu `Laporan WMS` (`onOpen`), kosongkan data |
| `src/Clearing_UI.gs` | Server untuk tab Clearing di aplikasi web (`clearingData`) |
| `src/Clearing_Sync.gs` | Sinkronisasi data Clearing dari spreadsheet sumber |
| `src/Telegram_Clearing.gs` | Pengiriman laporan Clearing ke Telegram |

## Deploy ke Apps Script
Salin isi file di `src/` ke project Apps Script (`Index.html` sebagai file HTML bernama `Index`), lalu deploy sebagai Web App.

## Catatan
- Hanya boleh ada satu `onOpen`; letaknya di `Riset_Lengkap.gs`.
- Script properties yang dipakai: `APP_PIN`, `VIEW_PIN`, `ALLOWED_EMAILS`, `SRC_ALLOW`, `SRC_ID`, `TG_TOKEN`, `TG_CHAT`, `TG_ON` (semua diisi di Project Settings, bukan di kode).

## Keamanan
- Token bot Telegram & chat ID disimpan di **Script Properties** (tidak ada di repo).
- Jangan commit token, PIN, atau file kredensial (lihat `.gitignore`).
