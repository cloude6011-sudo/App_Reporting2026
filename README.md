# App Reporting 2026 (WMS)

Aplikasi pelaporan WMS berbasis **Google Apps Script** + web app (HTML).

## Struktur
| File | Fungsi |
|---|---|
| `src/Index.html` | Antarmuka web app (frontend) |
| `src/Riset_Lengkap.gs` | Logika utama, setup sheet, menu laporan |
| `src/Clearing_Sync.gs` | Sinkronisasi data Clearing dari spreadsheet sumber |
| `src/Telegram_Clearing.gs` | Pengiriman laporan Clearing ke Telegram |

## Deploy ke Apps Script
Salin isi file di `src/` ke project Apps Script (`Index.html` sebagai file HTML bernama `Index`), lalu deploy sebagai Web App.

## Keamanan
- Token bot Telegram & chat ID disimpan di **Script Properties** (tidak ada di repo).
- Jangan commit token, PIN, atau file kredensial (lihat `.gitignore`).
