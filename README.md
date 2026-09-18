# Galeri Bahan Ajar Cyber AI

Website React + Vite berdasarkan PRD v3.0. Logo asli digunakan pada identitas situs dan favicon. Tidak ada data contoh yang dipublikasikan.

## Menjalankan

1. `npm ci`
2. Salin `.env.example` menjadi `.env`, lalu isi `VITE_APPS_SCRIPT_URL` dengan URL deployment Apps Script yang berakhiran `/exec`.
3. `npm run dev` untuk pratinjau; `npm run build` untuk menghasilkan `dist/`.
4. Hosting harus meneruskan rute yang tidak cocok dengan berkas ke `index.html` (SPA fallback).

## Backend Google Apps Script

Kode lengkap disediakan di `backend/Code.gs`. Kode ini belum otomatis dipasang pada akun Google Anda.

1. Cadangkan spreadsheet sebelum mengganti backend. Buka spreadsheet → Extensions → Apps Script.
2. Pasang `Code.gs` dan manifest `appsscript.json`. Untuk script standalone, set Script Property `SPREADSHEET_ID` ke spreadsheet yang dituju. Jangan menaruh ID ini di frontend.
3. Jalankan `setupDatabaseManual`. Ini memverifikasi header dan hanya mengisi header kosong bila kolom tersebut tidak memiliki data. Header yang berbeda akan menghentikan proses agar tidak menimpa skema lama.
4. Deploy sebagai Web App, Execute as pengelola, Who has access: Anyone. Aktifkan deployment versi baru ketika kode berubah.
5. Set URL `/exec` pada `.env` dan build ulang frontend. Vite menyimpan URL ini ke bundle; URL endpoint publik bukan secret.

Sheet wajib bernama **MateriGuruSD**, dengan urutan 17 kolom dari `HEADERS` di backend. Kolom O hanya disimpan server dan tidak disertakan dalam GET publik. Field formulir AI Studio dikirim pada POST create sesuai PRD; tidak disimpan dalam localStorage atau ditampilkan pada galeri.

## Keamanan dan operasional

- GET memakai daftar field publik eksplisit. POST create selalu membuat ID baru, mengabaikan ID/penghitung/unggulan dari client, dan menggunakan lock. Counter hanya memperbarui satu sel.
- Honeypot, payload maksimum 20 KB, validasi server, netralisasi formula, batas laju, serta deduplikasi 10 menit. Pengiriman ulang identik mengembalikan ID yang sudah tersimpan.
- Rate limit berdasarkan identitas ringan dapat dihindari; bukan autentikasi. Ada batas global create untuk mengurangi spam. Sesuaikan angka dengan trafik nyata.
- Suka menggunakan penanda browser setelah sukses; view dibatasi 30 menit per karya per browser. Angka bukan statistik audit.
- Karya hanya disimpan pada Google Sheets. Browser menyimpan respons di memori, bukan salinan permanen. Create berhasil memuat ulang daftar.
- Uji file Drive di jendela privat. Parser URL tidak dapat membuktikan file sudah dibagikan atau media eksternal selalu tersedia.
- Rollback: gunakan versi deployment frontend sebelumnya dan versi Apps Script sebelumnya; jangan mengubah urutan kolom. Cadangkan Sheet sebelum migrasi manual.

## Verifikasi

`npm test` menjalankan validator, filter, sort, sanitasi, serta simulasi Apps Script untuk privasi GET dan operasi tulis. Uji lokal tidak menggantikan uji Apps Script langsung. Pengujian produksi yang membuat karya nyata perlu dilakukan pengelola agar tidak menambahkan karya uji ke galeri.
