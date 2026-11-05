# Gudang Elektrik — Inventory Management System

Sistem Manajemen & Monitoring Stok Gudang Elektrik berbasis Google Apps Script, Progressive Web App (PWA), dan Aplikasi Android Native (APK).

---

## ⚡ Tautan Penting

- 📱 **Unduh Aplikasi Android (APK):** [Download GudangElektrik.apk (v1.0.0)](https://github.com/haysan10/gudang-elektrik/releases/download/v1.0.0/GudangElektrik.apk)
- 🌐 **PWA Live Web App (GitHub Pages):** [https://haysan10.github.io/gudang-elektrik/](https://haysan10.github.io/gudang-elektrik/)
- ☁️ **Google Apps Script Web App Direct:** [Buka Web App GAS](https://script.google.com/macros/s/AKfycbzMZQ1Qngh1r4ddx-adIPl1-UTiIgrb219ErnyYlXQiVVuL8mSoV5HVez7BgXosFpuG/exec)

---

## 📋 Fitur Utama

1. **Monitoring Stok Real-Time (Tab 1):**
   - Menampilkan 10 kolom data standar inventaris pabrik: `Rak`, `Item Code`, `Spesifikasi`, `Stok Awal`, `Masuk (In)`, `Keluar (Out)`, `Stok Akhir`, `Satuan`, `Minimal Stok`, dan `Remarks`.
   - Indikator visual stok aman (Hijau), menipis/warning (Kuning), dan habis/kritis (Merah).
   - Filter cepat berdasarkan lokasi Rak dan pencarian teks live.
   - Cetak Kartu Stok (Bincard A5) siap print.

2. **Form Mutasi Masuk/Keluar Otomatis (Tab 2):**
   - Transaksi mutasi barang dengan auto-resolve kode dan spesifikasi barang tanpa perlu memilih gudang manual.
   - Barcode & QR Code Scanner via kamera HP atau webcam laptop.
   - Stepper kuantitas praktis (+1, +5, +10, +50, MAX).
   - Validasi overdraft otomatis agar stok tidak minus.
   - Profil NIK operator dan Unit/Seksi tersimpan otomatis di perangkat.

3. **Laporan & Riwayat Mutasi (Tab 3):**
   - Riwayat transaksi mutasi lengkap dengan ID Transaksi, tanggal/waktu, NIK, dan perubahan stok sebelum & sesudah mutasi.
   - Filter rentang tanggal dan pencarian.
   - Ekspor data laporan ke format file CSV/Excel.

4. **Panel Administrator Terlindungi PIN 6 Digit (Tab 4):**
   - Akses aman dengan PIN 6-digit (Default: `100700` atau dapat diganti di menu Pengaturan).
   - Pengeditan master database barang (Tambah, Ubah, Hapus item).
   - Inisialisasi otomatis 1.006 data master barang dari spreadsheet kosong.
   - Pengaturan notifikasi alert ke Bot Telegram dan Email.

5. **Aplikasi Android & Progressive Web App (PWA):**
   - File APK mandiri (`GudangElektrik.apk`) siap pasang langsung di Android.
   - Dukungan PWA untuk instalasi instan di Google Chrome (Android/Desktop) dan Apple Safari (iOS).

---

## 📂 Struktur Proyek

```
gudang-elektrik/
├── code.gs                  # Backend Google Apps Script (CRUD, Mutasi, Notifikasi, Seeding)
├── index.html               # Frontend Web Application (HTML5, CSS3 Glassmorphism, JS Controller)
├── GudangElektrik.apk       # File installer aplikasi Android siap pakai (~12 MB)
├── PWA_DAN_APK_GUIDE.md     # Buku panduan lengkap PWA & cara pasang APK di ponsel
├── docs/                    # Berkas PWA untuk GitHub Pages (Live Web App)
│   ├── index.html           # PWA Shell responsive loader
│   ├── manifest.json        # Web App Manifest PWA
│   ├── sw.js                # Service Worker caching offline
│   └── icons/               # Ikon aplikasi 192x192 & 512x512
├── pwa/                     # Standalone PWA assets
├── android-source/          # Source code lengkap proyek Kotlin Android Studio
│   ├── app/
│   │   ├── src/main/AndroidManifest.xml
│   │   └── src/main/kotlin/.../MainActivity.kt
│   └── build.gradle.kts
└── gudang elektrik.xlsx     # Data master referensi 1.006 barang Gudang Elektrik
```

---

## 📱 Panduan Pemasangan

### Opsi A: Instal File APK di Android
1. Download file [`GudangElektrik.apk`](https://github.com/haysan10/gudang-elektrik/releases/download/v1.0.0/GudangElektrik.apk).
2. Ketuk file di HP Android Anda, lalu pilih **Instal** (*Pilih "Tetap Instal" jika muncul Google Play Protect*).

### Opsi B: Pasang via PWA Browser
1. Buka [https://haysan10.github.io/gudang-elektrik/](https://haysan10.github.io/gudang-elektrik/) di browser ponsel.
2. Di Chrome: Ketuk menu titik tiga (⋮) -> **"Instal Aplikasi"** atau **"Tambahkan ke Layar Utama"**.
3. Di Safari (iPhone): Ketuk tombol Share (⬆️) -> **"Tambah ke Layar Utama"**.

---

## 📄 Lisensi
Hak Cipta © 2026 PT Semen Padang / Tim Gudang Elektrik.
