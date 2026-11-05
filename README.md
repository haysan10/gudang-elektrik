# Gudang Elektrik — Inventory Management System

Sistem Manajemen & Monitoring Stok Gudang Elektrik berbasis Google Apps Script, Progressive Web App (PWA), dan Aplikasi Android Native (APK).

---

## ⚡ Tautan Penting

- 📱 **Unduh Aplikasi Android (APK):** [Download GudangElektrik.apk (v1.0.0)](https://github.com/haysan10/gudang-elektrik/releases/download/v1.0.0/GudangElektrik.apk)
- 🌐 **PWA Live Web App (GitHub Pages):** [https://haysan10.github.io/gudang-elektrik/](https://haysan10.github.io/gudang-elektrik/)
- ☁️ **Google Apps Script Web App Direct:** [Buka Web App GAS](https://script.google.com/macros/s/AKfycbzMZQ1Qngh1r4ddx-adIPl1-UTiIgrb219ErnyYlXQiVVuL8mSoV5HVez7BgXosFpuG/exec)

---

## 📋 Fitur Utama & Pembaruan Sistem

### 1. Monitoring Stok Seamless & Fluid (Tab 1)
- **Tampilan Tanpa Terpotong Frame:** Desain tabel fluida otomatis (`data-grid`) yang menyesuaikan lebar layar perangkat secara mulus tanpa terpotong bingkai (*card clipping*).
- **Sticky Freeze Columns:** 3 kolom pertama (`No`, `Lokasi Rak`, `Item Code`) tetap terkunci (*freeze*) di sisi kiri saat tabel digeser secara horizontal di layar kecil/ponsel, dengan bayangan pemisah lembut.
- **Dual View Mode Switcher:** Pilihan instan antara mode **[ ▤ Tabel ]** untuk audit data padat dan **[ ▦ Kartu Grid ]** untuk kartu visual interaktif di desktop, tablet, maupun ponsel.
- **10 Parameter Standar Database Gudang Elektrik:** Sesuai referensi master: `Rak`, `Item Code`, `Spesifikasi`, `Stok Awal`, `Masuk (In)`, `Keluar (Out)`, `Stok Akhir`, `Satuan`, `Minimal Stok`, dan `Remarks`.
- **Indikator Visual Dinamis:** Badge status otomatis `AMAN` (Hijau), `KRITIS` (Merah muda jika stok akhir ≤ batas minimal), dan `HABIS` (Abu-abu jika stok 0).

### 2. Form Mutasi Lapangan Otomatis (Tab 2)
- **Auto-Resolve Cerdas:** Lokasi Rak dan ketersediaan stok fisik terhubung otomatis saat barang dipilih, tanpa perlu memilih gudang manual.
- **Multi-Input Transaksi:** Menambah banyak baris barang sekaligus dalam satu transaksi mutasi.
- **Pencarian Cepat:** Pencarian live berbasis kata kunci jamak atau pemindai Barcode/QR Code via kamera HP/webcam.
- **Validasi Anti-Minus:** Peringatan otomatis jika kuantitas pengambilan (OUT) melebihi sisa stok fisik di gudang.

### 3. Sistem Bincard A5 Presisi & Analisis Geometri Duplex A4 (Potong 2)
- **Efisiensi Kertas A4:** 1 lembar kertas A4 Lanskap yang dicetak bolak-balik (*duplex*) dan dipotong dua tepat di garis tengah (`148.5 mm`) menghasilkan **2 Kartu Stok A5** siap pakai dengan 0% kertas terbuang.
- **Kalkulasi Geometri Duplex Akurat:**
  - Saat mencetak duplex mode lanskap dengan opsi printer *"Flip on Short Edge"*, kertas dibalik horizontal terhadap sumbu vertikal sehingga posisi kiri dan kanan bertukar fisik.
  - Sisi Depan (Sheet Front): `[Kiri: Item A Depan] |✂️| [Kanan: Item B Depan]`
  - Sisi Belakang (Sheet Back): `[Kiri: Item B Belakang] |✂️| [Kanan: Item A Belakang]`
  - Dengan formula ini, saat lembar dipotong di tengah, bagian belakang Item A tepat berada di balik Item A, dan bagian belakang Item B tepat berada di balik Item B!
- **Barcode Vektor Code128:** Barcode SVG resolusi tinggi tanpa ketergantungan API pihak ketiga, langsung terbaca oleh barcode scanner optik di gudang.
- **Format Standar ISO PT Semen Padang (`F-ELK-WHS-04 | Rev. 02`):**
  - **Sisi A:** Kop Resmi, Barcode Vektor, Grid Spesifikasi & Batas Min, serta 16 Baris Transaksi (Baris 1: Saldo Awal Sistem, Baris 2–16: Baris bertinta pulpen fisik).
  - **Sisi B:** Kop Lanjutan, 10 Baris Transaksi Fisik, **Tabel Audit Stock Opname Fisik** (4 kali rekonsiliasi SO dengan kolom Selisih & Paraf Auditor), serta **Kotak Pengesahan & Verifikasi** Petugas Gudang & Supervisor Pemeliharaan.

### 4. Pusat Laporan Data Eksekutif & Cetak Massal (Tab 3)
- **4 Kartu KPI Eksekutif:** Ringkasan instan `Total Mutasi Periode`, `Barang Masuk (IN)`, `Barang Keluar (OUT)`, dan `Item Defisit/Kritis`.
- **4 Pilihan Mode Laporan:**
  1. *Riwayat Transaksi Mutasi:* Log historis mutasi per rentang tanggal, lokasi rak, dan tipe transaksi.
  2. *Analisis Stok Kritis & Defisit:* Daftar barang yang membutuhkan restock segera dengan kalkulasi defisit kuantitas.
  3. *Rekapitulasi Stok per Rak:* Distribusi total SKU, total kuantitas fisik, dan tingkat kesehatan stok per rak.
  4. *Cetak Bincard Massal per Rak:* Memilih lokasi rak untuk mencetak seluruh bincard barang di rak tersebut sekaligus dalam format A4 duplex potong dua.
- **Cetak Laporan Resmi ISO (`LAP-ELK-SYS`):** Format dokumen cetak A4 ber-Kop resmi PT Semen Padang dengan 3 kolom tanda tangan pengesahan (*Dibuat Oleh Admin Gudang, Diperiksa Oleh Penyelia Listrik, Mengetahui Kepala Seksi Pemeliharaan*).
- **Export Excel Ramah Karakter (`.csv` UTF-8 BOM):** Disertai *Byte Order Mark* (`\uFEFF`) agar data terbuka rapi di Microsoft Excel tanpa masalah karakter aksen atau kolom acak.

### 5. Panel Administrator Terlindungi PIN 6 Digit (Tab 4)
- Akses aman dengan PIN 6-digit (Default: `100700` atau dapat diganti di menu Pengaturan).
- Manajemen master database barang (Tambah, Ubah, Hapus item).
- Inisialisasi otomatis 1.006 data master barang dari spreadsheet kosong.
- Pengaturan notifikasi alert ke Bot Telegram dan Email.

---

## 📂 Struktur Proyek

```
gudang-elektrik/
├── code.gs                  # Backend Google Apps Script (CRUD, Duplex Bincard, Barcode SVG, Laporan)
├── index.html               # Frontend Web Application (Seamless Data-Grid, View Switcher, PWA Modal)
├── GudangElektrik.apk       # File installer aplikasi Android native siap pakai (~12 MB)
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

## 🖨️ Panduan Cetak Bincard A4 Duplex (Potong 2 Jadi A5)

1. Buka menu **Laporan Data** -> Pilih **Cetak Bincard Massal** (atau klik ikon 🖨️ pada barang di Monitoring).
2. Pada dialog cetak peramban / printer:
   - **Ukuran Kertas (Paper Size):** `A4`
   - **Orientasi (Orientation):** `Landscape` (Mendatar)
   - **Cetak Bolak-Balik (Two-Sided / Duplex):** Centang `Aktif`
   - **Opsi Pembalikan (Duplex Flip):** Pilih **"Flip on Short Edge" (Balik pada sisi pendek)**
   - **Margin:** `None` / `Minimum` (0 mm)
   - **Skala:** `100%` / `Fit to Printable Area`
3. Ambil hasil cetak, potong kertas tepat di garis putus-putus vertikal tengah (`✂️`). Anda mendapatkan 2 kartu stok A5 bolak-balik yang presisi dan rapi!

---

## 📱 Panduan Pemasangan Aplikasi

### Opsi A: Instal File APK di Android
1. Download file [`GudangElektrik.apk`](https://github.com/haysan10/gudang-elektrik/releases/download/v1.0.0/GudangElektrik.apk).
2. Ketuk file di HP Android Anda, lalu pilih **Instal** (*Pilih "Tetap Instal" jika muncul notifikasi Google Play Protect*).

### Opsi B: Pasang via PWA Browser
1. Buka [https://haysan10.github.io/gudang-elektrik/](https://haysan10.github.io/gudang-elektrik/) di browser ponsel.
2. Di Chrome: Ketuk menu titik tiga (⋮) -> **"Instal Aplikasi"** atau **"Tambahkan ke Layar Utama"**.
3. Di Safari (iPhone): Ketuk tombol Share (⬆️) -> **"Tambah ke Layar Utama"**.

---

## 📄 Lisensi
Hak Cipta © 2026 PT Semen Padang / Tim Gudang Elektrik.
