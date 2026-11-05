# Panduan PWA & Aplikasi Android (APK) — Gudang Elektrik

Aplikasi Web Apps **Gudang Elektrik** kini telah dilengkapi dengan dukungan penuh **Progressive Web App (PWA)** dan **File APK Android Mandiri**.

---

## 🚀 1. File APK Android Siap Pakai (Instal Langsung di HP)

File APK yang siap diinstal langsung di smartphone/tablet Android operator dan teknisi sudah berhasil dibuat:

📍 **Lokasi File APK:**
`Gudang elektrik/GudangElektrik.apk` (Ukuran: ~12 MB)

### Cara Memasang di Ponsel Android:
1. Kirim file `GudangElektrik.apk` ke ponsel (bisa via WhatsApp, Google Drive, Telegram, atau transfer kabel USB).
2. Di ponsel Android, buka dan ketuk file **`GudangElektrik.apk`**.
3. Jika muncul konfirmasi keamanan:
   - Pilih **"Izinkan dari sumber ini"** (*Allow from this source*).
   - Jika Google Play Protect menampilkan pemberitahuan, ketuk **"Tetap Instal"** (*Install anyway*).
4. Aplikasi **Gudang Elektrik** akan langsung terpasang dengan ikon resmi, layar penuh (*fullscreen* tanpa kolom browser), izin kamera barcode scanner aktif, dan login operator tersimpan otomatis.

---

## 📱 2. Instalasi Langsung via Browser (PWA — Tanpa Download APK)

Jika teknisi membuka link Web App di browser:
🔗 **Link Web App:**
`https://script.google.com/macros/s/AKfycbzMZQ1Qngh1r4ddx-adIPl1-UTiIgrb219ErnyYlXQiVVuL8mSoV5HVez7BgXosFpuG/exec`

### Di Android (Google Chrome):
1. Buka link di atas pada Google Chrome HP.
2. Ketuk tombol **"App APK"** di header atas, atau ketuk menu titik tiga **(⋮)** di pojok kanan atas Chrome.
3. Pilih menu **"Instal Aplikasi"** atau **"Tambahkan ke Layar Utama"** (*Add to Home Screen*).
4. Aplikasi otomatis terinstal di daftar aplikasi HP Anda layaknya aplikasi native.

### Di iPhone / iPad (Apple Safari):
1. Buka link di atas menggunakan browser **Safari**.
2. Ketuk tombol **Bagikan (Share)** berupa ikon kotak panah atas **(⬆️)** di bilah bawah Safari.
3. Gulir ke bawah lalu ketuk **"Tambah ke Layar Utama"** (*Add to Home Screen*).
4. Ketuk **"Tambah"** di pojok kanan atas.

---

## 🌐 3. Paket PWA Mandiri (Untuk Hosting Sendiri & PWABuilder)

Folder `Gudang elektrik/pwa/` telah disiapkan lengkap:
- `index.html` (PWA Shell dengan loader, offline detection, dan responsive frame)
- `manifest.json` (Web App Manifest standar PWA)
- `sw.js` (Service Worker untuk caching offline)
- `icons/` (Ikon 192x192 dan 512x512)

### Jika Ingin Membangun APK via PWABuilder (Cloud):
1. Upload isi folder `Gudang elektrik/pwa/` ke hosting statis gratis (misalnya **GitHub Pages**, **Vercel**, atau **Firebase Hosting**).
2. Buka **[PWABuilder.com](https://www.pwabuilder.com)**.
3. Masukkan URL hosting Anda, lalu klik **Start**.
4. Klik **Package for Stores** -> pilih **Android** untuk mendownload APK / AAB resmi Google Play Store.

---

## 🛠️ 4. Source Code Android Studio

Source code lengkap aplikasi Android juga telah disimpan di:
📍 `Gudang elektrik/android-source/`
Dapat dibuka langsung menggunakan **Android Studio** kapan pun jika Anda ingin mengedit kode Kotlin atau mengompilasi ulang versi rilis (*release signed keystore*).

