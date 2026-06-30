# PayTrack — Panduan Deploy ke Cloudflare Pages (GRATIS)

## Isi Folder Ini
```
paytrack-deploy/
├── index.html          ← Entry point HTML
├── package.json        ← Dependensi & scripts
├── vite.config.js      ← Config build
├── .gitignore
└── src/
    ├── main.jsx        ← Entry point React
    └── App.jsx         ← Aplikasi utama (sudah ada password)
```

---

## 🔑 Ganti Password
Buka `src/App.jsx`, cari baris:
```js
const APP_PASSWORD = "payroll2025";
```
Ganti `"payroll2025"` dengan password pilihanmu, lalu simpan.

---

## 🚀 LANGKAH DEPLOY KE CLOUDFLARE PAGES (GRATIS)

### Persiapan (sekali saja)
1. Buat akun **GitHub** di https://github.com (gratis)
2. Buat akun **Cloudflare** di https://cloudflare.com (gratis)

### Upload ke GitHub
1. Buka https://github.com/new
2. Buat repository baru (misal: `paytrack`) → Private → klik **Create repository**
3. Di halaman repository, klik **uploading an existing file**
4. Upload **semua file** dari folder `paytrack-deploy/` ini
   - Pastikan struktur foldernya benar (ada folder `src/` di dalamnya)
5. Klik **Commit changes**

### Connect ke Cloudflare Pages
1. Login ke https://dash.cloudflare.com
2. Klik **Workers & Pages** → **Create application** → **Pages**
3. Klik **Connect to Git** → pilih repository `paytrack`
4. Isi pengaturan build:
   - **Framework preset**: Vite
   - **Build command**: `npm run build`
   - **Build output directory**: `dist`
5. Klik **Save and Deploy**
6. Tunggu 1-2 menit → websitemu online! 🎉

### URL Website
Cloudflare akan memberi URL seperti: `https://paytrack-xxxx.pages.dev`
Kamu bisa share URL ini ke orang yang tau password.

---

## 🔐 Cara Kerja Password
- Password **tersimpan di kode** (bukan database) — cukup aman untuk penggunaan internal
- Login bertahan selama tab browser tidak ditutup (session-based)
- Tutup browser = harus login lagi
- Kalau mau ganti password: ubah di `App.jsx` → push ke GitHub → otomatis deploy ulang

---

## 📱 Tes di Lokal (opsional)
Kalau mau coba dulu sebelum deploy:
1. Install Node.js dari https://nodejs.org
2. Buka terminal di folder ini
3. Jalankan: `npm install`
4. Jalankan: `npm run dev`
5. Buka `http://localhost:5173`
