# RoboMANTAP HTMX KUIS — Vercel test package

Folder ini adalah kandidat deployment terpisah untuk layanan siswa KUIS.

## Penting sebelum dipakai siswa
- Ini paket uji awal, bukan jaminan siap produksi. Aplikasi memakai sesi ujian dalam memori (`STUDENT_SESSIONS`), yang dapat hilang saat serverless instance berubah/dingin.
- Harus memakai PostgreSQL yang dapat diakses dari Vercel; set `DATABASE_URL` (atau `POSTGRES_URL`) sebagai Environment Variable. Jangan commit kredensial.
- Set `GEMINI_API_KEY` atau `GEMINI_API_KEYS` bila endpoint AI hint/generator membutuhkannya.
- Uji seluruh alur: login/token, mulai ujian, simpan jawaban, heartbeat, anti-cheat, submit, hasil, gambar (TKA), dan AI hint.
- Vercel Serverless Functions memiliki batas ukuran paket dan batas eksekusi; `pandas` + `matplotlib` + Google GenAI mungkin membuat build terlalu besar/lambat. Bila gagal build atau timeout, gunakan platform ASGI persistent yang cocok, atau refactor dependensi lebih lanjut.

## Deploy
1. Buat repository GitHub terpisah untuk layanan ini, lalu upload isi folder `kuis` sebagai root repository (jangan upload folder induk `RoboMANTAP_HTMX_Vercel_Starter`).
2. Di Vercel: **Add New → Project → Import Git Repository**.
3. Framework Preset: **Other**. Root Directory: `./` (root repository). Build Command dan Output Directory dikosongkan jika Vercel mengizinkan; Vercel mendeteksi `api/index.py`.
4. Tambahkan Environment Variables: `DATABASE_URL` (wajib, URL PostgreSQL lengkap), `GEMINI_API_KEY` (jika AI dipakai), dan `STREAMLIT_URL` (opsional untuk link kembali).
5. Deploy, lalu buka URL proyek. Untuk TKA, endpoint root menampilkan portal TKA; untuk Kuis, portal Kuis.

## Dependency
`requirements.txt` sengaja hanya memuat dependensi yang tampak diperlukan untuk runtime FastAPI/AI engine. Build harus diverifikasi di Vercel; belum diuji terhadap database produksi.
