# Rak JDC Update

Peta kapasitas rak gudang JDC dalam **tampak samping**, dibuat dengan Streamlit dari file Excel layout rak (`Rak_JDC.xlsx`, sheet `Rak Existing`).

## Isi aplikasi

- **Tampak samping**: satu strip per rak, tiap kotak satu posisi pallet, tinggi tumpukan = jumlah level. Bisa dilihat dari barat atau timur, diwarnai menurut status atau jumlah level.
- **Potongan melintang**: profil tinggi antar rak, per posisi utara-selatan atau siluet tertinggi.
- **3D isometrik**: gambaran umum (pelengkap).
- **Rekap PP**: PP existing, PP baru, dan pembanding dengan tabel PP di sheet.

Kotak abu-abu di Excel dibaca sebagai racking baru.

## Menjalankan

```
pip install -r requirements.txt
streamlit run app.py
```

Lalu unggah file Excel di panel kiri. Untuk uji lokal tanpa unggah, set `RAK_XLSX=/path/Rak_JDC.xlsx`.

## Deploy

Streamlit Community Cloud: pilih repo ini, branch `main`, file utama `app.py`.

## Struktur

| File | Fungsi |
|---|---|
| `app.py` | Aplikasi Streamlit (unggah Excel, tampilkan peta) |
| `baca_excel.py` | Membaca layout rak dari Excel (kode rak, level, warna abu-abu) |
| `tampak_samping.html` | Tampilan interaktif (SVG) yang disematkan ke Streamlit |

Layout Excel harus sama dengan `Rak_JDC.xlsx`: kode rak A01-D09 di baris 11 dan 49, angka level di baris 15-88.
