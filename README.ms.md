# HasilKira — Penjana Lembaran Kerja Cukai LHDN

🌐 [English](README.md) · [中文](README.zh.md) · **Bahasa Melayu**

> **HasilKira** = *Hasil* (LHDN · pendapatan) + *Kira* (mengira) — "kira cukai & hasil anda". Sesuai untuk **mana-mana** pembayar cukai individu Malaysia.

Alat satu klik yang menggunakan Python untuk menjana lembaran kerja cukai Excel **tiga bahasa** (中文 / Bahasa Melayu / English) yang **boleh diguna semula setiap tahun**, bagi membantu **individu Malaysia** menyusun dan menganggar cukai pendapatan — untuk mereka yang ada **pendapatan penggajian, pendapatan perniagaan (cth. pekebun kecil sawit), pendapatan lain, atau apa-apa gabungan**.

Fail output dinamakan **`LHDN_Tax_Worksheet_YA<tahun>.xlsx`**, dengan `<tahun>` diisi automatik kepada **tahun semasa** (cth. `LHDN_Tax_Worksheet_YA2026.xlsx`), dan disimpan dalam folder **`LHDN report`**.

---

## Apa ini?
Alat **satu klik**. Ia menggunakan Python untuk menjana lembaran kerja cukai Excel yang **boleh diguna setiap tahun**, disimpan dalam folder `LHDN report`. Menyokong **tiga bahasa (Cina / Melayu / Inggeris)** dengan suis bahasa (VLOOKUP).

## Untuk siapa?
Gaji sahaja, gaji + perniagaan (cth. sawit), perniagaan sahaja, perniagaan + lain-lain, atau pendapatan lain sahaja.

## Cara guna
- **Windows**: klik dua kali **`generate.cmd`**.
- **Mac**: klik dua kali **`generate.command`** (kali pertama, jika muncul amaran "pembangun tidak dikenali": klik kanan → Open → Open).
- **Linux**: jalankan **`./generate.sh`** dalam terminal (jika gagal: `chmod +x generate.sh` dahulu).
- Jika Python atau openpyxl tiada, ia **memasang secara automatik** (perlu internet).
- Selesai — folder `LHDN report` terbuka dengan fail **`LHDN_Tax_Worksheet_YA<tahun>.xlsx`**.
- Buka di Excel, **pilih bahasa di penjuru atas**, isi **sel KUNING** sahaja.

> ⚠️ Nama fail guna tahun semasa; **menjalankan semula pada tahun yang sama akan menulis ganti** fail itu. Simpan salinan yang telah diisi di tempat lain.

## Apa di dalam Excel? (13 helaian)
Muka depan, Maklumat peribadi, Pendapatan penggajian, Perniagaan sawit (P&L), Log bulanan sawit, Jadual elaun modal/pertanian, Pendapatan lain, Pelepasan cukai, Pengiraan cukai (pemastautin/bukan, berasingan/bersama automatik), Banding Berasingan vs Bersama, Ringkasan pelbagai tahun, Senarai semak dokumen, Lists.

## Petua
- Sel formula **dikunci**; untuk ubah: **Review → Unprotect Sheet** (tiada kata laluan).

## ⚠️ Penting
Untuk **menyusun & menganggar** sahaja; **sahkan angka dengan LHDN / ejen cukai berlesen**. Pendapatan perniagaan guna **Borang B** (e-Filing **15 Julai**); gaji sahaja Borang BE (15 Mei). Simpan resit **7 tahun**.

---

## 📂 Struktur projek
```
HasilKira/
├── generate.cmd            ← Windows: klik dua kali
├── generate.command        ← macOS: klik dua kali
├── generate.sh             ← Linux (atau Terminal macOS): ./generate.sh
├── build_tax_workbook.py   ← penjana (Python + openpyxl)
├── requirements.txt        ← openpyxl
├── README.md               ← English
├── README.zh.md            ← 中文
├── README.ms.md            ← fail ini (Bahasa Melayu)
├── .gitattributes          ← paksa LF pada skrip shell (merentas OS)
├── .gitignore
└── LHDN report/            ← dicipta pada larian pertama; simpan LHDN_Tax_Worksheet_YA<tahun>.xlsx
```

> 🛠 **Menyunting skrip pelancar?** `generate.sh` dan `generate.command` mesti kekal **LF** (line-ending) dan **boleh laksana** (`chmod +x`); `.gitattributes` menguatkuasakan bahagian LF. Menyimpannya semula dengan **CRLF** Windows akan merosakkannya pada macOS/Linux (`bad interpreter: bash^M`).

## 🔧 Keperluan
- Windows 10/11, macOS, atau Linux
- Python 3.9+ dan `openpyxl` — **dipasang secara automatik** oleh pelancar jika tiada.

## ❓ Penyelesaian masalah
- **SmartScreen / antivirus beri amaran tentang pelancar** — ia hanya menjalankan Python secara setempat; klik *More info → Run anyway*, atau jalankan `python build_tax_workbook.py` sendiri.
- **macOS — "pembangun tidak dikenali" / "tidak dapat dibuka"** — klik kanan (Control-klik) **`generate.command`** → **Open** → **Open** (sekali sahaja). Atau buang flag kuarantin: `xattr -d com.apple.quarantine generate.command`.
- **macOS — klik dua kali tiada apa-apa atau "permission denied"** — bit boleh-laksana telah hilang (biasa apabila folder tiba sebagai **ZIP/salinan dari Windows**; `git clone` mengekalkannya). Dalam Terminal, jalankan sekali: `chmod +x generate.command generate.sh`, kemudian klik dua kali semula.
- **"Python installed, run again" (Windows)** — tutup tetingkap dan klik dua kali `generate.cmd` sekali lagi (Windows perlu menyegarkan PATH).
- **Ralat `externally-managed-environment` (Linux/macOS)** — `generate.sh` sudah mengasingkan segalanya dalam `.venv` setempat, jadi ini tidak sepatutnya berlaku.
- **Tiada internet** — pasang Python dari python.org, jalankan `pip install openpyxl`, kemudian lancarkan semula.
- **Jalankan secara manual** — buka terminal dalam folder ini dan jalankan: `python build_tax_workbook.py`.

---

## 🤖 Penyelenggara / pembantu AI
Undang-undang cukai Malaysia berubah setiap tahun. Untuk **panduan penuh** mengemas kini data cukai (jadual kadar, pelepasan, rebat, tarikh akhir e-Filing, dan kekangan ketat supaya fail tidak rosak), lihat bahagian *For AI assistants* dalam [README.md](README.md) (Inggeris).

*Penafian: Tidak bergabung dengan LHDN. Untuk rujukan sahaja; bukan nasihat cukai profesional. Sentiasa sahkan dengan LHDN atau ejen cukai berlesen sebelum memfailkan.*
