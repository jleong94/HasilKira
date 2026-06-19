# HasilKira — LHDN Tax Worksheet Generator
**马来西亚个人所得税计算表生成器 · Penjana Lembaran Kerja Cukai LHDN**

> **HasilKira** = *Hasil* (LHDN / 收成 · income) + *Kira* (计算 · calculate) — "calculate your tax & your harvest". Works for **any** Malaysian individual taxpayer.

A one-click tool that generates a reusable, trilingual (中文 / Bahasa Melayu / English) Microsoft Excel worksheet to help **Malaysian individuals** organise and estimate their income tax — built for people with **employment income, business income (e.g. oil-palm smallholding), other income, or any mix**.

The output file is named **`LHDN_Tax_Worksheet_YA<year>.xlsx`**, where `<year>` is filled automatically with the **current year** (e.g. `LHDN_Tax_Worksheet_YA2026.xlsx`), and saved into the **`LHDN report`** subfolder.

> 🌐 Languages in this README: [中文](#中文) · [Bahasa Melayu](#bahasa-melayu) · [English](#english)
> 🤖 AI assistants: see [Keeping this up to date](#-for-ai-assistants--keeping-this-up-to-date) at the end.

---

<a name="中文"></a>
## 🇨🇳 中文

### 这是什么？
一个**一键生成**工具。运行后会用 Python 自动产生一份**可每年重复使用**的 Excel 报税计算表，存到 `LHDN report` 子文件夹里。表格**支持中／马／英三语切换**（含 VLOOKUP 清单）。

### 适合谁？
单一受雇收入、受雇＋业务（如棕油）、纯业务、业务＋其它、或只有其它收入——都适用。

### 怎么用？
- **Windows**:双击 **`generate.cmd`**。
- **Mac / Linux**:在终端机运行 **`./generate.sh`**（首次需 `chmod +x generate.sh`）。
- 第一次运行若缺 Python 或 openpyxl，它会**自动安装**（需要网络）。
  - Windows 若提示"装好 Python 后关窗再双击一次"，照做即可。
- 完成后会打开 `LHDN report` 文件夹，里面是 **`LHDN_Tax_Worksheet_YA<当年>.xlsx`**。
- 用 Excel 打开，**右上角选语言**，只填**黄色格**，灰色格自动计算。

> ⚠️ 文件名用当年年份;**同一年再跑一次会覆盖**同名文件。请把填好数据的表**另存到别处**。

### Excel 里有什么？（共 13 个分页）
封面、个人资料、就业收入、棕油业务 P&L、棕油月度记录、资本/农业免税额表、其它收入、扣税减免（全套）、税务计算（居民/非居民、单独/联合自动切换）、单独 vs 联合对比、多年汇总、单据清单、Lists（语言/税率引擎）。

### 小提示
- 公式格已**锁定**防误改;要改公式请到 Excel「审阅 → 撤销工作表保护」（无密码）。

### ⚠️ 重要
本工具仅供**整理与估算**，**数字以 LHDN 官方与持牌税务师为准**。有业务收入者用 **Borang B**（e-Filing 截止 **7 月 15 日**）;纯受雇用 Borang BE（5 月 15 日）。保留所有单据 **7 年**。

---

<a name="bahasa-melayu"></a>
## 🇲🇾 Bahasa Melayu

### Apa ini?
Alat **satu klik**. Ia menggunakan Python untuk menjana lembaran kerja cukai Excel yang **boleh diguna setiap tahun**, disimpan dalam folder `LHDN report`. Menyokong **tiga bahasa (Cina / Melayu / Inggeris)** dengan suis bahasa (VLOOKUP).

### Untuk siapa?
Gaji sahaja, gaji + perniagaan (cth. sawit), perniagaan sahaja, perniagaan + lain-lain, atau pendapatan lain sahaja.

### Cara guna
- **Windows**: klik dua kali **`generate.cmd`**.
- **Mac / Linux**: jalankan **`./generate.sh`** dalam terminal (kali pertama: `chmod +x generate.sh`).
- Jika Python atau openpyxl tiada, ia **memasang secara automatik** (perlu internet).
- Selesai — folder `LHDN report` terbuka dengan fail **`LHDN_Tax_Worksheet_YA<tahun>.xlsx`**.
- Buka di Excel, **pilih bahasa di penjuru atas**, isi **sel KUNING** sahaja.

> ⚠️ Nama fail guna tahun semasa; **menjalankan semula pada tahun yang sama akan menulis ganti** fail itu. Simpan salinan yang telah diisi di tempat lain.

### Apa di dalam Excel? (13 helaian)
Muka depan, Maklumat peribadi, Pendapatan penggajian, Perniagaan sawit (P&L), Log bulanan sawit, Jadual elaun modal/pertanian, Pendapatan lain, Pelepasan cukai, Pengiraan cukai (pemastautin/bukan, berasingan/bersama automatik), Banding Berasingan vs Bersama, Ringkasan pelbagai tahun, Senarai semak dokumen, Lists.

### Petua
- Sel formula **dikunci**; untuk ubah: **Review → Unprotect Sheet** (tiada kata laluan).

### ⚠️ Penting
Untuk **menyusun & menganggar** sahaja; **sahkan angka dengan LHDN / ejen cukai berlesen**. Pendapatan perniagaan guna **Borang B** (e-Filing **15 Julai**); gaji sahaja Borang BE (15 Mei). Simpan resit **7 tahun**.

---

<a name="english"></a>
## 🇬🇧 English

### What is this?
A **one-click** tool. It uses Python to generate a **reusable, yearly** Excel tax worksheet into the `LHDN report` subfolder. The worksheet is **trilingual (Chinese / Malay / English)** with a VLOOKUP-driven language switch.

### Who is it for?
Employment income only, employment + business (e.g. oil-palm smallholding), business only, business + other, or other income only — all supported.

### How to use
- **Windows**: double-click **`generate.cmd`**.
- **Mac / Linux**: run **`./generate.sh`** in a terminal (first time: `chmod +x generate.sh`).
- On first run, if Python or openpyxl is missing it will **install them automatically** (internet required).
  - On Windows, if it asks you to "close the window and double-click again" after installing Python, please do so.
- When done, the `LHDN report` folder opens with **`LHDN_Tax_Worksheet_YA<year>.xlsx`**.
- Open it in Excel, **pick a language at the top-right**, fill **YELLOW cells** only; grey cells auto-calculate.

> ⚠️ The filename uses the current year; **re-running in the same year overwrites** that file. Save your filled-in copy elsewhere.

### What's inside the Excel? (13 sheets)
Cover, Profile, Employment income, Palm-oil business P&L, Monthly FFB log, Capital/Agriculture allowance schedule, Other income, Tax reliefs (full list), Tax computation (resident/non-resident & separate/joint auto-switch), Separate-vs-Joint comparison, Multi-year summary, Documents checklist, Lists (language/tax engine).

### Tips
- Formula cells are **locked**; to change a formula use **Review → Unprotect Sheet** (no password).

### ⚠️ Important
For **organising & estimating** only — **confirm figures with LHDN / a licensed tax agent**. Business income files **Form B** (e-Filing deadline **15 July**); employment-only files Form BE (15 May). Keep all receipts for **7 years**.

---

## 📂 Project structure
```
HasilKira/
├── generate.cmd            ← Windows: double-click this
├── generate.sh             ← macOS / Linux: run ./generate.sh
├── build_tax_workbook.py   ← the generator (Python + openpyxl)
├── requirements.txt        ← openpyxl
├── README.md               ← this file
├── .gitignore
└── LHDN report/            ← created on first run; holds LHDN_Tax_Worksheet_YA<year>.xlsx
```

## 🔧 Requirements
- Windows 10/11, macOS, or Linux
- Python 3.9+ and `openpyxl` — **installed automatically** by the launcher if missing.

## ❓ Troubleshooting
- **SmartScreen / antivirus warns about the launcher** — it only runs Python locally; click *More info → Run anyway*, or run `python build_tax_workbook.py` yourself.
- **"Python installed, run again" (Windows)** — close the window and double-click `generate.cmd` once more (Windows needs to refresh PATH).
- **`externally-managed-environment` error (Linux/macOS)** — `generate.sh` already isolates everything in a local `.venv`, so this should not occur.
- **No internet** — install Python from python.org, then run `pip install openpyxl`, then launch again.
- **Run manually** — open a terminal in this folder and run: `python build_tax_workbook.py`

---

## 🤖 For AI assistants — keeping this up to date
*（此节是给 AI 大模型读的 / Bahagian ini untuk model AI）*

**Malaysian tax law changes every year** (Budget / *Belanjawan* announcements change relief items, caps, sub-limits, tax-bracket rates, rebates and deadlines). If you are an AI assistant asked to update or regenerate this project, **do not rely on your training data** — it may be stale. Instead:

1. **Fetch the latest official figures from the web**, prioritising:
   - **LHDN official site — https://www.hasil.gov.my** (Individual → Tax Rate / Relief / Rebate; e-Filing deadlines).
   - The latest **Malaysian Budget (Belanjawan)** announcement for the relevant Year of Assessment.
   - Cross-check with reputable secondary sources only to confirm, never as the primary source.
2. **Determine the correct Year of Assessment (YA)** being filed (usually the *previous* calendar year) and confirm which figures apply to it.
3. **Update `build_tax_workbook.py` precisely** — all tax data lives in clearly-marked Python structures; change values, do **not** restructure:
   - **Tax-bracket table** → the `bands` list (each tuple = `(lower_bound, cumulative_tax_at_lower_bound, marginal_rate)`). Recompute the cumulative `base` column if rates change.
   - **Reliefs** → the `reliefs` list (each tuple = `(label_key, max_cap, qty_editable, default_claim)`). Add/remove items and update caps & sub-limits.
   - **Relief / label wording** → the `LABELS` list (keep all three languages: ZH, BM, EN; keep the `key` stable).
   - **Rebates** → the `IF(...<=35000, 400, ...)` rebate logic and the RM35,000 threshold in the Tax-computation section.
   - **Non-resident flat rate** → the `0.30*` factor in the `tax_charged` formula.
   - **Deadlines & notes** → the `15 July` / `15 May` dates and any relief notes in this README and in the relevant `LABELS` note rows.
4. **Hard constraints — do not break the file** (these caused real corruption before):
   - Use **only legacy, widely-supported Excel functions** (SUM, IF, MIN, MAX, LOOKUP, VLOOKUP, MATCH, AND, IFERROR, ROUND, YEAR, TODAY).
   - **Never** introduce dynamic-array functions or spill syntax (`SEQUENCE`, `FILTER`, `XLOOKUP`, the `#` spill operator, `_xlfn.*`) — `openpyxl` does not serialise them correctly and Excel will report the file as corrupt.
   - Keep the trilingual VLOOKUP design and the yellow-input / grey-formula / locked-formula conventions intact.
5. **Verify before finishing**:
   - Run `python build_tax_workbook.py` and confirm it writes `LHDN report/LHDN_Tax_Worksheet_YA<year>.xlsx` with no error.
   - Unzip the `.xlsx` and confirm **every XML part parses** (well-formed) and contains **no** `SEQUENCE` / `_xlfn` / `ANCHORARRAY` tokens.
   - Sanity-check a few tax computations against the official rate table.
   - Update the YA / figures mentioned in this README to match.

*Disclaimer: Not affiliated with LHDN. For reference only; not professional tax advice. Always confirm with LHDN or a licensed tax agent before filing.*
