# HasilKira — LHDN Tax Worksheet Generator

🌐 **English** · [中文](README.zh.md) · [Bahasa Melayu](README.ms.md)

> **HasilKira** = *Hasil* (LHDN / 收成 · income) + *Kira* (计算 · calculate) — "calculate your tax & your harvest". Works for **any** Malaysian individual taxpayer.

A one-click tool that generates a reusable, trilingual (中文 / Bahasa Melayu / English) Microsoft Excel worksheet to help **Malaysian individuals** organise and estimate their income tax — built for people with **employment income, business income (e.g. oil-palm smallholding), other income, or any mix**.

The output file is named **`LHDN_Tax_Worksheet_YA<year>.xlsx`**, where `<year>` is filled automatically with the **current year** (e.g. `LHDN_Tax_Worksheet_YA2026.xlsx`), and saved into the **`LHDN report`** subfolder.

---

## What is this?
A **one-click** tool. It uses Python to generate a **reusable, yearly** Excel tax worksheet into the `LHDN report` subfolder. The worksheet is **trilingual (Chinese / Malay / English)** with a VLOOKUP-driven language switch.

## Who is it for?
Employment income only, employment + business (e.g. oil-palm smallholding), business only, business + other, or other income only — all supported.

## How to use
- **Windows**: double-click **`generate.cmd`**.
- **Mac**: double-click **`generate.command`** (first time, if macOS warns about an "unidentified developer": right-click the file → Open → Open).
- **Linux**: run **`./generate.sh`** in a terminal (if it won't run: `chmod +x generate.sh` first).
- On first run, if Python or openpyxl is missing it will **install them automatically** (internet required).
  - On Windows, if it asks you to "close the window and double-click again" after installing Python, please do so.
- When done, the `LHDN report` folder opens with **`LHDN_Tax_Worksheet_YA<year>.xlsx`**.
- Open it in Excel, **pick a language at the top-right**, fill **YELLOW cells** only; grey cells auto-calculate.

> ⚠️ The filename uses the current year; **re-running in the same year overwrites** that file. Save your filled-in copy elsewhere.

## What's inside the Excel? (13 sheets)
Cover, Profile, Employment income, Palm-oil business P&L, Monthly FFB log, Capital/Agriculture allowance schedule, Other income, Tax reliefs (full list), Tax computation (resident/non-resident & separate/joint auto-switch), Separate-vs-Joint comparison, Multi-year summary, Documents checklist, Lists (language/tax engine).

## Tips
- Formula cells are **locked**; to change a formula use **Review → Unprotect Sheet** (no password).

## ⚠️ Important
For **organising & estimating** only — **confirm figures with LHDN / a licensed tax agent**. Business income files **Form B** (e-Filing deadline **15 July**); employment-only files Form BE (15 May). Keep all receipts for **7 years**.

---

## 📂 Project structure
```
HasilKira/
├── generate.cmd            ← Windows: double-click this
├── generate.command        ← macOS: double-click this
├── generate.sh             ← Linux (or macOS Terminal): ./generate.sh
├── build_tax_workbook.py   ← the generator (Python + openpyxl)
├── requirements.txt        ← openpyxl
├── README.md               ← this file (English)
├── README.zh.md            ← 中文
├── README.ms.md            ← Bahasa Melayu
├── .gitattributes          ← forces LF on the shell scripts (cross-OS)
├── .gitignore
└── LHDN report/            ← created on first run; holds LHDN_Tax_Worksheet_YA<year>.xlsx
```

> 🛠 **Editing the launchers?** `generate.sh` and `generate.command` must stay **LF** line-endings and **executable** (`chmod +x`); `.gitattributes` enforces the LF part. Re-saving them with Windows **CRLF** breaks them on macOS/Linux (`bad interpreter: bash^M`).

## 🔧 Requirements
- Windows 10/11, macOS, or Linux
- Python 3.9+ and `openpyxl` — **installed automatically** by the launcher if missing.

## ❓ Troubleshooting
- **SmartScreen / antivirus warns about the launcher** — it only runs Python locally; click *More info → Run anyway*, or run `python build_tax_workbook.py` yourself.
- **macOS — "unidentified developer" / "cannot be opened"** — right-click (Control-click) **`generate.command`** → **Open** → **Open** (only needed once). Or clear the quarantine flag: `xattr -d com.apple.quarantine generate.command`.
- **macOS — double-click does nothing or "permission denied"** — the executable bit was lost (common when the folder arrives as a **ZIP/copy made on Windows**; a `git clone` preserves it). In Terminal, run once: `chmod +x generate.command generate.sh`, then double-click again.
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
   - **Launchers** — keep `generate.sh` and `generate.command` as **LF** line-endings and **executable**; never convert them to CRLF.
5. **Verify before finishing**:
   - Run `python build_tax_workbook.py` and confirm it writes `LHDN report/LHDN_Tax_Worksheet_YA<year>.xlsx` with no error.
   - Unzip the `.xlsx` and confirm **every XML part parses** (well-formed) and contains **no** `SEQUENCE` / `_xlfn` / `ANCHORARRAY` tokens.
   - Sanity-check a few tax computations against the official rate table.
   - Update the YA / figures mentioned in **all three** README files (`README.md`, `README.zh.md`, `README.ms.md`) to match.

*Disclaimer: Not affiliated with LHDN. For reference only; not professional tax advice. Always confirm with LHDN or a licensed tax agent before filing.*
