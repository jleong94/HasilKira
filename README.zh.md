# HasilKira — 马来西亚个人所得税计算表生成器

🌐 [English](README.md) · **中文** · [Bahasa Melayu](README.ms.md)

> **HasilKira** = *Hasil*（LHDN · 收成）+ *Kira*（计算）——“计算你的税务与收成”。适用于**任何**马来西亚个人纳税人。

一个一键工具，用 Python 生成一份**可每年重复使用**、三语（中文 / Bahasa Melayu / English）的 Microsoft Excel 报税计算表，帮助**马来西亚个人**整理与估算所得税——适用于**受雇收入、业务收入（如棕油小园主）、其它收入或任意组合**。

输出文件名为 **`LHDN_Tax_Worksheet_YA<年份>.xlsx`**，`<年份>` 自动填入**当年**（例如 `LHDN_Tax_Worksheet_YA2026.xlsx`），保存在 **`LHDN report`** 子文件夹。

---

## 这是什么？
一个**一键生成**工具。运行后会用 Python 自动产生一份**可每年重复使用**的 Excel 报税计算表，存到 `LHDN report` 子文件夹里。表格**支持中／马／英三语切换**（含 VLOOKUP 清单）。

## 适合谁？
单一受雇收入、受雇＋业务（如棕油）、纯业务、业务＋其它、或只有其它收入——都适用。

## 怎么用？
- **Windows**：双击 **`generate.cmd`**。
- **Mac**：双击 **`generate.command`**（首次若系统提示「身份不明的开发者」，右键点该文件 → 打开 → 打开）。
- **Linux**：在终端机运行 **`./generate.sh`**（若无法运行先 `chmod +x generate.sh`）。
- 第一次运行若缺 Python 或 openpyxl，它会**自动安装**（需要网络）。
  - Windows 若提示“装好 Python 后关窗再双击一次”，照做即可。
- 完成后会打开 `LHDN report` 文件夹，里面是 **`LHDN_Tax_Worksheet_YA<当年>.xlsx`**。
- 用 Excel 打开，**右上角选语言**，只填**黄色格**，灰色格自动计算。

> ⚠️ 文件名用当年年份；**同一年再跑一次会覆盖**同名文件。请把填好数据的表**另存到别处**。

## Excel 里有什么？（共 13 个分页）
封面、个人资料、就业收入、棕油业务 P&L、棕油月度记录、资本/农业免税额表、其它收入、扣税减免（全套）、税务计算（居民/非居民、单独/联合自动切换）、单独 vs 联合对比、多年汇总、单据清单、Lists（语言/税率引擎）。

## 小提示
- 公式格已**锁定**防误改；要改公式请到 Excel「审阅 → 撤销工作表保护」（无密码）。

## ⚠️ 重要
本工具仅供**整理与估算**，**数字以 LHDN 官方与持牌税务师为准**。有业务收入者用 **Borang B**（e-Filing 截止 **7 月 15 日**）；纯受雇用 Borang BE（5 月 15 日）。保留所有单据 **7 年**。

---

## 📂 项目结构
```
HasilKira/
├── generate.cmd            ← Windows：双击这个
├── generate.command        ← macOS：双击这个
├── generate.sh             ← Linux（或 macOS 终端机）：./generate.sh
├── build_tax_workbook.py   ← 生成器（Python + openpyxl）
├── requirements.txt        ← openpyxl
├── README.md               ← English
├── README.zh.md            ← 本文件（中文）
├── README.ms.md            ← Bahasa Melayu
├── .gitattributes          ← 强制 shell 脚本用 LF（跨系统）
├── .gitignore
└── LHDN report/            ← 首次运行时生成；存放 LHDN_Tax_Worksheet_YA<年份>.xlsx
```

> 🛠 **要修改启动脚本？** `generate.sh` 与 `generate.command` 必须保持 **LF** 换行并具**可执行权限**（`chmod +x`）；`.gitattributes` 负责强制 LF。若用 Windows 的 **CRLF** 换行重新保存，会导致它们在 macOS/Linux 上报错（`bad interpreter: bash^M`）。

## 🔧 系统需求
- Windows 10/11、macOS 或 Linux
- Python 3.9+ 与 `openpyxl`——若缺失，启动脚本会**自动安装**。

## ❓ 疑难排解
- **SmartScreen / 杀毒软件警告启动脚本** —— 它只在本机运行 Python；点 *More info → Run anyway*，或自行运行 `python build_tax_workbook.py`。
- **macOS —— “身份不明的开发者” / “无法打开”** —— 右键（按住 Control 点击）**`generate.command`** → **打开** → **打开**（只需一次）。或清除隔离属性：`xattr -d com.apple.quarantine generate.command`。
- **macOS —— 双击没反应或 “permission denied”** —— 可执行权限丢失了（常见于从 Windows 打包的 **ZIP/复制**过来；用 `git clone` 则会保留）。在终端机运行一次：`chmod +x generate.command generate.sh`，再双击。
- **“Python installed, run again”（Windows）** —— 关闭窗口，再双击一次 `generate.cmd`（Windows 需要刷新 PATH）。
- **`externally-managed-environment` 错误（Linux/macOS）** —— `generate.sh` 已在本地 `.venv` 里隔离一切，所以应该不会发生。
- **没有网络** —— 从 python.org 安装 Python，再运行 `pip install openpyxl`，然后重新启动。
- **手动运行** —— 在此文件夹打开终端机，运行：`python build_tax_workbook.py`。

---

## 🤖 维护者 / AI 助手
马来西亚税法每年都会变。更新税务数据的**完整指引**（税率表、减免、回扣、e-Filing 截止日期，以及不可破坏文件的硬性约束）请见英文版 [README.md](README.md) 的 *For AI assistants* 一节。

*免责声明：与 LHDN 无任何关联。仅供参考，非专业税务意见。报税前请务必向 LHDN 或持牌税务师确认。*
