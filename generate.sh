#!/usr/bin/env bash
# LHDN Tax Worksheet Generator - macOS / Linux launcher
# Equivalent of generate.cmd (which is for Windows).
set -e
cd "$(dirname "$0")"

echo "=================================================="
echo "  LHDN Tax Worksheet Generator (中 / BM / EN)"
echo "=================================================="
echo

# ---------- 1. Find Python 3 ----------
PY=""
if command -v python3 >/dev/null 2>&1; then PY=python3
elif command -v python >/dev/null 2>&1; then PY=python
fi

# ---------- 2. Install Python if missing ----------
if [ -z "$PY" ]; then
  echo "[INFO] Python not found. Trying to install..."
  if   command -v brew    >/dev/null 2>&1; then brew install python
  elif command -v apt-get >/dev/null 2>&1; then sudo apt-get update && sudo apt-get install -y python3 python3-venv python3-pip
  elif command -v dnf     >/dev/null 2>&1; then sudo dnf install -y python3 python3-pip
  elif command -v pacman  >/dev/null 2>&1; then sudo pacman -S --noconfirm python
  else
    echo "[ERROR] Could not auto-install. Please install Python 3 from https://www.python.org/downloads/"
    exit 1
  fi
  PY=python3
fi
echo "[OK] Using: $($PY --version)"
echo

# ---------- 3. Virtual env + openpyxl ----------
# A local venv avoids the 'externally-managed-environment' error on modern Linux/macOS.
if [ ! -d ".venv" ]; then
  echo "[INFO] Creating virtual environment (.venv) ..."
  "$PY" -m venv .venv
fi
VENV_PY=".venv/bin/python"
"$VENV_PY" -m pip install --upgrade pip >/dev/null
if ! "$VENV_PY" -m pip show openpyxl >/dev/null 2>&1; then
  echo "[INFO] Installing required library: openpyxl ..."
  "$VENV_PY" -m pip install openpyxl
fi

# ---------- 4. Generate the workbook ----------
echo "[INFO] Generating Excel file / 正在生成 Excel ..."
"$VENV_PY" build_tax_workbook.py

echo
echo "=================================================="
echo "  DONE / 完成 / SELESAI"
echo "  File saved in the 'LHDN report' folder."
echo "=================================================="

# open the output folder if a file manager is available
if   command -v open     >/dev/null 2>&1; then open "LHDN report"
elif command -v xdg-open >/dev/null 2>&1; then xdg-open "LHDN report" >/dev/null 2>&1 || true
fi
