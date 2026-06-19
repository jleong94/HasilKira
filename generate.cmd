@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ==================================================
echo   LHDN Tax Worksheet Generator
echo   生成马来西亚个人所得税表 / Penjana Lembaran Cukai
echo ==================================================
echo.

REM ---------- 1. Find a working Python ----------
set "PYCMD="
python --version >nul 2>&1 && set "PYCMD=python"
if not defined PYCMD (
    py -3 --version >nul 2>&1 && set "PYCMD=py -3"
)

REM ---------- 2. Install Python if missing (via winget) ----------
if not defined PYCMD (
    echo [INFO] Python not found / 找不到 Python / Python tidak dijumpai.
    winget --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] winget unavailable. Install Python manually:
        echo         https://www.python.org/downloads/
        echo         ^(tick "Add python.exe to PATH" during setup^)
        echo.
        pause
        exit /b 1
    )
    echo [INFO] Installing Python via winget ... please wait.
    winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
    REM re-probe (PATH may not refresh in this window)
    py -3 --version >nul 2>&1 && set "PYCMD=py -3"
    if not defined PYCMD ( python --version >nul 2>&1 && set "PYCMD=python" )
    if not defined PYCMD (
        echo.
        echo [INFO] Python installed. Please CLOSE this window and
        echo        double-click generate.cmd again.
        echo        Python 已安装,请关闭此窗口再双击 generate.cmd。
        echo.
        pause
        exit /b 0
    )
)
echo [OK] Using Python: %PYCMD%
echo.

REM ---------- 3. Ensure pip + openpyxl ----------
%PYCMD% -m pip --version >nul 2>&1 || %PYCMD% -m ensurepip --upgrade
%PYCMD% -m pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing required library: openpyxl ...
    %PYCMD% -m pip install --upgrade openpyxl
    if errorlevel 1 (
        echo [ERROR] Could not install openpyxl. Check your internet connection.
        echo         无法安装 openpyxl,请检查网络连接。
        pause
        exit /b 1
    )
)

REM ---------- 4. Generate the workbook ----------
echo [INFO] Generating Excel file / 正在生成 Excel ...
%PYCMD% "build_tax_workbook.py"
if errorlevel 1 (
    echo.
    echo [ERROR] Generation failed. See the message above.
    echo         生成失败,请看上方错误信息。
    pause
    exit /b 1
)

echo.
echo ==================================================
echo   DONE / 完成 / SELESAI
echo   File saved in the "LHDN report" folder.
echo   文件已保存在 "LHDN report" 文件夹。
echo ==================================================
start "" "LHDN report"
pause
endlocal
