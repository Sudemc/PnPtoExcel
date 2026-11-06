@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
title SMT PnP to Excel - AOI ve Canias BOM Entegrasyon Portali

echo ======================================================================
echo    SMT PICK AND PLACE TO EXCEL - AOI VE CANIAS BOM PORTALI
echo ======================================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Python ortami kontrol ediliyor...
    set "PY_CMD="

    python -c "import sys; assert sys.version_info >= (3, 9)" >nul 2>nul
    if not errorlevel 1 set "PY_CMD=python"

    if not defined PY_CMD (
        py -3 -c "import sys; assert sys.version_info >= (3, 9)" >nul 2>nul
        if not errorlevel 1 set "PY_CMD=py -3"
    )

    if not defined PY_CMD (
        for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*" "C:\Program Files\Python3*") do (
            if exist "%%~fD\python.exe" (
                set "PY_CMD="%%~fD\python.exe""
            )
        )
    )

    if not defined PY_CMD (
        echo Python bulunamadi. Windows uzerinden Python 3.11 otomatik kuruluyor...
        winget install -e --id Python.Python.3.11 --scope user --accept-package-agreements --accept-source-agreements
        for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*" "C:\Program Files\Python3*") do (
            if exist "%%~fD\python.exe" (
                set "PY_CMD="%%~fD\python.exe""
            )
        )
    )

    if not defined PY_CMD (
        echo.
        echo [HATA] Python otomatik kurulamadi!
        echo Lutfen https://www.python.org/downloads/ adresinden Python 3.11 indirin
        echo ve kurulumda "Add Python to PATH" kutucugunu isaretleyin.
        echo.
        pause
        exit /b 1
    )

    echo [1/3] Sanal ortam olusturuluyor: .venv ...
    !PY_CMD! -m venv .venv
    if not exist ".venv\Scripts\python.exe" (
        echo.
        echo [HATA] Sanal ortam olusturulamadi!
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" -c "import uvicorn, openpyxl, multipart, pp_to_excel" >nul 2>nul
if errorlevel 1 (
    echo [2/3] Gerekli kutuphaneler yukleniyor, lutfen bekleyin...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip --trusted-host pypi.org --trusted-host files.pythonhosted.org >nul 2>nul
    ".venv\Scripts\python.exe" -m pip install -e . --trusted-host pypi.org --trusted-host files.pythonhosted.org
    if errorlevel 1 (
        echo.
        echo [HATA] Paketler indirilemedi! Lutfen internet baglantinizi kontrol edin.
        pause
        exit /b 1
    )
)

echo [3/3] Web sunucusu hazirlaniyor...
echo.

REM 8000 portunu dinleyen eski bir islem varsa sonlandir
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8000 ^| findstr LISTENING') do (
    taskkill /f /pid %%a >nul 2>nul
)

echo Tarayiciniz aciliyor: http://127.0.0.1:8000
start http://127.0.0.1:8000

echo.
echo ======================================================================
echo Sunucu aktif! Bu siyah pencereyi kapatirsaniz uygulama durur.
echo ======================================================================
echo.

".venv\Scripts\python.exe" -m uvicorn pp_to_excel.web:app --host 127.0.0.1 --port 8000

if errorlevel 1 (
    echo.
    echo [HATA] Web sunucusu baslatilamadi!
    pause
)
