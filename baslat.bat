@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
title SMT PnP to Excel - AOI ve Canias BOM Entegrasyon Portali

echo ======================================================================
echo    SMT PICK AND PLACE TO EXCEL - AOI VE CANIAS BOM PORTALI
echo ======================================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Sanal ortam olusturuluyor: .venv ...
    python -m venv .venv 2>nul
    if not exist ".venv\Scripts\python.exe" (
        py -m venv .venv 2>nul
    )
    if not exist ".venv\Scripts\python.exe" (
        echo.
        echo [HATA] Python bulunamadi!
        echo Lutfen Python 3.10 veya uzerini yukleyin ve PATH secenegini isaretleyin.
        echo.
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" -c "import uvicorn, openpyxl, multipart, pp_to_excel" >nul 2>nul
if errorlevel 1 (
    echo [2/3] Gerekli paketler yukleniyor, lutfen bekleyin...
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
echo Sunucu aktif! Bu pencereyi kapatirsaniz uygulama sonlanir.
echo ======================================================================
echo.

".venv\Scripts\python.exe" -m uvicorn pp_to_excel.web:app --host 127.0.0.1 --port 8000

if errorlevel 1 (
    echo.
    echo [HATA] Web sunucusu baslatilamadi!
    pause
)
