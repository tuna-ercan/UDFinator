@echo off
rem UDFinator.exe derleme betigi.
rem Temiz bir sanal ortam (.venv) kullanir; boylece bilgisayardaki diger
rem Python paketleri exe'ye karismaz ve dosya kucuk kalir.
rem Klasorde image.jpg varsa arka plan olarak exe'nin icine gomulur.
chcp 65001 >nul
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
    echo Sanal ortam olusturuluyor...
    python -m venv .venv || goto :hata
)

echo Gerekli paketler yukleniyor...
.venv\Scripts\python.exe -m pip install --quiet --upgrade pip
.venv\Scripts\python.exe -m pip install --quiet -r requirements.txt || goto :hata

set EXTRA=
if exist image.jpg (
    echo Arka plan resmi bulundu: image.jpg exe'ye gomulecek.
    set EXTRA=--add-data "image.jpg;."
) else (
    echo Arka plan resmi yok: duz renkli tasarim kullanilacak.
)

rem tkinterdnd2: surukle-birak icin yerel tkdnd kutuphanesi de paketlenmeli
.venv\Scripts\python.exe -m PyInstaller --onefile --windowed --name UDFinator ^
    --collect-data docx --collect-data reportlab --collect-all tkinterdnd2 ^
    %EXTRA% --noconfirm app.py || goto :hata

echo.
echo Bitti: dist\UDFinator.exe
echo NOT: Bu exe aile fotografini icerebilir; GitHub'a yuklemeyin.
pause
exit /b 0

:hata
echo.
echo Derleme basarisiz oldu.
pause
exit /b 1
