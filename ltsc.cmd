@echo off
setlocal ENABLEDELAYEDEXPANSION

echo.
echo =======================================================
echo     Windows Enterprise LTSC Activation Script
echo =======================================================
echo.
echo    Versi Windows yang didukung:
echo    - Windows 11 Enterprise LTSC 2024
echo    - Windows 10 Enterprise LTSC 2021
echo    - Windows 10 Enterprise LTSC 2019
echo.
echo    code by ariphx
echo.
echo =======================================================
echo.

:: Periksa hak akses administrator
net session >nul 2>&1
if !errorLevel! neq 0 (
    echo.
    echo [GALAT] Skrip ini harus dijalankan sebagai administrator.
    echo Silakan klik kanan pada file .cmd dan pilih "Run as administrator".
    echo.
    pause
    exit /b
)

:: Direktori sementara
set "TEMP_DIR=%~dp0temp_activation"
set "SKUS_DIR=%WINDIR%\System32\spp\tokens\skus"

if exist "%TEMP_DIR%" (
    rmdir /s /q "%TEMP_DIR%"
)
mkdir "%TEMP_DIR%"
cd /d "%TEMP_DIR%"

echo [PROSES] Mengunduh file aktivasi...
curl -s -LJO https://raw.githubusercontent.com/ariphx/activator-repo/main/tes.rar
if !errorLevel! neq 0 (
    echo.
    echo [GALAT] Gagal mengunduh file. Periksa koneksi internet atau URL.
    echo.
    goto :cleanup_fail
)

echo [PROSES] Mengekstrak file...
"C:\Program Files\WinRAR\WinRAR.exe" x "tes.rar" >nul
if !errorLevel! neq 0 (
    echo.
    echo [GALAT] Gagal mengekstrak file. Pastikan WinRAR terinstal.
    echo.
    goto :cleanup_fail
)

echo [PROSES] Mengubah hak akses folder %SKUS_DIR%...
takeown /F "%SKUS_DIR%" /R /D Y >nul
if !errorLevel! neq 0 (
    echo.
    echo [GALAT] Gagal mengambil alih kepemilikan.
    echo.
    goto :cleanup_fail
)

icacls "%SKUS_DIR%" /grant Administrators:F /T >nul
if !errorLevel! neq 0 (
    echo.
    echo [GALAT] Gagal memberikan izin Kontrol Penuh.
    echo.
    goto :cleanup_fail
)

echo [PROSES] Menyalin file aktivasi...
xcopy "csvlk-pack" "%SKUS_DIR%\csvlk-pack" /E /I /Y >nul
xcopy "EnterpriseS" "%SKUS_DIR%\EnterpriseS" /E /I /Y >nul
if !errorLevel! neq 0 (
    echo.
    echo [GALAT] Gagal menyalin folder.
    echo.
    goto :cleanup_fail
)

echo [PROSES] Menjalankan perintah aktivasi Windows...
:: Menggunakan CALL untuk memastikan setiap perintah slmgr selesai sebelum skrip berlanjut
CALL cscript.exe %windir%\system32\slmgr.vbs /rilc >nul 2>&1
CALL cscript.exe %windir%\system32\slmgr.vbs /upk >nul 2>&1
CALL cscript.exe %windir%\system32\slmgr.vbs /ckms >nul 2>&1
CALL cscript.exe %windir%\system32\slmgr.vbs /cpky >nul 2>&1
CALL cscript.exe %windir%\system32\slmgr.vbs /ipk M7XTQ-FN8P6-TTKYV-9D4CC-J462D >nul 2>&1
CALL cscript.exe %windir%\system32\slmgr.vbs /skms kms.digiboy.ir >nul 2>&1
CALL cscript.exe %windir%\system32\slmgr.vbs /ato >nul 2>&1

if !errorLevel! neq 0 (
    echo.
    echo [GALAT] Proses aktivasi gagal.
    echo [INFO] Kemungkinan penyebab:
    echo - Versi Windows tidak cocok atau kunci produk tidak valid.
    echo - Masalah koneksi ke server KMS atau server sedang offline.
    echo.
    goto :cleanup_fail
)

echo.
echo [BERHASIL] Proses aktivasi berhasil diselesaikan.
echo.
echo # Please support this project: https://trakteer.id/ariphx
explorer "https://trakteer.id/ariphx"

goto :cleanup_end

:cleanup_fail
:: Pesan tambahan jika proses gagal
echo.
echo Pembersihan file sementara sedang dilakukan...

:cleanup_end
:: Membersihkan file sementara
cd ..
if exist "%TEMP_DIR%" (
    rmdir /s /q "%TEMP_DIR%"
)

echo.
echo Tekan tombol apa saja untuk keluar...
pause >nul
