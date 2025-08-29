@echo off
title Activate Microsoft Office 2016 ALL versions for FREE!
cls
setlocal ENABLEDELAYEDEXPANSION

echo.
echo ============================================================================
echo # Activating Microsoft software products for FREE without software by ariphx
echo ============================================================================
echo.
echo # Supported products:
echo - Microsoft Office Standard 2016
echo - Microsoft Office Professional Plus 2016
echo.
echo.
echo ============================================================================
echo [PROSES] Memulai proses aktivasi, silahkan tunggu...
echo ============================================================================
echo.

:: Temukan folder instalasi Office
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles%\Microsoft Office\Office16"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
) else (
    echo.
    echo [GALAT] File "ospp.vbs" tidak ditemukan.
    echo Pastikan Microsoft Office sudah terinstal dengan benar.
    echo.
    goto halt
)

:: Pasang lisensi KMS dan hapus kunci produk yang mungkin bertentangan
echo [PROSES] Menginstal lisensi KMS dan menghapus kunci produk yang tidak perlu...
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
cscript //nologo ospp.vbs /unpkey:WFG99 >nul
cscript //nologo ospp.vbs /unpkey:DRTFM >nul
cscript //nologo ospp.vbs /unpkey:BTDRB >nul
cscript //nologo ospp.vbs /unpkey:CPQVG >nul
cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul

set i=1
:server
if !i! EQU 1 (
    set KMS_Sev=kms7.MSGuides.com
) else if !i! EQU 2 (
    set KMS_Sev=kms8.MSGuides.com
) else if !i! EQU 3 (
    set KMS_Sev=kms9.MSGuides.com
) else if !i! EQU 4 (
    goto notsupported
)

echo.
echo ============================================================================
echo [PROSES] Menghubungkan ke server KMS: !KMS_Sev!...
echo.
cscript //nologo ospp.vbs /sethst:!KMS_Sev! >nul

cscript //nologo ospp.vbs /act | find /i "successful" && (
    echo.
    echo ============================================================================
    echo.
    echo [BERHASIL] Produk Anda berhasil diaktivasi!
    echo.
    echo # Please support this project: https://trakteer.id/ariphx
    echo.
    echo ============================================================================
    explorer "https://trakteer.id/ariphx"
    choice /n /c YN /m "do you want to donate? [Y,N]?"
    if !errorlevel! EQU 1 (
        goto halt
    ) else (
        goto halt
    )
) || (
    echo [GALAT] Koneksi ke server KMS gagal.
    echo [PROSES] Mencoba menghubungkan ke server lain... Mohon tunggu...
    echo.
    set /a i+=1
    timeout /t 10 >nul
    goto server
)

:notsupported
echo.
echo ============================================================================
echo [GALAT] Maaf! Versi Office 2016 Anda tidak didukung.
echo Silakan coba instal versi terbaru.
echo.
goto halt

:halt
pause >nul
