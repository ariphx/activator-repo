@echo off
title Activate Microsoft Office 2010 Volume for FREE!
cls
setlocal ENABLEDELAYEDEXPANSION

echo.
echo ============================================================================
echo                  Microsoft Office 2010 Activation Script
echo ============================================================================
echo.
echo # Supported products:
echo - Microsoft Office 2010 Standard Volume
echo - Microsoft Office 2010 Professional Plus Volume
echo.
echo.
echo ============================================================================
echo [PROSES] Memulai proses aktivasi, silahkan tunggu...
echo ============================================================================
echo.

:: Temukan folder instalasi Office
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" (
    cd /d "%ProgramFiles%\Microsoft Office\Office14"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" (
    cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14"
) else (
    echo.
    echo [GALAT] File "ospp.vbs" tidak ditemukan.
    echo Pastikan Microsoft Office sudah terinstal dengan benar.
    echo.
    goto halt
)

echo [PROSES] Menginstal dan menghapus kunci produk yang tidak perlu...
cscript //nologo ospp.vbs /unpkey:8R6BM >nul
cscript //nologo ospp.vbs /unpkey:H3GVB >nul
cscript //nologo ospp.vbs /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM >nul
cscript //nologo ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB >nul

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

echo [PROSES] Menghubungkan ke server KMS: !KMS_Sev!...
cscript //nologo ospp.vbs /sethst:!KMS_Sev! >nul
echo.
echo ============================================================================
echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (
    echo.
    echo ============================================================================
    echo.
    echo [BERHASIL] Produk Anda berhasil diaktivasi!
    echo.
    echo # Please support this project: https://trakteer.id/ariphx
    echo.
    echo ============================================================================
    choice /n /c YN /m "Do you want to donate? [Y,N]?" 
    if !errorlevel! EQU 1 (
        explorer "https://trakteer.id/ariphx"
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
echo [GALAT] Maaf! Versi Office 2010 Anda tidak didukung.
echo.
goto halt

:halt
pause >nul
