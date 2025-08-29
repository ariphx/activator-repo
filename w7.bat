@echo off
title Activate Windows 7 Professional/Enterprise for FREE!
cls
setlocal ENABLEDELAYEDEXPANSION

echo.
echo =====================================
echo     Windows 11 Activation Script
echo =====================================
echo.
echo # Supported products:
echo - Windows 7 Professional
echo - Windows 7 Professional N
echo - Windows 7 Professional E
echo - Windows 7 Enterprise
echo - Windows 7 Enterprise N
echo - Windows 7 Enterprise E
echo.
echo.
echo ============================================================================
echo [PROSES] Activating your Windows...
echo ============================================================================
echo.

:: Pindah ke direktori system32 untuk menjalankan slmgr.vbs
cd /d %windir%\system32

:: Hapus kunci produk yang ada
cscript //nologo slmgr.vbs /upk >nul
cscript //nologo slmgr.vbs /cpky >nul

set i=1
wmic os | findstr /I "enterprise" >nul
if !errorlevel! EQU 0 (
    echo [PROSES] Windows 7 Enterprise detected. Installing GVLK...
    cscript //nologo slmgr.vbs /ipk 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH >nul || (cscript //nologo slmgr.vbs /ipk YDRBP-3D83W-TY26F-D46B2-XCKRJ >nul || (cscript //nologo slmgr.vbs /ipk C29WB-22CC8-VJ326-GHFJW-H9DH4 >nul))
) else (
    echo [PROSES] Windows 7 Professional detected. Installing GVLK...
    cscript //nologo slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4 >nul || (cscript //nologo slmgr.vbs /ipk MRPKT-YTG23-K7D7T-X2JMM-QY7MG >nul || (cscript //nologo slmgr.vbs /ipk W82YF-2Q76Y-63HXB-FGJG9-GF7QX >nul))
)

:server
if !i! EQU 1 (
    set KMS_Sev=kms7.MSGuides.com
) else if !i! EQU 2 (
    set KMS_Sev=kms8.MSGuides.com
) else if !i! EQU 3 (
    set KMS_Sev=kms9.MSGuides.com
) else if !i! EQU 4 (
    goto unsupported
)
cscript //nologo slmgr.vbs /skms !KMS_Sev! >nul
echo.
echo [PROSES] Connecting to KMS server: !KMS_Sev!...

cscript //nologo slmgr.vbs /ato | find /i "successfully" && (
    echo.
    echo ============================================================================
    echo.
    echo [BERHASIL] Product activated successfully!
    echo.
    echo # Please support this project: https://trakteer.id/ariphx
    explorer "https://trakteer.id/ariphx"
    echo.
    echo ============================================================================
    choice /n /c YN /m "Do you want to exit? [Y,N]?"
    if !errorlevel! EQU 1 (
        echo.
        goto halt
    ) else (
        exit
    )
) || (
    echo.
    echo [GALAT] The connection to the KMS server failed.
    echo [PROSES] Trying to connect to another one... Please wait...
    echo.
    set /a i+=1
    timeout /t 10 >nul
    goto server
)

:unsupported
echo.
echo ============================================================================
echo.
echo [GALAT] Sorry! Your version is not supported.
echo.
goto halt

:halt
pause >nul
