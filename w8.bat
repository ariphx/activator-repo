@echo off
title Activate Windows 8 / Windows 8.1 
cls
setlocal ENABLEDELAYEDEXPANSION

echo.
echo ============================================================================
echo                   Windows 8 8.1 Activation Script
echo ============================================================================
echo.
echo # Supported products:
echo - Windows 8 Core
echo - Windows 8 Core Single Language
echo - Windows 8 Professional
echo - Windows 8 Professional N
echo - Windows 8 Professional WMC
echo - Windows 8 Enterprise
echo - Windows 8 Enterprise N
echo - Windows 8.1 Core
echo - Windows 8.1 Core N
echo - Windows 8.1 Core Single Language
echo - Windows 8.1 Professional
echo - Windows 8.1 Professional N
echo - Windows 8.1 Professional WMC
echo - Windows 8.1 Enterprise
echo - Windows 8.1 Enterprise N
echo.
echo.
echo ============================================================================
echo [PROSES] Activating your Windows...
echo ============================================================================
echo.

:: Hapus kunci produk yang ada
cscript //nologo slmgr.vbs /upk >nul
cscript //nologo slmgr.vbs /cpky >nul

set i=1
wmic os | findstr /I "enterprise" >nul
if !errorlevel! EQU 0 (
    echo [PROSES] Windows 8/8.1 Enterprise detected. Installing GVLK...
    cscript //nologo slmgr.vbs /ipk MHF9N-XY6XB-WVXMC-BTDCT-MKKG7 >nul || (cscript //nologo slmgr.vbs /ipk TT4HM-HN7YT-62K67-RGRQJ-JFFXW >nul || (cscript //nologo slmgr.vbs /ipk 32JNW-9KQ84-P47T8-D8GGY-CWCK7 >nul || (cscript //nologo slmgr.vbs /ipk JMNMF-RHW7P-DMY6X-RF3DR-X2BQT >nul)))
) else (
    echo [PROSES] Windows 8/8.1 non-Enterprise detected. Installing GVLK...
    cscript //nologo slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9 >nul || (cscript //nologo slmgr.vbs /ipk HMCNV-VVBFX-7HMBH-CTY9B-B4FXY >nul || (cscript //nologo slmgr.vbs /ipk NG4HW-VH26C-733KW-K6F98-J8CK4 >nul || (cscript //nologo slmgr.vbs /ipk XCVCF-2NXM9-723PB-MHCB7-2RYQQ >nul || (cscript //nologo slmgr.vbs /ipk BN3D2-R7TKB-3YPBD-8DRP2-27GG4 >nul || (cscript //nologo slmgr.vbs /ipk 2WN2H-YGCQR-KFX6K-CD6TF-84YXQ >nul || (cscript //nologo slmgr.vbs /ipk GNBB8-YVD74-QJHX6-27H4K-8QHDG >nul || (cscript //nologo slmgr.vbs /ipk M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK >nul || (cscript //nologo slmgr.vbs /ipk 7B9N3-D94CG-YTVHR-QBPX3-RJP64 >nul || (cscript //nologo slmgr.vbs /ipk BB6NG-PQ82V-VRDPW-8XVD2-V8P66 >nul || (cscript //nologo slmgr.vbs /ipk 789NJ-TQK6T-6XTH8-J39CJ-J8D3P >nul))))))))))
)

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
echo [PROSES] Connecting to KMS server: !KMS_Sev!...
cscript //nologo slmgr.vbs /skms !KMS_Sev! >nul
echo.
echo ============================================================================
echo.
cscript //nologo slmgr.vbs /ato | find /i "successfully" && (
    echo.
    echo ============================================================================
    echo.
    echo [BERHASIL] Product activated successfully!
    echo.
    echo # Please supporting this project: https://trakteer.id/ariphx
    echo.
    echo ============================================================================
    choice /n /c YN /m "do you want to donate? [Y,N]?"
    if !errorlevel! EQU 1 (
        explorer "https://trakteer.id/ariphx"
        goto halt
    ) else (
        goto halt
    )
) || (
    echo [GALAT] The connection to my KMS server failed! Trying to connect to another one...
    echo Please wait...
    echo.
    set /a i+=1
    timeout /t 10 >nul
    goto server
)

:notsupported
echo.
echo ============================================================================
echo [GALAT] Sorry! Your version is not supported.
echo.
goto halt

:halt
pause >nul
