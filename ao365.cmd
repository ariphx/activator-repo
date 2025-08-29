@echo off
title Activate Microsoft Office 365 Without Software - Ariphx
cls
setlocal ENABLEDELAYEDEXPANSION

echo.
echo =============================================
echo    Microsoft Office 365 Activation Script
echo =============================================
echo.
echo =====================================================================================
echo    [PROSES] Memulai proses aktivasi, silahkan tunggu...
echo =====================================================================================
echo.

:: Periksa hak akses administrator
net session >nul 2>&1
if !errorLevel! GTR 0 (
    echo.
    echo [GALAT] Aktivasi tidak bisa dilanjutkan.
    echo Skrip ini harus dijalankan sebagai administrator.
    echo Silakan klik kanan pada file .cmd dan pilih "Run as administrator".
    echo.
    goto halt
)

:: Temukan folder instalasi Office
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles%\Microsoft Office\Office16"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
) else (
    goto notsupported
)

:: Tutup aplikasi Office yang sedang berjalan
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE">NUL && taskkill /F /IM WINWORD.EXE >nul
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL && taskkill /F /IM EXCEL.EXE >nul
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>NUL | find /I /N "POWERPNT.EXE">NUL && taskkill /F /IM POWERPNT.EXE >nul

:: Hapus lisensi lama dan set port KMS
(REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing" /f >nul 2>NUL)
(cscript //nologo ospp.vbs /setprt:1688 >nul || (echo temp > temp.txt&goto wshdisabled))

:: Pasang lisensi KMS dan hapus kunci produk yang mungkin bertentangan
(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus*VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
cscript //nologo ospp.vbs /unpkey:WFG99 >nul
cscript //nologo ospp.vbs /unpkey:6MWKP >nul
cscript //nologo ospp.vbs /unpkey:6F7TH >nul

:ucpkey
cscript //nologo ospp.vbs /dstatus | find /i "Last 5 characters of installed product key" > temp.txt || goto skunpkey
set /P bpkey= < temp.txt
set pkeyt=%bpkey:~-5%
cscript //nologo ospp.vbs /unpkey:%pkeyt% >nul
goto ucpkey

:skunpkey
echo.
echo =====================================================================================
echo    [PROSES] Mengaktivasi Office 365, mohon tunggu...
echo =====================================================================================
echo.
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
set i=1
cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul||goto notsupported

:skms
if %i% GTR 5 (goto busy) else if %i% LEQ 5 (set KMS=s8.mshost.pro)
cscript //nologo ospp.vbs /sethst:%KMS% >nul

:ato
echo.
echo =====================================================================================
cscript //nologo ospp.vbs /act >nul
timeout /t 5 >nul
echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (
    echo.
    echo =====================================================================================
    echo.
    echo [BERHASIL] Produk Anda berhasil diaktivasi!
    echo.
    echo Dukung kami dengan donasi via trakteer.id/ariphx agar metode ini bisa terus digunakan.
    echo.
    timeout /t 7 >nul
    explorer "https://trakteer.id/ariphx"
    echo.
    goto halt
) || (
    echo Sepertinya butuh waktu lebih, silahkan tunggu...
    echo.
    echo.
    set /a i+=1
    timeout /t 10 >nul
    goto skms
)

:notsupported
echo =====================================================================================
echo.
echo [GALAT] Maaf, Office 365 Anda tidak didukung.
echo.
goto halt

:wshdisabled
echo =====================================================================================
echo.
echo [GALAT] Maaf, aktivasi gagal karena Windows Script Host dinonaktifkan.
echo.
goto halt

:busy
echo =====================================================================================
echo.
echo [GALAT] Maaf, aktivasi gagal karena koneksi internet Anda tidak stabil.
echo.
goto halt

:halt
del temp.txt >nul 2>nul
cd %~dp0
del %0 >nul 2>nul
pause >nul
