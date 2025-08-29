@echo off
title Activate Microsoft Office 2019 Without Software - ariphx
cls
setlocal ENABLEDELAYEDEXPANSION

echo.
echo =====================================================================================
echo                 Activate Microsoft Office 2019 Without Software
echo =====================================================================================
echo.
echo    - Microsoft Office Standard 2019
echo    - Microsoft Office Professional Plus 2019
echo.
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
echo [PROSES] Menemukan direktori instalasi Office...
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

:: Tutup aplikasi Office yang sedang berjalan
echo [PROSES] Menutup aplikasi Office yang sedang berjalan (Word, Excel, PowerPoint)...
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE">NUL && taskkill /F /IM WINWORD.EXE >nul
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL && taskkill /F /IM EXCEL.EXE >nul
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>NUL | find /I /N "POWERPNT.EXE">NUL && taskkill /F /IM POWERPNT.EXE >nul

:: Hapus lisensi lama dan set port KMS
echo [PROSES] Menghapus lisensi Office lama...
(REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing" /f >nul 2>NUL)
cscript //nologo ospp.vbs /setprt:1688 >nul || goto wshdisabled

:: Pasang lisensi KMS dan hapus kunci produk yang mungkin bertentangan
echo [PROSES] Menginstal lisensi KMS dan menghapus kunci produk yang tidak perlu...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
cscript //nologo ospp.vbs /unpkey:WFG99 >nul
cscript //nologo ospp.vbs /unpkey:6MWKP >nul
cscript //nologo ospp.vbs /unpkey:6F7TH >nul

:ucpkey
cscript //nologo ospp.vbs /dstatus | find /i "Last 5 characters of installed product key" > temp.txt || goto skunpkey
set /P bpkey= < temp.txt
set pkeyt=!bpkey:~-5!
cscript //nologo ospp.vbs /unpkey:!pkeyt! >nul
goto ucpkey

:skunpkey
echo.
echo =====================================================================================
echo [PROSES] Mengaktivasi Office 2019, mohon tunggu...
echo =====================================================================================
echo.
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul||goto notsupported
set i=1

:server
if !i! GTR 5 (goto busy) else if !i! LEQ 5 (set KMS=s1.mshost.pro)
echo [PROSES] Menghubungkan ke server KMS: !KMS!...
cscript //nologo ospp.vbs /sethst:!KMS! >nul

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
    goto server
)

:notsupported
echo.
echo =====================================================================================
echo [GALAT] Maaf, versi Office 2019 Anda tidak didukung oleh skrip ini.
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
goto halt

:halt
del temp.txt >nul 2>nul
cd %~dp0
del %0 >nul 2>nul
pause >nul
