@echo off
title Activate Microsoft Office Without Software - ariphx
cls
setlocal ENABLEDELAYEDEXPANSION

echo.
echo =====================================================================================
echo                 Activate Microsoft Office Without Software
echo =====================================================================================
echo.
echo    - Microsoft Office 2016
echo    - Microsoft Office 2019
echo    - Microsoft Office 2021
echo    - Microsoft Office 365
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
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles%\Microsoft Office\Office16"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
) else (
    goto notsupported
)

:: Tutup aplikasi Office yang sedang berjalan
echo [PROSES] Menutup aplikasi Office yang sedang berjalan (Word, Excel, PowerPoint)...
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE">NUL && taskkill /F /IM WINWORD.EXE >nul
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL && taskkill /F /IM EXCEL.EXE >nul
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>NUL | find /I /N "POWERPNT.EXE">NUL && taskkill /F /IM POWERPNT.EXE >nul

:: Hapus lisensi lama
echo [PROSES] Menghapus lisensi Office lama...
REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing" /f >nul 2>NUL

:: Dapatkan status lisensi dan deteksi versi Office
(cscript //nologo ospp.vbs /dstatus | find /i "license name" > temp.txt || (cscript //nologo ospp.vbs /dstatus | find /i "Windows Script Host access is disabled" > temp.txt && goto wshdisabled))
echo [PROSES] Mendeteksi versi Microsoft Office yang terinstal...
findstr /m "O365" temp.txt >nul || goto o21
set obrp=2021&goto bucpkey
:o21
findstr /m "office 21" temp.txt >nul || goto o19
set obrp=2021&goto bucpkey
:o19
findstr /m "office 19" temp.txt >nul || goto o16
set obrp=2019&goto bucpkey
:o16
findstr /m "office 16" temp.txt >nul || goto olain
set obrp=2016&goto bucpkey
:olain
echo.
echo [INFO] Versi Office tidak terdeteksi otomatis.
set /P obrp=Silahkan ketik versi Office yang Anda gunakan (2016/2019/2021) lalu tekan Enter: 
echo.

:bucpkey
if %obrp% EQU 2016 (set pko=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99) else if %obrp% EQU 2019 (set pko=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP) else if %obrp% EQU 2021 (set pko=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH) else (
    echo.
    echo =====================================================================================
    echo [GALAT] Maaf, metode ini hanya berfungsi untuk Microsoft Office 2016, 2019, 2021, dan 365.
    echo.
    echo - Untuk Office 2013, gunakan metode aktivasi berikut: bit.ly/acto2013
    echo - Untuk Office 2010, gunakan metode aktivasi berikut: bit.ly/acto2010
    echo.
    echo Salin atau ketik salah satu tautan di atas ke browser Anda untuk membuka tutorial.
    echo.
    echo Jika Anda menggunakan Microsoft Office 365, ketik 2021 di bidang di bawah ini.
    echo.
    goto olain
)

:: Hapus kunci produk yang mungkin bertentangan
echo [PROSES] Menghapus kunci produk lama yang terdeteksi...
cscript //nologo ospp.vbs /unpkey:WFG99 >nul
cscript //nologo ospp.vbs /unpkey:6MWKP >nul
cscript //nologo ospp.vbs /unpkey:6F7TH >nul

:ucpkey
cscript //nologo ospp.vbs /dstatus | find /i "Last 5 characters of installed product key" > temp.txt || goto skunpkey
set /P bpkey=<temp.txt
set pkeyt=!bpkey:~-5!
cscript //nologo ospp.vbs /unpkey:!pkeyt! >nul
goto ucpkey

:skunpkey
echo.
echo =====================================================================================
echo [PROSES] Mengaktivasi Office %obrp%, mohon tunggu...
echo =====================================================================================
echo.
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus*VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo [PROSES] Memasang kunci produk dan mengatur server KMS...
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /inpkey:%pko% >nul||goto notsupported
set i=1

:skms
if !i! GTR 5 (goto busy) else if !i! LEQ 5 (set KMS=s8.mshost.pro)
echo [PROSES] Menghubungkan ke server KMS: !KMS!...
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
    echo Dukung kami dengan donasi via trakteer.id/ariphx agar cara ini bisa terus digunakan.
    echo.
    timeout /t 7 >nul
    explorer "https://trakteer.id/ariphx"
    echo.
    echo Jika sudah donasi dan Office masih bermasalah, chat via Instagram @bagitekno_net
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
echo [GALAT] Maaf, versi Office Anda tidak didukung oleh skrip ini.
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
