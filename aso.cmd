@echo off
title Activate Microsoft Office Without Software - ariphx&cls&(net session >nul 2>&1)
if %errorLevel% GTR 0 (echo.&echo Activation cannot proceed. Please Run CMD as administrator...&echo.&echo temp > temp.txt&goto halt)
echo =====================================================================================&echo #Project: Activate Microsoft Office Without Software&echo =====================================================================================&echo.&echo #Supported products:&echo - Microsoft Office 2016&echo - Microsoft Office 2019&echo - Microsoft Office 2021&echo - Microsoft Office 365&echo.&echo.&echo Starting the activation process, please wait...&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office16") else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16") else (goto notsupported))&tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE">NUL && taskkill /F /IM WINWORD.EXE >nul&tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL && taskkill /F /IM EXCEL.EXE >nul&tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>NUL | find /I /N "POWERPNT.EXE">NUL && taskkill /F /IM POWERPNT.EXE >nul&(REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing" /f >nul 2>NUL)
(cscript //nologo ospp.vbs /dstatus | find /i "license name" > temp.txt || (cscript //nologo ospp.vbs /dstatus | find /i "Windows Script Host access is disabled" > temp.txt && goto wshdisabled))
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
set /P obrp=Type the version of Office you are using (2016/2019/2021) then press Enter: &echo.
:bucpkey
if %obrp% EQU 2016 (set pko=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99) else if %obrp% EQU 2019 (set pko=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP) else if %obrp% EQU 2021 (set pko=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH) else (echo =====================================================================================&echo.&echo This method only works for Microsoft Office 2016, 2019, 2021, and 365.&echo - For Office 2013, use the following activation method: bit.ly/acto2013&echo - For Office 2010, use the following activation method: bit.ly/acto2010&echo Copy or type one of the links above into your browser to open the tutorial.&echo.&echo If you are using Microsoft Office 365, type 2021 in the field below&echo.&goto olain)
cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:6MWKP >nul&cscript //nologo ospp.vbs /unpkey:6F7TH >nul
:ucpkey
cscript //nologo ospp.vbs /dstatus | find /i "Last 5 characters of installed product key" > temp.txt || goto skunpkey
set /P bpkey= < temp.txt
set pkeyt=%bpkey:~-5%
echo.&cscript //nologo ospp.vbs /unpkey:%pkeyt% >nul&goto ucpkey
:skunpkey
echo =====================================================================================&echo Activating your Office %obrp%, please wait...&echo =====================================================================================&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus*VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&echo.&cscript //nologo ospp.vbs /inpkey:%pko% >nul||goto notsupported&set i=1
:skms
if %i% GTR 5 (goto busy) else if %i% LEQ 5 (set KMS=s8.mshost.pro)
cscript //nologo ospp.vbs /sethst:%KMS% >nul
:ato
echo.&echo =====================================================================================&cscript //nologo ospp.vbs /act >nul&timeout /t 5 >nul&echo.&cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo =====================================================================================&echo.&echo.&echo Support us by donating via saweria.co/ariphx so this method can continue to be used.&echo.& if errorlevel 2 exit) || (echo It seems to take more time, please wait... & echo. & echo. & set /a i+=1 & timeout /t 10 >nul & goto skms)&timeout /t 7 >nul&explorer "https://saweria.co/ariphx"&echo.&echo If you have donated and Office still has issues, contact us via Instagram @ariphx_net&echo.&goto halt
:notsupported
echo =====================================================================================&echo.&echo Sorry, your Office is not supported.&echo.&echo - For Office 2013, use the following activation method: bit.ly/acto2013&echo.&echo - For Office 2010, use the following activation method: bit.ly/acto2010&echo.&echo Copy or type one of the links above into your browser to open the tutorial.&echo.&goto halt
:wshdisabled
echo =====================================================================================&echo.&echo Sorry, activation failed because Windows Script Host access is disabled.&echo.&echo The solution is to follow the guide in the video tutorial that opens to enable Windows Script.&echo.&timeout /t 7 >nul&explorer "https://youtu.be/vqGMSQnWMIY"&goto halt
:busy
echo =====================================================================================&echo.&echo.&echo Sorry, activation failed due to an unstable internet connection.&echo.&echo Please connect your laptop to a different WiFi network and try again.&echo.&goto halt
:halt
del temp.txt&cd %~dp0&del %0 >nul&pause >nul
