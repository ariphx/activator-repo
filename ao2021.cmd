@echo off
title Activate Microsoft Office 2021 Without Software - Ariphx&cls&(net session >nul 2>&1)
if %errorLevel% GTR 0 (echo.&echo Activation cannot proceed. Please Run CMD as administrator...&echo.&goto halt)
echo =====================================================================================&echo #Project: Activate Microsoft Office 2021 Without Software&echo =====================================================================================&echo.&echo #Supported products:&echo - Microsoft Office Standard 2021&echo - Microsoft Office Professional Plus 2021&echo.&echo.&echo Starting the activation process, please wait...&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE">NUL && taskkill /F /IM WINWORD.EXE >nul&tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL && taskkill /F /IM EXCEL.EXE >nul&tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>NUL | find /I /N "POWERPNT.EXE">NUL && taskkill /F /IM POWERPNT.EXE >nul&(REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing" /f >nul 2>NUL)&(cscript //nologo ospp.vbs /setprt:1688 >nul || goto wshdisabled)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:6MWKP >nul&cscript //nologo ospp.vbs /unpkey:6F7TH >nul
:ucpkey
cscript //nologo ospp.vbs /dstatus | find /i "Last 5 characters of installed product key" > temp.txt || goto skunpkey
set /P bpkey= < temp.txt
set pkeyt=%bpkey:~-5%
echo.&cscript //nologo ospp.vbs /unpkey:%pkeyt% >nul&goto ucpkey
:skunpkey
echo.&echo =====================================================================================&echo Activating your Office 2021, please wait...&echo =====================================================================================&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&set i=1&cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul||goto notsupported
:skms
if %i% GTR 5 (goto busy) else if %i% LEQ 5 (set KMS=s8.mshost.pro)
cscript //nologo ospp.vbs /sethst:%KMS% >nul
:ato
echo.&echo =====================================================================================&cscript //nologo ospp.vbs /act >nul&timeout /t 5 >nul&echo.&cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo =====================================================================================&echo.&echo.&echo Support us by donating via trakteer.id/ariphx so this method can continue to be used.&echo.& if errorlevel 2 exit) || (echo It seems to take more time, please wait... & echo. & echo. & set /a i+=1 & timeout /t 10 >nul & goto skms)&timeout /t 7 >nul&explorer "https://trakteer.id/ariphx"&echo.&echo.&goto halt
:notsupported
echo =====================================================================================&echo.&echo Sorry, your Office 2021 is not supported.&echo.&goto halt
:wshdisabled
echo =====================================================================================&echo.&echo Sorry, activation failed because Windows Script Host access is disabled.&echo.&echo The solution is to follow the guide in the video tutorial that opens to enable Windows Script.&echo.&timeout /t 7 >nul&explorer "https://youtu.be/vqGMSQnWMIY"&goto halt
:busy
echo =====================================================================================&echo.&echo.&echo Sorry, activation failed due to an unstable internet connection.&echo.&echo Please connect your laptop to a different WiFi network and try again.&echo.&goto halt
:halt
cd %~dp0&del %0 >nul&pause >nul
