@echo off
title Aktivace Office
set "params=%*"
cd /d "%~dp0" && ( if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs" ) && fsutil dirty query %systemdrive% 1>nul 2>nul || (  echo Set UAC = CreateObject^("Shell.Application"^) : UAC.ShellExecute "cmd.exe", "/k cd ""%~sdp0"" && %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs" && "%temp%\getadmin.vbs" && exit /B )
call :1

:1
echo Aktivuje se to, bude to chvilku trvat :)
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16 >nul
cd /d %ProgramFiles%\Microsoft Office\Office16 >nul
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x" >nul
cscript ospp.vbs /setprt:1688 >nul
cscript ospp.vbs /unpkey:6F7TH >nul >nul
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul
cscript ospp.vbs /sethst:s8.uk.to >nul
cscript ospp.vbs /act >nul
cls
echo Hotovo!
echo.
echo Stiskni mezernik :)
pause >nul
exit