@echo off
set ver=1.0
title Visual Studio Activation Scripts %ver%
cls
:init
 setlocal DisableDelayedExpansion
 set cmdInvoke=1
 set winSysFolder=System32
 set "batchPath=%~dpnx0"
 rem this works also from cmd shell, other than %~0
 for %%k in (%0) do set batchName=%%~nk
 set "vbsGetPrivileges=%temp%\OEgetPriv_%batchName%.vbs"
 setlocal EnableDelayedExpansion

:checkPrivileges
  whoami /groups /nh | find "S-1-16-12288" > nul
  if '%errorlevel%' == '0' ( goto checkPrivileges2 ) else ( goto getPrivileges )

:checkPrivileges2
  net session 1>nul 2>NUL
  if '%errorlevel%' == '0' ( goto gotPrivileges ) else ( goto getPrivileges )

:getPrivileges
  if '%1'=='ELEV' (echo ELEV & shift /1 & goto gotPrivileges)

  ECHO Set UAC = CreateObject^("Shell.Application"^) > "%vbsGetPrivileges%"
  ECHO args = "ELEV " >> "%vbsGetPrivileges%"
  ECHO For Each strArg in WScript.Arguments >> "%vbsGetPrivileges%"
  ECHO args = args ^& strArg ^& " "  >> "%vbsGetPrivileges%"
  ECHO Next >> "%vbsGetPrivileges%"
  
  if '%cmdInvoke%'=='1' goto InvokeCmd 

  ECHO UAC.ShellExecute "!batchPath!", args, "", "runas", 1 >> "%vbsGetPrivileges%"
  goto ExecElevation

:InvokeCmd
  ECHO args = "/c """ + "!batchPath!" + """ " + args >> "%vbsGetPrivileges%"
  ECHO UAC.ShellExecute "%SystemRoot%\%winSysFolder%\cmd.exe", args, "", "runas", 1 >> "%vbsGetPrivileges%"

:ExecElevation
 "%SystemRoot%\%winSysFolder%\WScript.exe" "%vbsGetPrivileges%" %*
 exit /B

:gotPrivileges
 setlocal & cd /d %~dp0
 if '%1'=='ELEV' (del "%vbsGetPrivileges%" 1>nul 2>nul  &  shift /1)

:checkVisualStudio
echo Visual Studio Activation Scripts %ver%
choice /T 3 /C m0 /D 0 /N /M "Press M for manual mode..."
if %errorlevel% == 1 goto :manual
cls
echo Checking for Visual Studio Installations...
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\StorePID.exe" goto :activateVS2019Enterprise
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\Common7\IDE\StorePID.exe" goto :activateVS2019Professional

goto unsupported

:manual
cls
set /p executable=Type the path to StorePID.exe and press enter (without quotas): 
set /p productkey=Type the product key and press enter: 
set /p mpc=Type the MPC and press enter: 
echo Activating...
"%executable%" %productkey% %mpc% >nul 2>&1
if "%errorlevel%" == "0" (goto :successful)
else (
	echo Error! Product Key has not been installed. (Error: "%errorlevel%")
	choice /C RA /N /M "Press R to retry or A to abort."
	if %errorlevel% 1 goto :checkVisualStudio
	else goto :eof
)

:unsupported
echo A supported Visual Studio version was not found!
echo Supported versions:
echo - Visual Studio 2019 (Enterprise and Professional)
echo If you have a supported installation, try again with manual mode.
choice /C RA /N /M "Press R to retry or A to abort."
if %errorlevel% 1 goto :checkVisualStudio
else goto :eof

:activateVS2019Professional
echo Activating...
"C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\Common7\IDE\StorePID.exe" NYWVH-HT4XC-R2WYW-9Y3CM-X4V3Y 09262 >nul 2>&1
if "%errorlevel%" == "0" (goto :successful)
else (
	echo Error! Product Key has not been installed. (Error: "%errorlevel%")
	choice /C RA /N /M "Press R to retry or A to abort."
	if %errorlevel% 1 goto :checkVisualStudio
	else goto :eof
)

:activateVS2019Enterprise
echo Activating...
"C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\StorePID.exe" BF8Y8-GN2QH-T84XB-QVY3B-RC4DF 09260 >nul 2>&1
if "%errorlevel%" == "0" (goto :successful)
else (
	echo Error! Product Key has not been installed. (Error: "%errorlevel%")
	choice /C RA /N /M "Press R to retry or A to abort."
	if %errorlevel% 1 goto :checkVisualStudio
	else goto :eof
)

:successful
echo Visual Studio Activation was successful!
echo Press any key to exit.
pause >nul
goto :eof

:eof