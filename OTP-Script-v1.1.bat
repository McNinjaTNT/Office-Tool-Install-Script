@echo off
net session >NUL 2>&1
if %errorlevel% LSS 1 (
GOTO processcheck
) ELSE (
echo.
echo  :: Office Tool Plus Install Script
echo.
echo     You must run this script as Administrator
echo.
echo  :: Press any key to exit...
pause > NUL
GOTO EOF
)
:processcheck
QPROCESS "officetoolplus.exe" >NUL 2>&1
IF %ERRORLEVEL% LSS 1 (
echo.
echo  :: Office Tool Plus Install Script
echo.
echo     Office Tool Plus is currently running
echo.
echo  :: Press any key to exit...
pause > NUL
GOTO EOF
) ELSE (
GOTO menu
)

:menu
cls
echo.
echo  :: Office Tool Plus Install Script
echo.
echo     1. Install Office Tool Plus
echo     2. Uninstall Office Tool Plus
echo.
echo  :: Type a 'number' and press ENTER
echo  :: Type 'exit' to quit
echo.

set /P menu=
	if %menu%==1 GOTO install1
	if %menu%==2 GOTO uninstall1
	if %menu%==exit GOTO EOF
else (
	cls
	echo.
	echo  :: Office Tool Plus Install Script
	echo.
	echo     Incorrect Input Entered
	echo.
	echo  :: Please type a 'number' or 'exit'
	echo  :: Press any key to return to the menu...
	echo.
	pause > NUL
	GOTO menu
)

:install1
cls
echo.
echo  :: Office Tool Plus Install Script
echo.
echo  :: Downloading...

wget -O \Office-Tool.zip https://otp.landian.vip/redirect/download.php?type=runtime

GOTO installcheck1

:installcheck1
if exist "\Program Files (x86)\Office Tool Plus" (
rmdir /s /q "\Program Files (x86)\Office Tool Plus"
GOTO installcheck2
) ELSE (
GOTO installcheck2
)

:installcheck2
if exist "\Program Files (x86)\Office Tool" (
rmdir /s /q "\Program Files (x86)\Office Tool"
GOTO install2
) ELSE (
GOTO install2
)

:install2
cls
echo.
echo  :: Office Tool Plus Install Script
echo.
echo  :: Extracting...

powershell -command "Expand-Archive -Path '\Office-Tool.zip' -DestinationPath '\Program Files (x86)'"
del \Office-Tool.zip
rename "\Program Files (x86)\Office Tool" "Office Tool Plus"
rename "\Program Files (x86)\Office Tool Plus\Office Tool Plus.exe" "OfficeToolPlus.exe"
GOTO install3

:install3

set TARGET='\Program Files (x86)\Office Tool Plus\OfficeToolPlus.exe'
set SHORTCUT='\ProgramData\Microsoft\Windows\Start Menu\Programs\Office Tool Plus.lnk'
set PWS=powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile

%PWS% -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut(%SHORTCUT%); $S.TargetPath = %TARGET%; $S.Save()"

start "Shungite" /b "C:\Program Files (x86)\Office Tool Plus\OfficeToolPlus.exe"
GOTO EOF

:uninstall1
cls
echo.
echo  :: Office Tool Plus Install Script
echo.
echo  :: Removing Files...
if exist "\Office-Tool.zip" del C:\Office-Tool.zip
if exist "\Program Files (x86)\Office Tool Plus" rmdir /s /q "\Program Files (x86)\Office Tool Plus"
if exist "\Program Files (x86)\Office Tool" rmdir /s /q "\Program Files (x86)\Office Tool"
if exist "\ProgramData\Microsoft\Windows\Start Menu\Programs\Office Tool Plus.lnk" del "\ProgramData\Microsoft\Windows\Start Menu\Programs\Office Tool Plus.lnk"
cls
echo.
echo  :: Office Tool Plus Install Script
echo.
echo     Office Tool Plus has been succesfully uninstalled
echo.
echo  :: Press any key to exit...
pause >nul
GOTO EOF
