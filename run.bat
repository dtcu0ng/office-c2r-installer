@echo off

:MAIN
echo "Office C2R Installer"
echo "github.com/dtcu0ng/office-c2r-installer"
echo "made by dtcu0ng - version 22.11230"
echo[
echo Your install option:
echo 1) Microsoft 365 Business x86
echo 2) Microsoft 365 Business x64
echo 3) Microsoft 365 Enterprise x86
echo 4) Microsoft 365 Enterprise x64
echo 5) Office 2021 LTSC x86
echo 6) Office 2021 LTSC x64
echo 7) Office 2019 x86
echo 8) Office 2019 x64
echo 9) Exit
CHOICE /N /C:123456789 /M "Enter your choice here:"%1	
IF ERRORLEVEL ==9 GOTO EXIT
IF ERRORLEVEL ==8 GOTO CHOICE8
IF ERRORLEVEL ==7 GOTO CHOICE7
IF ERRORLEVEL ==6 GOTO CHOICE6
IF ERRORLEVEL ==5 GOTO CHOICE5
IF ERRORLEVEL ==4 GOTO CHOICE4
IF ERRORLEVEL ==3 GOTO CHOICE3
IF ERRORLEVEL ==2 GOTO CHOICE2
IF ERRORLEVEL ==1 GOTO CHOICE1

:CHOICE1
echo Installing Office...
files\setup.exe /configure files\Configuration-365Business-x86.xml
echo Done.
pause
goto MAIN

:CHOICE2
echo Installing Office...
files\setup.exe /configure files\Configuration-365Business-x64.xml
echo Done.
pause
goto MAIN

:CHOICE3
echo Installing Office...
files\setup.exe /configure files\Configuration-365Enterprise-x86.xml
echo Done.
pause
goto MAIN

:CHOICE4
echo Installing Office...
files\setup.exe /configure files\Configuration-365Enterprise-x64.xml
echo Done.
pause
goto MAIN

:CHOICE5
echo Installing Office...
files\setup.exe /configure files\Configuration-2019ProPlus-x86.xml
echo Done.
pause
goto MAIN

:CHOICE6
echo Installing Office...
files\setup.exe /configure files\Configuration-2019ProPlus-x64.xml
echo Done.
pause
goto MAIN

:CHOICE7
echo Installing Office...
files\setup.exe /configure files\Configuration-2021LTSC-x86.xml
echo Done.
pause
goto MAIN

:CHOICE8
echo Installing Office...
files\setup.exe /configure files\Configuration-2021LTSC-x64.xml
echo Done.
pause
goto MAIN

:EXIT
echo Exiting...
exit