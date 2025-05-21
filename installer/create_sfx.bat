@echo off
echo Creating self-extracting installer...
copy /b 7zS.sfx + config.txt + ..\WindowsSystemOptimizer_Installer.zip WindowsSystemOptimizer_Setup.exe
echo Installer created at: installer\WindowsSystemOptimizer_Setup.exe
pause
