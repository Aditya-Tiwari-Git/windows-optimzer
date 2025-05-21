@echo off
echo =============================
echo Creating Windows System Optimizer Installer
echo =============================

:: Check if NSIS is installed (Nullsoft Scriptable Install System)
where makensis >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo NSIS not found. Please install NSIS first.
    echo Download from: https://nsis.sourceforge.io/Download
    pause
    exit /b 1
)

:: Create installer directory if it doesn't exist
if not exist "installer" mkdir installer

:: Create NSIS script
echo Creating NSIS script...
(
echo !include "MUI2.nsh"
echo Name "Windows System Optimizer"
echo OutFile "installer\WindowsSystemOptimizer_Setup.exe"
echo InstallDir "$PROGRAMFILES\Windows System Optimizer"
echo.
echo !define MUI_ICON "assets\app_icon.ico"
echo !define MUI_UNICON "assets\app_icon.ico"
echo !define MUI_WELCOMEFINISHPAGE_BITMAP "assets\installer-welcome.bmp"
echo !define MUI_HEADERIMAGE
echo !define MUI_HEADERIMAGE_BITMAP "assets\installer-header.bmp"
echo.
echo ; Modern UI pages
echo !insertmacro MUI_PAGE_WELCOME
echo !insertmacro MUI_PAGE_DIRECTORY
echo !insertmacro MUI_PAGE_INSTFILES
echo !insertmacro MUI_PAGE_FINISH
echo.
echo ; Uninstaller pages
echo !insertmacro MUI_UNPAGE_CONFIRM
echo !insertmacro MUI_UNPAGE_INSTFILES
echo.
echo ; Set UI language
echo !insertmacro MUI_LANGUAGE "English"
echo.
echo Section "Install"
echo     SetOutPath "$INSTDIR"
echo     File "dist\WindowsSystemOptimizer.exe"
echo     File "assets\app_icon.ico"
echo.
echo     ; Create Start Menu shortcuts
echo     CreateDirectory "$SMPROGRAMS\Windows System Optimizer"
echo     CreateShortcut "$SMPROGRAMS\Windows System Optimizer\Windows System Optimizer.lnk" "$INSTDIR\WindowsSystemOptimizer.exe"
echo     CreateShortcut "$SMPROGRAMS\Windows System Optimizer\Uninstall.lnk" "$INSTDIR\uninstall.exe"
echo.
echo     ; Create Desktop shortcut
echo     CreateShortcut "$DESKTOP\Windows System Optimizer.lnk" "$INSTDIR\WindowsSystemOptimizer.exe"
echo.
echo     ; Create uninstaller
echo     WriteUninstaller "$INSTDIR\uninstall.exe"
echo.
echo     ; Add uninstall information to Add/Remove Programs
echo     WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\WindowsSystemOptimizer" "DisplayName" "Windows System Optimizer"
echo     WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\WindowsSystemOptimizer" "UninstallString" "$\"$INSTDIR\uninstall.exe$\""
echo     WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\WindowsSystemOptimizer" "DisplayIcon" "$\"$INSTDIR\app_icon.ico$\""
echo     WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\WindowsSystemOptimizer" "Publisher" "WinOptimizer"
echo     WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\WindowsSystemOptimizer" "DisplayVersion" "1.0.0"
echo SectionEnd
echo.
echo Section "Uninstall"
echo     ; Remove files and shortcuts
echo     Delete "$INSTDIR\WindowsSystemOptimizer.exe"
echo     Delete "$INSTDIR\app_icon.ico"
echo     Delete "$INSTDIR\uninstall.exe"
echo     RMDir "$INSTDIR"
echo.
echo     ; Remove Start Menu shortcuts
echo     Delete "$SMPROGRAMS\Windows System Optimizer\Windows System Optimizer.lnk"
echo     Delete "$SMPROGRAMS\Windows System Optimizer\Uninstall.lnk"
echo     RMDir "$SMPROGRAMS\Windows System Optimizer"
echo.
echo     ; Remove Desktop shortcut
echo     Delete "$DESKTOP\Windows System Optimizer.lnk"
echo.
echo     ; Remove registry entries
echo     DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\WindowsSystemOptimizer"
echo SectionEnd
) > installer.nsi

:: Create placeholder images if they don't exist
echo Creating placeholder installer images if needed...
if not exist "assets\installer-welcome.bmp" (
    echo Warning: assets\installer-welcome.bmp not found. Using a placeholder.
    :: Create a placeholder welcome image (164x314 pixels)
    echo You should replace this with a proper welcome image (164x314 pixels) > assets\installer-welcome.bmp
)

if not exist "assets\installer-header.bmp" (
    echo Warning: assets\installer-header.bmp not found. Using a placeholder.
    :: Create a placeholder header image (150x57 pixels)
    echo You should replace this with a proper header image (150x57 pixels) > assets\installer-header.bmp
)

:: Run NSIS to create the installer
echo.
echo Building installer...
makensis installer.nsi

:: Check if installer was created successfully
if exist "installer\WindowsSystemOptimizer_Setup.exe" (
    echo.
    echo =============================
    echo Installer created successfully!
    echo Installer is located at: installer\WindowsSystemOptimizer_Setup.exe
    echo =============================
) else (
    echo.
    echo =============================
    echo Failed to create installer!
    echo =============================
)

pause 