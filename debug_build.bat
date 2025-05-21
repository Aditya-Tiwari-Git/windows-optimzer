@echo off
echo =============================
echo Building Debug Version of Windows System Optimizer
echo =============================

:: Clean build artifacts if they exist
if exist "dist\WindowsSystemOptimizerDebug.exe" del "dist\WindowsSystemOptimizerDebug.exe"
if exist "build\WindowsSystemOptimizerDebug" rmdir /s /q "build\WindowsSystemOptimizerDebug"
if exist "WindowsSystemOptimizerDebug.spec" del "WindowsSystemOptimizerDebug.spec"

:: Create executable with PyInstaller with console window
pyinstaller --name "WindowsSystemOptimizerDebug" ^
            --icon="assets\app_icon.ico" ^
            --onefile ^
            --add-data "assets;assets" ^
            --add-data "scripts;scripts" ^
            --add-data "ui;ui" ^
            --add-data "utils;utils" ^
            --add-data "services;services" ^
            --hidden-import=psutil ^
            --hidden-import=wmi ^
            --hidden-import=requests ^
            --hidden-import=matplotlib ^
            --hidden-import=numpy ^
            --hidden-import=PyQt5 ^
            --hidden-import=PyQt5.QtCore ^
            --hidden-import=PyQt5.QtGui ^
            --hidden-import=PyQt5.QtWidgets ^
            run_app.py

:: Check if build was successful
if exist "dist\WindowsSystemOptimizerDebug.exe" (
    echo.
    echo =============================
    echo Debug build successful!
    echo Executable is located at: dist\WindowsSystemOptimizerDebug.exe
    echo =============================
    echo Running the debug version...
    echo You will see console output that might help identify the issue.
    cd dist
    WindowsSystemOptimizerDebug.exe
) else (
    echo.
    echo =============================
    echo Debug build failed!
    echo =============================
)

pause 