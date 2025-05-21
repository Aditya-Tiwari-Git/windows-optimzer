@echo off
echo =============================
echo Building Windows System Optimizer
echo =============================

:: Clean build artifacts if they exist
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del *.spec

:: Create executable with PyInstaller
pyinstaller --name "WindowsSystemOptimizer" ^
            --icon="assets\app_icon.ico" ^
            --onefile ^
            --windowed ^
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
if exist "dist\WindowsSystemOptimizer.exe" (
    echo.
    echo =============================
    echo Build successful!
    echo Executable is located at: dist\WindowsSystemOptimizer.exe
    echo =============================
) else (
    echo.
    echo =============================
    echo Build failed!
    echo =============================
)

pause 