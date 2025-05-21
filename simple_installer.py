#!/usr/bin/env python3
"""
Simple installer script for Windows System Optimizer
"""

import os
import sys
import shutil
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox

def create_installer():
    # Define paths
    dist_dir = "dist"
    exe_file = os.path.join(dist_dir, "WindowsSystemOptimizer.exe")
    icon_file = os.path.join("assets", "app_icon.ico")
    
    # Check if executable exists
    if not os.path.exists(exe_file):
        print(f"ERROR: Executable not found at {exe_file}")
        print("Please build the application first using build.bat")
        return
    
    # Create a ZIP package
    zip_file = "WindowsSystemOptimizer_Installer.zip"
    
    print(f"Creating installer package: {zip_file}")
    with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Add the executable
        zipf.write(exe_file, os.path.basename(exe_file))
        
        # Add the icon if it exists
        if os.path.exists(icon_file):
            zipf.write(icon_file, os.path.basename(icon_file))
            
        # Create a simple README.txt file
        readme_content = """Windows System Optimizer

Thank you for downloading Windows System Optimizer!

Installation Instructions:
1. Extract all files from this ZIP archive to your desired location
2. Run WindowsSystemOptimizer.exe to start the application

For support, please contact the developer.
"""
        zipf.writestr("README.txt", readme_content)
        
        # Create a simple install.bat file that creates shortcuts
        install_bat = """@echo off
echo Creating shortcuts for Windows System Optimizer...
echo.

:: Create desktop shortcut
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%TEMP%\\CreateShortcut.vbs"
echo sLinkFile = "%USERPROFILE%\\Desktop\\Windows System Optimizer.lnk" >> "%TEMP%\\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%TEMP%\\CreateShortcut.vbs"
echo oLink.TargetPath = "%~dp0WindowsSystemOptimizer.exe" >> "%TEMP%\\CreateShortcut.vbs"
echo oLink.IconLocation = "%~dp0app_icon.ico" >> "%TEMP%\\CreateShortcut.vbs"
echo oLink.Save >> "%TEMP%\\CreateShortcut.vbs"
cscript //nologo "%TEMP%\\CreateShortcut.vbs"
del "%TEMP%\\CreateShortcut.vbs"

:: Create Start Menu shortcut
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%TEMP%\\CreateShortcut.vbs"
echo sLinkFile = "%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Windows System Optimizer.lnk" >> "%TEMP%\\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%TEMP%\\CreateShortcut.vbs"
echo oLink.TargetPath = "%~dp0WindowsSystemOptimizer.exe" >> "%TEMP%\\CreateShortcut.vbs"
echo oLink.IconLocation = "%~dp0app_icon.ico" >> "%TEMP%\\CreateShortcut.vbs"
echo oLink.Save >> "%TEMP%\\CreateShortcut.vbs"
cscript //nologo "%TEMP%\\CreateShortcut.vbs"
del "%TEMP%\\CreateShortcut.vbs"

echo.
echo Shortcuts created successfully!
echo.
echo Press any key to start Windows System Optimizer...
pause > nul
start "" "WindowsSystemOptimizer.exe"
"""
        zipf.writestr("install.bat", install_bat)
        
    print(f"Installer package created: {zip_file}")
    print("Instructions:")
    print("1. Share this ZIP file with users")
    print("2. Users should extract all files and run install.bat to set up shortcuts")
    print("3. The application can be run directly from WindowsSystemOptimizer.exe")
    
    # Ask if user wants to create a self-extracting archive
    if messagebox.askyesno("Create Installer", "Would you like to create a self-extracting installer?"):
        try:
            installer_dir = "installer"
            os.makedirs(installer_dir, exist_ok=True)
            
            # Create a simple SFX configuration file
            sfx_config = """
;!@Install@!UTF-8!
Title="Windows System Optimizer Setup"
BeginPrompt="Do you want to install Windows System Optimizer?"
RunProgram="install.bat"
;!@InstallEnd@!
"""
            with open(os.path.join(installer_dir, "config.txt"), "w") as f:
                f.write(sfx_config)
            
            # Create a batch file to create the SFX
            with open(os.path.join(installer_dir, "create_sfx.bat"), "w") as f:
                f.write(f"""@echo off
echo Creating self-extracting installer...
copy /b 7zS.sfx + config.txt + ..\\{zip_file} WindowsSystemOptimizer_Setup.exe
echo Installer created at: installer\\WindowsSystemOptimizer_Setup.exe
pause
""")
            
            messagebox.showinfo("SFX Creator", 
                "To create a self-extracting installer:\n\n"
                "1. Download 7zS.sfx from 7-Zip website\n"
                "2. Place it in the 'installer' folder\n"
                "3. Run create_sfx.bat in that folder\n\n"
                "The resulting .exe file will be a standalone installer.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create SFX configuration: {str(e)}")

def main():
    create_installer()
    messagebox.showinfo("Installer Created", 
                       "Windows System Optimizer installer package has been created.\n\n"
                       "You can find it in the project directory.")

if __name__ == "__main__":
    main() 