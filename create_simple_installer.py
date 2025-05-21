#!/usr/bin/env python3
"""
Script to create a simple self-extracting installer for Windows System Optimizer
This doesn't require NSIS, just standard Python libraries
"""

import os
import sys
import shutil
import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import base64
import zipfile
import traceback

def create_installer():
    """Create a simple self-extracting installer"""
    try:
        print("Creating simple installer for Windows System Optimizer...")
        
        # Paths
        dist_dir = "dist"
        installer_dir = "simple_installer"
        exe_file = os.path.join(dist_dir, "WindowsSystemOptimizer.exe")
        icon_file = os.path.join("assets", "app_icon.ico")
        
        # Check if the executable exists
        if not os.path.exists(exe_file):
            print(f"ERROR: Executable not found at {exe_file}.")
            print("Please build the application first using build.bat.")
            return False
        
        # Create installer directory
        os.makedirs(installer_dir, exist_ok=True)
        
        # Create installer Python script
        installer_script = os.path.join(installer_dir, "installer.py")
        
        # Write the installer script content
        installer_script_content = '''#!/usr/bin/env python3
"""
Self-extracting installer for Windows System Optimizer
"""

import os
import sys
import shutil
import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
from pathlib import Path
import base64
import tempfile
import zipfile
import io

try:
    import winreg
except ImportError:
    # Mock winreg for non-Windows platforms (for development)
    class WinregMock:
        HKEY_CURRENT_USER = None
        KEY_WRITE = None
        REG_SZ = None
        def CreateKeyEx(self, *args, **kwargs):
            return None
        def SetValueEx(self, *args, **kwargs):
            pass
        def CloseKey(self, *args, **kwargs):
            pass
        def DeleteKey(self, *args, **kwargs):
            pass
    winreg = WinregMock()

class InstallerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows System Optimizer Installer")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # Set icon if available
        try:
            # The icon will be embedded in the script as base64
            icon_data = """ICON_DATA_PLACEHOLDER"""
            if icon_data != "ICON_DATA_PLACEHOLDER":
                # Create a temporary icon file
                icon_data = base64.b64decode(icon_data)
                self.temp_icon = tempfile.NamedTemporaryFile(suffix='.ico', delete=False)
                self.temp_icon.write(icon_data)
                self.temp_icon.close()
                self.root.iconbitmap(self.temp_icon.name)
        except Exception as e:
            print(f"Could not set icon: {e}")
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Welcome label
        welcome_label = tk.Label(
            main_frame, 
            text="Welcome to Windows System Optimizer Setup",
            font=("Arial", 14, "bold")
        )
        welcome_label.pack(pady=(0, 20))
        
        # Description
        description = tk.Label(
            main_frame,
            text="This will install Windows System Optimizer on your computer.\\n"
                 "Click Install to begin the installation or Cancel to exit.",
            justify=tk.LEFT,
            wraplength=550
        )
        description.pack(pady=(0, 20), anchor=tk.W)
        
        # Installation path
        path_frame = tk.Frame(main_frame)
        path_frame.pack(fill=tk.X, pady=(0, 20))
        
        path_label = tk.Label(path_frame, text="Installation Directory:")
        path_label.pack(side=tk.LEFT, padx=(0, 10))
        
        default_path = os.path.join(os.environ.get("PROGRAMFILES", "C:\\Program Files"), "Windows System Optimizer")
        self.install_path = tk.StringVar(value=default_path)
        
        path_entry = tk.Entry(path_frame, textvariable=self.install_path, width=50)
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        browse_button = tk.Button(path_frame, text="Browse...", command=self.browse_location)
        browse_button.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Create desktop shortcut option
        self.create_shortcut = tk.BooleanVar(value=True)
        shortcut_check = tk.Checkbutton(
            main_frame, 
            text="Create Desktop Shortcut",
            variable=self.create_shortcut
        )
        shortcut_check.pack(anchor=tk.W, pady=(0, 10))
        
        # Create start menu shortcut option
        self.create_startmenu = tk.BooleanVar(value=True)
        startmenu_check = tk.Checkbutton(
            main_frame, 
            text="Create Start Menu Shortcut",
            variable=self.create_startmenu
        )
        startmenu_check.pack(anchor=tk.W, pady=(0, 20))
        
        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        cancel_button = tk.Button(
            button_frame, 
            text="Cancel",
            width=10,
            command=self.root.destroy
        )
        cancel_button.pack(side=tk.RIGHT, padx=(10, 0))
        
        install_button = tk.Button(
            button_frame, 
            text="Install",
            width=10,
            command=self.start_installation
        )
        install_button.pack(side=tk.RIGHT)
    
    def browse_location(self):
        from tkinter import filedialog
        directory = filedialog.askdirectory(
            initialdir=self.install_path.get(),
            title="Select Installation Directory"
        )
        if directory:
            self.install_path.set(directory)
    
    def start_installation(self):
        install_dir = self.install_path.get()
        
        # Create progress window
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Installing...")
        progress_window.geometry("500x150")
        progress_window.resizable(False, False)
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Center the progress window
        progress_window.update_idletasks()
        width = progress_window.winfo_width()
        height = progress_window.winfo_height()
        x = (progress_window.winfo_screenwidth() // 2) - (width // 2)
        y = (progress_window.winfo_screenheight() // 2) - (height // 2)
        progress_window.geometry(f"{width}x{height}+{x}+{y}")
        
        # Progress UI
        frame = tk.Frame(progress_window, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        status_label = tk.Label(frame, text="Installing Windows System Optimizer...")
        status_label.pack(pady=(0, 10))
        
        progress = ttk.Progressbar(frame, mode="indeterminate")
        progress.pack(fill=tk.X, pady=(0, 10))
        progress.start()
        
        detail_label = tk.Label(frame, text="")
        detail_label.pack(anchor=tk.W)
        
        # Schedule the installation process
        self.root.after(100, lambda: self.perform_installation(
            install_dir, 
            self.create_shortcut.get(),
            self.create_startmenu.get(),
            progress_window,
            status_label,
            detail_label
        ))
    
    def perform_installation(self, install_dir, create_shortcut, create_startmenu, 
                            progress_window, status_label, detail_label):
        try:
            # Create installation directory if it doesn't exist
            os.makedirs(install_dir, exist_ok=True)
            detail_label.config(text=f"Creating directory: {install_dir}")
            self.root.update()
            
            # Extract the embedded executable
            detail_label.config(text="Extracting application files...")
            self.root.update()
            
            # The application binary will be embedded in the script as base64
            app_data = """APP_DATA_PLACEHOLDER"""
            if app_data != "APP_DATA_PLACEHOLDER":
                app_data = base64.b64decode(app_data)
                zip_data = io.BytesIO(app_data)
                
                with zipfile.ZipFile(zip_data, 'r') as zip_ref:
                    zip_ref.extractall(install_dir)
            
            # Create shortcuts
            if create_shortcut:
                detail_label.config(text="Creating desktop shortcut...")
                self.root.update()
                self.create_desktop_shortcut(install_dir)
            
            if create_startmenu:
                detail_label.config(text="Creating Start Menu shortcut...")
                self.root.update()
                self.create_startmenu_shortcut(install_dir)
            
            # Add to registry for uninstallation
            detail_label.config(text="Registering application...")
            self.root.update()
            self.add_to_registry(install_dir)
            
            # Create uninstaller
            detail_label.config(text="Creating uninstaller...")
            self.root.update()
            self.create_uninstaller(install_dir)
            
            # Installation complete
            progress_window.destroy()
            messagebox.showinfo(
                "Installation Complete",
                f"Windows System Optimizer has been successfully installed to:\\n{install_dir}"
            )
            
            # Ask if the user wants to launch the application
            if messagebox.askyesno(
                "Launch Application",
                "Would you like to launch Windows System Optimizer now?"
            ):
                exe_path = os.path.join(install_dir, "WindowsSystemOptimizer.exe")
                subprocess.Popen([exe_path])
            
            self.root.destroy()
            
        except Exception as e:
            progress_window.destroy()
            messagebox.showerror(
                "Installation Error",
                f"An error occurred during installation:\\n{str(e)}"
            )
    
    def create_desktop_shortcut(self, install_dir):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        shortcut_path = os.path.join(desktop, "Windows System Optimizer.lnk")
        
        exe_path = os.path.join(install_dir, "WindowsSystemOptimizer.exe")
        icon_path = os.path.join(install_dir, "app_icon.ico")
        
        try:
            # Try to use pywin32
            try:
                from win32com.client import Dispatch
                shell = Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.TargetPath = exe_path
                if os.path.exists(icon_path):
                    shortcut.IconLocation = icon_path
                shortcut.WorkingDirectory = install_dir
                shortcut.Description = "Windows System Optimizer"
                shortcut.save()
            except ImportError:
                # Fallback to creating batch file as shortcut
                with open(shortcut_path + ".bat", "w") as f:
                    f.write(f'@echo off\\nstart "" "{exe_path}"')
        except Exception as e:
            print(f"Could not create desktop shortcut: {e}")
    
    def create_startmenu_shortcut(self, install_dir):
        start_menu = os.path.join(
            os.environ.get("APPDATA", os.path.expanduser("~")), 
            "Microsoft", "Windows", "Start Menu", "Programs",
            "Windows System Optimizer"
        )
        os.makedirs(start_menu, exist_ok=True)
        
        shortcut_path = os.path.join(start_menu, "Windows System Optimizer.lnk")
        uninstall_path = os.path.join(start_menu, "Uninstall.lnk")
        
        exe_path = os.path.join(install_dir, "WindowsSystemOptimizer.exe")
        uninstaller_path = os.path.join(install_dir, "uninstall.exe")
        icon_path = os.path.join(install_dir, "app_icon.ico")
        
        try:
            # Try to use pywin32
            try:
                from win32com.client import Dispatch
                shell = Dispatch("WScript.Shell")
                
                # Application shortcut
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.TargetPath = exe_path
                if os.path.exists(icon_path):
                    shortcut.IconLocation = icon_path
                shortcut.WorkingDirectory = install_dir
                shortcut.Description = "Windows System Optimizer"
                shortcut.save()
                
                # Uninstaller shortcut
                shortcut = shell.CreateShortCut(uninstall_path)
                shortcut.TargetPath = uninstaller_path
                shortcut.WorkingDirectory = install_dir
                shortcut.Description = "Uninstall Windows System Optimizer"
                shortcut.save()
            except ImportError:
                # Fallback to batch files
                with open(shortcut_path + ".bat", "w") as f:
                    f.write(f'@echo off\\nstart "" "{exe_path}"')
                with open(uninstall_path + ".bat", "w") as f:
                    f.write(f'@echo off\\nstart "" "{uninstaller_path}"')
        except Exception as e:
            print(f"Could not create start menu shortcuts: {e}")
    
    def add_to_registry(self, install_dir):
        try:
            exe_path = os.path.join(install_dir, "WindowsSystemOptimizer.exe")
            uninstaller_path = os.path.join(install_dir, "uninstall.exe")
            icon_path = os.path.join(install_dir, "app_icon.ico")
            
            registry_key = winreg.CreateKeyEx(
                winreg.HKEY_CURRENT_USER,
                r"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\WindowsSystemOptimizer",
                0,
                winreg.KEY_WRITE
            )
            
            winreg.SetValueEx(registry_key, "DisplayName", 0, winreg.REG_SZ, "Windows System Optimizer")
            winreg.SetValueEx(registry_key, "UninstallString", 0, winreg.REG_SZ, uninstaller_path)
            winreg.SetValueEx(registry_key, "DisplayIcon", 0, winreg.REG_SZ, icon_path)
            winreg.SetValueEx(registry_key, "DisplayVersion", 0, winreg.REG_SZ, "1.0.0")
            winreg.SetValueEx(registry_key, "Publisher", 0, winreg.REG_SZ, "WinOptimizer")
            winreg.SetValueEx(registry_key, "InstallLocation", 0, winreg.REG_SZ, install_dir)
            
            winreg.CloseKey(registry_key)
        except Exception as e:
            print(f"Could not add registry entry: {e}")
    
    def create_uninstaller(self, install_dir):
        uninstaller_path = os.path.join(install_dir, "uninstall.exe")
        uninstaller_script = os.path.join(install_dir, "uninstall.py")
        
        with open(uninstaller_script, "w") as f:
            f.write("""#!/usr/bin/env python3
import os
import sys
import shutil
import tkinter as tk
from tkinter import messagebox
import tempfile

try:
    import winreg
except ImportError:
    # Mock winreg for non-Windows platforms
    class WinregMock:
        HKEY_CURRENT_USER = None
        def DeleteKey(self, *args, **kwargs):
            pass
    winreg = WinregMock()

def remove_registry_key():
    try:
        winreg.DeleteKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\WindowsSystemOptimizer"
        )
    except Exception as e:
        print(f"Could not remove registry key: {e}")

def remove_shortcuts():
    # Remove desktop shortcut
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    desktop_shortcut = os.path.join(desktop, "Windows System Optimizer.lnk")
    if os.path.exists(desktop_shortcut):
        os.remove(desktop_shortcut)
    if os.path.exists(desktop_shortcut + ".bat"):
        os.remove(desktop_shortcut + ".bat")
    
    # Remove start menu shortcuts
    start_menu = os.path.join(
        os.environ.get("APPDATA", os.path.expanduser("~")), 
        "Microsoft", "Windows", "Start Menu", "Programs",
        "Windows System Optimizer"
    )
    if os.path.exists(start_menu):
        shutil.rmtree(start_menu)

def uninstall():
    # Get the installation directory
    install_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Create a temporary batch file to delete the installation directory
    # This is needed because we can't delete the directory while the uninstaller is running
    temp_dir = tempfile.gettempdir()
    batch_file = os.path.join(temp_dir, "uninstall_wso.bat")
    
    # Wait for a few seconds before deleting to ensure the uninstaller has exited
    with open(batch_file, "w") as f:
        f.write(f'''@echo off
timeout /t 2 /nobreak > nul
rmdir /S /Q "{install_dir}"
del "%~f0"
''')
    
    # First remove registry entries and shortcuts
    remove_registry_key()
    remove_shortcuts()
    
    # Ask for confirmation before uninstalling
    root = tk.Tk()
    root.withdraw()
    
    if messagebox.askyesno(
        "Confirm Uninstall",
        "Are you sure you want to uninstall Windows System Optimizer?"
    ):
        # Run the batch file to delete the installation directory
        os.system(f'start /b "" cmd /c "{batch_file}"')
        messagebox.showinfo(
            "Uninstall Complete",
            "Windows System Optimizer has been uninstalled."
        )
    
    root.destroy()

if __name__ == "__main__":
    uninstall()
""")
        
        # Create a batch file to run the uninstaller script
        with open(os.path.join(install_dir, "uninstall.bat"), "w") as f:
            f.write(f'''@echo off
python "{uninstaller_script}"
''')
        
        # Copy the uninstaller.bat to uninstall.exe
        # For a real uninstaller, you'd compile it, but for simplicity we'll just copy the batch file
        shutil.copy(os.path.join(install_dir, "uninstall.bat"), uninstaller_path)

def main():
    root = tk.Tk()
    app = InstallerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
'''
        
        with open(installer_script, "w") as f:
            f.write(installer_script_content)
        
        # Create a batch file to launch the installer
        installer_batch = os.path.join(installer_dir, "WindowsSystemOptimizer_Setup.bat")
        with open(installer_batch, "w") as f:
            f.write(f'''@echo off
python "{os.path.basename(installer_script)}"
''')
        
        # Now let's embed the executable in the installer script
        print("Embedding executable in installer script...")
        
        # First create a zip file of the necessary files
        zip_file = os.path.join(installer_dir, "app.zip")
        
        with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add the executable
            zipf.write(exe_file, os.path.basename(exe_file))
            
            # Add the icon
            if os.path.exists(icon_file):
                zipf.write(icon_file, os.path.basename(icon_file))
        
        # Read the zip file and encode it as base64
        with open(zip_file, 'rb') as f:
            zip_data = f.read()
        
        zip_base64 = base64.b64encode(zip_data).decode('utf-8')
        
        # Read the icon and encode it as base64
        icon_base64 = ""
        if os.path.exists(icon_file):
            with open(icon_file, 'rb') as f:
                icon_data = f.read()
            icon_base64 = base64.b64encode(icon_data).decode('utf-8')
        
        # Replace the placeholders in the installer script
        with open(installer_script, 'r') as f:
            content = f.read()
        
        content = content.replace('"""APP_DATA_PLACEHOLDER"""', f'"""{zip_base64}"""')
        content = content.replace('"""ICON_DATA_PLACEHOLDER"""', f'"""{icon_base64}"""')
        
        with open(installer_script, 'w') as f:
            f.write(content)
        
        # Clean up the temporary zip file
        os.remove(zip_file)
        
        print("===================================================")
        print("Simple installer created successfully!")
        print(f"Batch installer: {installer_batch}")
        print("===================================================")
        print("To use the installer:")
        print("1. Run the batch file directly")
        print("2. It will install the application and create shortcuts")
        print("===================================================")
        
        # Try to create standalone executable using PyInstaller
        try:
            print("Checking if PyInstaller is available to create standalone installer...")
            subprocess.run(["pyinstaller", "--version"], check=True, capture_output=True)
            
            print("Creating standalone installer executable with PyInstaller...")
            
            # Create a batch file to build the installer
            build_installer_batch = os.path.join(installer_dir, "build_installer.bat")
            with open(build_installer_batch, "w") as f:
                f.write(f'''@echo off
cd "{installer_dir}"
pyinstaller --onefile --windowed {"--icon=../" + icon_file if os.path.exists(icon_file) else ""} --name=WindowsSystemOptimizer_Setup installer.py
mkdir ..\installer
copy dist\WindowsSystemOptimizer_Setup.exe ..\installer\
echo Installer created at: installer\WindowsSystemOptimizer_Setup.exe
pause
''')
            
            print(f"Created build script: {build_installer_batch}")
            print("Run this script to create a standalone installer exe.")
            
            return True
            
        except Exception as e:
            print("PyInstaller not found or error creating standalone installer:")
            print(f"Error: {e}")
            print("You can still use the batch file for installation.")
            return True
            
    except Exception as e:
        print("Error creating installer:")
        print(f"Error: {e}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    create_installer() 