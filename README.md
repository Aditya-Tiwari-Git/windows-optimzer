# Windows System Optimizer

A modern Windows system optimization utility with cleaning, network diagnostics, and performance monitoring features.

![Windows System Optimizer](app_icon.png)

## Features

- **System Dashboard**: Monitor CPU, memory, disk usage, and system health.
- **System Cleaning**: Clean temporary files, browser caches, and registry issues.
- **Network Diagnostics**: Test connection speed, ping, traceroute, and DNS lookup.
- **Optimization Tools**: Quick fixes for Microsoft Teams and Outlook issues, driver updates.

## Installation

### Option 1: Using the Installer

1. Download the latest installer from the releases page.
2. Run the installer and follow the on-screen instructions.
3. Launch the application from the Start Menu or Desktop shortcut.

### Option 2: Portable Executable

1. Download the latest executable from the releases page.
2. Extract the ZIP file to a location of your choice.
3. Run `WindowsSystemOptimizer.exe` to start the application.

## Building from Source

### Prerequisites

- Python 3.8 or higher
- Required Python packages (install using `pip install -r requirements.txt`):
  - PyQt5
  - PyQtWebEngine
  - matplotlib
  - numpy
  - psutil
  - pywin32
  - wmi
  - qrcode
  - requests
  - pyinstaller

### Running the Application

1. Clone the repository:

   ```
   git clone https://github.com/Aditya-Tiwari-Git/windows-optimzer.git
   cd windows-optimzer
   ```

2. Install dependencies:

   ```
   pip install -r requirements.txt
   ```

3. Run the application:
   ```
   python run_app.py
   ```

### Building the Executable

#### Using the build script (recommended):

1. Run the build batch file:

   ```
   build.bat
   ```

2. The executable will be created in the `dist` folder.

#### Manually using PyInstaller:

```
pyinstaller --name "WindowsSystemOptimizer" --icon=assets\app_icon.ico --onefile --windowed run_app.py
```

### Creating an Installer

If you have NSIS (Nullsoft Scriptable Install System) installed:

1. Run the installer creation script:

   ```
   create_installer.bat
   ```

2. The installer will be created in the `installer` folder.

## Development

### Project Structure

- **assets/**: Icons and images used in the application
- **scripts/**: PowerShell and batch scripts for system operations
- **services/**: Backend service modules
- **ui/**: User interface components
- **utils/**: Utility functions and helpers

## License

Copyright © 2023 WinOptimizer. All rights reserved.
