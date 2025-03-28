#!/usr/bin/env python3
"""
Build script for Windows System Optimizer.
This script packages the application as an executable (.exe) file using PyInstaller.
"""

import os
import sys
import shutil
import subprocess
import argparse
import platform
from datetime import datetime

def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description='Build Windows System Optimizer executable')
    
    parser.add_argument('--name', default='WindowsSystemOptimizer',
                        help='Name of the output executable')
    
    parser.add_argument('--version', default='1.0.0',
                        help='Version number for the application')
    
    parser.add_argument('--icon', default='assets/app_icon.ico',
                        help='Path to icon file for the executable')
    
    parser.add_argument('--onefile', action='store_true',
                        help='Create a single executable file')
    
    parser.add_argument('--noconsole', action='store_true',
                        help='Do not show console window when running the app')
    
    parser.add_argument('--clean', action='store_true',
                        help='Clean build directories before building')
    
    parser.add_argument('--upx', action='store_true',
                        help='Use UPX to compress the executable')
    
    return parser.parse_args()

def check_requirements():
    """Check if all required packages are installed."""
    required_packages = [
        'PyQt5',
        'psutil',
        'matplotlib',
        'pyinstaller',
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print("The following required packages are missing:")
        for package in missing_packages:
            print(f"  - {package}")
        print("\nPlease install them using pip:")
        print(f"pip install {' '.join(missing_packages)}")
        return False
    
    return True

def clean_build_dirs():
    """Clean build and dist directories."""
    directories = ['build', 'dist']
    spec_files = [f for f in os.listdir('.') if f.endswith('.spec')]
    
    for directory in directories:
        if os.path.exists(directory):
            print(f"Removing {directory} directory...")
            shutil.rmtree(directory)
    
    for spec_file in spec_files:
        print(f"Removing {spec_file}...")
        os.remove(spec_file)

def create_version_file(version):
    """Create a version file for the application."""
    version_parts = version.split('.')
    while len(version_parts) < 3:
        version_parts.append('0')
    
    major = int(version_parts[0])
    minor = int(version_parts[1])
    patch = int(version_parts[2])
    
    version_info = f'''
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({major}, {minor}, {patch}, 0),
    prodvers=({major}, {minor}, {patch}, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
        StringTable(
          u'040904B0',
          [StringStruct(u'CompanyName', u'WinOptimizer'),
          StringStruct(u'FileDescription', u'Windows System Optimizer'),
          StringStruct(u'FileVersion', u'{version}'),
          StringStruct(u'InternalName', u'winoptimizer'),
          StringStruct(u'LegalCopyright', u'Copyright (C) 2023 WinOptimizer'),
          StringStruct(u'OriginalFilename', u'WindowsSystemOptimizer.exe'),
          StringStruct(u'ProductName', u'Windows System Optimizer'),
          StringStruct(u'ProductVersion', u'{version}')])
      ]
    ),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
'''
    
    with open('version_info.txt', 'w') as f:
        f.write(version_info)
    
    print(f"Created version info file with version {version}")
    return 'version_info.txt'

def build_executable(args):
    """Build the executable using PyInstaller."""
    # Create version file
    version_file = create_version_file(args.version)
    
    # Build command
    cmd = ['pyinstaller', 'main.py']
    
    # Add name
    cmd.extend(['--name', args.name])
    
    # Add icon if it exists
    if os.path.exists(args.icon):
        cmd.extend(['--icon', args.icon])
    else:
        print(f"Warning: Icon file {args.icon} not found, using default icon")
    
    # Add version file
    cmd.extend(['--version-file', version_file])
    
    # Onefile option
    if args.onefile:
        cmd.append('--onefile')
    else:
        cmd.append('--onedir')
    
    # No console option
    if args.noconsole:
        cmd.append('--noconsole')
    
    # UPX compression
    if args.upx:
        cmd.append('--upx-dir=upx')
    
    # Add data files
    cmd.extend(['--add-data', f'scripts{os.pathsep}scripts'])
    
    # Explicitly include problematic packages
    cmd.extend(['--hidden-import', 'pkg_resources.py2_warn'])
    
    # Run PyInstaller
    print("Building executable with PyInstaller...")
    print(f"Command: {' '.join(cmd)}")
    
    subprocess.run(cmd, check=True)
    
    # Clean up
    if os.path.exists(version_file):
        os.remove(version_file)
    
    return True

def main():
    """Main build function."""
    print("=" * 60)
    print("Windows System Optimizer - Build Script")
    print("=" * 60)
    
    # Check if running on Windows
    if platform.system() != "Windows":
        print("Error: This build script is only for Windows systems.")
        return 1
    
    # Parse arguments
    args = parse_arguments()
    
    # Check requirements
    if not check_requirements():
        return 1
    
    # Clean build directories if requested
    if args.clean:
        clean_build_dirs()
    
    # Build the executable
    try:
        if build_executable(args):
            output_path = os.path.join('dist', f"{args.name}.exe") if args.onefile else os.path.join('dist', args.name)
            
            print("\nBuild completed successfully!")
            print(f"Output: {output_path}")
            
            # Get file size
            if os.path.exists(output_path):
                if os.path.isfile(output_path):
                    size_bytes = os.path.getsize(output_path)
                    size_mb = size_bytes / (1024 * 1024)
                    print(f"File size: {size_mb:.2f} MB")
                else:
                    print("Directory created successfully")
            
            return 0
        else:
            print("\nBuild failed!")
            return 1
    except Exception as e:
        print(f"\nBuild failed with error: {str(e)}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
