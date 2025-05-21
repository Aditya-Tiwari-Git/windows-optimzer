#!/usr/bin/env python3
"""
Dependency checker for Windows System Optimizer.
This script checks if all required dependencies are installed.
"""

import os
import sys
import subprocess
import importlib.util

def is_module_installed(module_name):
    """Check if a Python module is installed.
    
    Args:
        module_name: Name of the module to check
        
    Returns:
        True if the module is installed, False otherwise
    """
    try:
        if module_name == 'PyQtWebEngine':
            # Special check for PyQtWebEngine
            import PyQt5.QtWebEngineWidgets
            return True
        elif module_name == 'pyinstaller':
            # Special check for pyinstaller
            try:
                import PyInstaller
                return True
            except ImportError:
                try:
                    result = subprocess.run(['pyinstaller', '--version'], 
                                           stdout=subprocess.PIPE, 
                                           stderr=subprocess.PIPE,
                                           text=True,
                                           check=False)
                    return result.returncode == 0
                except FileNotFoundError:
                    return False
        else:
            spec = importlib.util.find_spec(module_name)
            return spec is not None
    except ImportError:
        return False

def get_module_version(module_name):
    """Get the version of an installed module.
    
    Args:
        module_name: Name of the module to check
        
    Returns:
        Version string or 'Unknown' if unable to determine
    """
    try:
        if module_name == 'PyQtWebEngine':
            # Try to get the PyQtWebEngine version
            try:
                import PyQt5.QtWebEngine
                return getattr(PyQt5.QtWebEngine, '__version__', 'Unknown')
            except (ImportError, AttributeError):
                return 'Unknown'
        elif module_name == 'pyinstaller':
            # Special handling for pyinstaller
            try:
                import PyInstaller
                return getattr(PyInstaller, '__version__', 'Unknown')
            except ImportError:
                try:
                    result = subprocess.run(['pyinstaller', '--version'], 
                                           stdout=subprocess.PIPE, 
                                           stderr=subprocess.PIPE,
                                           text=True,
                                           check=False)
                    if result.returncode == 0:
                        return result.stdout.strip()
                    return 'Unknown'
                except FileNotFoundError:
                    return 'Not installed'
        
        module = __import__(module_name)
        if hasattr(module, '__version__'):
            return module.__version__
        if hasattr(module, 'version'):
            return module.version
        return 'Unknown'
    except (ImportError, AttributeError):
        return 'Not installed'

def check_dependencies():
    """Check all required dependencies for the application.
    
    Returns:
        Dict with results of the dependency check
    """
    required_packages = [
        'PyQt5',
        'PyQtWebEngine',
        'matplotlib',
        'numpy',
        'psutil',
        'win32con',  # Part of pywin32
        'wmi',
        'qrcode',
        'requests',
    ]
    
    optional_packages = [
        'pyinstaller',
    ]
    
    results = {
        'required': {},
        'optional': {},
        'missing_required': [],
        'missing_optional': []
    }
    
    # Check required packages
    for package in required_packages:
        installed = is_module_installed(package)
        version = get_module_version(package) if installed else 'Not installed'
        
        results['required'][package] = {
            'installed': installed,
            'version': version
        }
        
        if not installed:
            results['missing_required'].append(package)
    
    # Check optional packages
    for package in optional_packages:
        installed = is_module_installed(package)
        version = get_module_version(package) if installed else 'Not installed'
        
        results['optional'][package] = {
            'installed': installed,
            'version': version
        }
        
        if not installed:
            results['missing_optional'].append(package)
    
    # Additional check for pywin32 which has special import names
    if not results['required'].get('win32con', {}).get('installed', False):
        try:
            import win32api
            results['required']['win32con'] = {
                'installed': True,
                'version': 'Unknown'
            }
        except ImportError:
            pass
    
    return results

def main():
    """Main function to run the dependency check."""
    print("Windows System Optimizer - Dependency Checker")
    print("=" * 50)
    
    results = check_dependencies()
    
    # Display required packages
    print("\nRequired Packages:")
    print("-" * 50)
    for package, info in results['required'].items():
        status = "✓ Installed" if info['installed'] else "✗ Not installed"
        version = info['version']
        print(f"{package:15} | {status:15} | Version: {version}")
    
    # Display optional packages
    print("\nOptional Packages:")
    print("-" * 50)
    for package, info in results['optional'].items():
        status = "✓ Installed" if info['installed'] else "✗ Not installed"
        version = info['version']
        print(f"{package:15} | {status:15} | Version: {version}")
    
    # Summary
    if results['missing_required']:
        print("\nMissing required packages:")
        for package in results['missing_required']:
            print(f"  - {package}")
        
        print("\nTo install missing packages, run:")
        print("pip install -r requirements.txt")
        return 1
    else:
        print("\nAll required packages are installed! ✓")
        
        if results['missing_optional']:
            print("\nMissing optional packages:")
            for package in results['missing_optional']:
                print(f"  - {package}")
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 