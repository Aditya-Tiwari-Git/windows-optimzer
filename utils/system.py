"""
System utilities for the Windows System Optimizer.
This module provides helper functions for system operations.
"""

import os
import sys
import logging
import subprocess
import winreg
import ctypes
import platform
import tempfile
import shutil
from datetime import datetime

logger = logging.getLogger(__name__)

def is_admin():
    """
    Check if the current process has administrator privileges.
    
    Returns:
        bool: True if running as admin, False otherwise
    """
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception as e:
        logger.error(f"Error checking admin status: {str(e)}")
        return False

def restart_as_admin():
    """
    Restart the current script with administrator privileges.
    
    Returns:
        bool: True if restart initiated, False on error
    """
    try:
        if is_admin():
            return True  # Already running as admin
        
        script = sys.argv[0]
        args = ' '.join(sys.argv[1:])
        
        # Use ctypes to elevate privileges
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, f'"{script}" {args}', None, 1
        )
        
        # Exit current non-elevated process
        sys.exit()
        
    except Exception as e:
        logger.error(f"Error restarting as admin: {str(e)}")
        return False

def get_windows_version():
    """
    Get detailed Windows version information.
    
    Returns:
        dict: Windows version details
    """
    try:
        version_info = {
            'version': platform.version(),
            'win32_ver': platform.win32_ver(),
            'release': platform.release(),
            'build': '',
            'edition': '',
            'display_version': ''
        }
        
        # Get more detailed version info from registry
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion") as key:
                version_info['build'] = winreg.QueryValueEx(key, "CurrentBuild")[0]
                version_info['edition'] = winreg.QueryValueEx(key, "EditionID")[0]
                
                # DisplayVersion exists in Windows 10 version 2004 and later
                try:
                    version_info['display_version'] = winreg.QueryValueEx(key, "DisplayVersion")[0]
                except:
                    pass
        except Exception as e:
            logger.warning(f"Error getting detailed Windows version: {str(e)}")
        
        return version_info
        
    except Exception as e:
        logger.error(f"Error getting Windows version: {str(e)}")
        return {
            'version': 'Unknown',
            'win32_ver': ('Unknown', 'Unknown', 'Unknown'),
            'release': 'Unknown',
            'build': 'Unknown',
            'edition': 'Unknown',
            'display_version': 'Unknown'
        }

def get_system_directories():
    """
    Get important system directories.
    
    Returns:
        dict: Dictionary of system directory paths
    """
    try:
        user_profile = os.environ.get('USERPROFILE', '')
        windows_dir = os.environ.get('WINDIR', 'C:\\Windows')
        
        directories = {
            'user_profile': user_profile,
            'app_data': os.path.join(user_profile, 'AppData'),
            'local_app_data': os.path.join(user_profile, 'AppData', 'Local'),
            'roaming_app_data': os.path.join(user_profile, 'AppData', 'Roaming'),
            'windows': windows_dir,
            'windows_temp': os.path.join(windows_dir, 'Temp'),
            'user_temp': tempfile.gettempdir(),
            'program_files': os.environ.get('ProgramFiles', 'C:\\Program Files'),
            'program_files_x86': os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'),
            'documents': os.path.join(user_profile, 'Documents'),
            'downloads': os.path.join(user_profile, 'Downloads')
        }
        
        return directories
        
    except Exception as e:
        logger.error(f"Error getting system directories: {str(e)}")
        return {}

def get_windows_services(status=None):
    """
    Get Windows services and their status.
    
    Args:
        status (str, optional): Filter by status (running, stopped, etc.)
    
    Returns:
        list: List of services with details
    """
    try:
        # PowerShell command to get services
        ps_command = "Get-Service"
        
        # Add status filter if specified
        if status:
            ps_command += f" | Where-Object {{ $_.Status -eq '{status}' }}"
        
        # Add output formatting
        ps_command += " | Select-Object Name, DisplayName, Status | ConvertTo-Json"
        
        # Execute command
        result = subprocess.run(
            ['powershell.exe', '-Command', ps_command],
            capture_output=True,
            text=True,
            check=True
        )
        
        # Parse JSON output
        import json
        services = json.loads(result.stdout)
        
        # Ensure services is a list
        if not isinstance(services, list):
            services = [services]
        
        return services
        
    except Exception as e:
        logger.error(f"Error getting Windows services: {str(e)}")
        return []

def get_windows_features():
    """
    Get installed Windows features.
    
    Returns:
        list: List of Windows features with details
    """
    try:
        # PowerShell command to get Windows features
        ps_command = "Get-WindowsOptionalFeature -Online | Select-Object FeatureName, State | ConvertTo-Json"
        
        # Execute command
        result = subprocess.run(
            ['powershell.exe', '-Command', ps_command],
            capture_output=True,
            text=True,
            check=True
        )
        
        # Parse JSON output
        import json
        features = json.loads(result.stdout)
        
        # Ensure features is a list
        if not isinstance(features, list):
            features = [features]
        
        return features
        
    except Exception as e:
        logger.error(f"Error getting Windows features: {str(e)}")
        return []

def enable_windows_feature(feature_name):
    """
    Enable a Windows feature.
    
    Args:
        feature_name (str): Name of the feature to enable
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if not is_admin():
            logger.warning("Administrator privileges required to enable Windows features")
            return False
        
        # PowerShell command to enable feature
        ps_command = f"Enable-WindowsOptionalFeature -Online -FeatureName {feature_name} -NoRestart"
        
        # Execute command
        result = subprocess.run(
            ['powershell.exe', '-Command', ps_command],
            capture_output=True,
            text=True,
            check=True
        )
        
        return True
        
    except Exception as e:
        logger.error(f"Error enabling Windows feature: {str(e)}")
        return False

def disable_windows_feature(feature_name):
    """
    Disable a Windows feature.
    
    Args:
        feature_name (str): Name of the feature to disable
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if not is_admin():
            logger.warning("Administrator privileges required to disable Windows features")
            return False
        
        # PowerShell command to disable feature
        ps_command = f"Disable-WindowsOptionalFeature -Online -FeatureName {feature_name} -NoRestart"
        
        # Execute command
        result = subprocess.run(
            ['powershell.exe', '-Command', ps_command],
            capture_output=True,
            text=True,
            check=True
        )
        
        return True
        
    except Exception as e:
        logger.error(f"Error disabling Windows feature: {str(e)}")
        return False

def create_system_restore_point(description="Windows System Optimizer Restore Point"):
    """
    Create a system restore point.
    
    Args:
        description (str): Description for the restore point
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if not is_admin():
            logger.warning("Administrator privileges required to create restore points")
            return False
        
        # PowerShell command to create restore point
        ps_command = f'Checkpoint-Computer -Description "{description}" -RestorePointType "MODIFY_SETTINGS"'
        
        # Execute command
        result = subprocess.run(
            ['powershell.exe', '-Command', ps_command],
            capture_output=True,
            text=True,
            check=True
        )
        
        return True
        
    except Exception as e:
        logger.error(f"Error creating system restore point: {str(e)}")
        return False

def get_process_details(pid=None, name=None):
    """
    Get details about a process by PID or name.
    
    Args:
        pid (int, optional): Process ID
        name (str, optional): Process name
    
    Returns:
        dict or list: Process details
    """
    try:
        import psutil
        
        # If PID is specified
        if pid:
            try:
                process = psutil.Process(pid)
                return {
                    'pid': process.pid,
                    'name': process.name(),
                    'status': process.status(),
                    'cpu_percent': process.cpu_percent(),
                    'memory_percent': process.memory_percent(),
                    'username': process.username(),
                    'create_time': datetime.fromtimestamp(process.create_time()).strftime('%Y-%m-%d %H:%M:%S'),
                    'exe': process.exe(),
                    'cmdline': process.cmdline()
                }
            except psutil.NoSuchProcess:
                return {}
        
        # If name is specified
        elif name:
            processes = []
            for process in psutil.process_iter(['pid', 'name', 'status', 'cpu_percent', 'memory_percent', 'username', 'create_time', 'exe', 'cmdline']):
                if process.info['name'].lower() == name.lower():
                    process.info['create_time'] = datetime.fromtimestamp(process.info['create_time']).strftime('%Y-%m-%d %H:%M:%S')
                    processes.append(process.info)
            return processes
        
        # If neither is specified, return all processes
        else:
            processes = []
            for process in psutil.process_iter(['pid', 'name', 'status', 'cpu_percent', 'memory_percent', 'username', 'create_time']):
                process.info['create_time'] = datetime.fromtimestamp(process.info['create_time']).strftime('%Y-%m-%d %H:%M:%S')
                processes.append(process.info)
            return processes
        
    except Exception as e:
        logger.error(f"Error getting process details: {str(e)}")
        return [] if name or pid is None else {}

def is_process_running(process_name):
    """
    Check if a process is running by name.
    
    Args:
        process_name (str): Process name to check
    
    Returns:
        bool: True if running, False otherwise
    """
    try:
        import psutil
        
        for process in psutil.process_iter(['name']):
            if process.info['name'].lower() == process_name.lower():
                return True
        
        return False
        
    except Exception as e:
        logger.error(f"Error checking if process is running: {str(e)}")
        return False

def kill_process(pid=None, name=None, force=False):
    """
    Kill a process by PID or name.
    
    Args:
        pid (int, optional): Process ID
        name (str, optional): Process name
        force (bool): Force kill if True
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        import psutil
        
        # If PID is specified
        if pid:
            try:
                process = psutil.Process(pid)
                if force:
                    process.kill()
                else:
                    process.terminate()
                return True
            except psutil.NoSuchProcess:
                return False
        
        # If name is specified
        elif name:
            killed = False
            for process in psutil.process_iter(['pid', 'name']):
                if process.info['name'].lower() == name.lower():
                    try:
                        if force:
                            process.kill()
                        else:
                            process.terminate()
                        killed = True
                    except psutil.NoSuchProcess:
                        continue
            return killed
        
        return False
        
    except Exception as e:
        logger.error(f"Error killing process: {str(e)}")
        return False

def get_startup_items():
    """
    Get items configured to run at Windows startup.
    
    Returns:
        list: List of startup items with details
    """
    try:
        startup_items = []
        
        # Check registry run keys
        reg_locations = [
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run"),
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\RunOnce"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce")
        ]
        
        for reg_key, reg_path in reg_locations:
            try:
                with winreg.OpenKey(reg_key, reg_path) as key:
                    i = 0
                    while True:
                        try:
                            name, value, _ = winreg.EnumValue(key, i)
                            startup_items.append({
                                'name': name,
                                'command': value,
                                'location': f"{reg_key}\\{reg_path}",
                                'type': 'Registry'
                            })
                            i += 1
                        except WindowsError:
                            break
            except FileNotFoundError:
                continue
        
        # Check startup folders
        startup_folders = [
            os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup'),
            os.path.join(os.environ.get('ProgramData', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        ]
        
        for folder in startup_folders:
            if os.path.exists(folder):
                for item in os.listdir(folder):
                    item_path = os.path.join(folder, item)
                    startup_items.append({
                        'name': item,
                        'command': item_path,
                        'location': folder,
                        'type': 'Folder'
                    })
        
        return startup_items
        
    except Exception as e:
        logger.error(f"Error getting startup items: {str(e)}")
        return []

def set_application_startup(name, command, enable=True):
    """
    Add or remove application from Windows startup.
    
    Args:
        name (str): Name of the startup entry
        command (str): Command to execute at startup
        enable (bool): Add to startup if True, remove if False
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Path to Run key
        run_key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        
        if enable:
            # Add to startup
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key_path, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, command)
        else:
            # Remove from startup
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key_path, 0, winreg.KEY_WRITE) as key:
                try:
                    winreg.DeleteValue(key, name)
                except FileNotFoundError:
                    # Entry doesn't exist, nothing to remove
                    pass
        
        return True
        
    except Exception as e:
        logger.error(f"Error setting application startup: {str(e)}")
        return False
