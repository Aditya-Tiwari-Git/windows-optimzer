"""
Registry cleaning and maintenance utilities for Windows.
"""

import os
import sys
import re
import time
import logging
import winreg
import ctypes
import datetime
import tempfile
import subprocess
from threading import Lock
from pathlib import Path

logger = logging.getLogger(__name__)


class RegistryCleaner:
    """Utility for scanning and fixing Windows registry issues."""
    
    def __init__(self):
        """Initialize registry cleaner and set up target registry keys."""
        # Lock for thread-safe registry operations
        self.reg_lock = Lock()
        
        # Initialize registry keys to scan
        self.registry_keys = self._get_registry_keys()
        
        # Temporary backup location
        self.temp_dir = os.path.join(tempfile.gettempdir(), "RegistryBackup")
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir, exist_ok=True)
        
        # Found issues storage
        self.issues = []
    
    def _get_registry_keys(self):
        """Define registry keys and areas to scan.
        
        Returns:
            Dict of registry keys to scan by category
        """
        return {
            "Software and App Paths": [
                {
                    "root": winreg.HKEY_LOCAL_MACHINE,
                    "path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
                    "subkeys": True,
                    "check_values": True,
                    "check_file_exists": True,
                    "file_value_names": ["", "Path"]
                },
                {
                    "root": winreg.HKEY_CURRENT_USER,
                    "path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
                    "subkeys": True,
                    "check_values": True,
                    "check_file_exists": True,
                    "file_value_names": ["", "Path"] 
                },
                {
                    "root": winreg.HKEY_LOCAL_MACHINE,
                    "path": r"SOFTWARE\Classes\Applications",
                    "subkeys": True,
                    "check_values": True,
                    "check_file_exists": True,
                    "file_value_names": ["FriendlyAppName", "ApplicationName"]
                }
            ],
            "File Extensions": [
                {
                    "root": winreg.HKEY_CLASSES_ROOT,
                    "path": "",
                    "subkeys": True,
                    "check_values": True,
                    "include_pattern": r"^\.[a-zA-Z0-9]+$",  # Match file extensions like .txt
                    "exclude_pattern": r"\.dll$|\.exe$|\.com$|\.bat$",  # Skip binary file extensions
                    "check_file_exists": False
                },
                {
                    "root": winreg.HKEY_CURRENT_USER,
                    "path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts",
                    "subkeys": True,
                    "check_values": True,
                    "check_file_exists": False
                }
            ],
            "Startup Programs": [
                {
                    "root": winreg.HKEY_CURRENT_USER,
                    "path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
                    "subkeys": False,
                    "check_values": True,
                    "check_file_exists": True,
                    "file_value_names": ["*"]  # All values
                },
                {
                    "root": winreg.HKEY_LOCAL_MACHINE,
                    "path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
                    "subkeys": False,
                    "check_values": True,
                    "check_file_exists": True,
                    "file_value_names": ["*"]  # All values
                },
                {
                    "root": winreg.HKEY_CURRENT_USER,
                    "path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
                    "subkeys": False,
                    "check_values": True,
                    "check_file_exists": True,
                    "file_value_names": ["*"]  # All values
                },
                {
                    "root": winreg.HKEY_LOCAL_MACHINE,
                    "path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
                    "subkeys": False,
                    "check_values": True,
                    "check_file_exists": True,
                    "file_value_names": ["*"]  # All values
                }
            ],
            "Uninstall Information": [
                {
                    "root": winreg.HKEY_LOCAL_MACHINE,
                    "path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    "subkeys": True,
                    "check_values": True,
                    "check_file_exists": True,
                    "file_value_names": ["UninstallString", "DisplayIcon"]
                },
                {
                    "root": winreg.HKEY_LOCAL_MACHINE,
                    "path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                    "subkeys": True,
                    "check_values": True,
                    "check_file_exists": True,
                    "file_value_names": ["UninstallString", "DisplayIcon"]
                }
            ],
            "Shared DLLs": [
                {
                    "root": winreg.HKEY_LOCAL_MACHINE,
                    "path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs",
                    "subkeys": False,
                    "check_values": True,
                    "check_file_exists": True,
                    "file_value_names": ["*"]  # All keys are file paths
                }
            ]
        }
    
    def scan(self, selected_areas=None):
        """Scan registry for issues.
        
        Args:
            selected_areas: List of registry areas to scan (None for all)
        
        Returns:
            Dict with scan results
        """
        try:
            self.issues = []  # Reset issues
            scanned_keys = 0
            found_issues = 0
            
            # Filter registry keys based on selected areas
            keys_to_scan = {}
            if selected_areas:
                for area in selected_areas:
                    if area in self.registry_keys:
                        keys_to_scan[area] = self.registry_keys[area]
            else:
                # Scan all areas
                keys_to_scan = self.registry_keys
            
            # Scan each registry key area
            for area, keys in keys_to_scan.items():
                for key_info in keys:
                    area_result = self._scan_registry_area(area, key_info)
                    scanned_keys += area_result['scanned_keys']
                    found_issues += area_result['found_issues']
            
            return {
                'success': True,
                'scanned_keys': scanned_keys,
                'issue_count': found_issues,
                'issues': self.issues,
                'areas_scanned': list(keys_to_scan.keys())
            }
        
        except Exception as e:
            logger.error(f"Error scanning registry: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'scanned_keys': 0,
                'issue_count': 0,
                'issues': []
            }
    
    def _scan_registry_area(self, area, key_info):
        """Scan a specific registry area for issues.
        
        Args:
            area: Name of the area being scanned
            key_info: Dict with key information to scan
        
        Returns:
            Dict with scan results for this area
        """
        scanned_keys = 0
        found_issues = 0
        
        try:
            # Extract key information
            root = key_info['root']
            path = key_info['path']
            check_subkeys = key_info.get('subkeys', False)
            check_values = key_info.get('check_values', False)
            check_file_exists = key_info.get('check_file_exists', False)
            file_value_names = key_info.get('file_value_names', [])
            include_pattern = key_info.get('include_pattern', None)
            exclude_pattern = key_info.get('exclude_pattern', None)
            
            # Recursive function to scan keys
            def scan_key(root, path, depth=0):
                nonlocal scanned_keys, found_issues
                
                # Skip if max depth exceeded
                if depth > 100:  # Prevent too deep recursion
                    return
                
                # Skip keys with exclude pattern or without include pattern
                if path and exclude_pattern and re.search(exclude_pattern, path):
                    return
                if path and include_pattern and not re.search(include_pattern, path):
                    # Special case: if we're checking the root and have an include pattern,
                    # we still want to enumerate subkeys to find matches
                    if depth > 0:
                        return
                
                try:
                    # Open the key
                    key = winreg.OpenKey(root, path, 0, winreg.KEY_READ)
                    scanned_keys += 1
                    
                    # Check values for file existence
                    if check_values:
                        try:
                            # Get number of values
                            num_values = winreg.QueryInfoKey(key)[1]
                            
                            for i in range(num_values):
                                try:
                                    value_name, value_data, value_type = winreg.EnumValue(key, i)
                                    
                                    # Check if we should check this value for file existence
                                    if check_file_exists and (
                                        "*" in file_value_names or 
                                        value_name in file_value_names
                                    ):
                                        # Only check string type values (REG_SZ, REG_EXPAND_SZ)
                                        if value_type in (winreg.REG_SZ, winreg.REG_EXPAND_SZ) and value_data:
                                            # Extract file path from value
                                            file_path = self._extract_file_path(value_data)
                                            
                                            if file_path and not os.path.exists(file_path):
                                                # Found an issue - non-existent file path
                                                issue = {
                                                    'area': area,
                                                    'type': 'missing_file',
                                                    'key': path,
                                                    'value_name': value_name,
                                                    'data': value_data,
                                                    'file_path': file_path,
                                                    'fixable': True
                                                }
                                                self.issues.append(issue)
                                                found_issues += 1
                                
                                except WindowsError:
                                    continue
                        except WindowsError:
                            pass
                    
                    # Check subkeys
                    if check_subkeys:
                        try:
                            # Get number of subkeys
                            num_subkeys = winreg.QueryInfoKey(key)[0]
                            
                            for i in range(num_subkeys):
                                try:
                                    subkey_name = winreg.EnumKey(key, i)
                                    subkey_path = path + '\\' + subkey_name if path else subkey_name
                                    
                                    # Recursively scan subkey
                                    scan_key(root, subkey_path, depth + 1)
                                except WindowsError:
                                    continue
                        except WindowsError:
                            pass
                    
                    # Close the key
                    winreg.CloseKey(key)
                
                except WindowsError:
                    # Registry key couldn't be opened, skip
                    pass
            
            # Start scanning from the root key
            with self.reg_lock:
                scan_key(root, path)
            
            return {
                'scanned_keys': scanned_keys,
                'found_issues': found_issues
            }
        
        except Exception as e:
            logger.error(f"Error scanning registry area {area}: {str(e)}")
            return {
                'scanned_keys': scanned_keys,
                'found_issues': found_issues
            }
    
    def _extract_file_path(self, value_data):
        """Extract a file path from a registry value data string.
        
        Args:
            value_data: String data from registry value
        
        Returns:
            Extracted file path or None
        """
        if not value_data or not isinstance(value_data, str):
            return None
        
        # Expand environment variables
        expanded_data = os.path.expandvars(value_data)
        
        # Extract the file path (handling different formats)
        file_path = None
        
        # Format: Direct path "C:\Program Files\App\program.exe"
        if expanded_data.startswith('"'):
            # Extract path between quotes
            match = re.match(r'"([^"]+)"', expanded_data)
            if match:
                file_path = match.group(1)
        # Format: Command with arguments "C:\Program Files\App\program.exe" -arg
        elif ' ' in expanded_data and expanded_data.startswith('"'):
            # Extract path between quotes
            match = re.match(r'"([^"]+)"', expanded_data)
            if match:
                file_path = match.group(1)
        # Format: Direct path with arguments C:\Program Files\App\program.exe -arg
        elif ' ' in expanded_data:
            # Extract the path part (before the first space)
            file_path = expanded_data.split(' ')[0]
        # Format: Direct path without quotes or arguments
        else:
            file_path = expanded_data
        
        # Verify that the path looks like a file path
        if file_path and (
            re.match(r'^[a-zA-Z]:\\', file_path) or  # Windows absolute path
            file_path.startswith('\\\\')  # UNC path
        ):
            # Check for rundll32 and similar calls that specify DLL function
            if file_path.lower().endswith('.exe') and ',' in expanded_data:
                # Extract DLL path from rundll32 command
                dll_match = re.search(r',["\s]*([^,"\s]+\.dll)["\s]*,', expanded_data)
                if dll_match:
                    return dll_match.group(1)
            
            return file_path
        
        return None
    
    def fix_issues(self, create_backup=True):
        """Fix identified registry issues.
        
        Args:
            create_backup: Whether to create a registry backup before fixing
        
        Returns:
            Dict with fix results
        """
        try:
            # Check if there are issues to fix
            if not self.issues:
                return {
                    'success': True,
                    'fixed_count': 0,
                    'message': 'No issues to fix'
                }
            
            # Create backup if requested
            backup_path = None
            if create_backup:
                backup_path = self.create_registry_backup()
                if not backup_path:
                    return {
                        'success': False,
                        'error': 'Failed to create registry backup',
                        'fixed_count': 0
                    }
            
            # Fix issues
            fixed_count = 0
            failed_fixes = []
            
            with self.reg_lock:
                for issue in self.issues:
                    # Only attempt to fix fixable issues
                    if issue.get('fixable', False):
                        issue_type = issue.get('type')
                        
                        if issue_type == 'missing_file':
                            # Fix missing file issue by removing the registry value
                            try:
                                root = self._get_root_key_from_area(issue['area'])
                                key_path = issue['key']
                                value_name = issue['value_name']
                                
                                key = winreg.OpenKey(root, key_path, 0, winreg.KEY_WRITE)
                                winreg.DeleteValue(key, value_name)
                                winreg.CloseKey(key)
                                
                                fixed_count += 1
                            except WindowsError as e:
                                failed_fixes.append({
                                    'issue': issue,
                                    'error': str(e)
                                })
            
            return {
                'success': True,
                'fixed_count': fixed_count,
                'failed_fixes': failed_fixes,
                'backup_path': backup_path
            }
        
        except Exception as e:
            logger.error(f"Error fixing registry issues: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'fixed_count': 0
            }
    
    def create_registry_backup(self):
        """Create a backup of registry areas being modified.
        
        Returns:
            Path to backup file if successful, None otherwise
        """
        try:
            # Create timestamp for backup file
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = os.path.join(self.temp_dir, f"registry_backup_{timestamp}.reg")
            
            # Use reg.exe to export registry keys
            export_commands = []
            
            # Build export commands for each area with issues
            areas_with_issues = set()
            for issue in self.issues:
                areas_with_issues.add(issue['area'])
            
            for area in areas_with_issues:
                for key_info in self.registry_keys.get(area, []):
                    root = key_info['root']
                    path = key_info['path']
                    
                    # Only export if path is not empty (can't export entire hives)
                    if path:
                        root_name = self._get_root_name(root)
                        if root_name:
                            temp_file = os.path.join(self.temp_dir, f"{area.replace(' ', '_')}_{timestamp}.reg")
                            
                            # Build reg.exe command
                            cmd = [
                                "reg",
                                "export",
                                f"{root_name}\\{path}",
                                temp_file,
                                "/y"
                            ]
                            
                            export_commands.append((cmd, temp_file))
            
            # Execute export commands
            exported_files = []
            for cmd, temp_file in export_commands:
                try:
                    subprocess.run(cmd, capture_output=True, check=True)
                    if os.path.exists(temp_file):
                        exported_files.append(temp_file)
                except subprocess.CalledProcessError:
                    # Continue with next export if one fails
                    continue
            
            # Combine exported files into a single backup
            if exported_files:
                with open(backup_path, 'w', encoding='utf-16') as backup_file:
                    backup_file.write("Windows Registry Editor Version 5.00\n\n")
                    
                    for file in exported_files:
                        if os.path.exists(file):
                            with open(file, 'r', encoding='utf-16') as f:
                                # Skip the first line (header) for all but the first file
                                lines = f.readlines()
                                if lines:
                                    backup_file.writelines(lines[1:])
                            
                            # Remove temp file
                            try:
                                os.remove(file)
                            except:
                                pass
                
                return backup_path
            
            return None
        
        except Exception as e:
            logger.error(f"Error creating registry backup: {str(e)}")
            return None
    
    def restore_registry_backup(self, backup_path):
        """Restore registry from a backup file.
        
        Args:
            backup_path: Path to the backup .reg file
        
        Returns:
            Dict with restore results
        """
        try:
            # Check if backup file exists
            if not os.path.exists(backup_path):
                return {
                    'success': False,
                    'error': 'Backup file not found'
                }
            
            # Use reg.exe to import the backup
            cmd = [
                "reg",
                "import",
                backup_path
            ]
            
            result = subprocess.run(cmd, capture_output=True, check=False)
            
            if result.returncode == 0:
                return {
                    'success': True,
                    'message': 'Registry backup restored successfully'
                }
            else:
                error_msg = result.stderr.decode('utf-8', errors='ignore')
                return {
                    'success': False,
                    'error': f'Failed to restore registry backup: {error_msg}'
                }
        
        except Exception as e:
            logger.error(f"Error restoring registry backup: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def _get_root_key_from_area(self, area):
        """Get registry root key from area name.
        
        Args:
            area: Registry area name
        
        Returns:
            Registry root key (HKEY_*)
        """
        # Find the root key for this area
        for key_info in self.registry_keys.get(area, []):
            return key_info['root']
        
        # Default to HKEY_CURRENT_USER if not found
        return winreg.HKEY_CURRENT_USER
    
    def _get_root_name(self, root):
        """Get the string name of a registry root key.
        
        Args:
            root: Registry root key (HKEY_*)
        
        Returns:
            String name of the root key
        """
        if root == winreg.HKEY_CLASSES_ROOT:
            return "HKEY_CLASSES_ROOT"
        elif root == winreg.HKEY_CURRENT_USER:
            return "HKEY_CURRENT_USER"
        elif root == winreg.HKEY_LOCAL_MACHINE:
            return "HKEY_LOCAL_MACHINE"
        elif root == winreg.HKEY_USERS:
            return "HKEY_USERS"
        elif root == winreg.HKEY_CURRENT_CONFIG:
            return "HKEY_CURRENT_CONFIG"
        else:
            return None
