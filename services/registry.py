"""
Registry management service for the Windows System Optimizer.
This module provides functionality to scan, clean, and repair Windows registry issues.
"""

import os
import logging
import winreg
import subprocess
from pathlib import Path
import re

logger = logging.getLogger(__name__)

class RegistryManager:
    """Service class for registry management operations."""
    
    # Registry root keys mapping
    ROOT_KEYS = {
        "HKCR": winreg.HKEY_CLASSES_ROOT,
        "HKCU": winreg.HKEY_CURRENT_USER,
        "HKLM": winreg.HKEY_LOCAL_MACHINE,
        "HKU": winreg.HKEY_USERS,
        "HKCC": winreg.HKEY_CURRENT_CONFIG
    }
    
    def __init__(self):
        """Initialize the registry manager."""
        pass
    
    def scan_registry(self):
        """
        Scan the Windows registry for common issues.
        
        Returns:
            list: List of detected registry issues with details
        """
        issues = []
        
        # Add issues from various scan methods
        issues.extend(self._scan_invalid_file_associations())
        issues.extend(self._scan_orphaned_software())
        issues.extend(self._scan_startup_entries())
        issues.extend(self._scan_shell_extensions())
        issues.extend(self._scan_shared_dlls())
        
        return issues
    
    def fix_registry_issues(self, issues_to_fix):
        """
        Fix selected registry issues.
        
        Args:
            issues_to_fix (list): List of registry issues to fix
            
        Returns:
            dict: Results of fix operations
        """
        results = {
            "fixed": [],
            "failed": []
        }
        
        if not issues_to_fix:
            return results
        
        # Back up registry before making changes
        backup_success = self.backup_registry()
        
        if not backup_success:
            logger.error("Failed to create registry backup, aborting fixes")
            return {
                "fixed": [],
                "failed": ["Registry backup failed - no changes made"]
            }
        
        # Process each issue
        for issue in issues_to_fix:
            try:
                issue_type = issue.get("type")
                
                if issue_type == "invalid_file_association":
                    success = self._fix_file_association(issue)
                elif issue_type == "orphaned_software":
                    success = self._fix_orphaned_software(issue)
                elif issue_type == "startup_entry":
                    success = self._fix_startup_entry(issue)
                elif issue_type == "shell_extension":
                    success = self._fix_shell_extension(issue)
                elif issue_type == "shared_dll":
                    success = self._fix_shared_dll(issue)
                else:
                    success = False
                
                if success:
                    results["fixed"].append(issue)
                else:
                    results["failed"].append(issue)
                
            except Exception as e:
                logger.error(f"Error fixing registry issue: {str(e)}")
                results["failed"].append(issue)
        
        return results
    
    def backup_registry(self, backup_path=None):
        """
        Create a backup of the Windows registry.
        
        Args:
            backup_path (str, optional): Path where to save the backup.
                                        Defaults to user documents.
            
        Returns:
            bool: True if backup succeeded, False otherwise
        """
        try:
            # Default backup path is in user documents
            if not backup_path:
                documents_path = os.path.join(os.path.expanduser("~"), "Documents")
                backup_dir = os.path.join(documents_path, "WinOptimizer", "Backups")
                os.makedirs(backup_dir, exist_ok=True)
                
                # Create timestamp for filename
                import datetime
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_path = os.path.join(backup_dir, f"registry_backup_{timestamp}.reg")
            
            # Use reg.exe to export registry
            command = ["reg", "export", "HKLM", backup_path, "/y"]
            result = subprocess.run(command, 
                                   stdout=subprocess.PIPE, 
                                   stderr=subprocess.PIPE, 
                                   text=True, 
                                   check=True)
            
            logger.info(f"Registry backup created at: {backup_path}")
            return True
            
        except subprocess.CalledProcessError as e:
            logger.error(f"Error backing up registry: {str(e)}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error backing up registry: {str(e)}")
            return False
    
    def restore_registry(self, backup_path):
        """
        Restore Windows registry from a backup file.
        
        Args:
            backup_path (str): Path to the registry backup file
            
        Returns:
            bool: True if restore succeeded, False otherwise
        """
        try:
            if not os.path.exists(backup_path):
                logger.error(f"Registry backup file not found: {backup_path}")
                return False
            
            # Use reg.exe to import registry
            command = ["reg", "import", backup_path]
            result = subprocess.run(command, 
                                   stdout=subprocess.PIPE, 
                                   stderr=subprocess.PIPE, 
                                   text=True, 
                                   check=True)
            
            logger.info(f"Registry restored from: {backup_path}")
            return True
            
        except subprocess.CalledProcessError as e:
            logger.error(f"Error restoring registry: {str(e)}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error restoring registry: {str(e)}")
            return False
    
    def _scan_invalid_file_associations(self):
        """
        Scan for invalid file associations in the registry.
        
        Returns:
            list: List of invalid file association issues
        """
        issues = []
        
        try:
            # Open HKCR key
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "") as root_key:
                # Get number of subkeys
                num_subkeys = winreg.QueryInfoKey(root_key)[0]
                
                # Iterate through file extension subkeys
                for i in range(num_subkeys):
                    ext = winreg.EnumKey(root_key, i)
                    
                    # Check if it's a file extension (starts with .)
                    if ext.startswith("."):
                        try:
                            # Open the extension key
                            with winreg.OpenKey(root_key, ext) as ext_key:
                                try:
                                    # Get default value (file type)
                                    file_type, _ = winreg.QueryValueEx(ext_key, "")
                                    
                                    # If file type exists, check if it's valid
                                    if file_type:
                                        try:
                                            # Try to open the file type key
                                            with winreg.OpenKey(root_key, file_type):
                                                # Key exists, association is valid
                                                pass
                                        except FileNotFoundError:
                                            # File type key doesn't exist
                                            issues.append({
                                                "type": "invalid_file_association",
                                                "extension": ext,
                                                "file_type": file_type,
                                                "description": f"File extension {ext} is associated with non-existent file type {file_type}",
                                                "key_path": f"HKEY_CLASSES_ROOT\\{ext}"
                                            })
                                except (FileNotFoundError, OSError):
                                    # No default value, but not necessarily an issue
                                    pass
                        except (FileNotFoundError, OSError):
                            # Can't open key, skip it
                            continue
        except Exception as e:
            logger.error(f"Error scanning file associations: {str(e)}")
        
        return issues
    
    def _scan_orphaned_software(self):
        """
        Scan for orphaned software registry entries.
        
        Returns:
            list: List of orphaned software issues
        """
        issues = []
        
        try:
            # Check HKLM uninstall entries
            uninstall_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_path) as uninstall_key:
                num_subkeys = winreg.QueryInfoKey(uninstall_key)[0]
                
                for i in range(num_subkeys):
                    subkey_name = winreg.EnumKey(uninstall_key, i)
                    
                    try:
                        with winreg.OpenKey(uninstall_key, subkey_name) as app_key:
                            try:
                                # Get installation location
                                install_location, _ = winreg.QueryValueEx(app_key, "InstallLocation")
                                display_name, _ = winreg.QueryValueEx(app_key, "DisplayName")
                                
                                # Check if location exists
                                if install_location and len(install_location) > 0:
                                    if not os.path.exists(install_location):
                                        issues.append({
                                            "type": "orphaned_software",
                                            "name": display_name,
                                            "location": install_location,
                                            "description": f"Software {display_name} has invalid installation path: {install_location}",
                                            "key_path": f"HKEY_LOCAL_MACHINE\\{uninstall_path}\\{subkey_name}"
                                        })
                            except (FileNotFoundError, OSError):
                                # Missing values, not necessarily an issue
                                pass
                    except (FileNotFoundError, OSError):
                        continue
            
            # Also check HKCU uninstall entries
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, uninstall_path) as uninstall_key:
                    num_subkeys = winreg.QueryInfoKey(uninstall_key)[0]
                    
                    for i in range(num_subkeys):
                        subkey_name = winreg.EnumKey(uninstall_key, i)
                        
                        try:
                            with winreg.OpenKey(uninstall_key, subkey_name) as app_key:
                                try:
                                    # Get installation location
                                    install_location, _ = winreg.QueryValueEx(app_key, "InstallLocation")
                                    display_name, _ = winreg.QueryValueEx(app_key, "DisplayName")
                                    
                                    # Check if location exists
                                    if install_location and len(install_location) > 0:
                                        if not os.path.exists(install_location):
                                            issues.append({
                                                "type": "orphaned_software",
                                                "name": display_name,
                                                "location": install_location,
                                                "description": f"Software {display_name} has invalid installation path: {install_location}",
                                                "key_path": f"HKEY_CURRENT_USER\\{uninstall_path}\\{subkey_name}"
                                            })
                                except (FileNotFoundError, OSError):
                                    # Missing values, not necessarily an issue
                                    pass
                        except (FileNotFoundError, OSError):
                            continue
            except (FileNotFoundError, OSError):
                # HKCU uninstall key might not exist
                pass
                
        except Exception as e:
            logger.error(f"Error scanning orphaned software: {str(e)}")
        
        return issues
    
    def _scan_startup_entries(self):
        """
        Scan for invalid startup entries in the registry.
        
        Returns:
            list: List of invalid startup entry issues
        """
        issues = []
        
        try:
            # Check HKLM Run key
            run_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, run_path) as run_key:
                num_values = winreg.QueryInfoKey(run_key)[1]
                
                for i in range(num_values):
                    name, value, _ = winreg.EnumValue(run_key, i)
                    
                    # Extract executable path from command
                    exe_path = self._extract_exe_path(value)
                    
                    # Check if file exists
                    if exe_path and not os.path.exists(exe_path):
                        issues.append({
                            "type": "startup_entry",
                            "name": name,
                            "command": value,
                            "executable": exe_path,
                            "description": f"Startup entry '{name}' points to non-existent file: {exe_path}",
                            "key_path": f"HKEY_LOCAL_MACHINE\\{run_path}"
                        })
            
            # Check HKCU Run key
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_path) as run_key:
                num_values = winreg.QueryInfoKey(run_key)[1]
                
                for i in range(num_values):
                    name, value, _ = winreg.EnumValue(run_key, i)
                    
                    # Extract executable path from command
                    exe_path = self._extract_exe_path(value)
                    
                    # Check if file exists
                    if exe_path and not os.path.exists(exe_path):
                        issues.append({
                            "type": "startup_entry",
                            "name": name,
                            "command": value,
                            "executable": exe_path,
                            "description": f"Startup entry '{name}' points to non-existent file: {exe_path}",
                            "key_path": f"HKEY_CURRENT_USER\\{run_path}"
                        })
                
        except Exception as e:
            logger.error(f"Error scanning startup entries: {str(e)}")
        
        return issues
    
    def _scan_shell_extensions(self):
        """
        Scan for invalid shell extensions in the registry.
        
        Returns:
            list: List of invalid shell extension issues
        """
        issues = []
        
        try:
            # Path to shell extensions
            ext_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved"
            
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, ext_path) as ext_key:
                num_values = winreg.QueryInfoKey(ext_key)[1]
                
                for i in range(num_values):
                    clsid, name, _ = winreg.EnumValue(ext_key, i)
                    
                    # Check if CLSID exists
                    try:
                        clsid_path = f"CLSID\\{clsid}\\InProcServer32"
                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, clsid_path) as clsid_key:
                            # Get DLL path
                            dll_path, _ = winreg.QueryValueEx(clsid_key, "")
                            
                            # Check if DLL exists
                            if not os.path.exists(dll_path):
                                issues.append({
                                    "type": "shell_extension",
                                    "name": name,
                                    "clsid": clsid,
                                    "dll_path": dll_path,
                                    "description": f"Shell extension '{name}' ({clsid}) points to non-existent DLL: {dll_path}",
                                    "key_path": f"HKEY_LOCAL_MACHINE\\{ext_path}"
                                })
                    except (FileNotFoundError, OSError):
                        # CLSID doesn't exist or has no DLL path
                        issues.append({
                            "type": "shell_extension",
                            "name": name,
                            "clsid": clsid,
                            "dll_path": None,
                            "description": f"Shell extension '{name}' ({clsid}) has missing or invalid CLSID",
                            "key_path": f"HKEY_LOCAL_MACHINE\\{ext_path}"
                        })
                
        except Exception as e:
            logger.error(f"Error scanning shell extensions: {str(e)}")
        
        return issues
    
    def _scan_shared_dlls(self):
        """
        Scan for invalid shared DLLs in the registry.
        
        Returns:
            list: List of invalid shared DLL issues
        """
        issues = []
        
        try:
            # Path to shared DLLs
            dll_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs"
            
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, dll_path) as dll_key:
                num_values = winreg.QueryInfoKey(dll_key)[1]
                
                for i in range(num_values):
                    file_path, count, _ = winreg.EnumValue(dll_key, i)
                    
                    # Check if file exists
                    if not os.path.exists(file_path):
                        issues.append({
                            "type": "shared_dll",
                            "path": file_path,
                            "count": count,
                            "description": f"Shared DLL reference to non-existent file: {file_path} (ref count: {count})",
                            "key_path": f"HKEY_LOCAL_MACHINE\\{dll_path}"
                        })
                
        except Exception as e:
            logger.error(f"Error scanning shared DLLs: {str(e)}")
        
        return issues
    
    def _fix_file_association(self, issue):
        """
        Fix an invalid file association.
        
        Args:
            issue (dict): File association issue to fix
            
        Returns:
            bool: True if fixed successfully, False otherwise
        """
        try:
            # We're simply removing the invalid association
            ext = issue.get("extension")
            
            # Delete the default value for the extension
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, ext, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, "")
            
            logger.info(f"Fixed invalid file association for {ext}")
            return True
        
        except Exception as e:
            logger.error(f"Error fixing file association: {str(e)}")
            return False
    
    def _fix_orphaned_software(self, issue):
        """
        Fix an orphaned software entry.
        
        Args:
            issue (dict): Orphaned software issue to fix
            
        Returns:
            bool: True if fixed successfully, False otherwise
        """
        try:
            # We're removing the registry key for the orphaned software
            key_path = issue.get("key_path")
            
            if not key_path:
                return False
            
            # Parse root key and subpath
            parts = key_path.split('\\', 1)
            if len(parts) != 2:
                return False
            
            root_name, subpath = parts
            
            if root_name not in self.ROOT_KEYS:
                return False
            
            root_key = self.ROOT_KEYS[root_name]
            
            # Delete the subkey
            winreg.DeleteKey(root_key, subpath)
            
            logger.info(f"Removed orphaned software entry: {issue.get('name')}")
            return True
        
        except Exception as e:
            logger.error(f"Error fixing orphaned software entry: {str(e)}")
            return False
    
    def _fix_startup_entry(self, issue):
        """
        Fix an invalid startup entry.
        
        Args:
            issue (dict): Startup entry issue to fix
            
        Returns:
            bool: True if fixed successfully, False otherwise
        """
        try:
            # We're removing the invalid startup entry
            key_path = issue.get("key_path")
            name = issue.get("name")
            
            if not key_path or not name:
                return False
            
            # Parse root key and subpath
            parts = key_path.split('\\', 1)
            if len(parts) != 2:
                return False
            
            root_name, subpath = parts
            
            if root_name not in self.ROOT_KEYS:
                return False
            
            root_key = self.ROOT_KEYS[root_name]
            
            # Delete the value
            with winreg.OpenKey(root_key, subpath, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, name)
            
            logger.info(f"Removed invalid startup entry: {name}")
            return True
        
        except Exception as e:
            logger.error(f"Error fixing startup entry: {str(e)}")
            return False
    
    def _fix_shell_extension(self, issue):
        """
        Fix an invalid shell extension.
        
        Args:
            issue (dict): Shell extension issue to fix
            
        Returns:
            bool: True if fixed successfully, False otherwise
        """
        try:
            # We're removing the invalid shell extension
            key_path = issue.get("key_path")
            clsid = issue.get("clsid")
            
            if not key_path or not clsid:
                return False
            
            # Parse root key and subpath
            parts = key_path.split('\\', 1)
            if len(parts) != 2:
                return False
            
            root_name, subpath = parts
            
            if root_name not in self.ROOT_KEYS:
                return False
            
            root_key = self.ROOT_KEYS[root_name]
            
            # Delete the value
            with winreg.OpenKey(root_key, subpath, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, clsid)
            
            logger.info(f"Removed invalid shell extension: {issue.get('name')}")
            return True
        
        except Exception as e:
            logger.error(f"Error fixing shell extension: {str(e)}")
            return False
    
    def _fix_shared_dll(self, issue):
        """
        Fix an invalid shared DLL entry.
        
        Args:
            issue (dict): Shared DLL issue to fix
            
        Returns:
            bool: True if fixed successfully, False otherwise
        """
        try:
            # We're removing the invalid shared DLL reference
            key_path = issue.get("key_path")
            file_path = issue.get("path")
            
            if not key_path or not file_path:
                return False
            
            # Parse root key and subpath
            parts = key_path.split('\\', 1)
            if len(parts) != 2:
                return False
            
            root_name, subpath = parts
            
            if root_name not in self.ROOT_KEYS:
                return False
            
            root_key = self.ROOT_KEYS[root_name]
            
            # Delete the value
            with winreg.OpenKey(root_key, subpath, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, file_path)
            
            logger.info(f"Removed invalid shared DLL reference: {file_path}")
            return True
        
        except Exception as e:
            logger.error(f"Error fixing shared DLL entry: {str(e)}")
            return False
    
    def _extract_exe_path(self, command):
        """
        Extract executable path from a command string.
        
        Args:
            command (str): Command string
            
        Returns:
            str: Executable path or None if not found
        """
        if not command:
            return None
        
        # Handle quoted path
        if command.startswith('"'):
            # Find the closing quote
            end_quote = command.find('"', 1)
            if end_quote > 0:
                return command[1:end_quote]
        
        # No quotes, just split by space and take first part
        return command.split(' ')[0]
