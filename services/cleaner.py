"""
System cleaner service for the Windows System Optimizer.
This module provides functionality to clean browser caches, temporary files,
and scan/fix registry issues.
"""

import os
import shutil
import subprocess
import winreg
import tempfile
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

class SystemCleaner:
    """Service class for cleaning operations."""
    
    def __init__(self):
        """Initialize the cleaner with required paths."""
        self.user_profile = os.environ.get('USERPROFILE', '')
        self.windows_dir = os.environ.get('WINDIR', 'C:\\Windows')
        self.temp_dir = tempfile.gettempdir()
    
    def clean_chrome_cache(self):
        """Clean Google Chrome cache and temporary files."""
        try:
            # Kill Chrome processes first
            subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe'], 
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            # Chrome cache path
            chrome_cache_path = os.path.join(
                self.user_profile, 
                'AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache'
            )
            
            # Additional Chrome paths to clean
            chrome_paths = [
                os.path.join(self.user_profile, 'AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Google\\Chrome\\User Data\\Default\\Code Cache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Google\\Chrome\\User Data\\Default\\GPUCache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Google\\Chrome\\User Data\\Default\\Media Cache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Google\\Chrome\\User Data\\Default\\Application Cache'),
            ]
            
            # Clean each path
            for path in chrome_paths:
                if os.path.exists(path):
                    self._safe_clean_directory(path)
            
            logger.info("Chrome cache cleaned successfully")
            return True
        
        except Exception as e:
            logger.error(f"Error cleaning Chrome cache: {str(e)}")
            raise
    
    def clean_edge_cache(self):
        """Clean Microsoft Edge cache and temporary files."""
        try:
            # Kill Edge processes first
            subprocess.run(['taskkill', '/F', '/IM', 'msedge.exe'], 
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            # Edge cache paths
            edge_paths = [
                os.path.join(self.user_profile, 'AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Cache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Code Cache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\GPUCache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Media Cache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Application Cache'),
            ]
            
            # Clean each path
            for path in edge_paths:
                if os.path.exists(path):
                    self._safe_clean_directory(path)
            
            logger.info("Edge cache cleaned successfully")
            return True
        
        except Exception as e:
            logger.error(f"Error cleaning Edge cache: {str(e)}")
            raise
    
    def clean_temp_files(self):
        """Clean user temporary files."""
        try:
            # Clean user temp directory
            self._safe_clean_directory(self.temp_dir)
            
            logger.info("Temporary files cleaned successfully")
            return True
        
        except Exception as e:
            logger.error(f"Error cleaning temporary files: {str(e)}")
            raise
    
    def clean_windows_temp(self):
        """Clean Windows temporary files."""
        try:
            # Windows temp directory
            windows_temp = os.path.join(self.windows_dir, 'Temp')
            
            # Clean Windows temp directory
            self._safe_clean_directory(windows_temp)
            
            logger.info("Windows temporary files cleaned successfully")
            return True
        
        except Exception as e:
            logger.error(f"Error cleaning Windows temporary files: {str(e)}")
            raise
    
    def empty_recycle_bin(self):
        """Empty the Windows Recycle Bin."""
        try:
            # Use PowerShell to empty recycle bin
            subprocess.run([
                'powershell.exe', 
                '-Command', 
                'Clear-RecycleBin -Force'
            ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            logger.info("Recycle Bin emptied successfully")
            return True
        
        except Exception as e:
            logger.error(f"Error emptying Recycle Bin: {str(e)}")
            raise
    
    def scan_registry(self, fix=False):
        """
        Scan registry for common issues.
        If fix=True, attempts to fix found issues.
        Returns list of issues found.
        """
        try:
            issues = []
            
            # For demonstration, we'll check a few common registry issues
            # In a real app, this would be much more comprehensive
            
            # Check for software uninstall entries with missing paths
            issues.extend(self._scan_uninstall_entries(fix))
            
            # Check for file extension associations with missing handlers
            issues.extend(self._scan_file_associations(fix))
            
            # Check for startup entries pointing to non-existent files
            issues.extend(self._scan_startup_entries(fix))
            
            logger.info(f"Registry scan completed, found {len(issues)} issues")
            return issues
        
        except Exception as e:
            logger.error(f"Error scanning registry: {str(e)}")
            raise
    
    def _scan_uninstall_entries(self, fix=False):
        """Scan registry uninstall entries for issues."""
        issues = []
        
        try:
            uninstall_key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE, 
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            )
            
            # Count subkeys
            subkey_count = winreg.QueryInfoKey(uninstall_key)[0]
            
            for i in range(subkey_count):
                try:
                    subkey_name = winreg.EnumKey(uninstall_key, i)
                    subkey = winreg.OpenKey(uninstall_key, subkey_name)
                    
                    try:
                        install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                        display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                        
                        if install_location and not os.path.exists(install_location):
                            issue = f"Uninstall entry for '{display_name}' points to non-existent location: {install_location}"
                            issues.append(issue)
                    except (FileNotFoundError, OSError):
                        pass
                    
                    winreg.CloseKey(subkey)
                except (WindowsError, FileNotFoundError):
                    continue
            
            winreg.CloseKey(uninstall_key)
        
        except (WindowsError, FileNotFoundError) as e:
            logger.warning(f"Error accessing uninstall registry entries: {str(e)}")
        
        return issues
    
    def _scan_file_associations(self, fix=False):
        """Scan registry file associations for issues."""
        issues = []
        
        # This would be more complex in a real application
        # For demonstration, we'll just return a placeholder
        
        return issues
    
    def _scan_startup_entries(self, fix=False):
        """Scan registry startup entries for issues."""
        issues = []
        
        try:
            # Check HKCU Run key
            run_key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, 
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
            )
            
            # Get value count
            value_count = winreg.QueryInfoKey(run_key)[1]
            
            for i in range(value_count):
                try:
                    name, value, _ = winreg.EnumValue(run_key, i)
                    
                    # Extract the executable path
                    if value.startswith('"'):
                        exe_path = value.split('"')[1]
                    else:
                        exe_path = value.split(' ')[0]
                    
                    if not os.path.exists(exe_path):
                        issue = f"Startup entry '{name}' points to non-existent file: {exe_path}"
                        issues.append(issue)
                except (WindowsError, IndexError):
                    continue
            
            winreg.CloseKey(run_key)
        
        except (WindowsError, FileNotFoundError) as e:
            logger.warning(f"Error accessing startup registry entries: {str(e)}")
        
        return issues
    
    def _safe_clean_directory(self, directory):
        """
        Safely clean a directory by removing all files but keeping the directory.
        Handles file locks and permission errors gracefully.
        """
        if not os.path.exists(directory):
            return
        
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            try:
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    os.unlink(item_path)
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path, ignore_errors=True)
            except (PermissionError, OSError) as e:
                # Log error but continue with other files
                logger.warning(f"Could not delete {item_path}: {str(e)}")
    
    def get_chrome_cache_size(self):
        """Get the size of Chrome cache in bytes."""
        try:
            chrome_paths = [
                os.path.join(self.user_profile, 'AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Google\\Chrome\\User Data\\Default\\Code Cache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Google\\Chrome\\User Data\\Default\\GPUCache'),
            ]
            
            total_size = 0
            for path in chrome_paths:
                total_size += self._get_directory_size(path)
            
            return total_size
        
        except Exception as e:
            logger.error(f"Error calculating Chrome cache size: {str(e)}")
            return 0
    
    def get_edge_cache_size(self):
        """Get the size of Edge cache in bytes."""
        try:
            edge_paths = [
                os.path.join(self.user_profile, 'AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Cache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Code Cache'),
                os.path.join(self.user_profile, 'AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\GPUCache'),
            ]
            
            total_size = 0
            for path in edge_paths:
                total_size += self._get_directory_size(path)
            
            return total_size
        
        except Exception as e:
            logger.error(f"Error calculating Edge cache size: {str(e)}")
            return 0
    
    def get_temp_files_size(self):
        """Get the size of temporary files in bytes."""
        try:
            return self._get_directory_size(self.temp_dir)
        
        except Exception as e:
            logger.error(f"Error calculating temporary files size: {str(e)}")
            return 0
    
    def get_windows_temp_size(self):
        """Get the size of Windows temporary files in bytes."""
        try:
            windows_temp = os.path.join(self.windows_dir, 'Temp')
            return self._get_directory_size(windows_temp)
        
        except Exception as e:
            logger.error(f"Error calculating Windows temporary files size: {str(e)}")
            return 0
    
    def get_recycle_bin_size(self):
        """Get the size of the Recycle Bin in bytes."""
        try:
            # This is a simplified calculation - in a real app, we would need
            # to check all drives and hidden recycle bin folders
            recycler_folder = os.path.join(os.path.splitdrive(self.user_profile)[0], '$Recycle.Bin')
            return self._get_directory_size(recycler_folder)
        
        except Exception as e:
            logger.error(f"Error calculating Recycle Bin size: {str(e)}")
            return 0
    
    def _get_directory_size(self, path):
        """
        Calculate the total size of a directory in bytes.
        Handles permission errors gracefully.
        """
        total_size = 0
        
        if not os.path.exists(path):
            return 0
        
        for dirpath, dirnames, filenames in os.walk(path, topdown=True, onerror=None):
            for filename in filenames:
                try:
                    file_path = os.path.join(dirpath, filename)
                    if not os.path.islink(file_path):
                        total_size += os.path.getsize(file_path)
                except (OSError, FileNotFoundError):
                    # Skip files that can't be accessed
                    continue
        
        return total_size
