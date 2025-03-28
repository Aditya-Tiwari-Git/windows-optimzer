"""
System cleaner utility for cleaning temporary files and caches.
"""

import os
import sys
import re
import shutil
import logging
import winreg
import subprocess
from pathlib import Path
import json
import tempfile
import stat
import ctypes
from threading import Lock

logger = logging.getLogger(__name__)


class SystemCleaner:
    """Utility for cleaning temporary files and caches on Windows."""
    
    def __init__(self):
        """Initialize the system cleaner with cleanup targets."""
        self.temp_dirs = []
        self.browser_caches = {}
        self.find_cleanup_targets()
        
        # Thread-safe access for file operations
        self.file_lock = Lock()
        
        # Current user profile
        self.user_profile = os.path.expanduser("~")
    
    def find_cleanup_targets(self):
        """Identify temp directories and browser caches for cleanup."""
        try:
            # System temp directories
            self.temp_dirs = [
                os.environ.get('TEMP', os.path.join(os.environ.get('USERPROFILE', ''), 'AppData\\Local\\Temp')),
                os.path.join(os.environ.get('SYSTEMROOT', 'C:\\Windows'), 'Temp'),
                os.path.join(os.environ.get('SYSTEMROOT', 'C:\\Windows'), 'Prefetch'),
                os.path.join(os.environ.get('SYSTEMDRIVE', 'C:'), 'Windows\\SoftwareDistribution\\Download')
            ]
            
            # Check if directories exist
            self.temp_dirs = [d for d in self.temp_dirs if os.path.exists(d)]
            
            # Browser caches - Chrome
            chrome_cache = os.path.join(
                os.environ.get('LOCALAPPDATA', ''), 
                'Google\\Chrome\\User Data\\Default\\Cache'
            )
            
            chrome_cookies = os.path.join(
                os.environ.get('LOCALAPPDATA', ''), 
                'Google\\Chrome\\User Data\\Default\\Cookies'
            )
            
            # Browser caches - Edge
            edge_cache = os.path.join(
                os.environ.get('LOCALAPPDATA', ''), 
                'Microsoft\\Edge\\User Data\\Default\\Cache'
            )
            
            edge_cookies = os.path.join(
                os.environ.get('LOCALAPPDATA', ''), 
                'Microsoft\\Edge\\User Data\\Default\\Cookies'
            )
            
            # Store browser cache paths if they exist
            if os.path.exists(chrome_cache):
                self.browser_caches['Chrome Cache'] = chrome_cache
            
            if os.path.exists(chrome_cookies):
                self.browser_caches['Chrome Cookies'] = chrome_cookies
            
            if os.path.exists(edge_cache):
                self.browser_caches['Edge Cache'] = edge_cache
            
            if os.path.exists(edge_cookies):
                self.browser_caches['Edge Cookies'] = edge_cookies
            
            logger.info(f"Found {len(self.temp_dirs)} temp directories and {len(self.browser_caches)} browser caches")
        
        except Exception as e:
            logger.error(f"Error finding cleanup targets: {str(e)}")
    
    def get_disk_info(self):
        """Get disk information for C: drive.
        
        Returns:
            Dict with disk information
        """
        try:
            total, used, free = shutil.disk_usage("C:\\")
            
            return {
                "total": total,
                "used": used,
                "free": free,
                "total_gb": total / (1024**3),
                "used_gb": used / (1024**3),
                "free_gb": free / (1024**3),
                "percent": (used / total) * 100
            }
        except Exception as e:
            logger.error(f"Error getting disk info: {str(e)}")
            return {
                "total": 0,
                "used": 0,
                "free": 0,
                "total_gb": 0,
                "used_gb": 0,
                "free_gb": 0,
                "percent": 0
            }
    
    def calculate_directory_size(self, path):
        """Calculate the size of a directory recursively.
        
        Args:
            path: Directory path
        
        Returns:
            Directory size in bytes
        """
        try:
            with self.file_lock:
                total_size = 0
                
                # Use os.scandir for better performance
                for entry in os.scandir(path):
                    if entry.is_file(follow_symlinks=False):
                        try:
                            total_size += entry.stat().st_size
                        except (FileNotFoundError, PermissionError):
                            continue
                    elif entry.is_dir(follow_symlinks=False):
                        try:
                            total_size += self.calculate_directory_size(entry.path)
                        except (FileNotFoundError, PermissionError):
                            continue
                
                return total_size
        except (FileNotFoundError, PermissionError, OSError) as e:
            logger.debug(f"Error calculating size of {path}: {str(e)}")
            return 0
    
    def get_recycle_bin_size(self):
        """Calculate the size of the recycle bin.
        
        Returns:
            Size of recycle bin in bytes
        """
        try:
            # Use PowerShell to query recycle bin size
            result = subprocess.run(
                ["powershell", "-Command", 
                 "$shell = New-Object -ComObject Shell.Application;" +
                 "$recycle = $shell.NameSpace(10);" +
                 "$items = $recycle.Items();" +
                 "$size = 0;" +
                 "foreach($item in $items) { $size += $item.Size };" +
                 "Write-Output $size"],
                capture_output=True, text=True, check=True
            )
            
            size_str = result.stdout.strip()
            if size_str and size_str.isdigit():
                return int(size_str)
            
            # Fallback if PowerShell method fails
            return 0
        except Exception as e:
            logger.error(f"Error getting recycle bin size: {str(e)}")
            return 0
    
    def get_thumbnails_cache_size(self):
        """Calculate the size of the Windows thumbnails cache.
        
        Returns:
            Size of thumbnails cache in bytes
        """
        try:
            thumbs_db_path = os.path.join(
                os.environ.get('LOCALAPPDATA', ''),
                'Microsoft\\Windows\\Explorer'
            )
            
            if os.path.exists(thumbs_db_path):
                thumb_size = 0
                
                # Check for thumbcache_*.db files
                for root, _, files in os.walk(thumbs_db_path):
                    for file in files:
                        if file.startswith("thumbcache_") and file.endswith(".db"):
                            try:
                                file_path = os.path.join(root, file)
                                thumb_size += os.path.getsize(file_path)
                            except (FileNotFoundError, PermissionError):
                                continue
                
                return thumb_size
            
            return 0
        except Exception as e:
            logger.error(f"Error getting thumbnails cache size: {str(e)}")
            return 0
    
    def get_downloads_size(self, older_than_days=30):
        """Calculate the size of old files in the Downloads folder.
        
        Args:
            older_than_days: Only include files older than this many days
        
        Returns:
            Size of old downloads in bytes
        """
        try:
            downloads_path = os.path.join(self.user_profile, 'Downloads')
            
            if not os.path.exists(downloads_path):
                return 0
            
            total_size = 0
            cutoff_time = time.time() - (older_than_days * 24 * 3600)
            
            for entry in os.scandir(downloads_path):
                try:
                    if entry.is_file(follow_symlinks=False):
                        stat_result = entry.stat()
                        # Check if file is older than the cutoff
                        if stat_result.st_mtime < cutoff_time:
                            total_size += stat_result.st_size
                except (FileNotFoundError, PermissionError):
                    continue
            
            return total_size
        except Exception as e:
            logger.error(f"Error calculating downloads size: {str(e)}")
            return 0
    
    def analyze(self, selected_options):
        """Analyze system for cleanable files.
        
        Args:
            selected_options: List of cleanup options to analyze
        
        Returns:
            Dict mapping cleanup options to size in bytes
        """
        results = {}
        
        try:
            # Process selected options
            for option in selected_options:
                if option == "Temporary Files":
                    size = sum(self.calculate_directory_size(d) for d in self.temp_dirs)
                    results[option] = size
                
                elif option == "Windows Temp":
                    windows_temp = os.path.join(os.environ.get('SYSTEMROOT', 'C:\\Windows'), 'Temp')
                    size = self.calculate_directory_size(windows_temp)
                    results[option] = size
                
                elif option == "Recycle Bin":
                    size = self.get_recycle_bin_size()
                    results[option] = size
                
                elif option == "Chrome Cache":
                    size = 0
                    if 'Chrome Cache' in self.browser_caches:
                        size += self.calculate_directory_size(self.browser_caches['Chrome Cache'])
                    if 'Chrome Cookies' in self.browser_caches:
                        size += self.calculate_directory_size(self.browser_caches['Chrome Cookies'])
                    results[option] = size
                
                elif option == "Edge Cache":
                    size = 0
                    if 'Edge Cache' in self.browser_caches:
                        size += self.calculate_directory_size(self.browser_caches['Edge Cache'])
                    if 'Edge Cookies' in self.browser_caches:
                        size += self.calculate_directory_size(self.browser_caches['Edge Cookies'])
                    results[option] = size
                
                elif option == "Downloads Cleanup":
                    size = self.get_downloads_size(older_than_days=30)
                    results[option] = size
                
                elif option == "Thumbnails Cache":
                    size = self.get_thumbnails_cache_size()
                    results[option] = size
            
            return results
        
        except Exception as e:
            logger.error(f"Error analyzing system: {str(e)}")
            return results
    
    def clean(self, selected_options):
        """Clean selected items from the system.
        
        Args:
            selected_options: List of cleanup options to perform
        
        Returns:
            Dict with cleanup results
        """
        results = {
            'success': True,
            'cleaned_bytes': 0,
            'errors': []
        }
        
        try:
            # Process selected options
            for option in selected_options:
                try:
                    if option == "Temporary Files":
                        for temp_dir in self.temp_dirs:
                            cleaned = self._clean_directory(temp_dir)
                            results['cleaned_bytes'] += cleaned
                    
                    elif option == "Windows Temp":
                        windows_temp = os.path.join(os.environ.get('SYSTEMROOT', 'C:\\Windows'), 'Temp')
                        cleaned = self._clean_directory(windows_temp)
                        results['cleaned_bytes'] += cleaned
                    
                    elif option == "Recycle Bin":
                        cleaned = self._empty_recycle_bin()
                        results['cleaned_bytes'] += cleaned
                    
                    elif option == "Chrome Cache":
                        if 'Chrome Cache' in self.browser_caches:
                            cleaned = self._clean_directory(self.browser_caches['Chrome Cache'])
                            results['cleaned_bytes'] += cleaned
                        
                        if 'Chrome Cookies' in self.browser_caches:
                            # For cookies, we should use Chrome's API or clear specific files
                            # For simplicity, we'll just rename the cookies file, which forces Chrome to create a new one
                            cookie_path = self.browser_caches['Chrome Cookies']
                            if os.path.exists(cookie_path):
                                with self.file_lock:
                                    try:
                                        size = os.path.getsize(cookie_path)
                                        # Rename instead of delete to avoid corruption
                                        os.rename(cookie_path, cookie_path + ".old")
                                        results['cleaned_bytes'] += size
                                    except (FileNotFoundError, PermissionError, OSError):
                                        pass
                    
                    elif option == "Edge Cache":
                        if 'Edge Cache' in self.browser_caches:
                            cleaned = self._clean_directory(self.browser_caches['Edge Cache'])
                            results['cleaned_bytes'] += cleaned
                        
                        if 'Edge Cookies' in self.browser_caches:
                            # Handle cookies similarly to Chrome
                            cookie_path = self.browser_caches['Edge Cookies']
                            if os.path.exists(cookie_path):
                                with self.file_lock:
                                    try:
                                        size = os.path.getsize(cookie_path)
                                        os.rename(cookie_path, cookie_path + ".old")
                                        results['cleaned_bytes'] += size
                                    except (FileNotFoundError, PermissionError, OSError):
                                        pass
                    
                    elif option == "Downloads Cleanup":
                        downloads_path = os.path.join(self.user_profile, 'Downloads')
                        cutoff_time = time.time() - (30 * 24 * 3600)  # 30 days
                        
                        if os.path.exists(downloads_path):
                            for entry in os.scandir(downloads_path):
                                try:
                                    if entry.is_file(follow_symlinks=False):
                                        stat_result = entry.stat()
                                        # Check if file is older than the cutoff
                                        if stat_result.st_mtime < cutoff_time:
                                            with self.file_lock:
                                                try:
                                                    size = stat_result.st_size
                                                    os.remove(entry.path)
                                                    results['cleaned_bytes'] += size
                                                except (FileNotFoundError, PermissionError, OSError):
                                                    continue
                                except (FileNotFoundError, PermissionError):
                                    continue
                    
                    elif option == "Thumbnails Cache":
                        thumbs_db_path = os.path.join(
                            os.environ.get('LOCALAPPDATA', ''),
                            'Microsoft\\Windows\\Explorer'
                        )
                        
                        if os.path.exists(thumbs_db_path):
                            # Find and remove thumbnail cache files
                            for root, _, files in os.walk(thumbs_db_path):
                                for file in files:
                                    if file.startswith("thumbcache_") and file.endswith(".db"):
                                        try:
                                            file_path = os.path.join(root, file)
                                            with self.file_lock:
                                                try:
                                                    size = os.path.getsize(file_path)
                                                    os.remove(file_path)
                                                    results['cleaned_bytes'] += size
                                                except (FileNotFoundError, PermissionError, OSError):
                                                    continue
                                        except (FileNotFoundError, PermissionError):
                                            continue
                
                except Exception as e:
                    logger.error(f"Error cleaning {option}: {str(e)}")
                    results['errors'].append(f"Error cleaning {option}: {str(e)}")
            
            # Clear the icon cache by restarting Windows Explorer
            self._clear_icon_cache()
            
            return results
        
        except Exception as e:
            logger.error(f"Error during cleaning operation: {str(e)}")
            results['success'] = False
            results['errors'].append(str(e))
            return results
    
    def _clean_directory(self, directory):
        """Clean a directory by removing all files and subdirectories.
        
        Args:
            directory: Path to directory to clean
        
        Returns:
            Number of bytes cleaned
        """
        if not os.path.exists(directory):
            return 0
        
        cleaned_bytes = 0
        
        with self.file_lock:
            # First calculate size
            try:
                cleaned_bytes = self.calculate_directory_size(directory)
            except Exception:
                cleaned_bytes = 0
            
            # Then remove files
            for root, dirs, files in os.walk(directory, topdown=False):
                # Remove files
                for file in files:
                    try:
                        file_path = os.path.join(root, file)
                        # Change file attributes if needed
                        if os.path.exists(file_path):
                            try:
                                # Clear read-only attribute if present
                                attrs = ctypes.windll.kernel32.GetFileAttributesW(file_path)
                                if attrs & 1:  # FILE_ATTRIBUTE_READONLY
                                    ctypes.windll.kernel32.SetFileAttributesW(file_path, attrs & ~1)
                            except:
                                pass
                            
                            # Remove file
                            os.remove(file_path)
                    except (FileNotFoundError, PermissionError, OSError):
                        continue
                
                # Remove empty directories
                for dir_name in dirs:
                    try:
                        dir_path = os.path.join(root, dir_name)
                        if os.path.exists(dir_path):
                            # Only attempt to remove if it appears empty
                            if not os.listdir(dir_path):
                                os.rmdir(dir_path)
                    except (FileNotFoundError, PermissionError, OSError):
                        continue
        
        return cleaned_bytes
    
    def _empty_recycle_bin(self):
        """Empty the recycle bin.
        
        Returns:
            Number of bytes cleaned
        """
        try:
            # First get the size
            size = self.get_recycle_bin_size()
            
            # Use PowerShell to empty recycle bin
            subprocess.run(
                ["powershell", "-Command", "Clear-RecycleBin -Force"],
                capture_output=True, check=True
            )
            
            return size
        except Exception as e:
            logger.error(f"Error emptying recycle bin: {str(e)}")
            return 0
    
    def _clear_icon_cache(self):
        """Clear the Windows icon cache."""
        try:
            # This requires restarting Windows Explorer
            # First stop explorer
            subprocess.run(
                ["taskkill", "/f", "/im", "explorer.exe"],
                capture_output=True, check=False
            )
            
            # Delete icon cache files
            icon_cache_path = os.path.join(
                os.environ.get('LOCALAPPDATA', ''),
                'Microsoft\\Windows\\Explorer'
            )
            
            if os.path.exists(icon_cache_path):
                with self.file_lock:
                    for file in os.listdir(icon_cache_path):
                        if file.startswith("iconcache") or file.startswith("thumbcache"):
                            try:
                                file_path = os.path.join(icon_cache_path, file)
                                os.remove(file_path)
                            except (FileNotFoundError, PermissionError, OSError):
                                continue
            
            # Restart explorer
            subprocess.Popen("explorer.exe")
            
            return True
        except Exception as e:
            logger.error(f"Error clearing icon cache: {str(e)}")
            return False


# Import modules
import time
