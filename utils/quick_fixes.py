"""
Quick fix utilities for Microsoft Teams and Outlook applications.
"""

import os
import sys
import re
import shutil
import logging
import winreg
import subprocess
import ctypes
import time
import psutil
from pathlib import Path
from threading import Lock

logger = logging.getLogger(__name__)


class MSTeamsFixer:
    """Utility for diagnosing and fixing Microsoft Teams issues."""
    
    def __init__(self):
        """Initialize Teams fixer with paths to known Teams directories."""
        self.user_profile = os.path.expanduser("~")
        
        # Teams application paths
        self.teams_paths = {
            'app_data': os.path.join(self.user_profile, 'AppData', 'Local', 'Microsoft', 'Teams'),
            'roaming_data': os.path.join(self.user_profile, 'AppData', 'Roaming', 'Microsoft', 'Teams'),
            'cache': os.path.join(self.user_profile, 'AppData', 'Local', 'Microsoft', 'Teams', 'Cache'),
            'gpu_cache': os.path.join(self.user_profile, 'AppData', 'Local', 'Microsoft', 'Teams', 'GPUCache'),
            'blob_storage': os.path.join(self.user_profile, 'AppData', 'Local', 'Microsoft', 'Teams', 'blob_storage'),
            'databases': os.path.join(self.user_profile, 'AppData', 'Local', 'Microsoft', 'Teams', 'databases'),
            'local_storage': os.path.join(self.user_profile, 'AppData', 'Local', 'Microsoft', 'Teams', 'Local Storage'),
            'config': os.path.join(self.user_profile, 'AppData', 'Roaming', 'Microsoft', 'Teams', 'desktop-config.json')
        }
        
        # Thread-safe access for file operations
        self.file_lock = Lock()
    
    def check_status(self):
        """Check the status of Microsoft Teams.
        
        Returns:
            Dict with status information
        """
        try:
            status = "Healthy"
            issues = []
            
            # Check if Teams processes are running
            teams_processes = []
            for proc in psutil.process_iter(['name']):
                try:
                    if 'teams' in proc.info['name'].lower():
                        teams_processes.append(proc.info['name'])
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            # Check if Teams app data exists
            app_exists = os.path.exists(self.teams_paths['app_data'])
            
            # Check for common issues
            
            # Cache size issue
            if os.path.exists(self.teams_paths['cache']):
                cache_size = self.get_directory_size(self.teams_paths['cache'])
                # If cache is larger than 1GB, it might cause performance issues
                if cache_size > 1024 * 1024 * 1024:
                    issues.append("Large cache size")
                    status = "Issues Found"
            
            # Configuration file issues
            if os.path.exists(self.teams_paths['config']):
                try:
                    # Check if config file is corrupted
                    with open(self.teams_paths['config'], 'r') as f:
                        config_data = f.read().strip()
                        # Check if the file is valid JSON
                        if not config_data.startswith('{') or not config_data.endswith('}'):
                            issues.append("Corrupted configuration file")
                            status = "Issues Found"
                except Exception:
                    issues.append("Inaccessible configuration file")
                    status = "Issues Found"
            
            # GPU acceleration issues - often causes crashes on some systems
            if os.path.exists(self.teams_paths['gpu_cache']):
                # Check for graphics-related crash files
                for root, _, files in os.walk(self.teams_paths['app_data']):
                    for file in files:
                        if 'crash' in file.lower() and 'gpu' in file.lower():
                            issues.append("GPU-related crashes detected")
                            status = "Issues Found"
                            break
                    if "Issues Found" == status:
                        break
            
            # Check for multiple Teams processes (can indicate stuck processes)
            teams_exe_count = teams_processes.count('Teams.exe')
            if teams_exe_count > 1:
                issues.append("Multiple Teams processes running")
                status = "Issues Found"
            
            return {
                'status': status,
                'issues': issues,
                'processes': teams_processes,
                'app_exists': app_exists
            }
        
        except Exception as e:
            logger.error(f"Error checking Teams status: {str(e)}")
            return {
                'status': "Error",
                'issues': [str(e)],
                'processes': [],
                'app_exists': False
            }
    
    def analyze(self):
        """Analyze Microsoft Teams for issues and recommend fixes.
        
        Returns:
            Dict with analysis results and recommended fixes
        """
        try:
            # Check current status
            status = self.check_status()
            issues = status.get('issues', [])
            
            # Additional analysis
            found_issues = len(issues) > 0
            recommended_fixes = []
            
            # Check cache size
            if os.path.exists(self.teams_paths['cache']):
                cache_size = self.get_directory_size(self.teams_paths['cache'])
                if cache_size > 500 * 1024 * 1024:  # 500 MB
                    found_issues = True
                    recommended_fixes.append("Clear Cache")
            
            # Check for process issues
            if "Multiple Teams processes running" in issues:
                found_issues = True
                recommended_fixes.append("End Background Processes")
            
            # Check for config issues
            if "Corrupted configuration file" in issues or "Inaccessible configuration file" in issues:
                found_issues = True
                recommended_fixes.append("Reset Configuration")
            
            # Check for browser storage issues
            if os.path.exists(self.teams_paths['local_storage']):
                for root, _, files in os.walk(self.teams_paths['local_storage']):
                    if any(file.endswith('.log') for file in files):
                        found_issues = True
                        recommended_fixes.append("Clear Web Storage")
                        break
            
            # Check for network issues
            if self.check_teams_network_issues():
                found_issues = True
                recommended_fixes.append("Reset Network Settings")
            
            # Check for GPU issues
            if "GPU-related crashes detected" in issues:
                found_issues = True
                recommended_fixes.append("Disable GPU Hardware Acceleration")
            
            # If no specific issues found but status indicates problems, recommend basic fixes
            if found_issues and not recommended_fixes:
                recommended_fixes = ["Clear Cache", "End Background Processes"]
            
            return {
                'found_issues': found_issues,
                'issues': issues,
                'recommended_fixes': recommended_fixes
            }
        
        except Exception as e:
            logger.error(f"Error analyzing Teams: {str(e)}")
            return {
                'found_issues': True,
                'issues': [str(e)],
                'recommended_fixes': ["Clear Cache", "End Background Processes", "Reset Configuration"]
            }
    
    def check_teams_network_issues(self):
        """Check for Teams network connectivity issues.
        
        Returns:
            True if network issues detected, False otherwise
        """
        try:
            # Teams endpoints to check
            endpoints = [
                "teams.microsoft.com",
                "presence.teams.microsoft.com",
                "api.spaces.skype.com"
            ]
            
            issues_found = False
            
            # Test connectivity to Teams endpoints
            for endpoint in endpoints:
                cmd = ["ping", "-n", "1", "-w", "1000", endpoint]
                
                try:
                    result = subprocess.run(cmd, capture_output=True, text=True, check=False)
                    if result.returncode != 0 or "timed out" in result.stdout.lower():
                        issues_found = True
                        break
                except Exception:
                    issues_found = True
                    break
            
            return issues_found
        
        except Exception as e:
            logger.error(f"Error checking Teams network issues: {str(e)}")
            return False
    
    def apply_fixes(self, selected_fixes):
        """Apply selected fixes to Microsoft Teams.
        
        Args:
            selected_fixes: List of fix names to apply
        
        Returns:
            Dict with fix results
        """
        try:
            results = {
                'success': True,
                'applied_fixes': [],
                'failed_fixes': []
            }
            
            # Kill Teams processes first if needed
            kill_teams = "End Background Processes" in selected_fixes
            if kill_teams:
                self.kill_teams_processes()
                results['applied_fixes'].append("End Background Processes")
            
            # Apply each selected fix
            for fix in selected_fixes:
                if fix == "End Background Processes":
                    # Already handled above
                    continue
                
                elif fix == "Clear Cache":
                    success = self.clear_teams_cache()
                    if success:
                        results['applied_fixes'].append("Clear Cache")
                    else:
                        results['failed_fixes'].append("Clear Cache")
                        results['success'] = False
                
                elif fix == "Reset Configuration":
                    success = self.reset_teams_config()
                    if success:
                        results['applied_fixes'].append("Reset Configuration")
                    else:
                        results['failed_fixes'].append("Reset Configuration")
                        results['success'] = False
                
                elif fix == "Clear Web Storage":
                    success = self.clear_teams_web_storage()
                    if success:
                        results['applied_fixes'].append("Clear Web Storage")
                    else:
                        results['failed_fixes'].append("Clear Web Storage")
                        results['success'] = False
                
                elif fix == "Reset Network Settings":
                    success = self.reset_teams_network()
                    if success:
                        results['applied_fixes'].append("Reset Network Settings")
                    else:
                        results['failed_fixes'].append("Reset Network Settings")
                        results['success'] = False
                
                elif fix == "Disable GPU Hardware Acceleration":
                    success = self.disable_teams_gpu()
                    if success:
                        results['applied_fixes'].append("Disable GPU Hardware Acceleration")
                    else:
                        results['failed_fixes'].append("Disable GPU Hardware Acceleration")
                        results['success'] = False
            
            return results
        
        except Exception as e:
            logger.error(f"Error applying Teams fixes: {str(e)}")
            return {
                'success': False,
                'applied_fixes': [],
                'failed_fixes': selected_fixes,
                'error': str(e)
            }
    
    def kill_teams_processes(self):
        """Kill all Microsoft Teams processes.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            killed = False
            
            # Find and kill Teams processes
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    process_name = proc.info['name'].lower()
                    if 'teams' in process_name:
                        pid = proc.info['pid']
                        process = psutil.Process(pid)
                        process.terminate()
                        killed = True
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            # Allow time for processes to terminate
            if killed:
                time.sleep(2)
                
                # Check if any Teams processes are still running and kill them forcefully
                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        process_name = proc.info['name'].lower()
                        if 'teams' in process_name:
                            pid = proc.info['pid']
                            process = psutil.Process(pid)
                            process.kill()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
            
            return True
        
        except Exception as e:
            logger.error(f"Error killing Teams processes: {str(e)}")
            return False
    
    def clear_teams_cache(self):
        """Clear Teams cache directories.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            cache_dirs = [
                self.teams_paths['cache'],
                self.teams_paths['gpu_cache'],
                self.teams_paths['blob_storage']
            ]
            
            with self.file_lock:
                for cache_dir in cache_dirs:
                    if os.path.exists(cache_dir):
                        try:
                            # Remove all files in the directory
                            for item in os.listdir(cache_dir):
                                item_path = os.path.join(cache_dir, item)
                                try:
                                    if os.path.isfile(item_path):
                                        os.unlink(item_path)
                                    elif os.path.isdir(item_path):
                                        shutil.rmtree(item_path)
                                except Exception as e:
                                    logger.debug(f"Error removing {item_path}: {str(e)}")
                        except Exception as e:
                            logger.debug(f"Error clearing cache directory {cache_dir}: {str(e)}")
            
            return True
        
        except Exception as e:
            logger.error(f"Error clearing Teams cache: {str(e)}")
            return False
    
    def reset_teams_config(self):
        """Reset Teams configuration.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            config_file = self.teams_paths['config']
            
            if os.path.exists(config_file):
                with self.file_lock:
                    try:
                        # Backup the config file
                        backup_file = config_file + ".backup"
                        shutil.copy2(config_file, backup_file)
                        
                        # Remove the config file
                        os.unlink(config_file)
                    except Exception as e:
                        logger.debug(f"Error resetting Teams config: {str(e)}")
            
            return True
        
        except Exception as e:
            logger.error(f"Error resetting Teams configuration: {str(e)}")
            return False
    
    def clear_teams_web_storage(self):
        """Clear Teams web storage data.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            web_storage_dirs = [
                self.teams_paths['local_storage'],
                os.path.join(self.user_profile, 'AppData', 'Local', 'Microsoft', 'Teams', 'IndexedDB'),
                os.path.join(self.user_profile, 'AppData', 'Roaming', 'Microsoft', 'Teams', 'Cookies')
            ]
            
            with self.file_lock:
                for storage_dir in web_storage_dirs:
                    if os.path.exists(storage_dir):
                        try:
                            # Remove all files in the directory
                            for item in os.listdir(storage_dir):
                                item_path = os.path.join(storage_dir, item)
                                try:
                                    if os.path.isfile(item_path):
                                        os.unlink(item_path)
                                    elif os.path.isdir(item_path):
                                        shutil.rmtree(item_path)
                                except Exception as e:
                                    logger.debug(f"Error removing {item_path}: {str(e)}")
                        except Exception as e:
                            logger.debug(f"Error clearing web storage directory {storage_dir}: {str(e)}")
            
            return True
        
        except Exception as e:
            logger.error(f"Error clearing Teams web storage: {str(e)}")
            return False
    
    def reset_teams_network(self):
        """Reset Teams network settings.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Find and delete Teams network configuration files
            network_config_paths = [
                os.path.join(self.user_profile, 'AppData', 'Roaming', 'Microsoft', 'Teams', 'storage.json'),
                os.path.join(self.user_profile, 'AppData', 'Roaming', 'Microsoft', 'Teams', 'pre-IdentityCache.json'),
                os.path.join(self.user_profile, 'AppData', 'Roaming', 'Microsoft', 'Teams', 'IdentityCache.json')
            ]
            
            with self.file_lock:
                for config_file in network_config_paths:
                    if os.path.exists(config_file):
                        try:
                            # Backup the file
                            backup_file = config_file + ".backup"
                            shutil.copy2(config_file, backup_file)
                            
                            # Remove the file
                            os.unlink(config_file)
                        except Exception as e:
                            logger.debug(f"Error resetting Teams network config file {config_file}: {str(e)}")
            
            # Clear DNS cache
            try:
                subprocess.run(["ipconfig", "/flushdns"], check=False, capture_output=True)
            except Exception:
                pass
            
            # Reset Teams registry settings related to network
            try:
                reg_path = r"Software\Microsoft\Office\Teams"
                try:
                    # Try to open the key
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE)
                    
                    # Check if network-related values exist and delete them
                    try:
                        winreg.DeleteValue(key, "HomeNetworkName")
                    except:
                        pass
                    
                    try:
                        winreg.DeleteValue(key, "LastUsedNetworkCredential")
                    except:
                        pass
                    
                    # Close the key
                    winreg.CloseKey(key)
                except WindowsError:
                    # Key doesn't exist, that's fine
                    pass
            except Exception:
                pass
            
            return True
        
        except Exception as e:
            logger.error(f"Error resetting Teams network: {str(e)}")
            return False
    
    def disable_teams_gpu(self):
        """Disable Teams GPU hardware acceleration.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            config_file = self.teams_paths['config']
            
            if os.path.exists(config_file):
                with self.file_lock:
                    try:
                        # Read the config file
                        with open(config_file, 'r') as f:
                            config_data = f.read()
                        
                        # Check if the file is valid JSON
                        import json
                        try:
                            config_json = json.loads(config_data)
                            
                            # Update GPU acceleration setting
                            if 'appPreferenceSettings' not in config_json:
                                config_json['appPreferenceSettings'] = {}
                            
                            config_json['appPreferenceSettings']['hardwareAcceleration'] = False
                            
                            # Write updated config back
                            with open(config_file, 'w') as f:
                                json.dump(config_json, f, indent=2)
                            
                            return True
                        except json.JSONDecodeError:
                            # Config is corrupted, create a new one
                            new_config = {
                                "appPreferenceSettings": {
                                    "hardwareAcceleration": False
                                }
                            }
                            
                            with open(config_file, 'w') as f:
                                json.dump(new_config, f, indent=2)
                            
                            return True
                    except Exception as e:
                        logger.debug(f"Error updating Teams config for GPU: {str(e)}")
            else:
                # Config doesn't exist yet, create it
                with self.file_lock:
                    try:
                        # Create config with GPU acceleration disabled
                        new_config = {
                            "appPreferenceSettings": {
                                "hardwareAcceleration": False
                            }
                        }
                        
                        # Make sure the directory exists
                        os.makedirs(os.path.dirname(config_file), exist_ok=True)
                        
                        # Write config
                        import json
                        with open(config_file, 'w') as f:
                            json.dump(new_config, f, indent=2)
                        
                        return True
                    except Exception as e:
                        logger.debug(f"Error creating Teams config for GPU: {str(e)}")
            
            return False
        
        except Exception as e:
            logger.error(f"Error disabling Teams GPU acceleration: {str(e)}")
            return False
    
    def get_directory_size(self, path):
        """Calculate the size of a directory recursively.
        
        Args:
            path: Directory path
        
        Returns:
            Directory size in bytes
        """
        try:
            total_size = 0
            
            for dirpath, _, filenames in os.walk(path):
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    if os.path.exists(fp):
                        total_size += os.path.getsize(fp)
            
            return total_size
        except Exception as e:
            logger.debug(f"Error calculating directory size for {path}: {str(e)}")
            return 0
    
    def get_cache_size(self):
        """Get total size of Teams cache.
        
        Returns:
            Cache size in bytes
        """
        try:
            cache_dirs = [
                self.teams_paths['cache'],
                self.teams_paths['gpu_cache'],
                self.teams_paths['blob_storage'],
                self.teams_paths['databases'],
                self.teams_paths['local_storage']
            ]
            
            total_size = 0
            
            for cache_dir in cache_dirs:
                if os.path.exists(cache_dir):
                    total_size += self.get_directory_size(cache_dir)
            
            return total_size
        
        except Exception as e:
            logger.error(f"Error getting Teams cache size: {str(e)}")
            return 0


class OutlookFixer:
    """Utility for diagnosing and fixing Microsoft Outlook issues."""
    
    def __init__(self):
        """Initialize Outlook fixer with paths to known Outlook directories."""
        self.user_profile = os.path.expanduser("~")
        
        # Outlook application paths
        self.outlook_paths = {
            'app_data': os.path.join(self.user_profile, 'AppData', 'Local', 'Microsoft', 'Outlook'),
            'roaming_data': os.path.join(self.user_profile, 'AppData', 'Roaming', 'Microsoft', 'Outlook'),
            'ost_files': os.path.join(self.user_profile, 'AppData', 'Local', 'Microsoft', 'Outlook'),
            'pst_files': os.path.join(self.user_profile, 'Documents', 'Outlook Files'),
            'temp_folder': os.environ.get('TEMP', os.path.join(self.user_profile, 'AppData', 'Local', 'Temp')),
            'outlook_temp': os.path.join(os.environ.get('TEMP', os.path.join(self.user_profile, 'AppData', 'Local', 'Temp')), 'Outlook Temp')
        }
        
        # Thread-safe access for file operations
        self.file_lock = Lock()
    
    def check_status(self):
        """Check the status of Microsoft Outlook.
        
        Returns:
            Dict with status information
        """
        try:
            status = "Healthy"
            issues = []
            
            # Check if Outlook processes are running
            outlook_processes = []
            for proc in psutil.process_iter(['name']):
                try:
                    if 'outlook' in proc.info['name'].lower():
                        outlook_processes.append(proc.info['name'])
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            # Check if Outlook app data exists
            app_exists = os.path.exists(self.outlook_paths['app_data'])
            
            # Check for OST files (Offline Storage Tables)
            ost_files = []
            if os.path.exists(self.outlook_paths['ost_files']):
                for file in os.listdir(self.outlook_paths['ost_files']):
                    if file.endswith('.ost'):
                        ost_path = os.path.join(self.outlook_paths['ost_files'], file)
                        ost_files.append({
                            'path': ost_path,
                            'size': os.path.getsize(ost_path) if os.path.exists(ost_path) else 0
                        })
            
            # Check for common issues
            
            # OST file size issues (large OST files can cause performance problems)
            for ost_file in ost_files:
                if ost_file['size'] > 50 * 1024 * 1024 * 1024:  # 50 GB
                    issues.append("Large OST file")
                    status = "Issues Found"
                    break
            
            # Check for corrupted profile
            if self.check_outlook_profile_corruption():
                issues.append("Corrupted Outlook profile")
                status = "Issues Found"
            
            # Check for search index issues
            if self.check_outlook_search_issues():
                issues.append("Search index problems")
                status = "Issues Found"
            
            # Check for Outlook add-in issues
            if self.check_outlook_addin_issues():
                issues.append("Problematic add-ins detected")
                status = "Issues Found"
            
            # Check for multiple Outlook processes (can indicate stuck processes)
            outlook_exe_count = outlook_processes.count('OUTLOOK.EXE')
            if outlook_exe_count > 1:
                issues.append("Multiple Outlook processes running")
                status = "Issues Found"
            
            return {
                'status': status,
                'issues': issues,
                'processes': outlook_processes,
                'app_exists': app_exists,
                'ost_files': ost_files
            }
        
        except Exception as e:
            logger.error(f"Error checking Outlook status: {str(e)}")
            return {
                'status': "Error",
                'issues': [str(e)],
                'processes': [],
                'app_exists': False,
                'ost_files': []
            }
    
    def check_outlook_profile_corruption(self):
        """Check for Outlook profile corruption.
        
        Returns:
            True if corruption detected, False otherwise
        """
        try:
            # Check if Outlook profile registry keys are accessible
            profile_corrupted = False
            
            # Check Outlook profiles registry
            try:
                reg_path = r"Software\Microsoft\Office\16.0\Outlook\Profiles"
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                winreg.CloseKey(key)
            except WindowsError:
                # Try older versions
                try:
                    reg_path = r"Software\Microsoft\Office\15.0\Outlook\Profiles"
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                    winreg.CloseKey(key)
                except WindowsError:
                    try:
                        reg_path = r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
                        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                        winreg.CloseKey(key)
                    except WindowsError:
                        profile_corrupted = True
            
            # Check for problematic files in Outlook directories
            if os.path.exists(self.outlook_paths['app_data']):
                for file in os.listdir(self.outlook_paths['app_data']):
                    # Look for zero-byte OST/PST files or lock files that shouldn't be present when Outlook is closed
                    if file.endswith(('.ost', '.pst')):
                        file_path = os.path.join(self.outlook_paths['app_data'], file)
                        if os.path.exists(file_path) and os.path.getsize(file_path) == 0:
                            profile_corrupted = True
                            break
                    elif file.endswith('.lck'):
                        outlook_running = False
                        for proc in psutil.process_iter(['name']):
                            try:
                                if 'outlook' in proc.info['name'].lower():
                                    outlook_running = True
                                    break
                            except (psutil.NoSuchProcess, psutil.AccessDenied):
                                pass
                        
                        if not outlook_running:
                            profile_corrupted = True
                            break
            
            return profile_corrupted
        except Exception as e:
            logger.debug(f"Error checking Outlook profile corruption: {str(e)}")
            return False
    
    def check_outlook_search_issues(self):
        """Check for Outlook search index issues.
        
        Returns:
            True if issues detected, False otherwise
        """
        try:
            # Check if Windows Search service is running
            search_service_running = False
            
            try:
                cmd = [
                    "powershell",
                    "-Command",
                    "Get-Service -Name 'WSearch' | Select-Object -ExpandProperty Status"
                ]
                
                result = subprocess.run(cmd, capture_output=True, text=True, check=False)
                
                if result.returncode == 0 and "Running" in result.stdout:
                    search_service_running = True
            except Exception:
                pass
            
            if not search_service_running:
                return True  # Search issues detected
            
            # Check Windows Search registry settings
            try:
                reg_path = r"Software\Microsoft\Windows Search\CrawlScopeManager\Windows\SystemIndex\WorkingSetRules\OutlookExpressEmail"
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_READ)
                winreg.CloseKey(key)
            except WindowsError:
                return True  # Search issues detected
            
            return False  # No search issues detected
        except Exception as e:
            logger.debug(f"Error checking Outlook search issues: {str(e)}")
            return False
    
    def check_outlook_addin_issues(self):
        """Check for problematic Outlook add-ins.
        
        Returns:
            True if issues detected, False otherwise
        """
        try:
            # Check Outlook add-ins registry
            problematic_addins = False
            
            # Look for disabled add-ins which might indicate problems
            try:
                reg_path = r"Software\Microsoft\Office\Outlook\Addins"
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                
                # Count number of subkeys (add-ins)
                num_addins = winreg.QueryInfoKey(key)[0]
                
                for i in range(num_addins):
                    addin_name = winreg.EnumKey(key, i)
                    addin_key = winreg.OpenKey(key, addin_name, 0, winreg.KEY_READ)
                    
                    try:
                        # Check if add-in is disabled
                        load_behavior, _ = winreg.QueryValueEx(addin_key, "LoadBehavior")
                        if load_behavior == 0:  # Disabled add-in
                            problematic_addins = True
                    except WindowsError:
                        pass
                    
                    winreg.CloseKey(addin_key)
                
                winreg.CloseKey(key)
            except WindowsError:
                pass
            
            return problematic_addins
        except Exception as e:
            logger.debug(f"Error checking Outlook add-in issues: {str(e)}")
            return False
    
    def analyze(self):
        """Analyze Microsoft Outlook for issues and recommend fixes.
        
        Returns:
            Dict with analysis results and recommended fixes
        """
        try:
            # Check current status
            status = self.check_status()
            issues = status.get('issues', [])
            
            # Additional analysis
            found_issues = len(issues) > 0
            recommended_fixes = []
            
            # Check for profile issues
            if "Corrupted Outlook profile" in issues:
                found_issues = True
                recommended_fixes.append("Repair Outlook Profile")
            
            # Check for large OST files
            if "Large OST file" in issues:
                found_issues = True
                recommended_fixes.append("Clear Outlook Cache")
            
            # Check for search issues
            if "Search index problems" in issues:
                found_issues = True
                recommended_fixes.append("Rebuild Search Index")
            
            # Check for autodiscover issues
            if self.check_outlook_autodiscover_issues():
                found_issues = True
                recommended_fixes.append("Reset AutoDiscover")
            
            # Check for email rules issues
            if self.check_outlook_rules_issues():
                found_issues = True
                recommended_fixes.append("Repair Rules")
            
            # Check for add-in issues
            if "Problematic add-ins detected" in issues:
                found_issues = True
                recommended_fixes.append("Disable Add-ins")
            
            # If no specific issues found but status indicates problems, recommend basic fixes
            if found_issues and not recommended_fixes:
                recommended_fixes = ["Repair Outlook Profile", "Clear Outlook Cache"]
            
            return {
                'found_issues': found_issues,
                'issues': issues,
                'recommended_fixes': recommended_fixes
            }
        
        except Exception as e:
            logger.error(f"Error analyzing Outlook: {str(e)}")
            return {
                'found_issues': True,
                'issues': [str(e)],
                'recommended_fixes': ["Repair Outlook Profile", "Clear Outlook Cache", "Rebuild Search Index"]
            }
    
    def check_outlook_autodiscover_issues(self):
        """Check for Outlook autodiscover issues.
        
        Returns:
            True if issues detected, False otherwise
        """
        try:
            # Check for registry settings related to autodiscover
            autodiscover_issues = False
            
            # Check Outlook autodiscover registry settings
            try:
                reg_path = r"Software\Microsoft\Office\16.0\Outlook\AutoDiscover"
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                
                try:
                    exclude_http_redirect, _ = winreg.QueryValueEx(key, "ExcludeHttpRedirect")
                    if exclude_http_redirect == 1:
                        autodiscover_issues = True
                except WindowsError:
                    pass
                
                try:
                    exclude_scpLookup, _ = winreg.QueryValueEx(key, "ExcludeScpLookup")
                    if exclude_scpLookup == 1:
                        autodiscover_issues = True
                except WindowsError:
                    pass
                
                winreg.CloseKey(key)
            except WindowsError:
                # Key doesn't exist, check older versions
                try:
                    reg_path = r"Software\Microsoft\Office\15.0\Outlook\AutoDiscover"
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                    winreg.CloseKey(key)
                except WindowsError:
                    autodiscover_issues = True
            
            return autodiscover_issues
        except Exception as e:
            logger.debug(f"Error checking Outlook autodiscover issues: {str(e)}")
            return False
    
    def check_outlook_rules_issues(self):
        """Check for Outlook rules issues.
        
        Returns:
            True if issues detected, False otherwise
        """
        try:
            # Check for rules files
            rules_issues = False
            
            # Check for rules files in the Outlook profile folders
            rules_file_path = os.path.join(self.outlook_paths['roaming_data'], 'Outlook.rwz')
            
            if os.path.exists(rules_file_path):
                # Check if the file is locked or has zero size
                if os.path.getsize(rules_file_path) == 0:
                    rules_issues = True
                else:
                    try:
                        # Try to open the file to check if it's locked
                        with open(rules_file_path, 'rb') as f:
                            _ = f.read(1)  # Try to read a byte
                    except IOError:
                        rules_issues = True
            
            # Check registry for rules settings
            try:
                reg_path = r"Software\Microsoft\Office\16.0\Outlook\Options\Mail"
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                
                try:
                    # Check for disabled rules
                    rules_enabled, _ = winreg.QueryValueEx(key, "EnableRules")
                    if rules_enabled == 0:  # Rules disabled
                        rules_issues = True
                except WindowsError:
                    pass
                
                winreg.CloseKey(key)
            except WindowsError:
                pass
            
            return rules_issues
        except Exception as e:
            logger.debug(f"Error checking Outlook rules issues: {str(e)}")
            return False
    
    def apply_fixes(self, selected_fixes):
        """Apply selected fixes to Microsoft Outlook.
        
        Args:
            selected_fixes: List of fix names to apply
        
        Returns:
            Dict with fix results
        """
        try:
            results = {
                'success': True,
                'applied_fixes': [],
                'failed_fixes': []
            }
            
            # Kill Outlook processes first if needed
            kill_outlook = False
            for fix in selected_fixes:
                if fix in ["Repair Outlook Profile", "Clear Outlook Cache", "Disable Add-ins"]:
                    kill_outlook = True
                    break
            
            if kill_outlook:
                self.kill_outlook_processes()
            
            # Apply each selected fix
            for fix in selected_fixes:
                if fix == "Repair Outlook Profile":
                    success = self.repair_outlook_profile()
                    if success:
                        results['applied_fixes'].append("Repair Outlook Profile")
                    else:
                        results['failed_fixes'].append("Repair Outlook Profile")
                        results['success'] = False
                
                elif fix == "Clear Outlook Cache":
                    success = self.clear_outlook_cache()
                    if success:
                        results['applied_fixes'].append("Clear Outlook Cache")
                    else:
                        results['failed_fixes'].append("Clear Outlook Cache")
                        results['success'] = False
                
                elif fix == "Rebuild Search Index":
                    success = self.rebuild_search_index()
                    if success:
                        results['applied_fixes'].append("Rebuild Search Index")
                    else:
                        results['failed_fixes'].append("Rebuild Search Index")
                        results['success'] = False
                
                elif fix == "Reset AutoDiscover":
                    success = self.reset_autodiscover()
                    if success:
                        results['applied_fixes'].append("Reset AutoDiscover")
                    else:
                        results['failed_fixes'].append("Reset AutoDiscover")
                        results['success'] = False
                
                elif fix == "Repair Rules":
                    success = self.repair_outlook_rules()
                    if success:
                        results['applied_fixes'].append("Repair Rules")
                    else:
                        results['failed_fixes'].append("Repair Rules")
                        results['success'] = False
                
                elif fix == "Disable Add-ins":
                    success = self.disable_outlook_addins()
                    if success:
                        results['applied_fixes'].append("Disable Add-ins")
                    else:
                        results['failed_fixes'].append("Disable Add-ins")
                        results['success'] = False
            
            return results
        
        except Exception as e:
            logger.error(f"Error applying Outlook fixes: {str(e)}")
            return {
                'success': False,
                'applied_fixes': [],
                'failed_fixes': selected_fixes,
                'error': str(e)
            }
    
    def kill_outlook_processes(self):
        """Kill all Microsoft Outlook processes.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            killed = False
            
            # Find and kill Outlook processes
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    process_name = proc.info['name'].lower()
                    if 'outlook' in process_name:
                        pid = proc.info['pid']
                        process = psutil.Process(pid)
                        process.terminate()
                        killed = True
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            # Allow time for processes to terminate
            if killed:
                time.sleep(2)
                
                # Check if any Outlook processes are still running and kill them forcefully
                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        process_name = proc.info['name'].lower()
                        if 'outlook' in process_name:
                            pid = proc.info['pid']
                            process = psutil.Process(pid)
                            process.kill()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
            
            return True
        
        except Exception as e:
            logger.error(f"Error killing Outlook processes: {str(e)}")
            return False
    
    def repair_outlook_profile(self):
        """Repair Outlook profile.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Get the Outlook profile name
            profile_name = self.get_outlook_profile_name()
            
            if not profile_name:
                profile_name = "Outlook"  # Default profile name
            
            # Run Outlook.exe with /resetnavpane to repair the navigation pane
            try:
                outlook_path = self.find_outlook_executable()
                if outlook_path:
                    subprocess.run([outlook_path, "/resetnavpane"], check=False, capture_output=True)
                    time.sleep(2)
                    self.kill_outlook_processes()
            except Exception:
                pass
            
            # Clear Outlook profile registry entries
            try:
                # Try to find the correct registry key version
                office_versions = ["16.0", "15.0", "14.0"]
                
                for version in office_versions:
                    reg_path = f"Software\\Microsoft\\Office\\{version}\\Outlook\\Profiles\\{profile_name}"
                    try:
                        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                        winreg.CloseKey(key)
                        
                        # We found a valid profile, now remove problematic subkeys
                        try:
                            # Rename profile to backup
                            backup_reg_path = f"Software\\Microsoft\\Office\\{version}\\Outlook\\Profiles\\{profile_name}_Backup"
                            
                            # Check if backup already exists and delete it
                            try:
                                backup_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, backup_reg_path, 0, winreg.KEY_READ)
                                winreg.CloseKey(backup_key)
                                # Delete the backup key if it exists
                                self.delete_registry_key(winreg.HKEY_CURRENT_USER, backup_reg_path)
                            except WindowsError:
                                pass
                            
                            # Copy the profile to a backup
                            self.copy_registry_key(
                                winreg.HKEY_CURRENT_USER, 
                                reg_path, 
                                winreg.HKEY_CURRENT_USER, 
                                backup_reg_path
                            )
                            
                            # Delete problematic subkeys
                            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                            num_subkeys = winreg.QueryInfoKey(key)[0]
                            
                            # Get all subkey names
                            subkey_names = []
                            for i in range(num_subkeys):
                                subkey_names.append(winreg.EnumKey(key, i))
                            
                            winreg.CloseKey(key)
                            
                            # Now delete problematic subkeys
                            for subkey_name in subkey_names:
                                if "9375CFF0413111d3B88A00104B2A6676" in subkey_name:
                                    # This is the Exchange account settings
                                    subkey_path = f"{reg_path}\\{subkey_name}"
                                    self.delete_registry_key(winreg.HKEY_CURRENT_USER, subkey_path)
                        except Exception as e:
                            logger.debug(f"Error modifying profile registry: {str(e)}")
                    except WindowsError:
                        pass
            except Exception as e:
                logger.debug(f"Error clearing profile registry: {str(e)}")
            
            # Run Outlook.exe with /resetfolders to repair folder views
            try:
                if outlook_path:
                    subprocess.run([outlook_path, "/resetfolders"], check=False, capture_output=True)
                    time.sleep(2)
                    self.kill_outlook_processes()
            except Exception:
                pass
            
            # Run Outlook.exe with /cleanprofile to clean up the profile
            try:
                if outlook_path:
                    subprocess.run([outlook_path, "/cleanprofile", profile_name], check=False, capture_output=True)
                    time.sleep(2)
                    self.kill_outlook_processes()
            except Exception:
                pass
            
            return True
        
        except Exception as e:
            logger.error(f"Error repairing Outlook profile: {str(e)}")
            return False
    
    def clear_outlook_cache(self):
        """Clear Outlook cache and OST/PST files.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            with self.file_lock:
                # Clear Outlook temp files
                temp_dir = self.outlook_paths['temp_folder']
                outlook_temp = self.outlook_paths['outlook_temp']
                
                try:
                    if os.path.exists(temp_dir):
                        for item in os.listdir(temp_dir):
                            if 'outlook' in item.lower():
                                item_path = os.path.join(temp_dir, item)
                                try:
                                    if os.path.isfile(item_path):
                                        os.unlink(item_path)
                                    elif os.path.isdir(item_path):
                                        shutil.rmtree(item_path)
                                except Exception as e:
                                    logger.debug(f"Error removing temp file {item_path}: {str(e)}")
                except Exception as e:
                    logger.debug(f"Error clearing Outlook temp files: {str(e)}")
                
                # Clear Outlook temp folder
                try:
                    if os.path.exists(outlook_temp):
                        for item in os.listdir(outlook_temp):
                            item_path = os.path.join(outlook_temp, item)
                            try:
                                if os.path.isfile(item_path):
                                    os.unlink(item_path)
                                elif os.path.isdir(item_path):
                                    shutil.rmtree(item_path)
                            except Exception as e:
                                logger.debug(f"Error removing Outlook temp file {item_path}: {str(e)}")
                except Exception as e:
                    logger.debug(f"Error clearing Outlook temp folder: {str(e)}")
                
                # Reset OST files
                ost_path = self.outlook_paths['ost_files']
                try:
                    if os.path.exists(ost_path):
                        for file in os.listdir(ost_path):
                            if file.endswith('.ost'):
                                # We don't want to delete OST files, just rename them to force recreation
                                file_path = os.path.join(ost_path, file)
                                try:
                                    # Rename to .old extension
                                    backup_path = file_path + '.old'
                                    # Remove old backup if it exists
                                    if os.path.exists(backup_path):
                                        os.unlink(backup_path)
                                    # Rename current file
                                    os.rename(file_path, backup_path)
                                except Exception as e:
                                    logger.debug(f"Error resetting OST file {file_path}: {str(e)}")
                except Exception as e:
                    logger.debug(f"Error resetting OST files: {str(e)}")
            
            # Clear registry entries for Outlook autocomplete cache
            try:
                # Find Office version
                office_versions = ["16.0", "15.0", "14.0"]
                
                for version in office_versions:
                    reg_path = f"Software\\Microsoft\\Office\\{version}\\Outlook\\AutoComplete"
                    try:
                        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE)
                        
                        try:
                            # Clear roamed autocomplete
                            winreg.DeleteValue(key, "Roamed")
                        except WindowsError:
                            pass
                        
                        try:
                            # Clear stream autocomplete
                            winreg.DeleteValue(key, "Stream")
                        except WindowsError:
                            pass
                        
                        winreg.CloseKey(key)
                    except WindowsError:
                        pass
            except Exception as e:
                logger.debug(f"Error clearing autocomplete cache: {str(e)}")
            
            return True
        
        except Exception as e:
            logger.error(f"Error clearing Outlook cache: {str(e)}")
            return False
    
    def rebuild_search_index(self):
        """Rebuild Outlook search index.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Reset search index for Outlook
            
            # First, try using the Exchange rebuild command if Outlook is installed
            outlook_path = self.find_outlook_executable()
            if outlook_path:
                try:
                    # Run Outlook with the /resetfolderviews parameter to reset search
                    subprocess.run([outlook_path, "/resetfolderviews"], check=False, capture_output=True)
                    time.sleep(2)
                    self.kill_outlook_processes()
                except Exception:
                    pass
            
            # Reset Windows Search Service
            try:
                # Restart Windows Search service
                cmd = [
                    "powershell", 
                    "-Command",
                    "Stop-Service 'WSearch' -Force; Start-Service 'WSearch'"
                ]
                
                subprocess.run(cmd, check=False, capture_output=True)
            except Exception:
                pass
            
            # Reset search index registry settings for Outlook
            try:
                # Clear catalyst settings
                reg_path = r"Software\Microsoft\Office\16.0\Outlook\Search\Catalyst"
                try:
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE)
                    
                    try:
                        # Set ResetCatalystAPI to 1 to force rebuild
                        winreg.SetValueEx(key, "ResetCatalystAPI", 0, winreg.REG_DWORD, 1)
                    except WindowsError:
                        pass
                    
                    winreg.CloseKey(key)
                except WindowsError:
                    pass
            except Exception as e:
                logger.debug(f"Error resetting search registry: {str(e)}")
            
            # Force Windows to rebuild index
            try:
                # Add/modify registry value to rebuild search index
                reg_path = r"SOFTWARE\Microsoft\Windows Search\Gather\Windows\SystemIndex"
                try:
                    # First try to open with write access (requires admin)
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_WRITE)
                    
                    # Set SetupCompletedSuccessfully to 0 to force rebuild
                    winreg.SetValueEx(key, "SetupCompletedSuccessfully", 0, winreg.REG_DWORD, 0)
                    
                    winreg.CloseKey(key)
                except WindowsError:
                    # If that fails, try user-level indexing options
                    pass
            except Exception as e:
                logger.debug(f"Error forcing search index rebuild: {str(e)}")
            
            # Use Windows Search troubleshooter script
            try:
                # A PowerShell command to reset the Windows Search
                cmd = [
                    "powershell",
                    "-Command",
                    "$searchPath = 'HKLM:\\SOFTWARE\\Microsoft\\Windows Search'; "
                    "if (Test-Path $searchPath) { "
                    "  $searchKey = Get-Item $searchPath; "
                    "  $searchKey.SetValue('SetupCompletedSuccessfully', 0, 'DWord'); "
                    "  Restart-Service WSearch -Force; "
                    "}"
                ]
                
                subprocess.run(cmd, check=False, capture_output=True)
            except Exception:
                pass
            
            return True
        
        except Exception as e:
            logger.error(f"Error rebuilding search index: {str(e)}")
            return False
    
    def reset_autodiscover(self):
        """Reset Outlook autodiscover settings.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Reset autodiscover registry settings
            try:
                # Find Office version
                office_versions = ["16.0", "15.0", "14.0"]
                
                for version in office_versions:
                    reg_path = f"Software\\Microsoft\\Office\\{version}\\Outlook\\AutoDiscover"
                    try:
                        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE)
                        
                        # Reset all autodiscover settings
                        try:
                            winreg.DeleteValue(key, "ExcludeHttpRedirect")
                        except WindowsError:
                            pass
                        
                        try:
                            winreg.DeleteValue(key, "ExcludeHttpsAutoDiscoverDomain")
                        except WindowsError:
                            pass
                        
                        try:
                            winreg.DeleteValue(key, "ExcludeHttpsRootDomain")
                        except WindowsError:
                            pass
                        
                        try:
                            winreg.DeleteValue(key, "ExcludeScpLookup")
                        except WindowsError:
                            pass
                        
                        try:
                            winreg.DeleteValue(key, "ExcludeSrvRecord")
                        except WindowsError:
                            pass
                        
                        # Add PreferLocalXML setting
                        winreg.SetValueEx(key, "PreferLocalXML", 0, winreg.REG_DWORD, 1)
                        
                        winreg.CloseKey(key)
                    except WindowsError:
                        # Key doesn't exist, create it
                        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path)
                        winreg.SetValueEx(key, "PreferLocalXML", 0, winreg.REG_DWORD, 1)
                        winreg.CloseKey(key)
            except Exception as e:
                logger.debug(f"Error resetting autodiscover registry: {str(e)}")
            
            # Clear autodiscover cache
            autodiscover_cache_path = os.path.join(self.outlook_paths['roaming_data'], 'Autodiscover')
            
            try:
                if os.path.exists(autodiscover_cache_path):
                    with self.file_lock:
                        for item in os.listdir(autodiscover_cache_path):
                            item_path = os.path.join(autodiscover_cache_path, item)
                            try:
                                if os.path.isfile(item_path):
                                    os.unlink(item_path)
                                elif os.path.isdir(item_path):
                                    shutil.rmtree(item_path)
                            except Exception as e:
                                logger.debug(f"Error removing autodiscover cache item {item_path}: {str(e)}")
            except Exception as e:
                logger.debug(f"Error clearing autodiscover cache: {str(e)}")
            
            return True
        
        except Exception as e:
            logger.error(f"Error resetting autodiscover: {str(e)}")
            return False
    
    def repair_outlook_rules(self):
        """Repair Outlook rules.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # First, backup current rules file
            rules_file_path = os.path.join(self.outlook_paths['roaming_data'], 'Outlook.rwz')
            
            with self.file_lock:
                if os.path.exists(rules_file_path):
                    try:
                        # Create backup
                        backup_path = rules_file_path + '.backup'
                        if os.path.exists(backup_path):
                            os.unlink(backup_path)
                        shutil.copy2(rules_file_path, backup_path)
                        
                        # Remove original rules file to let Outlook recreate it
                        os.unlink(rules_file_path)
                    except Exception as e:
                        logger.debug(f"Error backing up rules file: {str(e)}")
            
            # Enable rules execution in registry
            try:
                # Find Office version
                office_versions = ["16.0", "15.0", "14.0"]
                
                for version in office_versions:
                    reg_path = f"Software\\Microsoft\\Office\\{version}\\Outlook\\Options\\Mail"
                    try:
                        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE)
                        
                        # Enable rules
                        winreg.SetValueEx(key, "EnableRules", 0, winreg.REG_DWORD, 1)
                        
                        winreg.CloseKey(key)
                    except WindowsError:
                        # Key doesn't exist, try to create it
                        try:
                            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path)
                            winreg.SetValueEx(key, "EnableRules", 0, winreg.REG_DWORD, 1)
                            winreg.CloseKey(key)
                        except WindowsError:
                            pass
            except Exception as e:
                logger.debug(f"Error enabling rules in registry: {str(e)}")
            
            # Clear rules registry cache
            try:
                # Find Rules registry key
                office_versions = ["16.0", "15.0", "14.0"]
                
                for version in office_versions:
                    reg_path = f"Software\\Microsoft\\Office\\{version}\\Outlook\\Rules"
                    try:
                        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                        winreg.CloseKey(key)
                        
                        # Delete the Rules key to reset
                        self.delete_registry_key(winreg.HKEY_CURRENT_USER, reg_path)
                    except WindowsError:
                        pass
            except Exception as e:
                logger.debug(f"Error clearing rules registry cache: {str(e)}")
            
            return True
        
        except Exception as e:
            logger.error(f"Error repairing Outlook rules: {str(e)}")
            return False
    
    def disable_outlook_addins(self):
        """Disable problematic Outlook add-ins.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Identify and disable problematic add-ins
            problematic_addins_disabled = False
            
            # List of known problematic add-ins (partial matches)
            known_problematic = [
                "gotomeeting", "webex", "zoom", "skype", "teams connector",
                "access connector", "bcm", "social connector", "dropbox",
                "evernote", "google calendar", "acrobat", "right signature",
                "salesforce", "crm", "amadeus", "sabre", "linkedin"
            ]
            
            # Find Office version
            office_versions = ["16.0", "15.0", "14.0"]
            
            for version in office_versions:
                reg_path = f"Software\\Microsoft\\Office\\{version}\\Outlook\\Addins"
                try:
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                    
                    # Count number of subkeys (add-ins)
                    num_addins = winreg.QueryInfoKey(key)[0]
                    
                    # Get all add-in names
                    addin_names = []
                    for i in range(num_addins):
                        addin_names.append(winreg.EnumKey(key, i))
                    
                    winreg.CloseKey(key)
                    
                    # Process each add-in
                    for addin_name in addin_names:
                        addin_key_path = f"{reg_path}\\{addin_name}"
                        addin_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, addin_key_path, 0, winreg.KEY_READ | winreg.KEY_WRITE)
                        
                        # Check if add-in is in known problematic list
                        is_problematic = False
                        addin_lower = addin_name.lower()
                        
                        for problem_name in known_problematic:
                            if problem_name in addin_lower:
                                is_problematic = True
                                break
                        
                        # Get current load behavior
                        try:
                            load_behavior, _ = winreg.QueryValueEx(addin_key, "LoadBehavior")
                            
                            # Check if already disabled (0) or if problematic
                            if load_behavior != 0 and is_problematic:
                                # Disable the add-in by setting LoadBehavior to 0
                                winreg.SetValueEx(addin_key, "LoadBehavior", 0, winreg.REG_DWORD, 0)
                                problematic_addins_disabled = True
                        except WindowsError:
                            pass
                        
                        winreg.CloseKey(addin_key)
                except WindowsError:
                    pass
            
            # If no problematic add-ins were found, try to disable any crashed add-ins
            if not problematic_addins_disabled:
                for version in office_versions:
                    crash_reg_path = f"Software\\Microsoft\\Office\\{version}\\Outlook\\Resiliency"
                    try:
                        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, crash_reg_path, 0, winreg.KEY_READ)
                        winreg.CloseKey(key)
                        
                        # Clear the crash detection data to reset add-ins
                        self.delete_registry_key(winreg.HKEY_CURRENT_USER, crash_reg_path)
                    except WindowsError:
                        pass
            
            return True
        
        except Exception as e:
            logger.error(f"Error disabling Outlook add-ins: {str(e)}")
            return False
    
    def find_outlook_executable(self):
        """Find Outlook executable path.
        
        Returns:
            Path to Outlook.exe if found, None otherwise
        """
        try:
            # Common Outlook installation paths
            possible_paths = [
                os.path.join(os.environ.get('PROGRAMFILES', 'C:\\Program Files'), 'Microsoft Office\\root\\Office16\\OUTLOOK.EXE'),
                os.path.join(os.environ.get('PROGRAMFILES', 'C:\\Program Files'), 'Microsoft Office\\Office16\\OUTLOOK.EXE'),
                os.path.join(os.environ.get('PROGRAMFILES', 'C:\\Program Files'), 'Microsoft Office\\Office15\\OUTLOOK.EXE'),
                os.path.join(os.environ.get('PROGRAMFILES', 'C:\\Program Files'), 'Microsoft Office\\Office14\\OUTLOOK.EXE'),
                os.path.join(os.environ.get('PROGRAMFILES(X86)', 'C:\\Program Files (x86)'), 'Microsoft Office\\root\\Office16\\OUTLOOK.EXE'),
                os.path.join(os.environ.get('PROGRAMFILES(X86)', 'C:\\Program Files (x86)'), 'Microsoft Office\\Office16\\OUTLOOK.EXE'),
                os.path.join(os.environ.get('PROGRAMFILES(X86)', 'C:\\Program Files (x86)'), 'Microsoft Office\\Office15\\OUTLOOK.EXE'),
                os.path.join(os.environ.get('PROGRAMFILES(X86)', 'C:\\Program Files (x86)'), 'Microsoft Office\\Office14\\OUTLOOK.EXE')
            ]
            
            for path in possible_paths:
                if os.path.exists(path):
                    return path
            
            # Try to find using registry
            try:
                # Check common registry paths for Outlook installation
                reg_paths = [
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
                    r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
                ]
                
                for reg_path in reg_paths:
                    try:
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_READ)
                        path, _ = winreg.QueryValueEx(key, "")
                        winreg.CloseKey(key)
                        
                        if os.path.exists(path):
                            return path
                    except WindowsError:
                        pass
            except Exception:
                pass
            
            return None
        
        except Exception as e:
            logger.debug(f"Error finding Outlook executable: {str(e)}")
            return None
    
    def get_outlook_profile_name(self):
        """Get the current Outlook profile name.
        
        Returns:
            Profile name if found, None otherwise
        """
        try:
            profile_name = None
            
            # Try to find the current profile from registry
            office_versions = ["16.0", "15.0", "14.0"]
            
            for version in office_versions:
                reg_path = f"Software\\Microsoft\\Office\\{version}\\Outlook"
                try:
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
                    
                    try:
                        # Get default profile name
                        profile_name, _ = winreg.QueryValueEx(key, "DefaultProfile")
                    except WindowsError:
                        pass
                    
                    winreg.CloseKey(key)
                    
                    if profile_name:
                        break
                except WindowsError:
                    pass
            
            return profile_name
        
        except Exception as e:
            logger.debug(f"Error getting Outlook profile name: {str(e)}")
            return None
    
    def delete_registry_key(self, root, key_path):
        """Recursively delete a registry key and all its subkeys.
        
        Args:
            root: Registry root key (HKEY_CURRENT_USER, etc.)
            key_path: Path to the key to delete
        """
        try:
            # Open the key to get subkey information
            key = winreg.OpenKey(root, key_path, 0, winreg.KEY_READ)
            
            # Get number of subkeys
            info = winreg.QueryInfoKey(key)
            num_subkeys = info[0]
            
            # Collect subkey names (can't delete while enumerating)
            subkey_names = []
            for i in range(num_subkeys):
                subkey_names.append(winreg.EnumKey(key, i))
            
            winreg.CloseKey(key)
            
            # Recursively delete subkeys
            for subkey_name in subkey_names:
                subkey_path = f"{key_path}\\{subkey_name}"
                self.delete_registry_key(root, subkey_path)
            
            # Delete the key itself
            winreg.DeleteKey(root, key_path)
        except WindowsError:
            pass
    
    def copy_registry_key(self, src_root, src_path, dst_root, dst_path):
        """Recursively copy a registry key and all its values and subkeys.
        
        Args:
            src_root: Source registry root key
            src_path: Source key path
            dst_root: Destination registry root key
            dst_path: Destination key path
        """
        try:
            # Create destination key
            dst_key = winreg.CreateKey(dst_root, dst_path)
            
            # Open source key
            src_key = winreg.OpenKey(src_root, src_path, 0, winreg.KEY_READ)
            
            # Copy values
            try:
                info = winreg.QueryInfoKey(src_key)
                for i in range(info[1]):
                    name, value, type_id = winreg.EnumValue(src_key, i)
                    winreg.SetValueEx(dst_key, name, 0, type_id, value)
            except WindowsError:
                pass
            
            # Copy subkeys recursively
            try:
                info = winreg.QueryInfoKey(src_key)
                for i in range(info[0]):
                    subkey_name = winreg.EnumKey(src_key, i)
                    self.copy_registry_key(
                        src_root, f"{src_path}\\{subkey_name}",
                        dst_root, f"{dst_path}\\{subkey_name}"
                    )
            except WindowsError:
                pass
            
            winreg.CloseKey(src_key)
            winreg.CloseKey(dst_key)
        except WindowsError:
            pass
    
    def get_cache_size(self):
        """Get total size of Outlook cache.
        
        Returns:
            Cache size in bytes
        """
        try:
            # Paths to check for cache files
            cache_paths = [
                self.outlook_paths['app_data'],
                self.outlook_paths['outlook_temp'],
                os.path.join(self.outlook_paths['temp_folder'], 'Outlook Logging')
            ]
            
            total_size = 0
            
            # Check OST/PST files
            if os.path.exists(self.outlook_paths['ost_files']):
                for file in os.listdir(self.outlook_paths['ost_files']):
                    if file.endswith(('.ost', '.pst')):
                        file_path = os.path.join(self.outlook_paths['ost_files'], file)
                        if os.path.exists(file_path):
                            total_size += os.path.getsize(file_path)
            
            # Check temp and cache directories
            for path in cache_paths:
                if os.path.exists(path):
                    total_size += self.get_directory_size(path)
            
            return total_size
        
        except Exception as e:
            logger.error(f"Error getting Outlook cache size: {str(e)}")
            return 0
    
    def get_directory_size(self, path):
        """Calculate the size of a directory recursively.
        
        Args:
            path: Directory path
        
        Returns:
            Directory size in bytes
        """
        try:
            total_size = 0
            
            for dirpath, _, filenames in os.walk(path):
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    if os.path.exists(fp):
                        total_size += os.path.getsize(fp)
            
            return total_size
        except Exception as e:
            logger.debug(f"Error calculating directory size for {path}: {str(e)}")
            return 0
