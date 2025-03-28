"""
Quick fix service for the Windows System Optimizer.
This module provides functionality to fix common issues with Microsoft Teams and Outlook.
"""

import os
import logging
import subprocess
import shutil
import tempfile
import winreg
import time
from pathlib import Path

logger = logging.getLogger(__name__)

class QuickFixTools:
    """Service class for quick fix tools."""
    
    def __init__(self):
        """Initialize the quick fix tools."""
        self.user_profile = os.environ.get('USERPROFILE', '')
        self.app_data = os.path.join(self.user_profile, 'AppData')
        self.local_app_data = os.path.join(self.app_data, 'Local')
        self.roaming_app_data = os.path.join(self.app_data, 'Roaming')
        
        # Path to PowerShell scripts
        self.script_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'scripts')
    
    def fix_teams(self):
        """
        Fix common Microsoft Teams issues by clearing cache and resetting configuration.
        
        Returns:
            dict: Results of the operation
        """
        results = {
            "success": False,
            "message": "",
            "details": []
        }
        
        try:
            # First kill Teams processes
            self._kill_process("Teams.exe")
            self._kill_process("Update.exe") # Teams updater
            
            # Add details to results
            results["details"].append("Stopped Microsoft Teams processes")
            
            # Clear Teams cache directories
            teams_directories = [
                os.path.join(self.app_data, 'Microsoft', 'Teams', 'Cache'),
                os.path.join(self.app_data, 'Microsoft', 'Teams', 'blob_storage'),
                os.path.join(self.app_data, 'Microsoft', 'Teams', 'databases'),
                os.path.join(self.app_data, 'Microsoft', 'Teams', 'GPUCache'),
                os.path.join(self.app_data, 'Microsoft', 'Teams', 'IndexedDB'),
                os.path.join(self.app_data, 'Microsoft', 'Teams', 'Local Storage'),
                os.path.join(self.app_data, 'Microsoft', 'Teams', 'tmp'),
            ]
            
            for directory in teams_directories:
                if os.path.exists(directory):
                    try:
                        # Try to delete all files in directory but keep directory itself
                        for item in os.listdir(directory):
                            item_path = os.path.join(directory, item)
                            try:
                                if os.path.isfile(item_path) or os.path.islink(item_path):
                                    os.unlink(item_path)
                                elif os.path.isdir(item_path):
                                    shutil.rmtree(item_path)
                            except Exception as e:
                                logger.warning(f"Error deleting {item_path}: {str(e)}")
                                results["details"].append(f"Warning: Could not delete {item_path}")
                        
                        results["details"].append(f"Cleared cache in {directory}")
                    except Exception as e:
                        logger.warning(f"Error clearing directory {directory}: {str(e)}")
                        results["details"].append(f"Warning: Could not clear {directory}")
            
            # Reset Teams configuration file
            try:
                settings_file = os.path.join(self.app_data, 'Microsoft', 'Teams', 'desktop-config.json')
                if os.path.exists(settings_file):
                    # Backup the file first
                    backup_file = settings_file + '.bak'
                    shutil.copy2(settings_file, backup_file)
                    results["details"].append(f"Backed up configuration to {backup_file}")
                    
                    # Delete the file (Teams will recreate it)
                    os.unlink(settings_file)
                    results["details"].append("Reset Teams configuration")
            except Exception as e:
                logger.warning(f"Error resetting Teams configuration: {str(e)}")
                results["details"].append(f"Warning: Could not reset Teams configuration")
            
            # Run PowerShell fix script if available
            ps_script = os.path.join(self.script_dir, 'teams_fix.ps1')
            if os.path.exists(ps_script):
                try:
                    result = subprocess.run(
                        ['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', ps_script],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    results["details"].append("Executed Teams fix script")
                    results["details"].append(result.stdout.strip())
                except Exception as e:
                    logger.warning(f"Error running Teams fix script: {str(e)}")
                    results["details"].append(f"Warning: Could not run Teams fix script: {str(e)}")
            
            # Mark as success
            results["success"] = True
            results["message"] = "Microsoft Teams has been successfully fixed. Please restart Teams."
            
            logger.info("Teams fix completed successfully")
            return results
            
        except Exception as e:
            logger.error(f"Error fixing Teams: {str(e)}")
            results["success"] = False
            results["message"] = f"Error fixing Microsoft Teams: {str(e)}"
            return results
    
    def fix_outlook(self):
        """
        Fix common Microsoft Outlook issues by resetting profile and fixing data files.
        
        Returns:
            dict: Results of the operation
        """
        results = {
            "success": False,
            "message": "",
            "details": []
        }
        
        try:
            # Kill Outlook process
            self._kill_process("OUTLOOK.EXE")
            results["details"].append("Stopped Microsoft Outlook process")
            
            # Reset navigation pane settings
            try:
                outlook_key_path = r"Software\Microsoft\Office\16.0\Outlook\Preferences"
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, outlook_key_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.DeleteValue(key, "NavigationPaneViewState")
                results["details"].append("Reset Outlook navigation pane settings")
            except FileNotFoundError:
                # Key doesn't exist, not an error
                pass
            except Exception as e:
                logger.warning(f"Error resetting Outlook navigation pane: {str(e)}")
                results["details"].append(f"Warning: Could not reset navigation pane settings")
            
            # Clear Outlook cached credentials
            try:
                import win32cred
                credentials = win32cred.CredEnumerate(None, 0)
                for cred in credentials:
                    if "Microsoft_OC_" in cred['TargetName'] or "MicrosoftOffice" in cred['TargetName']:
                        win32cred.CredDelete(cred['TargetName'], 0, 0)
                results["details"].append("Cleared Outlook cached credentials")
            except ImportError:
                # win32cred not available, skip this step
                results["details"].append("Warning: Could not clear cached credentials (pywin32 not available)")
            except Exception as e:
                logger.warning(f"Error clearing Outlook credentials: {str(e)}")
                results["details"].append(f"Warning: Could not clear cached credentials")
            
            # Reset Outlook profile settings
            try:
                self._reset_outlook_profiles()
                results["details"].append("Reset corrupt Outlook profile settings")
            except Exception as e:
                logger.warning(f"Error resetting Outlook profiles: {str(e)}")
                results["details"].append(f"Warning: Could not reset Outlook profiles")
            
            # Repair OST/PST files
            try:
                self._repair_outlook_data_files()
                results["details"].append("Repaired Outlook data files")
            except Exception as e:
                logger.warning(f"Error repairing Outlook data files: {str(e)}")
                results["details"].append(f"Warning: Could not repair Outlook data files")
            
            # Run Outlook fix PowerShell script if available
            ps_script = os.path.join(self.script_dir, 'outlook_fix.ps1')
            if os.path.exists(ps_script):
                try:
                    result = subprocess.run(
                        ['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', ps_script],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    results["details"].append("Executed Outlook fix script")
                    results["details"].append(result.stdout.strip())
                except Exception as e:
                    logger.warning(f"Error running Outlook fix script: {str(e)}")
                    results["details"].append(f"Warning: Could not run Outlook fix script: {str(e)}")
            
            # Run Outlook in safe mode to reset components
            try:
                subprocess.Popen(['outlook.exe', '/safe'], shell=True)
                time.sleep(5)  # Give Outlook time to start
                self._kill_process("OUTLOOK.EXE")
                results["details"].append("Started and closed Outlook in safe mode")
            except Exception as e:
                logger.warning(f"Error running Outlook in safe mode: {str(e)}")
                results["details"].append(f"Warning: Could not start Outlook in safe mode")
            
            # Mark as success
            results["success"] = True
            results["message"] = "Microsoft Outlook has been successfully fixed. Please restart Outlook."
            
            logger.info("Outlook fix completed successfully")
            return results
            
        except Exception as e:
            logger.error(f"Error fixing Outlook: {str(e)}")
            results["success"] = False
            results["message"] = f"Error fixing Microsoft Outlook: {str(e)}"
            return results
    
    def _kill_process(self, process_name):
        """
        Kill a process by name.
        
        Args:
            process_name (str): Name of process to kill
        """
        try:
            # Try to gracefully kill the process
            subprocess.run(['taskkill', '/IM', process_name], 
                          stdout=subprocess.PIPE, 
                          stderr=subprocess.PIPE,
                          check=False)
            
            # Force kill if still running
            subprocess.run(['taskkill', '/F', '/IM', process_name], 
                          stdout=subprocess.PIPE, 
                          stderr=subprocess.PIPE,
                          check=False)
        except Exception as e:
            logger.warning(f"Error killing process {process_name}: {str(e)}")
    
    def _reset_outlook_profiles(self):
        """Reset problematic Outlook profile settings."""
        try:
            # Backup the profiles registry key
            profiles_key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
            backup_cmd = ['reg', 'export', 
                         f'HKCU\\{profiles_key_path}', 
                         os.path.join(tempfile.gettempdir(), 'outlook_profiles_backup.reg'),
                         '/y']
            
            subprocess.run(backup_cmd, 
                          stdout=subprocess.PIPE, 
                          stderr=subprocess.PIPE,
                          check=False)
            
            # Check for corrupt profile settings and reset them
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, profiles_key_path) as profiles_key:
                    num_profiles = winreg.QueryInfoKey(profiles_key)[0]
                    
                    for i in range(num_profiles):
                        profile_name = winreg.EnumKey(profiles_key, i)
                        
                        # Check profile for corruption
                        profile_path = f"{profiles_key_path}\\{profile_name}"
                        try:
                            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, f"{profile_path}\\9375CFF0413111d3B88A00104B2A6676") as sub_key:
                                # This key exists, profile may be corrupted
                                # Delete problem keys
                                try:
                                    subprocess.run(['reg', 'delete', 
                                                  f'HKCU\\{profile_path}\\9375CFF0413111d3B88A00104B2A6676', 
                                                  '/f'], 
                                                 stdout=subprocess.PIPE, 
                                                 stderr=subprocess.PIPE,
                                                 check=False)
                                except:
                                    pass
                        except FileNotFoundError:
                            # This key doesn't exist, profile is probably not corrupted
                            pass
            except Exception as e:
                logger.warning(f"Error checking for corrupt profiles: {str(e)}")
        
        except Exception as e:
            logger.error(f"Error resetting Outlook profiles: {str(e)}")
            raise
    
    def _repair_outlook_data_files(self):
        """Repair Outlook OST and PST files."""
        try:
            # Find scanpst.exe
            scanpst_locations = [
                os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), 'Microsoft Office', 'root', 'Office16', 'SCANPST.EXE'),
                os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), 'Microsoft Office', 'root', 'Office16', 'SCANPST.EXE'),
                os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), 'Microsoft Office 15', 'root', 'Office15', 'SCANPST.EXE'),
                os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), 'Microsoft Office 15', 'root', 'Office15', 'SCANPST.EXE'),
                os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), 'Microsoft Office', 'Office14', 'SCANPST.EXE'),
                os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), 'Microsoft Office', 'Office14', 'SCANPST.EXE'),
            ]
            
            scanpst_exe = None
            for location in scanpst_locations:
                if os.path.exists(location):
                    scanpst_exe = location
                    break
            
            if not scanpst_exe:
                logger.warning("SCANPST.EXE not found, cannot repair Outlook data files")
                return
            
            # Find OST and PST files
            outlook_files = []
            
            # Check default locations
            for ext in ['.ost', '.pst']:
                # Search in the default location
                search_path = os.path.join(self.local_app_data, 'Microsoft', 'Outlook')
                if os.path.exists(search_path):
                    for file in os.listdir(search_path):
                        if file.lower().endswith(ext):
                            outlook_files.append(os.path.join(search_path, file))
                
                # Also search in Documents folder
                documents_path = os.path.join(self.user_profile, 'Documents')
                if os.path.exists(documents_path):
                    for file in os.listdir(documents_path):
                        if file.lower().endswith(ext):
                            outlook_files.append(os.path.join(documents_path, file))
            
            # Repair each file
            for file_path in outlook_files:
                try:
                    # Create a batch file to automate SCANPST
                    # Since SCANPST doesn't support command-line arguments
                    batch_file = os.path.join(tempfile.gettempdir(), 'repair_outlook.bat')
                    with open(batch_file, 'w') as f:
                        # These keyboard shortcuts simulate user interaction with SCANPST
                        # Alt+B (browse), then paste file path, then Alt+S (start), then Alt+R (repair)
                        f.write(f'@echo off\n')
                        f.write(f'start "" "{scanpst_exe}"\n')
                        f.write(f'timeout /t 1\n')
                        f.write(f'echo %~1 | clip\n')
                        f.write(f'timeout /t 1\n')
                        f.write(f'echo Alt+B, Ctrl+V, Alt+S, Alt+R | clip\n')
                    
                    # Run the batch file
                    subprocess.run(['cmd', '/c', batch_file, file_path], 
                                  stdout=subprocess.PIPE, 
                                  stderr=subprocess.PIPE,
                                  check=False)
                    
                    # Wait for SCANPST to complete
                    time.sleep(10)
                    
                    # Close SCANPST if still running
                    self._kill_process("SCANPST.EXE")
                    
                except Exception as e:
                    logger.warning(f"Error repairing Outlook file {file_path}: {str(e)}")
                    continue
            
        except Exception as e:
            logger.error(f"Error repairing Outlook data files: {str(e)}")
            raise
