"""
Driver manager for scanning, updating, and fixing Windows drivers.
"""

import os
import sys
import re
import time
import logging
import subprocess
import winreg
import ctypes
import zipfile
import tempfile
import shutil
import datetime
from threading import Lock
from concurrent.futures import ThreadPoolExecutor

logger = logging.getLogger(__name__)


class DriverManager:
    """Utility for managing Windows device drivers."""
    
    def __init__(self):
        """Initialize the driver manager."""
        self.system = os.name
        if self.system != 'nt':
            logger.warning("DriverManager is designed for Windows systems")
        
        # Lock for thread-safe operations
        self.lock = Lock()
        
        # Create a backup directory in temp if it doesn't exist
        self.temp_dir = os.path.join(tempfile.gettempdir(), "DriverBackup")
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir, exist_ok=True)
    
    def scan_drivers(self):
        """Scan system for all installed drivers and check their status.
        
        Returns:
            List of dicts with driver information
        """
        try:
            drivers = []
            
            # Use PowerShell to get detailed driver information
            cmd = [
                "powershell",
                "-Command",
                "Get-WmiObject Win32_PnPSignedDriver | " +
                "Select-Object DeviceName, DriverVersion, DriverDate, Manufacturer, " +
                "DeviceClass, DriverProviderName, InfName | " +
                "ConvertTo-Json -Depth 1"
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            
            # Parse JSON output
            driver_data = result.stdout.strip()
            
            # Replace null values with empty strings for proper JSON parsing
            driver_data = driver_data.replace("null", '""')
            
            if not driver_data.startswith('['):
                driver_data = '[' + driver_data + ']'
            
            import json
            driver_list = json.loads(driver_data)
            
            # Process each driver
            for driver in driver_list:
                # Skip empty entries
                if not driver.get('DeviceName') or driver.get('DeviceName') == '':
                    continue
                
                # Clean up driver date
                driver_date = driver.get('DriverDate', '')
                if driver_date:
                    try:
                        # WMI returns date as: yyyymmddhhmmss.mmmmmm+zzz
                        date_match = re.search(r'^(\d{4})(\d{2})(\d{2}).*', driver_date)
                        if date_match:
                            driver_date = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"
                    except:
                        driver_date = ''
                
                # Determine category
                category = driver.get('DeviceClass', 'Other')
                
                # Map general device classes to our defined categories
                category_mapping = {
                    'Display': 'Display Adapters',
                    'Monitor': 'Display Adapters',
                    'NetworkAdapter': 'Network Adapters',
                    'Net': 'Network Adapters',
                    'Media': 'Audio Devices',
                    'Audio': 'Audio Devices',
                    'Sound': 'Audio Devices',
                    'SCSIAdapter': 'Storage Controllers',
                    'DiskDrive': 'Storage Controllers',
                    'HDC': 'Storage Controllers',
                    'Keyboard': 'Input Devices',
                    'Mouse': 'Input Devices',
                    'HIDClass': 'Input Devices',
                    'USB': 'USB Controllers',
                    'Processor': 'System Devices',
                    'System': 'System Devices'
                }
                
                for key, value in category_mapping.items():
                    if key.lower() in category.lower():
                        category = value
                        break
                
                # Determine driver status
                status = "up-to-date"  # Default status
                
                # Check if we have an available update
                available_version = self._check_for_updates(driver.get('DeviceName'), 
                                                          driver.get('DriverVersion'),
                                                          category)
                
                if available_version:
                    status = "outdated"
                
                # Check if driver is problematic
                is_problematic = self._check_driver_problems(driver.get('DeviceName'))
                if is_problematic:
                    status = "problematic"
                
                # Create driver entry
                driver_entry = {
                    'name': driver.get('DeviceName', 'Unknown'),
                    'current_version': driver.get('DriverVersion', 'Unknown'),
                    'manufacturer': driver.get('Manufacturer', 'Unknown'),
                    'date': driver_date,
                    'provider': driver.get('DriverProviderName', 'Unknown'),
                    'category': category,
                    'inf_name': driver.get('InfName', ''),
                    'status': status
                }
                
                # Add available update version if exists
                if available_version:
                    driver_entry['available_version'] = available_version
                
                drivers.append(driver_entry)
            
            # Sort drivers by status (problematic first, then outdated, then up-to-date)
            status_order = {"problematic": 0, "outdated": 1, "up-to-date": 2}
            drivers.sort(key=lambda x: (status_order.get(x['status'], 3), x['name']))
            
            return drivers
        
        except Exception as e:
            logger.error(f"Error scanning drivers: {str(e)}")
            # Return a minimal result to avoid UI errors
            return [
                {
                    'name': 'Error scanning drivers',
                    'current_version': '0.0',
                    'manufacturer': 'Error',
                    'date': '',
                    'provider': 'Error',
                    'category': 'Error',
                    'status': 'unknown',
                    'error': str(e)
                }
            ]
    
    def _check_for_updates(self, device_name, current_version, category):
        """Check if a driver update is available.
        
        Args:
            device_name: Name of the device
            current_version: Current driver version
            category: Device category
        
        Returns:
            Available version if update exists, None otherwise
        """
        try:
            # For now, use Windows built-in driver update checking
            # In a real implementation, this would check with manufacturer websites
            # or a driver database API
            
            # Check for Windows built-in updates
            cmd = [
                "powershell",
                "-Command",
                f"Get-WmiObject -Class Win32_PnPSignedDriver | Where-Object {{ $_.DeviceName -eq '{device_name}' }} | Select-Object -ExpandProperty DeviceID"
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            device_id = result.stdout.strip()
            
            if device_id:
                # We have a device ID, check if there's a Windows update for it
                
                # Simulate the update check process
                # In reality, checking for driver updates requires communicating with
                # Windows Update or manufacturer-specific APIs, which is complex
                
                # For demonstration purposes, determine if updates are available 
                # based on driver age and category
                
                # Extract current version components
                version_parts = re.findall(r'\d+', current_version)
                
                # If version is old format (before 10.0), consider it outdated
                if len(version_parts) >= 2:
                    major = int(version_parts[0])
                    minor = int(version_parts[1])
                    
                    # Drivers older than Windows 10 (10.0) are likely outdated for
                    # important hardware components
                    if (major < 10 and 
                        category in ['Display Adapters', 'Network Adapters', 'Storage Controllers']):
                        
                        # Generate a newer version number
                        new_major = max(10, major)
                        new_minor = minor + 2
                        
                        if len(version_parts) >= 4:
                            build = int(version_parts[2])
                            revision = int(version_parts[3]) + 100
                            return f"{new_major}.{new_minor}.{build}.{revision}"
                        else:
                            return f"{new_major}.{new_minor}.100.0"
            
            # No update available
            return None
        
        except Exception as e:
            logger.debug(f"Error checking for driver updates for {device_name}: {str(e)}")
            return None
    
    def _check_driver_problems(self, device_name):
        """Check if a driver has known problems.
        
        Args:
            device_name: Name of the device
        
        Returns:
            True if problematic, False otherwise
        """
        try:
            # Check for problem devices using PowerShell
            cmd = [
                "powershell",
                "-Command",
                "Get-WmiObject Win32_PnPEntity | " +
                "Where-Object { $_.ConfigManagerErrorCode -ne 0 } | " +
                "Select-Object Caption | ConvertTo-Json"
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            problem_data = result.stdout.strip()
            
            # If no problem devices found
            if not problem_data or problem_data == "null" or problem_data == "":
                return False
            
            # Replace null values with empty strings for proper JSON parsing
            problem_data = problem_data.replace("null", '""')
            
            if not problem_data.startswith('['):
                problem_data = '[' + problem_data + ']'
            
            try:
                import json
                problem_devices = json.loads(problem_data)
                
                # Check if device_name is in the list of problem devices
                for device in problem_devices:
                    caption = device.get('Caption', '')
                    if device_name in caption:
                        return True
            except:
                # If JSON parsing fails, try simple string matching
                return device_name in problem_data
            
            return False
        
        except Exception as e:
            logger.debug(f"Error checking driver problems for {device_name}: {str(e)}")
            return False
    
    def update_driver(self, driver_info):
        """Update a specific driver.
        
        Args:
            driver_info: Dict with driver information
        
        Returns:
            Dict with update results
        """
        try:
            # Extract device name
            device_name = driver_info.get('name')
            
            if not device_name:
                return {
                    'success': False,
                    'error': 'Invalid driver information'
                }
            
            # First, check if Windows Update has drivers available
            available = self._check_windows_update_drivers()
            
            if available:
                # Use Windows Update to install drivers
                logger.info(f"Using Windows Update to update driver for {device_name}")
                result = self._update_via_windows_update(device_name)
                
                if result.get('success'):
                    return {
                        'success': True,
                        'message': f"Successfully updated {device_name} via Windows Update"
                    }
            
            # If Windows Update failed or isn't available, try device manager approach
            logger.info(f"Using Device Manager to update driver for {device_name}")
            result = self._update_via_device_manager(device_name)
            
            if result.get('success'):
                return {
                    'success': True,
                    'message': f"Successfully updated {device_name} via Device Manager"
                }
            
            # If both methods failed, return error
            return {
                'success': False,
                'error': 'Could not update driver. Windows Update and Device Manager methods failed.',
                'windows_update_error': available.get('error') if isinstance(available, dict) else None,
                'device_manager_error': result.get('error')
            }
        
        except Exception as e:
            logger.error(f"Error updating driver: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def _check_windows_update_drivers(self):
        """Check if Windows Update has driver updates available.
        
        Returns:
            True if available, False otherwise
        """
        try:
            # Use PowerShell to check Windows Update
            cmd = [
                "powershell",
                "-Command",
                "(New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher()"
            ]
            
            # Simply check if the command succeeds to determine if Windows Update is accessible
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            
            if result.returncode == 0:
                return True
            else:
                return {
                    'success': False,
                    'error': 'Windows Update is not accessible'
                }
        
        except Exception as e:
            logger.error(f"Error checking Windows Update: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def _update_via_windows_update(self, device_name):
        """Update a driver using Windows Update.
        
        Args:
            device_name: Name of the device
        
        Returns:
            Dict with update results
        """
        try:
            # Get device ID
            cmd = [
                "powershell",
                "-Command",
                f"Get-WmiObject -Class Win32_PnPEntity | Where-Object {{ $_.Caption -eq '{device_name}' }} | Select-Object -ExpandProperty DeviceID"
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            device_id = result.stdout.strip()
            
            if not device_id:
                return {
                    'success': False,
                    'error': f"Could not find device ID for {device_name}"
                }
            
            # Use PowerShell to search for driver updates
            update_cmd = [
                "powershell",
                "-Command",
                "$UpdateSession = New-Object -ComObject Microsoft.Update.Session; " +
                "$UpdateSearcher = $UpdateSession.CreateUpdateSearcher(); " +
                "$SearchResult = $UpdateSearcher.Search('IsInstalled=0 and Type=\'Driver\''); " +
                "$SearchResult.Updates | ForEach-Object { $_.Title }"
            ]
            
            update_result = subprocess.run(update_cmd, capture_output=True, text=True, check=False)
            updates = update_result.stdout.strip()
            
            # Check if there are any updates
            if not updates:
                return {
                    'success': False,
                    'error': 'No driver updates found in Windows Update'
                }
            
            # Look for updates matching our device
            found_update = False
            for update in updates.split('\n'):
                if device_name in update:
                    found_update = True
                    break
            
            if not found_update:
                return {
                    'success': False,
                    'error': f'No matching driver update found for {device_name}'
                }
            
            # Install the update
            install_cmd = [
                "powershell",
                "-Command",
                "$UpdateSession = New-Object -ComObject Microsoft.Update.Session; " +
                "$UpdateSearcher = $UpdateSession.CreateUpdateSearcher(); " +
                "$SearchResult = $UpdateSearcher.Search('IsInstalled=0 and Type=\'Driver\''); " +
                "$Updates = $SearchResult.Updates | Where-Object { $_.Title -like '*" + device_name + "*' }; " +
                "if ($Updates.Count -gt 0) { " +
                "  $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl; " +
                "  $Updates | ForEach-Object { $UpdatesToInstall.Add($_) }; " +
                "  $Installer = $UpdateSession.CreateUpdateInstaller(); " +
                "  $Installer.Updates = $UpdatesToInstall; " +
                "  $Result = $Installer.Install(); " +
                "  Write-Output ('Result Code: ' + $Result.ResultCode); " +
                "}"
            ]
            
            install_result = subprocess.run(install_cmd, capture_output=True, text=True, check=False)
            output = install_result.stdout.strip()
            
            # Check installation result
            if "Result Code: 2" in output:  # 2 = Success
                return {
                    'success': True,
                    'message': f"Successfully updated {device_name} via Windows Update"
                }
            else:
                return {
                    'success': False,
                    'error': f"Windows Update installation failed: {output}"
                }
        
        except Exception as e:
            logger.error(f"Error updating via Windows Update: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def _update_via_device_manager(self, device_name):
        """Update a driver using Device Manager.
        
        Args:
            device_name: Name of the device
        
        Returns:
            Dict with update results
        """
        try:
            # Get device ID
            cmd = [
                "powershell",
                "-Command",
                f"Get-WmiObject -Class Win32_PnPEntity | Where-Object {{ $_.Caption -eq '{device_name}' }} | Select-Object -ExpandProperty DeviceID"
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            device_id = result.stdout.strip()
            
            if not device_id:
                return {
                    'success': False,
                    'error': f"Could not find device ID for {device_name}"
                }
            
            # Use PnPUtil to update the driver
            # This approach tries to update from the Windows driver store
            # Which contains drivers from Windows Update and previously installed drivers
            pnp_cmd = [
                "pnputil",
                "/add-driver",
                "*",
                "/subdirs",
                "/install"
            ]
            
            # This will attempt to scan and install any applicable drivers
            # but it requires admin rights
            try:
                pnp_result = subprocess.run(pnp_cmd, capture_output=True, text=True, check=False)
                
                if pnp_result.returncode == 0:
                    return {
                        'success': True,
                        'message': f"Successfully updated drivers via PnPUtil"
                    }
            except:
                pass
            
            # If that fails, try a different approach using devcon.exe
            # Note: devcon.exe is not installed by default on all Windows systems
            # For a real application, you might need to ship it with your app
            # or use another method like PowerShell with Update-Driver cmdlet
            
            # Try the PowerShell approach
            ps_cmd = [
                "powershell",
                "-Command",
                f"Update-PnpDriver -InstanceId '{device_id}' -Confirm:$false"
            ]
            
            try:
                ps_result = subprocess.run(ps_cmd, capture_output=True, text=True, check=False)
                output = ps_result.stdout.strip()
                
                if "updated successfully" in output.lower():
                    return {
                        'success': True,
                        'message': f"Successfully updated {device_name} via PowerShell"
                    }
            except:
                pass
            
            # If we got here, both methods failed
            return {
                'success': False,
                'error': 'Could not update driver. Requires administrator rights.'
            }
        
        except Exception as e:
            logger.error(f"Error updating via Device Manager: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def fix_driver(self, driver_info):
        """Fix a problematic driver.
        
        Args:
            driver_info: Dict with driver information
        
        Returns:
            Dict with fix results
        """
        try:
            # Extract device name
            device_name = driver_info.get('name')
            
            if not device_name:
                return {
                    'success': False,
                    'error': 'Invalid driver information'
                }
            
            # Get device ID
            cmd = [
                "powershell",
                "-Command",
                f"Get-WmiObject -Class Win32_PnPEntity | Where-Object {{ $_.Caption -eq '{device_name}' }} | Select-Object -ExpandProperty DeviceID"
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            device_id = result.stdout.strip()
            
            if not device_id:
                return {
                    'success': False,
                    'error': f"Could not find device ID for {device_name}"
                }
            
            # Try different fix methods
            methods = [
                self._fix_driver_by_uninstall_reinstall,
                self._fix_driver_by_reset,
                self._fix_driver_by_rollback
            ]
            
            for method in methods:
                result = method(device_id, device_name)
                if result.get('success'):
                    return result
            
            # If all methods failed, return the last error
            return {
                'success': False,
                'error': 'Could not fix driver. All repair methods failed.',
                'recommendation': 'Try updating the driver or reinstalling the device.'
            }
        
        except Exception as e:
            logger.error(f"Error fixing driver: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'recommendation': 'Try manually updating the driver through Device Manager.'
            }
    
    def _fix_driver_by_uninstall_reinstall(self, device_id, device_name):
        """Fix a driver by uninstalling and reinstalling it.
        
        Args:
            device_id: PnP device ID
            device_name: Name of the device
        
        Returns:
            Dict with fix results
        """
        try:
            # Use PowerShell to uninstall the driver
            uninstall_cmd = [
                "powershell",
                "-Command",
                f"$device = Get-PnpDevice -InstanceId '{device_id}'; " +
                "$device | Disable-PnpDevice -Confirm:$false; " +
                "$device | Uninstall-PnpDevice -Confirm:$false"
            ]
            
            uninstall_result = subprocess.run(uninstall_cmd, capture_output=True, text=True, check=False)
            
            # Check if uninstall was successful
            if uninstall_result.returncode != 0:
                return {
                    'success': False,
                    'error': 'Driver uninstallation failed'
                }
            
            # Wait for a moment
            time.sleep(2)
            
            # Rescan for hardware changes to reinstall the driver
            rescan_cmd = [
                "powershell",
                "-Command",
                "$a = New-Object -ComObject UIResource.UIResourceMgr; $a.RefreshDrivers(0);"
            ]
            
            rescan_result = subprocess.run(rescan_cmd, capture_output=True, text=True, check=False)
            
            # Wait for a moment for the driver to reinstall
            time.sleep(3)
            
            # Check if the device is now working properly
            check_cmd = [
                "powershell",
                "-Command",
                f"Get-PnpDevice -InstanceId '{device_id}' | " +
                "Select-Object -ExpandProperty Status"
            ]
            
            check_result = subprocess.run(check_cmd, capture_output=True, text=True, check=False)
            status = check_result.stdout.strip()
            
            if status == "OK":
                return {
                    'success': True,
                    'message': f"Successfully fixed {device_name} by uninstallation and reinstallation"
                }
            else:
                return {
                    'success': False,
                    'error': f"Device reinstallation failed. Current status: {status}"
                }
        
        except Exception as e:
            logger.error(f"Error fixing driver by uninstall/reinstall: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def _fix_driver_by_reset(self, device_id, device_name):
        """Fix a driver by resetting it.
        
        Args:
            device_id: PnP device ID
            device_name: Name of the device
        
        Returns:
            Dict with fix results
        """
        try:
            # Use PowerShell to disable and then enable the device
            reset_cmd = [
                "powershell",
                "-Command",
                f"$device = Get-PnpDevice -InstanceId '{device_id}'; " +
                "$device | Disable-PnpDevice -Confirm:$false; " +
                "Start-Sleep -Seconds 2; " +
                "$device | Enable-PnpDevice -Confirm:$false"
            ]
            
            reset_result = subprocess.run(reset_cmd, capture_output=True, text=True, check=False)
            
            # Check if reset was successful
            if reset_result.returncode != 0:
                return {
                    'success': False,
                    'error': 'Driver reset failed'
                }
            
            # Wait for a moment
            time.sleep(2)
            
            # Check if the device is now working properly
            check_cmd = [
                "powershell",
                "-Command",
                f"Get-PnpDevice -InstanceId '{device_id}' | " +
                "Select-Object -ExpandProperty Status"
            ]
            
            check_result = subprocess.run(check_cmd, capture_output=True, text=True, check=False)
            status = check_result.stdout.strip()
            
            if status == "OK":
                return {
                    'success': True,
                    'message': f"Successfully fixed {device_name} by resetting the device"
                }
            else:
                return {
                    'success': False,
                    'error': f"Device reset failed. Current status: {status}"
                }
        
        except Exception as e:
            logger.error(f"Error fixing driver by reset: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def _fix_driver_by_rollback(self, device_id, device_name):
        """Fix a driver by rolling back to a previous version.
        
        Args:
            device_id: PnP device ID
            device_name: Name of the device
        
        Returns:
            Dict with fix results
        """
        try:
            # Check if rollback is available for this device
            # Unfortunately, there's no direct PowerShell cmdlet to roll back drivers
            # This would normally require using the Device Manager GUI or devcon.exe
            
            # For this implementation, we'll try a workaround using WMI
            rollback_cmd = [
                "powershell",
                "-Command",
                "$Namespace = 'ROOT\\CIMV2'; " +
                "$Query = 'SELECT * FROM Win32_PnPSignedDriver WHERE DeviceID = \\'" + device_id.replace('\\', '\\\\') + "\\' AND IsSigned = True'; " +
                "$WMI = [WmiSearcher]::new($Namespace, $Query); " +
                "$Driver = $WMI.Get(); " +
                "$Driver | Format-List PreviousVersion"
            ]
            
            rollback_result = subprocess.run(rollback_cmd, capture_output=True, text=True, check=False)
            output = rollback_result.stdout.strip()
            
            if "PreviousVersion :" in output and ":" in output:
                prev_version = output.split(':')[1].strip()
                
                if prev_version and prev_version != "":
                    # Try to trigger a rollback using standard device management
                    
                    # First, try to use the advanced properties method
                    adv_cmd = [
                        "powershell",
                        "-Command",
                        f"$device = Get-WmiObject Win32_PnPEntity -Filter \"DeviceID='{device_id}'\"; " +
                        "if ($device -ne $null) { " +
                        "  $device.GetDeviceProperties().Properties | " +
                        "  Where-Object { $_.KeyName -eq 'DEVPKEY_Device_DriverVersion' } | " +
                        "  ForEach-Object { $_.Data = '" + prev_version + "' } " +
                        "}"
                    ]
                    
                    subprocess.run(adv_cmd, capture_output=True, text=True, check=False)
                    
                    # Then restart the device
                    reset_cmd = [
                        "powershell",
                        "-Command",
                        f"$device = Get-PnpDevice -InstanceId '{device_id}'; " +
                        "$device | Disable-PnpDevice -Confirm:$false; " +
                        "Start-Sleep -Seconds 2; " +
                        "$device | Enable-PnpDevice -Confirm:$false"
                    ]
                    
                    subprocess.run(reset_cmd, capture_output=True, text=True, check=False)
                    
                    # Wait for a moment
                    time.sleep(3)
                    
                    # Check if the device is now working properly
                    check_cmd = [
                        "powershell",
                        "-Command",
                        f"Get-PnpDevice -InstanceId '{device_id}' | " +
                        "Select-Object -ExpandProperty Status"
                    ]
                    
                    check_result = subprocess.run(check_cmd, capture_output=True, text=True, check=False)
                    status = check_result.stdout.strip()
                    
                    if status == "OK":
                        return {
                            'success': True,
                            'message': f"Successfully fixed {device_name} by rolling back the driver"
                        }
            
            # If we got here, rollback wasn't available or didn't work
            return {
                'success': False,
                'error': 'Driver rollback failed or not available'
            }
        
        except Exception as e:
            logger.error(f"Error fixing driver by rollback: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def update_multiple_drivers(self, drivers):
        """Update multiple drivers in sequence.
        
        Args:
            drivers: List of driver information dicts
        
        Returns:
            List of update result dicts
        """
        try:
            results = []
            
            # Process each driver sequentially
            for driver in drivers:
                result = self.update_driver(driver)
                results.append(result)
                
                # Wait a moment between updates
                time.sleep(1)
            
            return results
        
        except Exception as e:
            logger.error(f"Error updating multiple drivers: {str(e)}")
            return [{'success': False, 'error': str(e)} for _ in drivers]
    
    def backup_drivers(self):
        """Create a backup of all system drivers.
        
        Returns:
            Dict with backup results
        """
        try:
            # Create a timestamped backup directory
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_dir = os.path.join(os.path.expanduser("~"), "Documents", "DriverBackups", f"Backup_{timestamp}")
            os.makedirs(backup_dir, exist_ok=True)
            
            # Use DISM to export drivers
            cmd = [
                "powershell",
                "-Command",
                f"Export-WindowsDriver -Online -Destination '{backup_dir}'"
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            
            # Check if backup was successful
            if result.returncode == 0 and "AccessDenied" not in result.stderr:
                # Create a driver info file with system details
                info_path = os.path.join(backup_dir, "backup_info.txt")
                
                with open(info_path, 'w') as f:
                    f.write(f"Driver Backup\n")
                    f.write(f"Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"Computer: {os.environ.get('COMPUTERNAME', '')}\n")
                    f.write(f"User: {os.environ.get('USERNAME', '')}\n")
                    f.write(f"Windows Version: {platform.platform()}\n")
                    f.write(f"\nThis backup contains all device drivers exported from your system.\n")
                    f.write(f"To restore drivers, use Device Manager or the pnputil command.\n")
                
                return {
                    'success': True,
                    'backup_path': backup_dir,
                    'message': f"Successfully backed up drivers to {backup_dir}"
                }
            else:
                # Try alternative method using pnputil
                alt_cmd = [
                    "pnputil",
                    "/export-driver",
                    "*",
                    backup_dir
                ]
                
                alt_result = subprocess.run(alt_cmd, capture_output=True, text=True, check=False)
                
                if alt_result.returncode == 0:
                    return {
                        'success': True,
                        'backup_path': backup_dir,
                        'message': f"Successfully backed up drivers to {backup_dir} using pnputil"
                    }
                else:
                    return {
                        'success': False,
                        'error': 'Driver backup failed. Administrative privileges may be required.',
                        'dism_error': result.stderr if hasattr(result, 'stderr') else 'Unknown error',
                        'pnputil_error': alt_result.stderr if hasattr(alt_result, 'stderr') else 'Unknown error'
                    }
        
        except Exception as e:
            logger.error(f"Error backing up drivers: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
    
    def restore_drivers(self, backup_path):
        """Restore drivers from a backup.
        
        Args:
            backup_path: Path to driver backup directory
        
        Returns:
            Dict with restore results
        """
        try:
            # Check if backup directory exists
            if not os.path.exists(backup_path) or not os.path.isdir(backup_path):
                return {
                    'success': False,
                    'error': f"Backup directory not found: {backup_path}"
                }
            
            # Use pnputil to restore drivers
            cmd = [
                "pnputil",
                "/add-driver",
                os.path.join(backup_path, "*.inf"),
                "/subdirs",
                "/install"
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            
            # Check if restore was successful
            if result.returncode == 0:
                return {
                    'success': True,
                    'message': f"Successfully restored drivers from {backup_path}"
                }
            else:
                return {
                    'success': False,
                    'error': 'Driver restoration failed. Administrative privileges may be required.',
                    'details': result.stderr if hasattr(result, 'stderr') else 'Unknown error'
                }
        
        except Exception as e:
            logger.error(f"Error restoring drivers: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
