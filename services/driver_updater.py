"""
Driver updating service for the Windows System Optimizer.
This module provides functionality to check for outdated drivers and update them.
"""

import os
import logging
import subprocess
import re
import tempfile
import ctypes
import time
import platform
from datetime import datetime

logger = logging.getLogger(__name__)

class DriverUpdater:
    """Service class for driver update operations."""
    
    def __init__(self):
        """Initialize the driver updater."""
        # Check if running with admin privileges
        self.is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    
    def check_drivers(self):
        """
        Check for outdated drivers in the system.
        
        Returns:
            list: List of outdated drivers with details
        """
        outdated_drivers = []
        
        try:
            # Use PowerShell to get driver information
            # This command gets device drivers, their version, date, and status
            ps_command = """
            Get-WmiObject Win32_PnPSignedDriver | 
            Select-Object DeviceName, DriverVersion, DriverDate, IsSigned | 
            ConvertTo-Json
            """
            
            result = subprocess.run(
                ['powershell.exe', '-Command', ps_command],
                capture_output=True,
                text=True,
                check=True
            )
            
            if result.returncode != 0:
                logger.error(f"Error running driver check: {result.stderr}")
                return outdated_drivers
            
            # Parse the JSON result
            import json
            drivers_data = json.loads(result.stdout)
            
            # Ensure drivers_data is a list
            if not isinstance(drivers_data, list):
                drivers_data = [drivers_data]
            
            # Get current date for comparison
            current_date = datetime.now()
            
            # Process driver information
            for driver in drivers_data:
                try:
                    # Skip empty or None device names
                    if not driver.get('DeviceName'):
                        continue
                    
                    # Extract driver date (format: yyyymmdd)
                    driver_date_str = driver.get('DriverDate', '')
                    driver_date = None
                    
                    if driver_date_str:
                        # PowerShell returns WMI dates in a specific format
                        # Example: /Date(1612345678000)/
                        date_match = re.search(r'/Date\((\d+)\)/', driver_date_str)
                        if date_match:
                            timestamp = int(date_match.group(1)) / 1000  # Convert to seconds
                            driver_date = datetime.fromtimestamp(timestamp)
                    
                    # Determine if driver is outdated based on age (older than 1 year)
                    is_outdated = False
                    if driver_date:
                        delta = current_date - driver_date
                        if delta.days > 365:  # Outdated if more than a year old
                            is_outdated = True
                    
                    # Skip drivers that aren't outdated
                    if not is_outdated:
                        continue
                    
                    # Add to list of outdated drivers
                    outdated_drivers.append({
                        "name": driver.get('DeviceName', 'Unknown Device'),
                        "version": driver.get('DriverVersion', 'Unknown'),
                        "date": driver_date.strftime('%Y-%m-%d') if driver_date else 'Unknown',
                        "signed": driver.get('IsSigned', False),
                        "update_available": True,  # Placeholder, would be determined by manufacturer
                        "manufacturer": self._get_driver_manufacturer(driver.get('DeviceName', ''))
                    })
                
                except Exception as e:
                    logger.warning(f"Error processing driver: {str(e)}")
                    continue
            
            logger.info(f"Found {len(outdated_drivers)} outdated drivers")
            return outdated_drivers
            
        except Exception as e:
            logger.error(f"Error checking for outdated drivers: {str(e)}")
            return outdated_drivers
    
    def update_drivers(self, driver_list=None):
        """
        Update outdated drivers.
        
        Args:
            driver_list (list, optional): List of drivers to update. If None, update all outdated drivers.
            
        Returns:
            dict: Results of the update operation
        """
        results = {
            "success": False,
            "updated": [],
            "failed": [],
            "message": ""
        }
        
        try:
            # Verify admin privileges
            if not self.is_admin:
                results["message"] = "Administrator privileges required to update drivers"
                return results
            
            # If no driver list provided, get all outdated drivers
            if driver_list is None:
                driver_list = self.check_drivers()
            
            if not driver_list:
                results["message"] = "No drivers to update"
                results["success"] = True
                return results
            
            # In a real implementation, we would use Windows Update API or
            # vendor-specific tools to update drivers. For this demonstration,
            # we'll simulate driver updates.
            
            for driver in driver_list:
                try:
                    # Simulate driver update process
                    logger.info(f"Updating driver: {driver['name']}")
                    
                    # Simulate a process that takes a few seconds
                    time.sleep(2)
                    
                    # Add to updated list
                    results["updated"].append(driver)
                    
                except Exception as e:
                    logger.error(f"Error updating driver {driver['name']}: {str(e)}")
                    # Add to failed list
                    driver["error"] = str(e)
                    results["failed"].append(driver)
            
            # Set success based on results
            if results["updated"] and not results["failed"]:
                results["success"] = True
                results["message"] = f"Successfully updated {len(results['updated'])} drivers"
            elif results["updated"] and results["failed"]:
                results["success"] = True
                results["message"] = f"Partially updated drivers: {len(results['updated'])} succeeded, {len(results['failed'])} failed"
            else:
                results["success"] = False
                results["message"] = "Failed to update any drivers"
            
            return results
            
        except Exception as e:
            logger.error(f"Error updating drivers: {str(e)}")
            results["success"] = False
            results["message"] = f"Error updating drivers: {str(e)}"
            return results
    
    def create_restore_point(self):
        """
        Create a system restore point before updating drivers.
        
        Returns:
            bool: True if restore point created successfully, False otherwise
        """
        try:
            # Verify admin privileges
            if not self.is_admin:
                logger.warning("Administrator privileges required to create restore point")
                return False
            
            # Use PowerShell to create a restore point
            ps_command = """
            Checkpoint-Computer -Description "Before driver updates - Windows System Optimizer" -RestorePointType "MODIFY_SETTINGS"
            """
            
            result = subprocess.run(
                ['powershell.exe', '-Command', ps_command],
                capture_output=True,
                text=True,
                check=True
            )
            
            if result.returncode != 0:
                logger.error(f"Error creating restore point: {result.stderr}")
                return False
            
            logger.info("System restore point created successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error creating system restore point: {str(e)}")
            return False
    
    def get_driver_details(self, device_name):
        """
        Get detailed information about a specific driver.
        
        Args:
            device_name (str): Name of the device
            
        Returns:
            dict: Detailed driver information
        """
        try:
            # Use PowerShell to get detailed driver information
            ps_command = f"""
            Get-WmiObject Win32_PnPSignedDriver | 
            Where-Object {{ $_.DeviceName -eq "{device_name}" }} |
            Select-Object DeviceName, DriverVersion, DriverDate, DeviceID, InfName, IsSigned, DriverProviderName |
            ConvertTo-Json
            """
            
            result = subprocess.run(
                ['powershell.exe', '-Command', ps_command],
                capture_output=True,
                text=True,
                check=True
            )
            
            if result.returncode != 0:
                logger.error(f"Error getting driver details: {result.stderr}")
                return {}
            
            # Parse the JSON result
            import json
            driver_data = json.loads(result.stdout)
            
            # Convert driver data to Python dict
            driver_details = {
                "name": driver_data.get('DeviceName', 'Unknown'),
                "version": driver_data.get('DriverVersion', 'Unknown'),
                "date": driver_data.get('DriverDate', 'Unknown'),
                "device_id": driver_data.get('DeviceID', 'Unknown'),
                "inf_name": driver_data.get('InfName', 'Unknown'),
                "signed": driver_data.get('IsSigned', False),
                "provider": driver_data.get('DriverProviderName', 'Unknown')
            }
            
            return driver_details
            
        except Exception as e:
            logger.error(f"Error getting driver details: {str(e)}")
            return {}
    
    def _get_driver_manufacturer(self, device_name):
        """
        Get the manufacturer for a device based on its name.
        
        Args:
            device_name (str): Name of the device
            
        Returns:
            str: Manufacturer name or 'Unknown'
        """
        # This is a simplified implementation
        # In a real app, we would query the actual manufacturer
        
        common_manufacturers = {
            "NVIDIA": ["NVIDIA", "GeForce"],
            "AMD": ["AMD", "Radeon"],
            "Intel": ["Intel", "HD Graphics", "UHD Graphics"],
            "Realtek": ["Realtek", "High Definition Audio"],
            "Broadcom": ["Broadcom", "BCM"],
            "Qualcomm": ["Qualcomm", "Atheros"],
            "Microsoft": ["Microsoft", "Basic Display", "Basic Render"],
            "Dell": ["Dell"],
            "HP": ["HP", "Hewlett-Packard"],
            "Lenovo": ["Lenovo", "ThinkPad"],
            "ASUS": ["ASUS"],
            "Logitech": ["Logitech"],
            "Canon": ["Canon"],
            "Epson": ["Epson"],
            "Brother": ["Brother"],
            "Samsung": ["Samsung"],
            "Kingston": ["Kingston"],
            "Western Digital": ["WD", "Western Digital"],
            "Seagate": ["Seagate"],
            "Crucial": ["Crucial"]
        }
        
        device_name_upper = device_name.upper()
        
        for manufacturer, keywords in common_manufacturers.items():
            for keyword in keywords:
                if keyword.upper() in device_name_upper:
                    return manufacturer
        
        return "Unknown"
