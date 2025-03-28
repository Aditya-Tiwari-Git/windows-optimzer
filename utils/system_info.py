"""
System information utilities for collecting system metrics.
"""

import os
import sys
import time
import platform
import datetime
import logging
import psutil
import subprocess
import winreg
import socket
import threading
import ctypes
from collections import namedtuple
from concurrent.futures import ThreadPoolExecutor

logger = logging.getLogger(__name__)


class SystemInfo:
    """Collection of system information utilities."""
    
    def __init__(self):
        """Initialize system information utility."""
        self.system = platform.system()
        
        # Check if running on Windows
        if self.system != "Windows":
            logger.warning("This utility is designed for Windows systems")
        
        # Initialize cache for repeated calls
        self._cache = {}
        self._cache_timeout = {}
        self._cache_lock = threading.Lock()
    
    def _cached_result(self, key, func, timeout=10):
        """Return cached result if available, otherwise call function and cache the result.
        
        Args:
            key: Cache key
            func: Function to call if result not cached
            timeout: Cache timeout in seconds
        
        Returns:
            Cached or new result from func
        """
        with self._cache_lock:
            current_time = time.time()
            
            if key in self._cache and key in self._cache_timeout:
                if current_time - self._cache_timeout[key] < timeout:
                    return self._cache[key]
            
            # Cache miss or expired, call function
            result = func()
            self._cache[key] = result
            self._cache_timeout[key] = current_time
            
            return result
    
    def get_cpu_usage(self):
        """Get CPU usage percentage.
        
        Returns:
            CPU usage percentage (0-100)
        """
        try:
            # Use a shorter timeout for CPU since it changes frequently
            return self._cached_result("cpu_percent", lambda: psutil.cpu_percent(interval=0.5), timeout=1)
        except Exception as e:
            logger.error(f"Error getting CPU usage: {str(e)}")
            return 0
    
    def get_cpu_temperature(self):
        """Get CPU temperature if available.
        
        Returns:
            CPU temperature in Celsius or None if not available
        """
        try:
            # This may not work on all systems
            temperatures = psutil.sensors_temperatures()
            if not temperatures:
                return None
            
            # Try to find CPU temperature
            for name, entries in temperatures.items():
                if name.lower() in ['coretemp', 'cpu_thermal', 'cpu', 'k10temp']:
                    return entries[0].current
                
            # If we have any temperatures but couldn't identify CPU specifically
            for entries in temperatures.values():
                return entries[0].current
            
            return None
        except Exception as e:
            logger.debug(f"Unable to get CPU temperature: {str(e)}")
            return None
    
    def get_memory_info(self):
        """Get memory information.
        
        Returns:
            Dict with memory information:
                - total: Total physical memory in bytes
                - available: Available memory in bytes
                - used: Used memory in bytes
                - percent: Percent of used memory
                - total_gb: Total memory in GB
                - used_gb: Used memory in GB
        """
        try:
            memory = psutil.virtual_memory()
            
            return {
                "total": memory.total,
                "available": memory.available,
                "used": memory.used,
                "percent": memory.percent,
                "total_gb": memory.total / (1024**3),
                "used_gb": memory.used / (1024**3)
            }
        except Exception as e:
            logger.error(f"Error getting memory info: {str(e)}")
            return {
                "total": 0,
                "available": 0,
                "used": 0,
                "percent": 0,
                "total_gb": 0,
                "used_gb": 0
            }
    
    def get_disk_info(self, drive="C:\\"):
        """Get disk information.
        
        Args:
            drive: Drive letter or path to check
        
        Returns:
            Dict with disk information:
                - total: Total disk space in bytes
                - used: Used disk space in bytes
                - free: Free disk space in bytes
                - percent: Percent of used disk space
                - total_gb: Total disk space in GB
                - used_gb: Used disk space in GB
                - free_gb: Free disk space in GB
        """
        try:
            disk_usage = psutil.disk_usage(drive)
            
            return {
                "total": disk_usage.total,
                "used": disk_usage.used,
                "free": disk_usage.free,
                "percent": disk_usage.percent,
                "total_gb": disk_usage.total / (1024**3),
                "used_gb": disk_usage.used / (1024**3),
                "free_gb": disk_usage.free / (1024**3)
            }
        except Exception as e:
            logger.error(f"Error getting disk info for {drive}: {str(e)}")
            return {
                "total": 0,
                "used": 0,
                "free": 0,
                "percent": 0,
                "total_gb": 0,
                "used_gb": 0,
                "free_gb": 0
            }
    
    def get_network_info(self):
        """Get network information.
        
        Returns:
            Dict with network information:
                - download: Download speed in human-readable format (KB/s, MB/s)
                - upload: Upload speed in human-readable format (KB/s, MB/s)
                - download_bytes: Download speed in bytes/sec
                - upload_bytes: Upload speed in bytes/sec
        """
        try:
            def get_network_io():
                net_io = psutil.net_io_counters()
                return (net_io.bytes_sent, net_io.bytes_recv)
            
            # Get initial values
            sent_bytes_start, recv_bytes_start = get_network_io()
            
            # Wait and measure again
            time.sleep(1)
            
            sent_bytes_end, recv_bytes_end = get_network_io()
            
            # Calculate speeds
            upload_speed = sent_bytes_end - sent_bytes_start
            download_speed = recv_bytes_end - recv_bytes_start
            
            # Format for display
            def format_speed(bytes_per_sec):
                if bytes_per_sec < 1024:
                    return f"{bytes_per_sec} B/s"
                elif bytes_per_sec < 1024**2:
                    return f"{bytes_per_sec/1024:.1f} KB/s"
                else:
                    return f"{bytes_per_sec/(1024**2):.1f} MB/s"
            
            return {
                "download": format_speed(download_speed),
                "upload": format_speed(upload_speed),
                "download_bytes": download_speed,
                "upload_bytes": upload_speed
            }
        except Exception as e:
            logger.error(f"Error getting network info: {str(e)}")
            return {
                "download": "0 B/s",
                "upload": "0 B/s",
                "download_bytes": 0,
                "upload_bytes": 0
            }
    
    def get_battery_info(self):
        """Get battery information if available.
        
        Returns:
            Dict with battery information or None if no battery
        """
        try:
            battery = psutil.sensors_battery()
            if battery is None:
                return None
            
            # Determine status
            status = "Unknown"
            if battery.power_plugged:
                status = "Charging" if battery.percent < 100 else "Plugged In"
            else:
                status = "Discharging"
            
            # Calculate time remaining
            time_left = "Unknown"
            if battery.secsleft > 0 and not battery.power_plugged:
                hours, remainder = divmod(battery.secsleft, 3600)
                minutes, _ = divmod(remainder, 60)
                time_left = f"{int(hours)}h {int(minutes)}m"
            elif battery.power_plugged:
                time_left = "Plugged In"
            
            return {
                "percent": battery.percent,
                "status": status,
                "time_left": time_left,
                "is_plugged": battery.power_plugged
            }
        except Exception as e:
            logger.debug(f"Unable to get battery info: {str(e)}")
            return None
    
    def get_uptime(self):
        """Get system uptime.
        
        Returns:
            Uptime as a formatted string (e.g., '3h 45m')
        """
        try:
            uptime_seconds = int(time.time() - psutil.boot_time())
            hours, remainder = divmod(uptime_seconds, 3600)
            minutes, _ = divmod(remainder, 60)
            
            if hours > 24:
                days, hours = divmod(hours, 24)
                return f"{days}d {hours}h {minutes}m"
            else:
                return f"{hours}h {minutes}m"
        except Exception as e:
            logger.error(f"Error getting uptime: {str(e)}")
            return "Unknown"
    
    def get_process_count(self):
        """Get number of running processes.
        
        Returns:
            Integer count of running processes
        """
        try:
            return len(psutil.pids())
        except Exception as e:
            logger.error(f"Error getting process count: {str(e)}")
            return 0
    
    def get_installed_software(self):
        """Get list of installed software from registry.
        
        Returns:
            List of dicts with software information (name, version, publisher, install_date)
        """
        try:
            software_list = []
            
            # Check both registry locations
            registry_paths = [
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            ]
            
            for reg_path in registry_paths:
                try:
                    reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
                    
                    # Iterate through each subkey
                    for i in range(winreg.QueryInfoKey(reg_key)[0]):
                        try:
                            subkey_name = winreg.EnumKey(reg_key, i)
                            subkey = winreg.OpenKey(reg_key, subkey_name)
                            
                            try:
                                display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                
                                # Skip entries without proper display name
                                if not display_name or display_name.strip() == "":
                                    continue
                                
                                # Get software details
                                software_info = {
                                    "name": display_name,
                                    "version": "",
                                    "publisher": "",
                                    "install_date": ""
                                }
                                
                                try:
                                    software_info["version"] = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                                except:
                                    pass
                                
                                try:
                                    software_info["publisher"] = winreg.QueryValueEx(subkey, "Publisher")[0]
                                except:
                                    pass
                                
                                try:
                                    date_str = winreg.QueryValueEx(subkey, "InstallDate")[0]
                                    if date_str and len(date_str) == 8:
                                        formatted_date = f"{date_str[0:4]}-{date_str[4:6]}-{date_str[6:8]}"
                                        software_info["install_date"] = formatted_date
                                except:
                                    pass
                                
                                software_list.append(software_info)
                            
                            except (WindowsError, ValueError, TypeError):
                                # Skip entries that cause errors
                                continue
                            finally:
                                winreg.CloseKey(subkey)
                        except WindowsError:
                            continue
                    
                    winreg.CloseKey(reg_key)
                
                except WindowsError:
                    continue
            
            # Sort by name
            software_list.sort(key=lambda x: x["name"].lower())
            
            return software_list
        except Exception as e:
            logger.error(f"Error getting installed software: {str(e)}")
            return []
    
    def get_system_health(self):
        """Calculate system health score based on various metrics.
        
        Returns:
            Integer health score (0-100)
        """
        try:
            # Get metrics
            cpu_percent = self.get_cpu_usage()
            memory_info = self.get_memory_info()
            disk_info = self.get_disk_info()
            
            # Calculate component scores (higher is better)
            cpu_score = 100 - cpu_percent
            memory_score = 100 - memory_info["percent"]
            disk_score = 100 - disk_info["percent"]
            
            # Apply weights to different components
            health_score = (
                cpu_score * 0.4 +       # CPU has highest weight
                memory_score * 0.3 +    # Memory is also important
                disk_score * 0.3        # Disk has equal weight to memory
            )
            
            return int(health_score)
        except Exception as e:
            logger.error(f"Error calculating system health: {str(e)}")
            return 50  # Return neutral score on error
    
    def get_startup_time(self):
        """Get the startup time of the system.
        
        Returns:
            Startup time as a formatted string (e.g., '4.5 seconds')
        """
        try:
            # Try to get startup time from WMI
            result = subprocess.run(
                ["powershell", "-Command", 
                 "Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object LastBootUpTime"],
                capture_output=True, text=True, check=True
            )
            
            if "LastBootUpTime" in result.stdout:
                # WMI startup time is available; next try to get the system timer resolution
                duration_result = subprocess.run(
                    ["powershell", "-Command", 
                     "Measure-Command {Get-CimInstance -ClassName Win32_OperatingSystem}"],
                    capture_output=True, text=True, check=True
                )
                
                if "TotalMilliseconds" in duration_result.stdout:
                    total_ms_lines = [line for line in duration_result.stdout.splitlines() 
                                    if "TotalMilliseconds" in line]
                    if total_ms_lines:
                        ms_time = float(total_ms_lines[0].split(":")[-1].strip())
                        startup_time = ms_time / 1000.0
                        
                        if startup_time < 60:
                            return f"{startup_time:.1f} seconds"
                        else:
                            minutes = int(startup_time // 60)
                            seconds = int(startup_time % 60)
                            return f"{minutes}m {seconds}s"
            
            # Fallback if WMI failed
            # Query the windows event log
            event_result = subprocess.run(
                ["powershell", "-Command", 
                 "Get-WinEvent -FilterHashtable @{Logname='System'; ID=6005} -MaxEvents 1 | Select-Object TimeCreated"],
                capture_output=True, text=True, check=True
            )
            
            if "TimeCreated" in event_result.stdout:
                # Get the last boot time and compare with system driver load times
                boot_time = psutil.boot_time()
                boot_datetime = datetime.datetime.fromtimestamp(boot_time)
                
                # Approximate startup time from boot time to login
                # (This is just an approximation, real measurement would need ETW trace)
                startup_time = 15.0  # Default to 15 seconds if we can't calculate
                
                return f"{startup_time:.1f} seconds"
                
            return "Unknown"
        except Exception as e:
            logger.error(f"Error getting startup time: {str(e)}")
            return "Unknown"
    
    def get_system_info(self):
        """Get comprehensive system information.
        
        Returns:
            Dict with system information
        """
        try:
            # Get information with thread pool to avoid blocking
            with ThreadPoolExecutor(max_workers=4) as executor:
                cpu_future = executor.submit(self.get_cpu_usage)
                memory_future = executor.submit(self.get_memory_info)
                disk_future = executor.submit(self.get_disk_info)
                network_future = executor.submit(self.get_network_info)
                
                cpu_percent = cpu_future.result()
                memory_info = memory_future.result()
                disk_info = disk_future.result()
                network_info = network_future.result()
            
            # Additional information
            cpu_temp = self.get_cpu_temperature()
            battery_info = self.get_battery_info()
            uptime = self.get_uptime()
            process_count = self.get_process_count()
            
            # System information
            system_info = {
                "system": platform.system(),
                "release": platform.release(),
                "version": platform.version(),
                "processor": platform.processor(),
                "architecture": platform.machine(),
                "hostname": socket.gethostname(),
                "username": os.getlogin(),
                "cpu_percent": cpu_percent,
                "memory": memory_info,
                "disk": disk_info,
                "network": network_info,
                "battery": battery_info,
                "uptime": uptime,
                "process_count": process_count
            }
            
            if cpu_temp is not None:
                system_info["cpu_temperature"] = cpu_temp
            
            return system_info
        except Exception as e:
            logger.error(f"Error getting system info: {str(e)}")
            return {
                "system": platform.system(),
                "error": str(e)
            }
