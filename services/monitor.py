"""
System monitoring service for the Windows System Optimizer.
This module provides functionality to monitor system performance metrics
such as CPU, memory, disk usage, and network statistics.
"""

import os
import psutil
import platform
import tempfile
import logging
import time
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)

class SystemMonitor:
    """Service class for system monitoring operations."""
    
    def get_cpu_info(self):
        """Get CPU information and usage."""
        try:
            cpu_count = psutil.cpu_count(logical=True)
            cpu_physical_count = psutil.cpu_count(logical=False)
            cpu_percent = psutil.cpu_percent(interval=0.5)
            cpu_freq = psutil.cpu_freq()
            
            # Format frequency
            freq_current = "N/A"
            freq_max = "N/A"
            
            if cpu_freq:
                if cpu_freq.current:
                    freq_current = f"{cpu_freq.current:.2f} MHz"
                if cpu_freq.max:
                    freq_max = f"{cpu_freq.max:.2f} MHz"
            
            return {
                "cpu_count": cpu_count,
                "cpu_physical_count": cpu_physical_count,
                "cpu_percent": cpu_percent,
                "cpu_freq_current": freq_current,
                "cpu_freq_max": freq_max
            }
        
        except Exception as e:
            logger.error(f"Error getting CPU info: {str(e)}")
            return {
                "cpu_count": "N/A",
                "cpu_physical_count": "N/A",
                "cpu_percent": 0,
                "cpu_freq_current": "N/A",
                "cpu_freq_max": "N/A"
            }
    
    def get_memory_info(self):
        """Get memory information and usage."""
        try:
            memory = psutil.virtual_memory()
            
            # Convert to readable format
            total_gb = memory.total / (1024 ** 3)
            available_gb = memory.available / (1024 ** 3)
            used_gb = memory.used / (1024 ** 3)
            
            return {
                "memory_total": f"{total_gb:.2f} GB",
                "memory_available": f"{available_gb:.2f} GB",
                "memory_used": f"{used_gb:.2f} GB",
                "memory_percent": memory.percent
            }
        
        except Exception as e:
            logger.error(f"Error getting memory info: {str(e)}")
            return {
                "memory_total": "N/A",
                "memory_available": "N/A",
                "memory_used": "N/A",
                "memory_percent": 0
            }
    
    def get_disk_info(self):
        """Get disk information and usage."""
        try:
            # Get all disk partitions
            partitions = psutil.disk_partitions()
            disk_info = []
            
            for partition in partitions:
                # Skip CD-ROM and similar
                if 'cdrom' in partition.opts or partition.fstype == '':
                    continue
                
                try:
                    usage = psutil.disk_usage(partition.mountpoint)
                    
                    # Convert to readable format
                    total_gb = usage.total / (1024 ** 3)
                    used_gb = usage.used / (1024 ** 3)
                    free_gb = usage.free / (1024 ** 3)
                    
                    disk_info.append({
                        "device": partition.device,
                        "mountpoint": partition.mountpoint,
                        "fstype": partition.fstype,
                        "total": f"{total_gb:.2f} GB",
                        "used": f"{used_gb:.2f} GB",
                        "free": f"{free_gb:.2f} GB",
                        "percent": usage.percent
                    })
                except PermissionError:
                    # Skip partitions we can't read
                    continue
            
            return disk_info
        
        except Exception as e:
            logger.error(f"Error getting disk info: {str(e)}")
            return []
    
    def get_network_info(self):
        """Get network information and statistics."""
        try:
            # Get network IO counters
            net_io = psutil.net_io_counters()
            
            # Convert to readable format
            bytes_sent_mb = net_io.bytes_sent / (1024 ** 2)
            bytes_recv_mb = net_io.bytes_recv / (1024 ** 2)
            
            return {
                "bytes_sent": f"{bytes_sent_mb:.2f} MB",
                "bytes_recv": f"{bytes_recv_mb:.2f} MB",
                "packets_sent": net_io.packets_sent,
                "packets_recv": net_io.packets_recv,
                "errin": net_io.errin,
                "errout": net_io.errout,
                "dropin": net_io.dropin,
                "dropout": net_io.dropout
            }
        
        except Exception as e:
            logger.error(f"Error getting network info: {str(e)}")
            return {
                "bytes_sent": "N/A",
                "bytes_recv": "N/A",
                "packets_sent": 0,
                "packets_recv": 0,
                "errin": 0,
                "errout": 0,
                "dropin": 0,
                "dropout": 0
            }
    
    def get_system_info(self):
        """Get general system information."""
        try:
            boot_time = datetime.fromtimestamp(psutil.boot_time())
            uptime = datetime.now() - boot_time
            
            # Format uptime
            days = uptime.days
            hours, remainder = divmod(uptime.seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            uptime_str = f"{days}d {hours:02}:{minutes:02}:{seconds:02}"
            
            return {
                "system": platform.system(),
                "node": platform.node(),
                "release": platform.release(),
                "version": platform.version(),
                "machine": platform.machine(),
                "processor": platform.processor(),
                "boot_time": boot_time.strftime("%Y-%m-%d %H:%M:%S"),
                "uptime": uptime_str
            }
        
        except Exception as e:
            logger.error(f"Error getting system info: {str(e)}")
            return {
                "system": "N/A",
                "node": "N/A",
                "release": "N/A",
                "version": "N/A",
                "machine": "N/A",
                "processor": "N/A",
                "boot_time": "N/A",
                "uptime": "N/A"
            }
    
    def get_process_info(self, sort_by="memory_percent", limit=10):
        """
        Get information about running processes.
        
        Args:
            sort_by (str): Field to sort by (memory_percent, cpu_percent, name)
            limit (int): Maximum number of processes to return
            
        Returns:
            list: List of process dictionaries
        """
        try:
            processes = []
            
            for proc in psutil.process_iter(['pid', 'name', 'username', 'memory_percent', 'cpu_percent']):
                # Get process info
                proc_info = proc.info
                
                # Add to list
                processes.append({
                    "pid": proc_info['pid'],
                    "name": proc_info['name'],
                    "username": proc_info['username'],
                    "memory_percent": proc_info['memory_percent'],
                    "cpu_percent": proc_info['cpu_percent']
                })
            
            # Sort processes
            if sort_by == "memory_percent":
                processes.sort(key=lambda x: x['memory_percent'], reverse=True)
            elif sort_by == "cpu_percent":
                processes.sort(key=lambda x: x['cpu_percent'], reverse=True)
            elif sort_by == "name":
                processes.sort(key=lambda x: x['name'])
            
            # Limit number of processes
            return processes[:limit]
        
        except Exception as e:
            logger.error(f"Error getting process info: {str(e)}")
            return []
    
    def get_battery_info(self):
        """Get battery information if available."""
        try:
            if not hasattr(psutil, "sensors_battery"):
                return None
            
            battery = psutil.sensors_battery()
            
            if battery:
                return {
                    "percent": battery.percent,
                    "power_plugged": battery.power_plugged,
                    "secsleft": battery.secsleft
                }
            
            return None
        
        except Exception as e:
            logger.error(f"Error getting battery info: {str(e)}")
            return None
    
    def get_temp_files_size(self):
        """Get the total size of temporary files in bytes."""
        try:
            temp_dir = tempfile.gettempdir()
            total_size = 0
            
            for dirpath, dirnames, filenames in os.walk(temp_dir):
                for f in filenames:
                    try:
                        fp = os.path.join(dirpath, f)
                        if os.path.isfile(fp):
                            total_size += os.path.getsize(fp)
                    except (PermissionError, FileNotFoundError):
                        # Skip files we can't access
                        continue
            
            return total_size
        
        except Exception as e:
            logger.error(f"Error calculating temp files size: {str(e)}")
            return 0
    
    def get_all_metrics(self):
        """Get all system metrics in one call."""
        return {
            "cpu": self.get_cpu_info(),
            "memory": self.get_memory_info(),
            "disk": self.get_disk_info(),
            "network": self.get_network_info(),
            "system": self.get_system_info(),
            "battery": self.get_battery_info(),
            "temp_files_size": self.get_temp_files_size()
        }
