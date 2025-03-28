"""
Network diagnostics service for the Windows System Optimizer.
This module provides functionality for network diagnostics such as
ping tests, traceroute, DNS lookups, and port scans.
"""

import os
import subprocess
import socket
import logging
import re
import time
from datetime import datetime

logger = logging.getLogger(__name__)

class NetworkDiagnostics:
    """Service class for network diagnostic operations."""
    
    def ping_test(self, target, count=4, timeout=1000):
        """
        Perform a ping test to the specified target.
        
        Args:
            target (str): Domain or IP address to ping
            count (int): Number of echo requests to send
            timeout (int): Timeout in milliseconds
            
        Returns:
            str: Formatted ping test results
        """
        try:
            # Validate input
            if not target:
                return "Error: No target specified"
            
            # Execute ping command
            process = subprocess.run(
                ["ping", "-n", str(count), "-w", str(timeout), target],
                capture_output=True,
                text=True
            )
            
            output = process.stdout
            
            # Add timestamp
            result = f"Ping test to {target} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            result += "=" * 50 + "\n"
            result += output
            
            return result
        
        except Exception as e:
            logger.error(f"Error during ping test: {str(e)}")
            return f"Error during ping test: {str(e)}"
    
    def traceroute(self, target, max_hops=30, timeout=1000):
        """
        Perform a traceroute to the specified target.
        
        Args:
            target (str): Domain or IP address to trace
            max_hops (int): Maximum number of hops
            timeout (int): Timeout in milliseconds
            
        Returns:
            str: Formatted traceroute results
        """
        try:
            # Validate input
            if not target:
                return "Error: No target specified"
            
            # Execute tracert command
            process = subprocess.run(
                ["tracert", "-h", str(max_hops), "-w", str(timeout), target],
                capture_output=True,
                text=True
            )
            
            output = process.stdout
            
            # Add timestamp
            result = f"Traceroute to {target} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            result += "=" * 50 + "\n"
            result += output
            
            return result
        
        except Exception as e:
            logger.error(f"Error during traceroute: {str(e)}")
            return f"Error during traceroute: {str(e)}"
    
    def dns_lookup(self, target):
        """
        Perform a DNS lookup for the specified target.
        
        Args:
            target (str): Domain to lookup
            
        Returns:
            str: Formatted DNS lookup results
        """
        try:
            # Validate input
            if not target:
                return "Error: No target specified"
            
            # Execute nslookup command
            process = subprocess.run(
                ["nslookup", target],
                capture_output=True,
                text=True
            )
            
            output = process.stdout
            
            # Add timestamp and additional info
            result = f"DNS lookup for {target} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            result += "=" * 50 + "\n"
            result += output
            
            # Try to get additional information
            try:
                # Get IPv4 and IPv6 addresses
                ipv4_info = socket.getaddrinfo(target, None, socket.AF_INET)
                ipv6_info = socket.getaddrinfo(target, None, socket.AF_INET6)
                
                result += "\nAdditional Information:\n"
                result += "-" * 50 + "\n"
                
                # IPv4 addresses
                result += "IPv4 Addresses:\n"
                for info in ipv4_info:
                    result += f"  {info[4][0]}\n"
                
                # IPv6 addresses
                result += "\nIPv6 Addresses:\n"
                for info in ipv6_info:
                    result += f"  {info[4][0]}\n"
                
            except (socket.gaierror, IndexError):
                # If additional lookup fails, ignore and continue
                pass
            
            return result
        
        except Exception as e:
            logger.error(f"Error during DNS lookup: {str(e)}")
            return f"Error during DNS lookup: {str(e)}"
    
    def port_scan(self, target, ports):
        """
        Perform a basic port scan on the specified target.
        
        Args:
            target (str): Domain or IP address to scan
            ports (list): List of ports to scan
            
        Returns:
            str: Formatted port scan results
        """
        try:
            # Validate input
            if not target:
                return "Error: No target specified"
            
            if not ports:
                return "Error: No ports specified"
            
            # Try to resolve the hostname to an IP address
            try:
                ip = socket.gethostbyname(target)
            except socket.gaierror:
                return f"Error: Could not resolve hostname {target}"
            
            # Add timestamp
            result = f"Port scan for {target} ({ip}) at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            result += "=" * 50 + "\n"
            result += "PORT     STATE    SERVICE\n"
            
            # Check each port
            for port in ports:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(1)  # 1 second timeout
                
                service = self._get_service_name(port)
                
                try:
                    # Try to connect to the port
                    connection = sock.connect_ex((ip, port))
                    if connection == 0:
                        result += f"{port:5d}    open     {service}\n"
                    else:
                        result += f"{port:5d}    closed   {service}\n"
                except:
                    result += f"{port:5d}    error    {service}\n"
                
                sock.close()
            
            return result
        
        except Exception as e:
            logger.error(f"Error during port scan: {str(e)}")
            return f"Error during port scan: {str(e)}"
    
    def capture_network_log(self, target, duration=10):
        """
        Capture network activity log for the specified target.
        
        Args:
            target (str): Domain or IP to filter for, or "*" for all
            duration (int): Duration in seconds to capture
            
        Returns:
            str: Formatted network log results
        """
        try:
            # Add timestamp
            result = f"Network log at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            result += "=" * 50 + "\n"
            
            # Use netstat to capture current connections
            result += "Current Connections:\n"
            
            # Run netstat
            process = subprocess.run(
                ["netstat", "-ano"],
                capture_output=True,
                text=True
            )
            
            output = process.stdout
            
            # Filter output if target is specified and not wildcard
            if target and target != "*":
                filtered_lines = []
                for line in output.splitlines():
                    if target in line:
                        filtered_lines.append(line)
                
                if filtered_lines:
                    result += "\n".join(filtered_lines)
                else:
                    result += f"No connections found for {target}\n"
            else:
                # Add all output
                result += output
            
            # Add ipconfig information
            result += "\n\nNetwork Configuration:\n"
            result += "-" * 50 + "\n"
            
            # Run ipconfig
            process = subprocess.run(
                ["ipconfig", "/all"],
                capture_output=True,
                text=True
            )
            
            result += process.stdout
            
            return result
        
        except Exception as e:
            logger.error(f"Error during network log capture: {str(e)}")
            return f"Error during network log capture: {str(e)}"
    
    def _get_service_name(self, port):
        """Get the service name for a well-known port."""
        common_ports = {
            21: "FTP",
            22: "SSH",
            23: "Telnet",
            25: "SMTP",
            53: "DNS",
            80: "HTTP",
            110: "POP3",
            143: "IMAP",
            443: "HTTPS",
            465: "SMTPS",
            587: "SMTP",
            993: "IMAPS",
            995: "POP3S",
            3306: "MySQL",
            3389: "RDP",
            5900: "VNC",
            8080: "HTTP-Alt",
            8443: "HTTPS-Alt"
        }
        
        return common_ports.get(port, "Unknown")
