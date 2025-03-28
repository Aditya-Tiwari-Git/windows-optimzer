"""
Network diagnostic tools for network connectivity testing and analysis.
"""

import os
import re
import time
import socket
import logging
import subprocess
import threading
import ipaddress
import platform
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

logger = logging.getLogger(__name__)


class NetworkTools:
    """Network diagnostic utilities for Windows systems."""
    
    def __init__(self):
        """Initialize NetworkTools with default values."""
        self.system = platform.system()
        if self.system != "Windows":
            logger.warning("NetworkTools is optimized for Windows systems")
        
        # Default ping parameters
        self.default_ping_count = 4
        self.default_ping_timeout = 1000  # milliseconds
        
        # Default traceroute parameters
        self.default_max_hops = 30
        self.default_tracert_timeout = 1000  # milliseconds
        
        # Capture parameters
        self.is_capturing = False
        self.capture_thread = None
    
    def get_network_info(self):
        """Get detailed network interface information.
        
        Returns:
            Dict with network interface information
        """
        try:
            # Get network interface information using ipconfig
            result = subprocess.run(
                ["ipconfig", "/all"],
                capture_output=True, 
                text=True, 
                check=True
            )
            
            output = result.stdout
            
            # Default values
            info = {
                "ip_address": "Not available",
                "mac_address": "Not available",
                "default_gateway": "Not available",
                "dns_servers": [],
                "adapter_name": "Not available",
                "connection_type": "Unknown"
            }
            
            # Parse ipconfig output to extract network information
            # Find active interface
            active_interface = None
            interfaces = output.split("\r\n\r\n")
            
            for interface in interfaces:
                # Skip empty interfaces
                if not interface.strip():
                    continue
                
                if "IPv4 Address" in interface and "Disconnected" not in interface:
                    active_interface = interface
                    
                    # Get adapter name
                    adapter_match = re.search(r"^(.*?):", interface.strip(), re.MULTILINE)
                    if adapter_match:
                        info["adapter_name"] = adapter_match.group(1).strip()
                    
                    # Get IP address
                    ip_match = re.search(r"IPv4 Address.*?:\s*([\d\.]+)", interface)
                    if ip_match:
                        info["ip_address"] = ip_match.group(1)
                    
                    # Get MAC address
                    mac_match = re.search(r"Physical Address.*?:\s*([0-9A-F\-]+)", interface)
                    if mac_match:
                        info["mac_address"] = mac_match.group(1)
                    
                    # Get default gateway
                    gateway_match = re.search(r"Default Gateway.*?:\s*([\d\.]+)", interface)
                    if gateway_match:
                        info["default_gateway"] = gateway_match.group(1)
                    
                    # Get DNS servers
                    dns_servers = re.findall(r"DNS Servers.*?:\s*([\d\.]+)", interface)
                    if dns_servers:
                        info["dns_servers"] = dns_servers
                    
                    # Determine connection type
                    if "Wireless" in interface or "Wi-Fi" in interface:
                        info["connection_type"] = "Wireless"
                    elif "Ethernet" in interface:
                        info["connection_type"] = "Ethernet"
                    else:
                        info["connection_type"] = "Unknown"
                    
                    # Found active interface, stop searching
                    break
            
            # If no active interface found, get at least some information
            if not active_interface:
                # Try the first interface that doesn't show disconnected
                for interface in interfaces:
                    if "Disconnected" not in interface and interface.strip():
                        adapter_match = re.search(r"^(.*?):", interface.strip(), re.MULTILINE)
                        if adapter_match:
                            info["adapter_name"] = adapter_match.group(1).strip()
                            
                            # Try to get MAC at minimum
                            mac_match = re.search(r"Physical Address.*?:\s*([0-9A-F\-]+)", interface)
                            if mac_match:
                                info["mac_address"] = mac_match.group(1)
                            
                            break
            
            return info
        except Exception as e:
            logger.error(f"Error getting network info: {str(e)}")
            return {
                "ip_address": "Error",
                "mac_address": "Error",
                "default_gateway": "Error",
                "dns_servers": [],
                "adapter_name": "Error",
                "connection_type": "Error",
                "error": str(e)
            }
    
    def get_interface_list(self):
        """Get list of network interfaces.
        
        Returns:
            List of interface names
        """
        try:
            result = subprocess.run(
                ["ipconfig"], 
                capture_output=True, 
                text=True, 
                check=True
            )
            
            output = result.stdout
            interfaces = []
            
            # Extract interface names
            for line in output.split('\n'):
                if ':' in line and 'adapter' in line.lower():
                    # Extract interface name from line like "Ethernet adapter Local Area Connection:"
                    interface_name = line.split('adapter')[1].split(':')[0].strip()
                    interfaces.append(interface_name)
            
            return interfaces
        except Exception as e:
            logger.error(f"Error getting interface list: {str(e)}")
            return []
    
    def ping(self, host, count=None, timeout=None, continuous=False):
        """Ping a host and return results.
        
        Args:
            host: Host or IP address to ping
            count: Number of pings to send
            timeout: Timeout in milliseconds
            continuous: Whether to ping continuously until stopped
        
        Yields:
            Ping output lines as they are received
        
        Returns:
            List of summary lines
        """
        if count is None:
            count = self.default_ping_count
        
        if timeout is None:
            timeout = self.default_ping_timeout
        
        # Validate host
        if not host:
            yield "Error: No host specified"
            return ["Error: No host specified"]
        
        try:
            # Build ping command
            if continuous:
                cmd = ["ping", "-t", host, "-w", str(timeout)]
            else:
                cmd = ["ping", "-n", str(count), host, "-w", str(timeout)]
            
            # Run ping process
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True
            )
            
            output_lines = []
            summary_lines = []
            in_summary = False
            
            # Read and yield output lines
            for line in iter(process.stdout.readline, ""):
                line = line.strip()
                
                # Skip empty lines
                if not line:
                    continue
                
                # Detect summary section
                if "Ping statistics" in line:
                    in_summary = True
                
                if in_summary:
                    summary_lines.append(line)
                
                output_lines.append(line)
                yield line
                
                # Check if we've been told to stop (for continuous mode)
                if not continuous and len(output_lines) >= count + 5:  # Count + headers and summary
                    break
            
            # Wait for process to finish
            process.terminate()
            process.wait()
            
            # Return summary lines
            return summary_lines
        
        except Exception as e:
            logger.error(f"Error during ping: {str(e)}")
            yield f"Error: {str(e)}"
            return [f"Error: {str(e)}"]
    
    def traceroute(self, host, max_hops=None, timeout=None):
        """Perform traceroute to a host.
        
        Args:
            host: Host or IP address to trace
            max_hops: Maximum number of hops
            timeout: Timeout in milliseconds
        
        Yields:
            Traceroute output lines as they are received
        
        Returns:
            Completion message
        """
        if max_hops is None:
            max_hops = self.default_max_hops
        
        if timeout is None:
            timeout = self.default_tracert_timeout
        
        # Validate host
        if not host:
            yield "Error: No host specified"
            return "Error: No host specified"
        
        try:
            # Build tracert command
            cmd = ["tracert", "-h", str(max_hops), "-w", str(timeout), host]
            
            # Run tracert process
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True
            )
            
            # Read and yield output lines
            for line in iter(process.stdout.readline, ""):
                line = line.strip()
                
                # Skip empty lines
                if not line:
                    continue
                
                yield line
                
                # Check for completion
                if "Trace complete" in line:
                    break
            
            # Wait for process to finish
            process.terminate()
            process.wait()
            
            return "Trace complete"
        
        except Exception as e:
            logger.error(f"Error during traceroute: {str(e)}")
            yield f"Error: {str(e)}"
            return f"Error: {str(e)}"
    
    def capture_traffic(self, duration, interface=None, include_dns=True, include_http=True):
        """Capture network traffic using netsh or PowerShell.
        
        Args:
            duration: Duration in seconds to capture
            interface: Network interface to capture on (None for default)
            include_dns: Whether to include DNS traffic
            include_http: Whether to include HTTP traffic
        
        Returns:
            List of captured packet dictionaries
        """
        try:
            # Check if we're already capturing
            if self.is_capturing:
                return [{"time": "", "source": "", "destination": "", "protocol": "Error", 
                         "info": "A capture is already in progress"}]
            
            self.is_capturing = True
            packets = []
            
            try:
                # We'll use PowerShell with Get-NetTCPConnection and Get-NetUDPEndpoint
                # First prepare the capture commands
                tcp_cmd = [
                    "powershell", 
                    "-Command",
                    f"Get-NetTCPConnection | Select-Object LocalAddress,LocalPort,RemoteAddress,RemotePort,State,OwningProcess"
                ]
                
                udp_cmd = [
                    "powershell", 
                    "-Command",
                    f"Get-NetUDPEndpoint | Select-Object LocalAddress,LocalPort,OwningProcess"
                ]
                
                # Get process information to resolve PIDs to names
                process_cmd = [
                    "powershell", 
                    "-Command",
                    "Get-Process | Select-Object Id,ProcessName"
                ]
                
                # Run initial commands to get baseline
                tcp_initial = subprocess.run(tcp_cmd, capture_output=True, text=True, check=True).stdout
                udp_initial = subprocess.run(udp_cmd, capture_output=True, text=True, check=True).stdout
                
                # Wait for the specified duration
                start_time = datetime.now()
                time.sleep(duration)
                
                # Run commands again to see what's changed
                tcp_final = subprocess.run(tcp_cmd, capture_output=True, text=True, check=True).stdout
                udp_final = subprocess.run(udp_cmd, capture_output=True, text=True, check=True).stdout
                process_info = subprocess.run(process_cmd, capture_output=True, text=True, check=True).stdout
                
                # Parse process information
                processes = {}
                for line in process_info.splitlines():
                    line = line.strip()
                    if line and "Id" not in line:  # Skip header
                        parts = line.split()
                        if parts and len(parts) >= 2:
                            try:
                                pid = int(parts[0])
                                process_name = parts[1]
                                processes[pid] = process_name
                            except:
                                continue
                
                # Parse TCP connections
                tcp_connections = {}
                for line in tcp_final.splitlines():
                    line = line.strip()
                    if line and "LocalAddress" not in line:  # Skip header
                        parts = line.split()
                        if len(parts) >= 6:
                            try:
                                local_addr = parts[0]
                                local_port = parts[1]
                                remote_addr = parts[2]
                                remote_port = parts[3]
                                state = parts[4]
                                try:
                                    pid = int(parts[5])
                                    process_name = processes.get(pid, "Unknown")
                                except:
                                    process_name = "Unknown"
                                
                                # Skip localhost connections unless requested
                                if local_addr == "127.0.0.1" or remote_addr == "127.0.0.1":
                                    continue
                                
                                # Skip non-HTTP if requested
                                if not include_http and remote_port not in ("80", "443", "8080"):
                                    continue
                                
                                key = f"{local_addr}:{local_port}-{remote_addr}:{remote_port}"
                                tcp_connections[key] = {
                                    "time": datetime.now().strftime("%H:%M:%S"),
                                    "source": f"{local_addr}:{local_port}",
                                    "destination": f"{remote_addr}:{remote_port}",
                                    "protocol": "TCP",
                                    "info": f"{state} ({process_name})"
                                }
                            except:
                                continue
                
                # Parse UDP endpoints
                udp_endpoints = {}
                for line in udp_final.splitlines():
                    line = line.strip()
                    if line and "LocalAddress" not in line:  # Skip header
                        parts = line.split()
                        if len(parts) >= 3:
                            try:
                                local_addr = parts[0]
                                local_port = parts[1]
                                
                                try:
                                    pid = int(parts[2])
                                    process_name = processes.get(pid, "Unknown")
                                except:
                                    process_name = "Unknown"
                                
                                # Skip localhost unless requested
                                if local_addr == "127.0.0.1":
                                    continue
                                
                                # Skip non-DNS if requested
                                if not include_dns and local_port != "53":
                                    continue
                                
                                key = f"{local_addr}:{local_port}"
                                udp_endpoints[key] = {
                                    "time": datetime.now().strftime("%H:%M:%S"),
                                    "source": f"{local_addr}:{local_port}",
                                    "destination": "Multiple",
                                    "protocol": "UDP",
                                    "info": f"Listening ({process_name})"
                                }
                            except:
                                continue
                
                # Combine results
                for conn in tcp_connections.values():
                    packets.append(conn)
                
                for endpoint in udp_endpoints.values():
                    packets.append(endpoint)
                
                # Sort by time
                packets.sort(key=lambda x: x["time"])
                
                return packets
            
            finally:
                self.is_capturing = False
        
        except Exception as e:
            logger.error(f"Error capturing traffic: {str(e)}")
            self.is_capturing = False
            return [{"time": "", "source": "", "destination": "", "protocol": "Error", 
                     "info": f"Error capturing traffic: {str(e)}"}]
    
    def dns_lookup(self, hostname):
        """Perform DNS lookup for a hostname.
        
        Args:
            hostname: Hostname to look up
        
        Returns:
            Dict with DNS information
        """
        try:
            # Get IP address from hostname
            ip_address = socket.gethostbyname(hostname)
            
            # Get additional DNS information using nslookup
            cmd = ["nslookup", hostname]
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            
            output = result.stdout
            
            # Parse nslookup output
            name_servers = []
            addresses = []
            aliases = []
            
            for line in output.split('\n'):
                line = line.strip()
                
                if "Server:" in line:
                    server_line = line.split("Server:")[1].strip()
                    name_servers.append(server_line)
                
                if "Address:" in line and "Server:" not in line:
                    addr = line.split("Address:")[1].strip()
                    addresses.append(addr)
                
                if "Aliases:" in line:
                    alias = line.split("Aliases:")[1].strip()
                    aliases.append(alias)
            
            return {
                "hostname": hostname,
                "ip_address": ip_address,
                "addresses": addresses,
                "name_servers": name_servers,
                "aliases": aliases,
                "raw_output": output
            }
        
        except socket.gaierror:
            return {
                "hostname": hostname,
                "error": "DNS resolution failed",
                "status": "error"
            }
        except Exception as e:
            logger.error(f"Error performing DNS lookup: {str(e)}")
            return {
                "hostname": hostname,
                "error": str(e),
                "status": "error"
            }
    
    def check_port(self, host, port, timeout=2):
        """Check if a port is open on a host.
        
        Args:
            host: Host or IP address to check
            port: Port number to check
            timeout: Timeout in seconds
        
        Returns:
            Dict with port status information
        """
        try:
            # Create socket
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(timeout)
            
            # Try to connect
            start_time = time.time()
            result = sock.connect_ex((host, port))
            elapsed_time = time.time() - start_time
            
            # Check result
            if result == 0:
                status = "open"
            else:
                status = "closed"
            
            sock.close()
            
            return {
                "host": host,
                "port": port,
                "status": status,
                "response_time": elapsed_time,
                "error": None
            }
        
        except socket.gaierror:
            return {
                "host": host,
                "port": port,
                "status": "error",
                "response_time": None,
                "error": "DNS resolution failed"
            }
        except Exception as e:
            logger.error(f"Error checking port: {str(e)}")
            return {
                "host": host,
                "port": port,
                "status": "error",
                "response_time": None,
                "error": str(e)
            }
    
    def scan_network(self, subnet="192.168.1.0/24"):
        """Scan local network for active hosts.
        
        Args:
            subnet: Subnet to scan in CIDR notation (e.g., 192.168.1.0/24)
        
        Returns:
            List of dicts with active host information
        """
        try:
            # Parse subnet
            network = ipaddress.ip_network(subnet, strict=False)
            
            active_hosts = []
            
            # Use thread pool to scan in parallel
            with ThreadPoolExecutor(max_workers=50) as executor:
                # Submit ping tasks
                future_to_ip = {
                    executor.submit(self._ping_host, str(ip)): str(ip) 
                    for ip in network.hosts()
                }
                
                # Collect results as they complete
                for future in future_to_ip:
                    ip = future_to_ip[future]
                    try:
                        result = future.result()
                        if result["status"] == "active":
                            active_hosts.append(result)
                    except Exception as e:
                        logger.error(f"Error scanning host {ip}: {str(e)}")
            
            # Sort by IP address
            active_hosts.sort(key=lambda x: [int(octet) for octet in x["ip"].split('.')])
            
            return active_hosts
        
        except Exception as e:
            logger.error(f"Error scanning network: {str(e)}")
            return [{"ip": "Error", "hostname": "", "status": "error", "response_time": 0, "error": str(e)}]
    
    def _ping_host(self, ip, timeout=500):
        """Ping a host to check if it's active.
        
        Args:
            ip: IP address to ping
            timeout: Timeout in milliseconds
        
        Returns:
            Dict with host status information
        """
        try:
            # Use ping command with 1 packet and short timeout
            cmd = ["ping", "-n", "1", "-w", str(timeout), ip]
            
            start_time = time.time()
            result = subprocess.run(cmd, capture_output=True, text=True)
            elapsed_time = (time.time() - start_time) * 1000  # Convert to ms
            
            # Check if ping was successful
            if result.returncode == 0 and "Reply from" in result.stdout:
                # Try to get hostname
                hostname = ""
                try:
                    hostname = socket.gethostbyaddr(ip)[0]
                except socket.herror:
                    pass
                
                return {
                    "ip": ip,
                    "hostname": hostname,
                    "status": "active",
                    "response_time": elapsed_time
                }
            else:
                return {
                    "ip": ip,
                    "hostname": "",
                    "status": "inactive",
                    "response_time": 0
                }
        
        except Exception as e:
            logger.error(f"Error pinging host {ip}: {str(e)}")
            return {
                "ip": ip,
                "hostname": "",
                "status": "error",
                "response_time": 0,
                "error": str(e)
            }
