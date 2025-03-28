"""
Helper utilities for the Windows System Optimizer.
This module provides general helper functions used throughout the application.
"""

import os
import sys
import logging
import subprocess
import shutil
import json
import tempfile
import re
from datetime import datetime
import hashlib

logger = logging.getLogger(__name__)

def format_bytes(size, decimal_places=2):
    """
    Format bytes to human-readable size.
    
    Args:
        size (int): Size in bytes
        decimal_places (int): Number of decimal places
    
    Returns:
        str: Formatted size string
    """
    for unit in ['B', 'KB', 'MB', 'GB', 'TB', 'PB']:
        if size < 1024.0 or unit == 'PB':
            break
        size /= 1024.0
    
    return f"{size:.{decimal_places}f} {unit}"

def format_time_delta(seconds):
    """
    Format a time delta in seconds to human-readable format.
    
    Args:
        seconds (int): Time delta in seconds
    
    Returns:
        str: Formatted time string
    """
    days, remainder = divmod(seconds, 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, seconds = divmod(remainder, 60)
    
    if days > 0:
        return f"{int(days)}d {int(hours):02}:{int(minutes):02}:{int(seconds):02}"
    else:
        return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"

def safe_delete_file(file_path):
    """
    Safely delete a file.
    
    Args:
        file_path (str): Path to file
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if os.path.exists(file_path):
            os.unlink(file_path)
            return True
        return False
    except Exception as e:
        logger.warning(f"Error deleting file {file_path}: {str(e)}")
        return False

def safe_delete_directory(dir_path, recursive=True):
    """
    Safely delete a directory.
    
    Args:
        dir_path (str): Path to directory
        recursive (bool): Recursively delete contents if True
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if os.path.exists(dir_path):
            if recursive:
                shutil.rmtree(dir_path, ignore_errors=True)
            else:
                os.rmdir(dir_path)
            return True
        return False
    except Exception as e:
        logger.warning(f"Error deleting directory {dir_path}: {str(e)}")
        return False

def safe_create_directory(dir_path):
    """
    Safely create a directory.
    
    Args:
        dir_path (str): Path to directory
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)
        return True
    except Exception as e:
        logger.warning(f"Error creating directory {dir_path}: {str(e)}")
        return False

def calculate_file_hash(file_path, algorithm='sha256'):
    """
    Calculate the hash of a file.
    
    Args:
        file_path (str): Path to file
        algorithm (str): Hash algorithm to use
    
    Returns:
        str: Hash digest as hexadecimal string
    """
    try:
        if not os.path.exists(file_path):
            return None
        
        hash_obj = None
        if algorithm == 'md5':
            hash_obj = hashlib.md5()
        elif algorithm == 'sha1':
            hash_obj = hashlib.sha1()
        elif algorithm == 'sha256':
            hash_obj = hashlib.sha256()
        else:
            hash_obj = hashlib.sha256()
        
        with open(file_path, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b''):
                hash_obj.update(chunk)
        
        return hash_obj.hexdigest()
    except Exception as e:
        logger.warning(f"Error calculating file hash for {file_path}: {str(e)}")
        return None

def backup_file(file_path, backup_dir=None):
    """
    Create a backup of a file.
    
    Args:
        file_path (str): Path to file
        backup_dir (str, optional): Directory to store backup
    
    Returns:
        str: Path to backup file if successful, None otherwise
    """
    try:
        if not os.path.exists(file_path):
            return None
        
        # Default backup directory
        if not backup_dir:
            backup_dir = os.path.join(tempfile.gettempdir(), "WinOptimizerBackups")
        
        # Ensure backup directory exists
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        
        # Create backup filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        filename = os.path.basename(file_path)
        backup_filename = f"{filename}.{timestamp}.bak"
        backup_path = os.path.join(backup_dir, backup_filename)
        
        # Copy file
        shutil.copy2(file_path, backup_path)
        
        return backup_path
    except Exception as e:
        logger.warning(f"Error backing up file {file_path}: {str(e)}")
        return None

def find_files(directory, pattern=None, recursive=True, max_size=None, min_size=None):
    """
    Find files matching criteria.
    
    Args:
        directory (str): Directory to search
        pattern (str, optional): Filename pattern (regex)
        recursive (bool): Search subdirectories if True
        max_size (int, optional): Maximum file size in bytes
        min_size (int, optional): Minimum file size in bytes
    
    Returns:
        list: List of matching file paths
    """
    try:
        if not os.path.exists(directory):
            return []
        
        matching_files = []
        
        # Compile regex pattern if provided
        regex = re.compile(pattern) if pattern else None
        
        # Walk directory
        if recursive:
            for root, _, files in os.walk(directory):
                for filename in files:
                    file_path = os.path.join(root, filename)
                    
                    # Check if filename matches pattern
                    if regex and not regex.search(filename):
                        continue
                    
                    # Check file size
                    if max_size or min_size:
                        try:
                            file_size = os.path.getsize(file_path)
                            if max_size and file_size > max_size:
                                continue
                            if min_size and file_size < min_size:
                                continue
                        except OSError:
                            continue
                    
                    matching_files.append(file_path)
        else:
            # Non-recursive search
            for filename in os.listdir(directory):
                file_path = os.path.join(directory, filename)
                
                # Skip directories
                if os.path.isdir(file_path):
                    continue
                
                # Check if filename matches pattern
                if regex and not regex.search(filename):
                    continue
                
                # Check file size
                if max_size or min_size:
                    try:
                        file_size = os.path.getsize(file_path)
                        if max_size and file_size > max_size:
                            continue
                        if min_size and file_size < min_size:
                            continue
                    except OSError:
                        continue
                
                matching_files.append(file_path)
        
        return matching_files
    except Exception as e:
        logger.warning(f"Error finding files in {directory}: {str(e)}")
        return []

def run_command(command, shell=False, timeout=None):
    """
    Run a command and return the output.
    
    Args:
        command (list or str): Command to run
        shell (bool): Use shell if True
        timeout (int, optional): Timeout in seconds
    
    Returns:
        dict: Command result with stdout, stderr, and return code
    """
    try:
        result = subprocess.run(
            command,
            shell=shell,
            capture_output=True,
            text=True,
            timeout=timeout
        )
        
        return {
            'stdout': result.stdout,
            'stderr': result.stderr,
            'returncode': result.returncode,
            'success': result.returncode == 0
        }
    except subprocess.TimeoutExpired:
        return {
            'stdout': '',
            'stderr': 'Command timed out',
            'returncode': -1,
            'success': False
        }
    except Exception as e:
        logger.warning(f"Error running command {command}: {str(e)}")
        return {
            'stdout': '',
            'stderr': str(e),
            'returncode': -1,
            'success': False
        }

def load_json_file(file_path, default=None):
    """
    Load data from a JSON file.
    
    Args:
        file_path (str): Path to JSON file
        default (any, optional): Default value if file doesn't exist or is invalid
    
    Returns:
        any: Loaded JSON data or default value
    """
    try:
        if not os.path.exists(file_path):
            return default
        
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.warning(f"Error loading JSON file {file_path}: {str(e)}")
        return default

def save_json_file(file_path, data, indent=4):
    """
    Save data to a JSON file.
    
    Args:
        file_path (str): Path to JSON file
        data (any): Data to save
        indent (int): JSON indentation
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        directory = os.path.dirname(file_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=indent)
        
        return True
    except Exception as e:
        logger.warning(f"Error saving JSON file {file_path}: {str(e)}")
        return False

def is_valid_url(url):
    """
    Check if a URL is valid.
    
    Args:
        url (str): URL to check
    
    Returns:
        bool: True if valid, False otherwise
    """
    try:
        import re
        url_pattern = re.compile(
            r'^(?:http|ftp)s?://'  # http:// or https://
            r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|'  # domain
            r'localhost|'  # localhost
            r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'  # or IP
            r'(?::\d+)?'  # optional port
            r'(?:/?|[/?]\S+)$', re.IGNORECASE)
        
        return bool(url_pattern.match(url))
    except Exception:
        return False

def is_valid_ip(ip):
    """
    Check if an IP address is valid.
    
    Args:
        ip (str): IP address to check
    
    Returns:
        bool: True if valid, False otherwise
    """
    try:
        import ipaddress
        ipaddress.ip_address(ip)
        return True
    except ValueError:
        return False
    except Exception:
        return False

def is_valid_domain(domain):
    """
    Check if a domain name is valid.
    
    Args:
        domain (str): Domain name to check
    
    Returns:
        bool: True if valid, False otherwise
    """
    try:
        import re
        domain_pattern = re.compile(
            r'^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z0-9][a-z0-9-]{0,61}[a-z0-9]$',
            re.IGNORECASE
        )
        
        return bool(domain_pattern.match(domain))
    except Exception:
        return False

def get_file_properties(file_path):
    """
    Get file properties like version, description, etc.
    
    Args:
        file_path (str): Path to file
    
    Returns:
        dict: File properties
    """
    try:
        if not os.path.exists(file_path):
            return {}
        
        # Use PowerShell to get file version info
        ps_command = f"""
        $fileInfo = New-Object System.IO.FileInfo -ArgumentList '{file_path}'
        $fileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo('{file_path}')
        $properties = @{{
            FileName = $fileInfo.Name
            FileSize = $fileInfo.Length
            CreationTime = $fileInfo.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
            LastWriteTime = $fileInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            FileVersion = $fileVersion.FileVersion
            ProductVersion = $fileVersion.ProductVersion
            ProductName = $fileVersion.ProductName
            CompanyName = $fileVersion.CompanyName
            FileDescription = $fileVersion.FileDescription
            InternalName = $fileVersion.InternalName
            OriginalFileName = $fileVersion.OriginalFilename
            LegalCopyright = $fileVersion.LegalCopyright
        }}
        $properties | ConvertTo-Json
        """
        
        ps_command = ps_command.replace('\n', ' ')
        
        result = subprocess.run(
            ['powershell.exe', '-Command', ps_command],
            capture_output=True,
            text=True,
            check=True
        )
        
        # Parse JSON output
        return json.loads(result.stdout)
    except Exception as e:
        logger.warning(f"Error getting file properties for {file_path}: {str(e)}")
        
        # Fallback to basic file info
        try:
            file_stat = os.stat(file_path)
            return {
                'FileName': os.path.basename(file_path),
                'FileSize': file_stat.st_size,
                'CreationTime': datetime.fromtimestamp(file_stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                'LastWriteTime': datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            }
        except Exception:
            return {}
