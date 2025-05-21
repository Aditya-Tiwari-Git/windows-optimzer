"""
Dashboard widget for the Windows System Optimizer.
This module provides the main dashboard with system performance metrics
and quick access to main features.
"""

import os
import psutil
import platform
import concurrent.futures
from datetime import datetime

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QFrame, QGridLayout, QPushButton, QProgressBar,
                            QSpacerItem, QSizePolicy)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QFont

import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from services.monitor import SystemMonitor
from .styles import COLORS

class MetricCard(QFrame):
    """Card widget displaying a single system metric."""
    def __init__(self, title, value="", icon="fa.info-circle", parent=None):
        super().__init__(parent)
        self.setObjectName("metricCard")
        self.setMinimumHeight(120)
        
        layout = QVBoxLayout(self)
        
        # Card title
        title_layout = QHBoxLayout()
        title_label = QLabel(title)
        title_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        layout.addLayout(title_layout)
        
        # Value
        self.value_label = QLabel(value)
        self.value_label.setFont(QFont("Segoe UI", 20, QFont.Bold))
        self.value_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.value_label)
        
        layout.addStretch()
    
    def update_value(self, value):
        """Update the displayed value."""
        self.value_label.setText(str(value))
        

class CPUUsageChart(QWidget):
    """Widget displaying CPU usage over time in a chart."""
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        
        # Create the matplotlib figure and canvas
        self.figure = Figure(figsize=(5, 3), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
        
        # Set up the plot
        self.ax = self.figure.add_subplot(111)
        self.cpu_data = [0] * 60  # Store 60 seconds of data
        self.time_data = list(range(60))
        
        # Initial plot
        self.line, = self.ax.plot(self.time_data, self.cpu_data)
        self.ax.set_ylim(0, 100)
        self.ax.set_xlim(0, 59)
        self.ax.set_ylabel('CPU Usage (%)')
        self.ax.set_title('CPU Usage Over Time')
        self.ax.grid(True)
        
        self.figure.tight_layout()
    
    def update_data(self, value):
        """Update the chart with new CPU usage data."""
        # Add new data point and remove oldest one
        self.cpu_data.append(value)
        self.cpu_data.pop(0)
        
        # Update the line plot
        self.line.set_ydata(self.cpu_data)
        
        # Redraw the canvas
        self.canvas.draw_idle()


class MemoryUsageChart(QWidget):
    """Widget displaying memory usage over time in a chart."""
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        
        # Create the matplotlib figure and canvas
        self.figure = Figure(figsize=(5, 3), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
        
        # Set up the plot
        self.ax = self.figure.add_subplot(111)
        self.memory_data = [0] * 60  # Store 60 seconds of data
        self.time_data = list(range(60))
        
        # Initial plot
        self.line, = self.ax.plot(self.time_data, self.memory_data)
        self.ax.set_ylim(0, 100)
        self.ax.set_xlim(0, 59)
        self.ax.set_ylabel('Memory Usage (%)')
        self.ax.set_title('Memory Usage Over Time')
        self.ax.grid(True)
        
        self.figure.tight_layout()
    
    def update_data(self, value):
        """Update the chart with new memory usage data."""
        # Add new data point and remove oldest one
        self.memory_data.append(value)
        self.memory_data.pop(0)
        
        # Update the line plot
        self.line.set_ydata(self.memory_data)
        
        # Redraw the canvas
        self.canvas.draw_idle()


class DashboardWidget(QWidget):
    """Main dashboard widget showing system metrics and charts."""
    
    refresh_signal = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        # Create system monitor instance
        self.monitor = SystemMonitor()
        
        # Create main layout
        main_layout = QVBoxLayout(self)
        
        # Header with system info
        header_layout = QHBoxLayout()
        
        # System name and version
        system_info_layout = QVBoxLayout()
        
        self.system_name_label = QLabel(f"Windows {platform.win32_ver()[0]}")
        self.system_name_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        system_info_layout.addWidget(self.system_name_label)
        
        computer_name = os.environ.get('COMPUTERNAME', 'Unknown')
        self.computer_name_label = QLabel(f"Computer: {computer_name}")
        self.computer_name_label.setFont(QFont("Segoe UI", 10))
        system_info_layout.addWidget(self.computer_name_label)
        
        header_layout.addLayout(system_info_layout)
        header_layout.addStretch()
        
        # Last scan time
        scan_info_layout = QVBoxLayout()
        scan_label = QLabel("Last System Scan:")
        scan_label.setAlignment(Qt.AlignRight)
        scan_info_layout.addWidget(scan_label)
        
        self.last_scan_label = QLabel("Never")
        self.last_scan_label.setAlignment(Qt.AlignRight)
        scan_info_layout.addWidget(self.last_scan_label)
        
        header_layout.addLayout(scan_info_layout)
        
        # Scan button
        scan_button = QPushButton("Scan System")
        scan_button.setMinimumHeight(40)
        scan_button.clicked.connect(self.scan_system)
        header_layout.addWidget(scan_button)
        
        main_layout.addLayout(header_layout)
        
        # Horizontal separator
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(separator)
        
        # Metric cards
        metrics_layout = QGridLayout()
        
        # Create metric cards
        self.cpu_card = MetricCard("CPU Usage", "0%")
        self.memory_card = MetricCard("Memory Usage", "0%")
        self.disk_card = MetricCard("Disk Usage", "0%")
        self.temp_files_card = MetricCard("Temp Files", "0 MB")
        self.system_uptime_card = MetricCard("System Uptime", "0:00:00")
        self.battery_card = MetricCard("Battery", "N/A")
        
        # Add cards to grid
        metrics_layout.addWidget(self.cpu_card, 0, 0)
        metrics_layout.addWidget(self.memory_card, 0, 1)
        metrics_layout.addWidget(self.disk_card, 0, 2)
        metrics_layout.addWidget(self.temp_files_card, 1, 0)
        metrics_layout.addWidget(self.system_uptime_card, 1, 1)
        metrics_layout.addWidget(self.battery_card, 1, 2)
        
        main_layout.addLayout(metrics_layout)
        
        # Charts
        charts_layout = QHBoxLayout()
        
        # CPU usage chart
        self.cpu_chart = CPUUsageChart()
        charts_layout.addWidget(self.cpu_chart)
        
        # Memory usage chart
        self.memory_chart = MemoryUsageChart()
        charts_layout.addWidget(self.memory_chart)
        
        main_layout.addLayout(charts_layout)
        
        # Setup refresh timer (update every second)
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.update_metrics)
        self.refresh_timer.start(1000)  # 1 second interval
        
        # Initial update
        self.update_metrics()
        
        # Initialize thread pool for background tasks
        self.thread_pool = concurrent.futures.ThreadPoolExecutor(max_workers=2)
    
    def update_metrics(self):
        """Update all displayed metrics with current system information."""
        # Update basic metrics
        cpu_percent = psutil.cpu_percent()
        self.cpu_card.update_value(f"{cpu_percent}%")
        self.cpu_chart.update_data(cpu_percent)
        
        memory_info = psutil.virtual_memory()
        memory_percent = memory_info.percent
        self.memory_card.update_value(f"{memory_percent}%")
        self.memory_chart.update_data(memory_percent)
        
        # Disk usage (C: drive)
        disk_usage = psutil.disk_usage('C:')
        self.disk_card.update_value(f"{disk_usage.percent}%")
        
        # System uptime
        uptime = datetime.now() - datetime.fromtimestamp(psutil.boot_time())
        hours, remainder = divmod(uptime.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        uptime_str = f"{uptime.days}d {hours:02}:{minutes:02}:{seconds:02}"
        self.system_uptime_card.update_value(uptime_str)
        
        # Battery information (if available)
        if hasattr(psutil, "sensors_battery"):
            battery = psutil.sensors_battery()
            if battery:
                battery_percent = battery.percent
                status = "Charging" if battery.power_plugged else "Discharging"
                self.battery_card.update_value(f"{battery_percent}% ({status})")
            else:
                self.battery_card.update_value("No battery")
    
    def scan_system(self):
        """Perform a full system scan in the background."""
        # Submit the scan task to the thread pool
        self.thread_pool.submit(self._background_scan)
    
    def _background_scan(self):
        """Background task for system scanning."""
        # Update temp files size
        temp_size = self.monitor.get_temp_files_size()
        temp_size_mb = temp_size / (1024 * 1024)  # Convert to MB
        
        # Update UI from the main thread
        self.temp_files_card.update_value(f"{temp_size_mb:.1f} MB")
        
        # Update last scan time
        now = datetime.now()
        self.last_scan_label.setText(now.strftime("%Y-%m-%d %H:%M:%S"))
