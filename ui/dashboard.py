"""
Dashboard widget showing system performance metrics and status.
"""

import os
import time
import psutil
import matplotlib
matplotlib.use('Qt5Agg')  # Use Qt5 backend for matplotlib
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QFrame, QGridLayout, QSizePolicy, QProgressBar)
from PyQt5.QtGui import QFont, QColor, QPainter, QPen
from PyQt5.QtCore import Qt, QTimer, pyqtSlot, QRectF

from ui.theme import COLORS
from ui.components import Card, HeaderLabel, SubHeaderLabel, InfoCard
from utils.system_info import SystemInfo


class MetricsChart(FigureCanvas):
    """Custom Matplotlib chart for system metrics visualization."""
    
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.fig.set_facecolor(COLORS['background'])
        
        # Create subplots
        self.axes = self.fig.add_subplot(111)
        self.axes.set_facecolor(COLORS['background'])
        self.axes.spines['top'].set_visible(False)
        self.axes.spines['right'].set_visible(False)
        self.axes.spines['bottom'].set_color(COLORS['text'])
        self.axes.spines['left'].set_color(COLORS['text'])
        self.axes.tick_params(colors=COLORS['text'])
        self.axes.xaxis.label.set_color(COLORS['text'])
        self.axes.yaxis.label.set_color(COLORS['text'])
        self.axes.grid(True, linestyle='--', alpha=0.7)
        
        # Initialize data
        self.timestamps = []
        self.cpu_values = []
        self.memory_values = []
        
        # Setup plot elements
        self.cpu_line, = self.axes.plot([], [], lw=2, label='CPU %', color=COLORS['primary'])
        self.memory_line, = self.axes.plot([], [], lw=2, label='Memory %', color=COLORS['secondary'])
        
        self.axes.set_ylim(0, 100)
        self.axes.set_title('Real-time System Performance', color=COLORS['text'])
        self.axes.set_xlabel('Time', color=COLORS['text'])
        self.axes.set_ylabel('Usage %', color=COLORS['text'])
        self.axes.legend(loc='upper left')
        
        # Adjust layout
        self.fig.tight_layout()
        
        super(MetricsChart, self).__init__(self.fig)
        self.setMinimumSize(400, 300)
    
    def update_plot(self, cpu_percent, memory_percent):
        """Update the plot with new data points."""
        current_time = datetime.now()
        
        # Add new data
        self.timestamps.append(current_time)
        self.cpu_values.append(cpu_percent)
        self.memory_values.append(memory_percent)
        
        # Limit data to the last 60 seconds
        max_points = 60
        if len(self.timestamps) > max_points:
            self.timestamps = self.timestamps[-max_points:]
            self.cpu_values = self.cpu_values[-max_points:]
            self.memory_values = self.memory_values[-max_points:]
        
        # Format time axis
        self.axes.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
        
        # Update plot data
        self.cpu_line.set_data(self.timestamps, self.cpu_values)
        self.memory_line.set_data(self.timestamps, self.memory_values)
        
        # Adjust x-axis limits to show only the last minute
        if len(self.timestamps) > 1:
            self.axes.set_xlim(
                min(self.timestamps), 
                max(self.timestamps) + timedelta(seconds=2)
            )
        
        # Redraw
        self.fig.canvas.draw_idle()
        self.fig.canvas.flush_events()


class SystemHealthWidget(QFrame):
    """Widget displaying system health status as a gauge."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.health_score = 85  # Initial value
        self.setMinimumSize(200, 200)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        
        title = QLabel("System Health")
        title.setAlignment(Qt.AlignCenter)
        title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        
        self.score_label = QLabel(f"{self.health_score}")
        self.score_label.setAlignment(Qt.AlignCenter)
        self.score_label.setFont(QFont("Segoe UI", 24, QFont.Bold))
        
        self.status_label = QLabel("Good")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setFont(QFont("Segoe UI", 10))
        
        layout.addWidget(title)
        layout.addWidget(self.score_label)
        layout.addWidget(self.status_label)
        layout.addStretch()
        
        self.update_health(self.health_score)
    
    def update_health(self, score):
        """Update the health score and status label."""
        self.health_score = score
        self.score_label.setText(f"{score}")
        
        if score >= 80:
            status = "Excellent"
            self.status_label.setStyleSheet(f"color: {COLORS['secondary']}")
        elif score >= 60:
            status = "Good"
            self.status_label.setStyleSheet(f"color: {COLORS['primary']}")
        elif score >= 40:
            status = "Fair"
            self.status_label.setStyleSheet(f"color: {COLORS['warning']}")
        else:
            status = "Poor"
            self.status_label.setStyleSheet(f"color: {COLORS['error']}")
        
        self.status_label.setText(status)
        self.update()
    
    def paintEvent(self, event):
        """Draw the gauge visualization."""
        super().paintEvent(event)
        
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Calculate gauge dimensions
        width = self.width()
        height = self.height()
        center_x = width / 2
        center_y = height * 0.6
        radius = min(width, height) * 0.35
        
        # Draw gauge background
        painter.setPen(QPen(QColor(COLORS['text']), 2))
        painter.drawArc(int(center_x - radius), int(center_y - radius), 
                       int(radius * 2), int(radius * 2), 30 * 16, 120 * 16)
        
        # Draw gauge value
        angle = 210 - (self.health_score / 100 * 150)
        angle_rad = angle * 3.14159 / 180
        
        end_x = center_x + radius * -1 * 0.8 * round(1 - 2 * (angle_rad / 3.14159), 2)
        end_y = center_y - radius * 0.8 * round(1 - 2 * (angle_rad / 3.14159 + 0.5), 2)
        
        if self.health_score >= 80:
            color = QColor(COLORS['secondary'])
        elif self.health_score >= 60:
            color = QColor(COLORS['primary'])
        elif self.health_score >= 40:
            color = QColor(COLORS['warning'])
        else:
            color = QColor(COLORS['error'])
        
        painter.setPen(QPen(color, 3))
        painter.drawLine(int(center_x), int(center_y), int(end_x), int(end_y))
        
        # Draw center point
        painter.setBrush(color)
        painter.drawEllipse(int(center_x - 5), int(center_y - 5), 10, 10)


class DashboardWidget(QWidget):
    """Main dashboard widget displaying system metrics and health status."""
    
    def __init__(self):
        super().__init__()
        self.system_info = SystemInfo()
        self.initUI()
        
        # Set up timer for real-time updates
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_metrics)
        self.update_timer.start(1000)  # Update every second
    
    def initUI(self):
        """Initialize the dashboard UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)
        
        # Header
        header = HeaderLabel("System Dashboard")
        subheader = SubHeaderLabel("Real-time monitoring and system statistics")
        
        layout.addWidget(header)
        layout.addWidget(subheader)
        
        # Top metrics cards
        metrics_layout = QGridLayout()
        metrics_layout.setSpacing(20)
        
        # CPU usage card
        self.cpu_card = InfoCard(
            title="CPU Usage",
            icon_name="cpu",
            value="0%",
            description="Current CPU utilization"
        )
        
        # Memory usage card
        self.memory_card = InfoCard(
            title="Memory Usage",
            icon_name="memory",
            value="0 / 0 GB",
            description="RAM utilization"
        )
        
        # Disk usage card
        self.disk_card = InfoCard(
            title="Disk Usage",
            icon_name="disk",
            value="0 / 0 GB",
            description="Primary disk (C:)"
        )
        
        # Temperature card
        self.temp_card = InfoCard(
            title="Temperature",
            icon_name="temperature",
            value="--°C",
            description="CPU temperature"
        )
        
        # Add cards to grid
        metrics_layout.addWidget(self.cpu_card, 0, 0)
        metrics_layout.addWidget(self.memory_card, 0, 1)
        metrics_layout.addWidget(self.disk_card, 0, 2)
        metrics_layout.addWidget(self.temp_card, 0, 3)
        
        layout.addLayout(metrics_layout)
        
        # Charts section
        charts_layout = QHBoxLayout()
        
        # Performance chart
        chart_card = Card()
        chart_layout = QVBoxLayout(chart_card)
        self.performance_chart = MetricsChart(self, width=8, height=4)
        chart_layout.addWidget(self.performance_chart)
        
        # System health gauge
        health_card = Card()
        health_layout = QVBoxLayout(health_card)
        self.health_widget = SystemHealthWidget()
        health_layout.addWidget(self.health_widget)
        
        # Add widgets to charts layout
        charts_layout.addWidget(chart_card, 7)
        charts_layout.addWidget(health_card, 3)
        
        layout.addLayout(charts_layout, 1)
        
        # Bottom metrics
        bottom_layout = QGridLayout()
        bottom_layout.setSpacing(20)
        
        # Network speed card
        self.network_card = InfoCard(
            title="Network",
            icon_name="network",
            value="↓ 0 KB/s ↑ 0 KB/s",
            description="Current network speed"
        )
        
        # Uptime card
        self.uptime_card = InfoCard(
            title="Uptime",
            icon_name="time",
            value="0h 0m",
            description="System uptime"
        )
        
        # Process count card
        self.process_card = InfoCard(
            title="Processes",
            icon_name="process",
            value="0",
            description="Running processes"
        )
        
        # Battery card (if applicable)
        self.battery_card = InfoCard(
            title="Battery",
            icon_name="battery",
            value="--",
            description="Battery status"
        )
        
        # Add cards to grid
        bottom_layout.addWidget(self.network_card, 0, 0)
        bottom_layout.addWidget(self.uptime_card, 0, 1)
        bottom_layout.addWidget(self.process_card, 0, 2)
        bottom_layout.addWidget(self.battery_card, 0, 3)
        
        layout.addLayout(bottom_layout)
        
        # Perform initial update
        self.update_metrics()
    
    @pyqtSlot()
    def update_metrics(self):
        """Update all system metrics displayed on the dashboard."""
        # Get system information
        cpu_percent = self.system_info.get_cpu_usage()
        memory_info = self.system_info.get_memory_info()
        disk_info = self.system_info.get_disk_info()
        network_info = self.system_info.get_network_info()
        uptime_info = self.system_info.get_uptime()
        process_count = self.system_info.get_process_count()
        battery_info = self.system_info.get_battery_info()
        
        # Calculate system health score
        health_score = self.calculate_health_score(
            cpu_percent, 
            memory_info["percent"], 
            disk_info["percent"]
        )
        
        # Update the UI elements
        self.cpu_card.update_value(f"{cpu_percent}%")
        
        self.memory_card.update_value(
            f"{memory_info['used_gb']:.1f} / {memory_info['total_gb']:.1f} GB"
        )
        
        self.disk_card.update_value(
            f"{disk_info['used_gb']:.1f} / {disk_info['total_gb']:.1f} GB"
        )
        
        # Try to get CPU temperature if available
        try:
            temp = self.system_info.get_cpu_temperature()
            self.temp_card.update_value(f"{temp}°C")
        except:
            self.temp_card.update_value("N/A")
        
        self.network_card.update_value(
            f"↓ {network_info['download']} ↑ {network_info['upload']}"
        )
        
        self.uptime_card.update_value(uptime_info)
        self.process_card.update_value(str(process_count))
        
        if battery_info:
            self.battery_card.update_value(f"{battery_info['percent']}% ({battery_info['status']})")
        else:
            self.battery_card.update_value("AC Power")
        
        # Update performance chart
        self.performance_chart.update_plot(cpu_percent, memory_info["percent"])
        
        # Update health gauge
        self.health_widget.update_health(health_score)
    
    def calculate_health_score(self, cpu_percent, memory_percent, disk_percent):
        """Calculate a simple system health score based on key metrics."""
        # Lower usage is better, so calculate scores as inverse percentages
        cpu_score = 100 - cpu_percent
        memory_score = 100 - memory_percent
        disk_score = 100 - disk_percent
        
        # Weight the different components
        health_score = (
            cpu_score * 0.4 +    # CPU has the highest weight
            memory_score * 0.4 +  # Memory is equally important
            disk_score * 0.2      # Disk space is less critical immediately
        )
        
        return int(health_score)
    
    def showEvent(self, event):
        """When widget becomes visible, start the update timer."""
        super().showEvent(event)
        self.update_timer.start(1000)
    
    def hideEvent(self, event):
        """When widget is hidden, stop the update timer."""
        super().hideEvent(event)
        self.update_timer.stop()
