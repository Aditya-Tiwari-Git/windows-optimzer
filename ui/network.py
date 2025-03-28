"""
Network diagnostics widget for performing network tests and analysis.
"""

import os
import time
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QFrame, QLineEdit, QTextEdit,
                             QTabWidget, QGridLayout, QGroupBox, QComboBox,
                             QSpinBox, QCheckBox, QScrollArea, QSizePolicy,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtGui import QIcon, QFont, QPalette, QColor, QTextCursor
from PyQt5.QtCore import Qt, pyqtSignal, pyqtSlot, QThread

from ui.theme import COLORS
from ui.components import Card, HeaderLabel, SubHeaderLabel, ActionButton
from ui.loading_screen import LoadingScreen
from utils.network_tools import NetworkTools
from utils.background_worker import BackgroundWorker


class ConsoleOutput(QTextEdit):
    """Custom text area for console-like output."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setFont(QFont("Consolas", 9))
        self.setStyleSheet(
            f"background-color: #1E1E1E; color: #F1F1F1; "
            f"border-radius: 4px; padding: 5px;"
        )
    
    def append_output(self, text, color=None):
        """Append text to the console with optional color."""
        self.moveCursor(QTextCursor.End)
        
        if color:
            self.insertHtml(f'<span style="color:{color}">{text}</span><br>')
        else:
            self.insertPlainText(text + "\n")
        
        # Auto scroll to bottom
        self.moveCursor(QTextCursor.End)
        self.ensureCursorVisible()


class NetworkWidget(QWidget):
    """Widget for network diagnostics tools."""
    
    def __init__(self):
        super().__init__()
        self.network_tools = NetworkTools()
        self.initUI()
        
        # Loading screen instance
        self.loading_screen = None
    
    def initUI(self):
        """Initialize the network diagnostics UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)
        
        # Header
        header = HeaderLabel("Network Diagnostics")
        subheader = SubHeaderLabel("Analyze and troubleshoot your network connections")
        
        layout.addWidget(header)
        layout.addWidget(subheader)
        
        # Network info card
        info_card = Card()
        info_layout = QGridLayout(info_card)
        
        # IP Address
        ip_label = QLabel("IP Address:")
        ip_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.ip_value = QLabel("Fetching...")
        info_layout.addWidget(ip_label, 0, 0)
        info_layout.addWidget(self.ip_value, 0, 1)
        
        # MAC Address
        mac_label = QLabel("MAC Address:")
        mac_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.mac_value = QLabel("Fetching...")
        info_layout.addWidget(mac_label, 0, 2)
        info_layout.addWidget(self.mac_value, 0, 3)
        
        # Default Gateway
        gateway_label = QLabel("Default Gateway:")
        gateway_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.gateway_value = QLabel("Fetching...")
        info_layout.addWidget(gateway_label, 1, 0)
        info_layout.addWidget(self.gateway_value, 1, 1)
        
        # DNS Servers
        dns_label = QLabel("DNS Servers:")
        dns_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.dns_value = QLabel("Fetching...")
        info_layout.addWidget(dns_label, 1, 2)
        info_layout.addWidget(self.dns_value, 1, 3)
        
        # Network Adapter
        adapter_label = QLabel("Network Adapter:")
        adapter_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.adapter_value = QLabel("Fetching...")
        info_layout.addWidget(adapter_label, 2, 0)
        info_layout.addWidget(self.adapter_value, 2, 1)
        
        # Connection Type
        conn_label = QLabel("Connection Type:")
        conn_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.conn_value = QLabel("Fetching...")
        info_layout.addWidget(conn_label, 2, 2)
        info_layout.addWidget(self.conn_value, 2, 3)
        
        # Refresh button
        refresh_btn = ActionButton("Refresh Network Info", "refresh", "primary")
        refresh_btn.clicked.connect(self.update_network_info)
        info_layout.addWidget(refresh_btn, 3, 0, 1, 4, Qt.AlignCenter)
        
        layout.addWidget(info_card)
        
        # Tabs for different diagnostic tools
        tabs = QTabWidget()
        tabs.setObjectName("network_tabs")
        
        # Ping Tab
        ping_tab = QWidget()
        ping_layout = QVBoxLayout(ping_tab)
        ping_layout.setContentsMargins(15, 15, 15, 15)
        
        # Target input and options
        target_group = QGroupBox("Ping Target")
        target_layout = QGridLayout(target_group)
        
        target_layout.addWidget(QLabel("Host/IP:"), 0, 0)
        self.ping_host_input = QLineEdit()
        self.ping_host_input.setText("8.8.8.8")  # Default to Google DNS
        target_layout.addWidget(self.ping_host_input, 0, 1)
        
        target_layout.addWidget(QLabel("Count:"), 1, 0)
        self.ping_count = QSpinBox()
        self.ping_count.setRange(1, 100)
        self.ping_count.setValue(4)
        target_layout.addWidget(self.ping_count, 1, 1)
        
        target_layout.addWidget(QLabel("Timeout (ms):"), 2, 0)
        self.ping_timeout = QSpinBox()
        self.ping_timeout.setRange(100, 10000)
        self.ping_timeout.setValue(1000)
        self.ping_timeout.setSingleStep(100)
        target_layout.addWidget(self.ping_timeout, 2, 1)
        
        self.ping_continuous = QCheckBox("Continuous ping")
        self.ping_continuous.setChecked(False)
        target_layout.addWidget(self.ping_continuous, 3, 0, 1, 2)
        
        ping_layout.addWidget(target_group)
        
        # Ping output
        self.ping_output = ConsoleOutput()
        ping_layout.addWidget(self.ping_output, 1)
        
        # Ping actions
        ping_actions = QHBoxLayout()
        
        self.ping_start_btn = ActionButton("Start Ping", "play", "primary")
        self.ping_start_btn.clicked.connect(self.start_ping)
        
        self.ping_stop_btn = ActionButton("Stop", "stop", "error")
        self.ping_stop_btn.clicked.connect(self.stop_ping)
        self.ping_stop_btn.setEnabled(False)
        
        ping_clear_btn = ActionButton("Clear Results", "clear", "warning")
        ping_clear_btn.clicked.connect(lambda: self.ping_output.clear())
        
        ping_actions.addStretch()
        ping_actions.addWidget(self.ping_start_btn)
        ping_actions.addWidget(self.ping_stop_btn)
        ping_actions.addWidget(ping_clear_btn)
        
        ping_layout.addLayout(ping_actions)
        
        # Traceroute Tab
        tracert_tab = QWidget()
        tracert_layout = QVBoxLayout(tracert_tab)
        tracert_layout.setContentsMargins(15, 15, 15, 15)
        
        # Traceroute target input
        tracert_group = QGroupBox("Traceroute Target")
        tracert_group_layout = QGridLayout(tracert_group)
        
        tracert_group_layout.addWidget(QLabel("Host/IP:"), 0, 0)
        self.tracert_host_input = QLineEdit()
        self.tracert_host_input.setText("www.google.com")
        tracert_group_layout.addWidget(self.tracert_host_input, 0, 1)
        
        tracert_group_layout.addWidget(QLabel("Max Hops:"), 1, 0)
        self.tracert_max_hops = QSpinBox()
        self.tracert_max_hops.setRange(1, 64)
        self.tracert_max_hops.setValue(30)
        tracert_group_layout.addWidget(self.tracert_max_hops, 1, 1)
        
        tracert_group_layout.addWidget(QLabel("Timeout (ms):"), 2, 0)
        self.tracert_timeout = QSpinBox()
        self.tracert_timeout.setRange(100, 10000)
        self.tracert_timeout.setValue(1000)
        self.tracert_timeout.setSingleStep(100)
        tracert_group_layout.addWidget(self.tracert_timeout, 2, 1)
        
        tracert_layout.addWidget(tracert_group)
        
        # Traceroute output
        self.tracert_output = ConsoleOutput()
        tracert_layout.addWidget(self.tracert_output, 1)
        
        # Traceroute actions
        tracert_actions = QHBoxLayout()
        
        self.tracert_start_btn = ActionButton("Start Traceroute", "play", "primary")
        self.tracert_start_btn.clicked.connect(self.start_traceroute)
        
        tracert_clear_btn = ActionButton("Clear Results", "clear", "warning")
        tracert_clear_btn.clicked.connect(lambda: self.tracert_output.clear())
        
        tracert_actions.addStretch()
        tracert_actions.addWidget(self.tracert_start_btn)
        tracert_actions.addWidget(tracert_clear_btn)
        
        tracert_layout.addLayout(tracert_actions)
        
        # Network Log Tab
        log_tab = QWidget()
        log_layout = QVBoxLayout(log_tab)
        log_layout.setContentsMargins(15, 15, 15, 15)
        
        # Network log options
        log_group = QGroupBox("Network Traffic Log")
        log_group_layout = QGridLayout(log_group)
        
        log_group_layout.addWidget(QLabel("Interface:"), 0, 0)
        self.interface_combo = QComboBox()
        self.interface_combo.addItem("Default")
        log_group_layout.addWidget(self.interface_combo, 0, 1)
        
        log_group_layout.addWidget(QLabel("Duration (seconds):"), 1, 0)
        self.capture_duration = QSpinBox()
        self.capture_duration.setRange(5, 300)
        self.capture_duration.setValue(30)
        log_group_layout.addWidget(self.capture_duration, 1, 1)
        
        self.capture_dns = QCheckBox("Include DNS queries")
        self.capture_dns.setChecked(True)
        log_group_layout.addWidget(self.capture_dns, 2, 0)
        
        self.capture_http = QCheckBox("Include HTTP traffic")
        self.capture_http.setChecked(True)
        log_group_layout.addWidget(self.capture_http, 2, 1)
        
        log_layout.addWidget(log_group)
        
        # Network log table
        self.log_table = QTableWidget()
        self.log_table.setColumnCount(5)
        self.log_table.setHorizontalHeaderLabels(["Time", "Source", "Destination", "Protocol", "Info"])
        self.log_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        log_layout.addWidget(self.log_table, 1)
        
        # Network log actions
        log_actions = QHBoxLayout()
        
        self.capture_start_btn = ActionButton("Start Capture", "play", "primary")
        self.capture_start_btn.clicked.connect(self.start_network_capture)
        
        log_clear_btn = ActionButton("Clear Log", "clear", "warning")
        log_clear_btn.clicked.connect(lambda: self.log_table.setRowCount(0))
        
        log_export_btn = ActionButton("Export Log", "download", "secondary")
        log_export_btn.clicked.connect(self.export_network_log)
        
        log_actions.addStretch()
        log_actions.addWidget(self.capture_start_btn)
        log_actions.addWidget(log_clear_btn)
        log_actions.addWidget(log_export_btn)
        
        log_layout.addLayout(log_actions)
        
        # Add tabs
        tabs.addTab(ping_tab, "Ping Test")
        tabs.addTab(tracert_tab, "Traceroute")
        tabs.addTab(log_tab, "Network Log")
        
        layout.addWidget(tabs, 1)
        
        # Initialize network information
        self.update_network_info()
        self.update_interfaces()
        
        # Set up ping worker attributes
        self.ping_worker = None
        self.is_ping_running = False
    
    def update_network_info(self):
        """Update network information display."""
        self.show_loading("Fetching Network Information...", "Querying system...")
        
        # Create background worker
        worker = BackgroundWorker(self.network_tools.get_network_info)
        worker.taskFinished.connect(self.on_network_info_update)
        worker.start()
    
    def on_network_info_update(self, info):
        """Handle completion of network info update."""
        self.hide_loading()
        
        # Update UI with network information
        self.ip_value.setText(info.get('ip_address', 'N/A'))
        self.mac_value.setText(info.get('mac_address', 'N/A'))
        self.gateway_value.setText(info.get('default_gateway', 'N/A'))
        self.dns_value.setText(', '.join(info.get('dns_servers', ['N/A'])))
        self.adapter_value.setText(info.get('adapter_name', 'N/A'))
        self.conn_value.setText(info.get('connection_type', 'N/A'))
    
    def update_interfaces(self):
        """Update network interfaces list."""
        interfaces = self.network_tools.get_interface_list()
        
        self.interface_combo.clear()
        self.interface_combo.addItem("Default")
        
        for interface in interfaces:
            self.interface_combo.addItem(interface)
    
    def start_ping(self):
        """Start ping operation."""
        host = self.ping_host_input.text().strip()
        if not host:
            self.ping_output.append_output("Error: Please enter a valid host or IP address.", COLORS['error'])
            return
        
        count = self.ping_count.value()
        timeout = self.ping_timeout.value()
        continuous = self.ping_continuous.isChecked()
        
        self.ping_output.append_output(f"Starting ping to {host}...", COLORS['primary'])
        self.ping_start_btn.setEnabled(False)
        self.ping_stop_btn.setEnabled(True)
        
        # Create worker thread for ping
        self.is_ping_running = True
        
        # Create background worker
        worker = BackgroundWorker(
            self.network_tools.ping,
            host,
            count,
            timeout,
            continuous
        )
        worker.taskUpdate.connect(self.on_ping_update)
        worker.taskFinished.connect(self.on_ping_complete)
        worker.start()
        
        self.ping_worker = worker
    
    def on_ping_update(self, data):
        """Handle ping update."""
        if "Reply from" in data:
            self.ping_output.append_output(data, COLORS['secondary'])
        elif "Request timed out" in data:
            self.ping_output.append_output(data, COLORS['warning'])
        else:
            self.ping_output.append_output(data, COLORS['text'])
    
    def on_ping_complete(self, summary):
        """Handle ping completion."""
        if self.is_ping_running:  # Only update if not manually stopped
            self.ping_output.append_output("\nPing statistics:", COLORS['primary'])
            
            for line in summary:
                if "loss" in line and "0%" not in line:
                    self.ping_output.append_output(line, COLORS['warning'])
                else:
                    self.ping_output.append_output(line)
        
        self.ping_start_btn.setEnabled(True)
        self.ping_stop_btn.setEnabled(False)
        self.is_ping_running = False
    
    def stop_ping(self):
        """Stop the ongoing ping operation."""
        if self.ping_worker and self.is_ping_running:
            self.is_ping_running = False
            self.ping_worker.terminate()
            self.ping_output.append_output("\nPing operation stopped by user.", COLORS['error'])
            
            self.ping_start_btn.setEnabled(True)
            self.ping_stop_btn.setEnabled(False)
    
    def start_traceroute(self):
        """Start traceroute operation."""
        host = self.tracert_host_input.text().strip()
        if not host:
            self.tracert_output.append_output("Error: Please enter a valid host or IP address.", COLORS['error'])
            return
        
        max_hops = self.tracert_max_hops.value()
        timeout = self.tracert_timeout.value()
        
        self.tracert_output.append_output(f"Starting traceroute to {host}...", COLORS['primary'])
        self.tracert_start_btn.setEnabled(False)
        
        # Create background worker
        worker = BackgroundWorker(
            self.network_tools.traceroute,
            host,
            max_hops,
            timeout
        )
        worker.taskUpdate.connect(self.on_traceroute_update)
        worker.taskFinished.connect(self.on_traceroute_complete)
        worker.start()
    
    def on_traceroute_update(self, data):
        """Handle traceroute update."""
        if "Request timed out" in data:
            self.tracert_output.append_output(data, COLORS['warning'])
        elif data.startswith("Tracing") or data.startswith("Trace"):
            self.tracert_output.append_output(data, COLORS['primary'])
        else:
            self.tracert_output.append_output(data)
    
    def on_traceroute_complete(self, message):
        """Handle traceroute completion."""
        self.tracert_output.append_output("\nTraceroute complete.", COLORS['secondary'])
        self.tracert_start_btn.setEnabled(True)
    
    def start_network_capture(self):
        """Start network traffic capture."""
        duration = self.capture_duration.value()
        interface = self.interface_combo.currentText()
        include_dns = self.capture_dns.isChecked()
        include_http = self.capture_http.isChecked()
        
        if interface == "Default":
            interface = None
        
        self.show_loading(f"Capturing Network Traffic...", f"Duration: {duration} seconds")
        self.capture_start_btn.setEnabled(False)
        
        # Clear existing table
        self.log_table.setRowCount(0)
        
        # Create background worker
        worker = BackgroundWorker(
            self.network_tools.capture_traffic,
            duration,
            interface,
            include_dns,
            include_http
        )
        worker.taskFinished.connect(self.on_capture_complete)
        worker.start()
    
    def on_capture_complete(self, packets):
        """Handle completion of network capture."""
        self.hide_loading()
        self.capture_start_btn.setEnabled(True)
        
        # Fill table with packet data
        self.log_table.setRowCount(len(packets))
        
        for i, packet in enumerate(packets):
            self.log_table.setItem(i, 0, QTableWidgetItem(packet.get('time', '')))
            self.log_table.setItem(i, 1, QTableWidgetItem(packet.get('source', '')))
            self.log_table.setItem(i, 2, QTableWidgetItem(packet.get('destination', '')))
            self.log_table.setItem(i, 3, QTableWidgetItem(packet.get('protocol', '')))
            self.log_table.setItem(i, 4, QTableWidgetItem(packet.get('info', '')))
    
    def export_network_log(self):
        """Export network log to a file."""
        if self.log_table.rowCount() == 0:
            return
        
        try:
            path = os.path.join(os.path.expanduser("~"), "Documents", "NetworkCapture.csv")
            
            with open(path, 'w') as f:
                f.write("Time,Source,Destination,Protocol,Info\n")
                
                for row in range(self.log_table.rowCount()):
                    line = []
                    for col in range(self.log_table.columnCount()):
                        item = self.log_table.item(row, col)
                        text = item.text() if item else ""
                        # Escape commas and quotes for CSV
                        if "," in text or '"' in text:
                            text = f'"{text.replace(\'"\', \'""\')}\"'
                        line.append(text)
                    
                    f.write(','.join(line) + '\n')
            
            self.show_message(f"Network log exported to {path}", "success")
        except Exception as e:
            self.show_message(f"Error exporting log: {str(e)}", "error")
    
    def show_message(self, message, status="info"):
        """Show a brief message at the bottom of the screen."""
        color = COLORS['text']
        
        if status == "success":
            color = COLORS['secondary']
        elif status == "error":
            color = COLORS['error']
        elif status == "warning":
            color = COLORS['warning']
        elif status == "info":
            color = COLORS['primary']
        
        # In a real implementation, this would show a toast message
        # For now, we'll just print to console
        print(f"[{status.upper()}] {message}")
    
    def show_loading(self, title, message):
        """Show loading screen."""
        if not self.loading_screen:
            self.loading_screen = LoadingScreen(self)
        
        self.loading_screen.set_message(title, message)
        self.loading_screen.show()
    
    def hide_loading(self):
        """Hide loading screen."""
        if self.loading_screen:
            self.loading_screen.hide()
