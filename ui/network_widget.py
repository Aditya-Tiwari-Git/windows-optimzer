"""
Network widget for the Windows System Optimizer.
This module provides the interface for network diagnostic operations
such as ping tests, traceroute, and network logs.
"""

import concurrent.futures
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QFrame, QPushButton, QLineEdit, QTextEdit,
                            QComboBox, QGroupBox, QFormLayout, QSpinBox,
                            QTabWidget, QMessageBox, QFileDialog)
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from PyQt5.QtGui import QFont

from ..services.network import NetworkDiagnostics
from .loading_screen import LoadingOverlay

class NetworkTask(QThread):
    """Background thread for executing network diagnostic operations."""
    result_ready = pyqtSignal(str)
    task_completed = pyqtSignal(bool, str)  # Success, message
    
    def __init__(self, task_type, target, args=None, parent=None):
        super().__init__(parent)
        self.task_type = task_type
        self.target = target
        self.args = args or {}
        self.network = NetworkDiagnostics()
    
    def run(self):
        """Execute the network diagnostic task."""
        try:
            result = ""
            
            if self.task_type == "ping":
                count = self.args.get("count", 4)
                timeout = self.args.get("timeout", 1000)
                result = self.network.ping_test(self.target, count, timeout)
            
            elif self.task_type == "traceroute":
                max_hops = self.args.get("max_hops", 30)
                timeout = self.args.get("timeout", 1000)
                result = self.network.traceroute(self.target, max_hops, timeout)
            
            elif self.task_type == "dns_lookup":
                result = self.network.dns_lookup(self.target)
            
            elif self.task_type == "port_scan":
                ports = self.args.get("ports", [80, 443, 22, 21])
                result = self.network.port_scan(self.target, ports)
            
            elif self.task_type == "network_log":
                duration = self.args.get("duration", 10)
                result = self.network.capture_network_log(self.target, duration)
            
            self.result_ready.emit(result)
            self.task_completed.emit(True, "Task completed successfully")
        
        except Exception as e:
            error_msg = f"Error during {self.task_type}: {str(e)}"
            self.result_ready.emit(error_msg)
            self.task_completed.emit(False, error_msg)


class NetworkWidget(QWidget):
    """Widget for network diagnostic operations."""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.network = NetworkDiagnostics()
        
        # Set up main layout
        main_layout = QVBoxLayout(self)
        
        # Header
        header_label = QLabel("Network Diagnostics")
        header_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        main_layout.addWidget(header_label)
        
        description_label = QLabel(
            "Diagnose network issues with these tools. Enter a domain name or IP address to start."
        )
        description_label.setWordWrap(True)
        main_layout.addWidget(description_label)
        
        # Target input
        target_layout = QHBoxLayout()
        target_label = QLabel("Target:")
        target_layout.addWidget(target_label)
        
        self.target_input = QLineEdit()
        self.target_input.setPlaceholderText("Enter domain or IP (e.g. google.com, 8.8.8.8)")
        target_layout.addWidget(self.target_input)
        
        main_layout.addLayout(target_layout)
        
        # Tab widget for different tools
        self.tab_widget = QTabWidget()
        
        # Ping tab
        self.ping_tab = self.create_ping_tab()
        self.tab_widget.addTab(self.ping_tab, "Ping Test")
        
        # Traceroute tab
        self.traceroute_tab = self.create_traceroute_tab()
        self.tab_widget.addTab(self.traceroute_tab, "Traceroute")
        
        # DNS tab
        self.dns_tab = self.create_dns_tab()
        self.tab_widget.addTab(self.dns_tab, "DNS Lookup")
        
        # Port scan tab
        self.port_scan_tab = self.create_port_scan_tab()
        self.tab_widget.addTab(self.port_scan_tab, "Port Scan")
        
        # Network log tab
        self.network_log_tab = self.create_network_log_tab()
        self.tab_widget.addTab(self.network_log_tab, "Network Log")
        
        main_layout.addWidget(self.tab_widget)
        
        # Loading overlay
        self.loading_overlay = LoadingOverlay(self)
        self.loading_overlay.hide()
    
    def create_ping_tab(self):
        """Create the ping test tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Options group
        options_group = QGroupBox("Ping Options")
        options_layout = QFormLayout(options_group)
        
        self.ping_count = QSpinBox()
        self.ping_count.setRange(1, 100)
        self.ping_count.setValue(4)
        options_layout.addRow("Packet Count:", self.ping_count)
        
        self.ping_timeout = QSpinBox()
        self.ping_timeout.setRange(100, 10000)
        self.ping_timeout.setValue(1000)
        self.ping_timeout.setSingleStep(100)
        options_layout.addRow("Timeout (ms):", self.ping_timeout)
        
        layout.addWidget(options_group)
        
        # Results area
        self.ping_results = QTextEdit()
        self.ping_results.setReadOnly(True)
        self.ping_results.setMinimumHeight(200)
        layout.addWidget(self.ping_results)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.ping_button = QPushButton("Start Ping Test")
        self.ping_button.clicked.connect(self.start_ping_test)
        button_layout.addWidget(self.ping_button)
        
        self.save_ping_button = QPushButton("Save Results")
        self.save_ping_button.clicked.connect(lambda: self.save_results(self.ping_results))
        button_layout.addWidget(self.save_ping_button)
        
        layout.addLayout(button_layout)
        
        return tab
    
    def create_traceroute_tab(self):
        """Create the traceroute tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Options group
        options_group = QGroupBox("Traceroute Options")
        options_layout = QFormLayout(options_group)
        
        self.traceroute_max_hops = QSpinBox()
        self.traceroute_max_hops.setRange(1, 64)
        self.traceroute_max_hops.setValue(30)
        options_layout.addRow("Max Hops:", self.traceroute_max_hops)
        
        self.traceroute_timeout = QSpinBox()
        self.traceroute_timeout.setRange(100, 10000)
        self.traceroute_timeout.setValue(1000)
        self.traceroute_timeout.setSingleStep(100)
        options_layout.addRow("Timeout (ms):", self.traceroute_timeout)
        
        layout.addWidget(options_group)
        
        # Results area
        self.traceroute_results = QTextEdit()
        self.traceroute_results.setReadOnly(True)
        self.traceroute_results.setMinimumHeight(200)
        layout.addWidget(self.traceroute_results)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.traceroute_button = QPushButton("Start Traceroute")
        self.traceroute_button.clicked.connect(self.start_traceroute)
        button_layout.addWidget(self.traceroute_button)
        
        self.save_traceroute_button = QPushButton("Save Results")
        self.save_traceroute_button.clicked.connect(lambda: self.save_results(self.traceroute_results))
        button_layout.addWidget(self.save_traceroute_button)
        
        layout.addLayout(button_layout)
        
        return tab
    
    def create_dns_tab(self):
        """Create the DNS lookup tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Results area
        self.dns_results = QTextEdit()
        self.dns_results.setReadOnly(True)
        self.dns_results.setMinimumHeight(200)
        layout.addWidget(self.dns_results)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.dns_button = QPushButton("Start DNS Lookup")
        self.dns_button.clicked.connect(self.start_dns_lookup)
        button_layout.addWidget(self.dns_button)
        
        self.save_dns_button = QPushButton("Save Results")
        self.save_dns_button.clicked.connect(lambda: self.save_results(self.dns_results))
        button_layout.addWidget(self.save_dns_button)
        
        layout.addLayout(button_layout)
        
        return tab
    
    def create_port_scan_tab(self):
        """Create the port scan tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Options group
        options_group = QGroupBox("Port Scan Options")
        options_layout = QFormLayout(options_group)
        
        self.port_scan_combo = QComboBox()
        self.port_scan_combo.addItem("Common Ports", [80, 443, 22, 21, 25, 110, 143, 3389])
        self.port_scan_combo.addItem("Web Ports", [80, 443, 8080, 8443])
        self.port_scan_combo.addItem("Mail Ports", [25, 110, 143, 465, 587, 993, 995])
        self.port_scan_combo.addItem("Remote Access", [22, 23, 3389, 5900])
        
        options_layout.addRow("Port Selection:", self.port_scan_combo)
        
        self.custom_ports = QLineEdit()
        self.custom_ports.setPlaceholderText("Enter custom ports separated by commas (e.g. 80,443,8080)")
        options_layout.addRow("Custom Ports:", self.custom_ports)
        
        layout.addWidget(options_group)
        
        # Results area
        self.port_scan_results = QTextEdit()
        self.port_scan_results.setReadOnly(True)
        self.port_scan_results.setMinimumHeight(200)
        layout.addWidget(self.port_scan_results)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.port_scan_button = QPushButton("Start Port Scan")
        self.port_scan_button.clicked.connect(self.start_port_scan)
        button_layout.addWidget(self.port_scan_button)
        
        self.save_port_scan_button = QPushButton("Save Results")
        self.save_port_scan_button.clicked.connect(lambda: self.save_results(self.port_scan_results))
        button_layout.addWidget(self.save_port_scan_button)
        
        layout.addLayout(button_layout)
        
        return tab
    
    def create_network_log_tab(self):
        """Create the network log tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Options group
        options_group = QGroupBox("Network Log Options")
        options_layout = QFormLayout(options_group)
        
        self.log_duration = QSpinBox()
        self.log_duration.setRange(5, 300)
        self.log_duration.setValue(10)
        options_layout.addRow("Duration (seconds):", self.log_duration)
        
        layout.addWidget(options_group)
        
        # Results area
        self.network_log_results = QTextEdit()
        self.network_log_results.setReadOnly(True)
        self.network_log_results.setMinimumHeight(200)
        layout.addWidget(self.network_log_results)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.network_log_button = QPushButton("Start Network Log")
        self.network_log_button.clicked.connect(self.start_network_log)
        button_layout.addWidget(self.network_log_button)
        
        self.save_network_log_button = QPushButton("Save Results")
        self.save_network_log_button.clicked.connect(lambda: self.save_results(self.network_log_results))
        button_layout.addWidget(self.save_network_log_button)
        
        layout.addLayout(button_layout)
        
        return tab
    
    def start_ping_test(self):
        """Start a ping test."""
        target = self.target_input.text().strip()
        if not target:
            QMessageBox.warning(self, "Missing Target", "Please enter a domain or IP address.")
            return
        
        self.loading_overlay.show()
        self.loading_overlay.set_message(f"Pinging {target}...")
        
        count = self.ping_count.value()
        timeout = self.ping_timeout.value()
        
        self.task = NetworkTask("ping", target, {"count": count, "timeout": timeout})
        self.task.result_ready.connect(self.ping_results.setText)
        self.task.task_completed.connect(self.task_finished)
        self.task.start()
    
    def start_traceroute(self):
        """Start a traceroute."""
        target = self.target_input.text().strip()
        if not target:
            QMessageBox.warning(self, "Missing Target", "Please enter a domain or IP address.")
            return
        
        self.loading_overlay.show()
        self.loading_overlay.set_message(f"Tracing route to {target}...")
        
        max_hops = self.traceroute_max_hops.value()
        timeout = self.traceroute_timeout.value()
        
        self.task = NetworkTask("traceroute", target, {"max_hops": max_hops, "timeout": timeout})
        self.task.result_ready.connect(self.traceroute_results.setText)
        self.task.task_completed.connect(self.task_finished)
        self.task.start()
    
    def start_dns_lookup(self):
        """Start a DNS lookup."""
        target = self.target_input.text().strip()
        if not target:
            QMessageBox.warning(self, "Missing Target", "Please enter a domain or IP address.")
            return
        
        self.loading_overlay.show()
        self.loading_overlay.set_message(f"Looking up DNS for {target}...")
        
        self.task = NetworkTask("dns_lookup", target)
        self.task.result_ready.connect(self.dns_results.setText)
        self.task.task_completed.connect(self.task_finished)
        self.task.start()
    
    def start_port_scan(self):
        """Start a port scan."""
        target = self.target_input.text().strip()
        if not target:
            QMessageBox.warning(self, "Missing Target", "Please enter a domain or IP address.")
            return
        
        self.loading_overlay.show()
        self.loading_overlay.set_message(f"Scanning ports on {target}...")
        
        # Get selected ports
        ports = []
        if self.custom_ports.text().strip():
            try:
                ports = [int(p.strip()) for p in self.custom_ports.text().split(",")]
            except ValueError:
                QMessageBox.warning(self, "Invalid Ports", "Please enter valid port numbers separated by commas.")
                self.loading_overlay.hide()
                return
        else:
            # Use the selected preset
            ports = self.port_scan_combo.currentData()
        
        self.task = NetworkTask("port_scan", target, {"ports": ports})
        self.task.result_ready.connect(self.port_scan_results.setText)
        self.task.task_completed.connect(self.task_finished)
        self.task.start()
    
    def start_network_log(self):
        """Start a network log capture."""
        target = self.target_input.text().strip()
        if not target:
            # For network log, target is optional
            target = "*"  # Capture all network activity
        
        self.loading_overlay.show()
        self.loading_overlay.set_message("Capturing network log...")
        
        duration = self.log_duration.value()
        
        self.task = NetworkTask("network_log", target, {"duration": duration})
        self.task.result_ready.connect(self.network_log_results.setText)
        self.task.task_completed.connect(self.task_finished)
        self.task.start()
    
    def task_finished(self, success, message):
        """Handle completion of network task."""
        self.loading_overlay.hide()
        
        if not success:
            QMessageBox.warning(self, "Error", message)
    
    def save_results(self, text_edit):
        """Save the results to a file."""
        if not text_edit.toPlainText():
            QMessageBox.warning(self, "No Results", "There are no results to save.")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Results", "", "Text Files (*.txt);;All Files (*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    f.write(text_edit.toPlainText())
                QMessageBox.information(self, "Save Successful", f"Results saved to {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Save Failed", f"Error saving file: {str(e)}")
