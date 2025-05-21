"""
Optimization widget for the Windows System Optimizer.
This module provides the interface for system optimization operations
including quick fix tools for Teams and Outlook, and driver updates.
"""

import concurrent.futures
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QFrame, QPushButton, QProgressBar,
                            QGroupBox, QGridLayout, QSpacerItem, QSizePolicy,
                            QTabWidget, QMessageBox, QCheckBox, QListWidget,
                            QListWidgetItem)
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QIcon

from services.quickfix import QuickFixTools
from services.driver_updater import DriverUpdater
from services.monitor import SystemMonitor
from .loading_screen import LoadingOverlay

class OptimizationTask(QThread):
    """Background thread for executing optimization tasks."""
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    completed = pyqtSignal(bool, str)  # Success, message
    
    def __init__(self, task_type, parent=None):
        super().__init__(parent)
        self.task_type = task_type
        self.quick_fix = QuickFixTools()
        self.driver_updater = DriverUpdater()
    
    def run(self):
        """Execute the optimization task."""
        try:
            self.status_updated.emit(f"Starting {self.task_type}...")
            
            if self.task_type == "fix_teams":
                self.progress_updated.emit(10)
                self.status_updated.emit("Stopping Microsoft Teams processes...")
                self.quick_fix.fix_teams()
                self.progress_updated.emit(100)
            
            elif self.task_type == "fix_outlook":
                self.progress_updated.emit(10)
                self.status_updated.emit("Repairing Microsoft Outlook profile...")
                self.quick_fix.fix_outlook()
                self.progress_updated.emit(100)
            
            elif self.task_type == "scan_drivers":
                self.status_updated.emit("Scanning for outdated drivers...")
                self.progress_updated.emit(20)
                outdated_drivers = self.driver_updater.check_drivers()
                self.progress_updated.emit(100)
                self.completed.emit(True, str(outdated_drivers))
                return
            
            elif self.task_type == "update_drivers":
                self.status_updated.emit("Updating drivers...")
                self.progress_updated.emit(5)
                
                # This would be split into steps in a real implementation
                self.driver_updater.update_drivers()
                self.progress_updated.emit(100)
            
            self.completed.emit(True, f"{self.task_type} completed successfully")
        
        except Exception as e:
            self.completed.emit(False, f"Error during {self.task_type}: {str(e)}")


class OptimizationWidget(QWidget):
    """Widget for system optimization operations."""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.monitor = SystemMonitor()
        self.quick_fix = QuickFixTools()
        self.driver_updater = DriverUpdater()
        
        # Set up main layout
        main_layout = QVBoxLayout(self)
        
        # Header
        header_label = QLabel("System Optimization")
        header_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        main_layout.addWidget(header_label)
        
        description_label = QLabel(
            "Optimize your system performance and fix common issues with Microsoft applications."
        )
        description_label.setWordWrap(True)
        main_layout.addWidget(description_label)
        
        # Tab widget
        self.tab_widget = QTabWidget()
        
        # Quick fix tab
        self.quick_fix_tab = self.create_quick_fix_tab()
        self.tab_widget.addTab(self.quick_fix_tab, "Quick Fix Tools")
        
        # Driver updater tab
        self.driver_tab = self.create_driver_tab()
        self.tab_widget.addTab(self.driver_tab, "Driver Updates")
        
        main_layout.addWidget(self.tab_widget)
        
        # Progress area
        progress_group = QGroupBox("Task Progress")
        progress_layout = QVBoxLayout(progress_group)
        
        self.status_label = QLabel("Ready")
        progress_layout.addWidget(self.status_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        main_layout.addWidget(progress_group)
        
        # Loading overlay
        self.loading_overlay = LoadingOverlay(self)
        self.loading_overlay.hide()
        
        # Scan drivers on initialization
        # self.scan_drivers()
    
    def create_quick_fix_tab(self):
        """Create the quick fix tools tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Microsoft Teams fix card
        teams_group = QGroupBox("Microsoft Teams Quick Fix")
        teams_layout = QVBoxLayout(teams_group)
        
        teams_desc = QLabel(
            "Fix common Microsoft Teams issues by resetting the app cache and configuration. "
            "This will close Teams if it's running."
        )
        teams_desc.setWordWrap(True)
        teams_layout.addWidget(teams_desc)
        
        teams_issues_label = QLabel("Issues this may fix:")
        teams_layout.addWidget(teams_issues_label)
        
        teams_issues = QLabel(
            "• Teams won't start or crashes\n"
            "• Video/audio not working correctly\n"
            "• Missing messages or contacts\n"
            "• General performance issues"
        )
        teams_issues.setIndent(20)
        teams_layout.addWidget(teams_issues)
        
        teams_button = QPushButton("Fix Microsoft Teams")
        teams_button.clicked.connect(self.fix_teams)
        teams_layout.addWidget(teams_button)
        
        layout.addWidget(teams_group)
        
        # Microsoft Outlook fix card
        outlook_group = QGroupBox("Microsoft Outlook Quick Fix")
        outlook_layout = QVBoxLayout(outlook_group)
        
        outlook_desc = QLabel(
            "Fix common Microsoft Outlook issues by repairing the Outlook profile and data files. "
            "This will close Outlook if it's running."
        )
        outlook_desc.setWordWrap(True)
        outlook_layout.addWidget(outlook_desc)
        
        outlook_issues_label = QLabel("Issues this may fix:")
        outlook_layout.addWidget(outlook_issues_label)
        
        outlook_issues = QLabel(
            "• Outlook won't start or crashes\n"
            "• Search not working correctly\n"
            "• Calendar issues\n"
            "• Slow performance or hanging"
        )
        outlook_issues.setIndent(20)
        outlook_layout.addWidget(outlook_issues)
        
        outlook_button = QPushButton("Fix Microsoft Outlook")
        outlook_button.clicked.connect(self.fix_outlook)
        outlook_layout.addWidget(outlook_button)
        
        layout.addWidget(outlook_group)
        layout.addStretch()
        
        return tab
    
    def create_driver_tab(self):
        """Create the driver updates tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Driver updater introduction
        intro_label = QLabel(
            "The driver updater scans your system for outdated drivers "
            "and provides updates to improve system stability and performance."
        )
        intro_label.setWordWrap(True)
        layout.addWidget(intro_label)
        
        # Warning label
        warning_label = QLabel(
            "Warning: Updating drivers can sometimes cause system instability. "
            "It's recommended to create a system restore point before updating drivers."
        )
        warning_label.setStyleSheet("color: #FF9500;")  # Warning color
        warning_label.setWordWrap(True)
        layout.addWidget(warning_label)
        
        # Driver list
        self.driver_list_label = QLabel("Available Driver Updates:")
        layout.addWidget(self.driver_list_label)
        
        self.driver_list = QListWidget()
        self.driver_list.setMinimumHeight(200)
        layout.addWidget(self.driver_list)
        
        # Action buttons
        button_layout = QHBoxLayout()
        
        self.scan_drivers_button = QPushButton("Scan for Driver Updates")
        self.scan_drivers_button.clicked.connect(self.scan_drivers)
        button_layout.addWidget(self.scan_drivers_button)
        
        self.update_drivers_button = QPushButton("Update Selected Drivers")
        self.update_drivers_button.clicked.connect(self.update_drivers)
        self.update_drivers_button.setEnabled(False)
        button_layout.addWidget(self.update_drivers_button)
        
        layout.addLayout(button_layout)
        
        return tab
    
    def fix_teams(self):
        """Start the Microsoft Teams quick fix."""
        confirm = QMessageBox.question(
            self,
            "Confirm Teams Fix",
            "This will close Microsoft Teams if it's running and reset its cache. Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if confirm == QMessageBox.No:
            return
        
        self.loading_overlay.show()
        self.loading_overlay.set_message("Fixing Microsoft Teams...")
        self.progress_bar.setValue(0)
        
        # Start task
        self.task = OptimizationTask("fix_teams")
        self.task.progress_updated.connect(self.progress_bar.setValue)
        self.task.status_updated.connect(self.status_label.setText)
        self.task.completed.connect(self.task_completed)
        self.task.start()
    
    def fix_outlook(self):
        """Start the Microsoft Outlook quick fix."""
        confirm = QMessageBox.question(
            self,
            "Confirm Outlook Fix",
            "This will close Microsoft Outlook if it's running and repair its profile. Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if confirm == QMessageBox.No:
            return
        
        self.loading_overlay.show()
        self.loading_overlay.set_message("Fixing Microsoft Outlook...")
        self.progress_bar.setValue(0)
        
        # Start task
        self.task = OptimizationTask("fix_outlook")
        self.task.progress_updated.connect(self.progress_bar.setValue)
        self.task.status_updated.connect(self.status_label.setText)
        self.task.completed.connect(self.task_completed)
        self.task.start()
    
    def scan_drivers(self):
        """Scan for outdated drivers."""
        self.loading_overlay.show()
        self.loading_overlay.set_message("Scanning for driver updates...")
        self.progress_bar.setValue(0)
        
        # Clear the list
        self.driver_list.clear()
        self.update_drivers_button.setEnabled(False)
        
        # Start task
        self.task = OptimizationTask("scan_drivers")
        self.task.progress_updated.connect(self.progress_bar.setValue)
        self.task.status_updated.connect(self.status_label.setText)
        self.task.completed.connect(self.driver_scan_completed)
        self.task.start()
    
    def driver_scan_completed(self, success, message):
        """Handle completion of driver scan."""
        self.loading_overlay.hide()
        
        if not success:
            QMessageBox.warning(self, "Error", message)
            return
        
        try:
            # Parse the result message as a list of drivers
            driver_list = eval(message)  # Safe in this context as we control the input
            
            if not driver_list:
                self.driver_list.addItem("No outdated drivers found.")
                return
            
            # Add drivers to the list
            for driver in driver_list:
                item = QListWidgetItem(driver["name"])
                item.setData(Qt.UserRole, driver)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)
                self.driver_list.addItem(item)
            
            self.update_drivers_button.setEnabled(True)
            self.status_label.setText(f"Found {len(driver_list)} outdated drivers")
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error processing driver list: {str(e)}")
    
    def update_drivers(self):
        """Update the selected drivers."""
        selected_count = 0
        for i in range(self.driver_list.count()):
            item = self.driver_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_count += 1
        
        if selected_count == 0:
            QMessageBox.warning(self, "No Selection", "Please select at least one driver to update.")
            return
        
        confirm = QMessageBox.question(
            self,
            "Confirm Driver Update",
            f"Are you sure you want to update {selected_count} drivers?\n"
            "It's recommended to create a system restore point before proceeding.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if confirm == QMessageBox.No:
            return
        
        self.loading_overlay.show()
        self.loading_overlay.set_message("Updating drivers...")
        self.progress_bar.setValue(0)
        
        # Start task
        self.task = OptimizationTask("update_drivers")
        self.task.progress_updated.connect(self.progress_bar.setValue)
        self.task.status_updated.connect(self.status_label.setText)
        self.task.completed.connect(self.task_completed)
        self.task.start()
    
    def task_completed(self, success, message):
        """Handle completion of optimization task."""
        self.loading_overlay.hide()
        self.progress_bar.setValue(100 if success else 0)
        self.status_label.setText(message)
        
        if success:
            QMessageBox.information(self, "Task Complete", message)
            
            # If driver update was successful, rescan
            if self.tab_widget.currentIndex() == 1:  # Driver tab
                self.scan_drivers()
        else:
            QMessageBox.critical(self, "Error", message)
