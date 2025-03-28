"""
Driver Updates widget for managing device drivers.
"""

import os
import time
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QFrame, QProgressBar, QTableWidget,
                             QTableWidgetItem, QHeaderView, QSpacerItem,
                             QSizePolicy, QCheckBox, QComboBox, QGroupBox,
                             QMessageBox)
from PyQt5.QtGui import QIcon, QFont, QPixmap, QColor
from PyQt5.QtCore import Qt, pyqtSignal, pyqtSlot, QThread, QSize

from ui.theme import COLORS
from ui.components import Card, HeaderLabel, SubHeaderLabel, ActionButton, InfoCard
from ui.loading_screen import LoadingScreen
from utils.driver_manager import DriverManager
from utils.background_worker import BackgroundWorker
from assets.icons import get_icon


class DriversWidget(QWidget):
    """Widget for driver updates and management."""
    
    def __init__(self):
        super().__init__()
        self.driver_manager = DriverManager()
        self.initUI()
        
        # Loading screen instance
        self.loading_screen = None
        
        # Driver data
        self.driver_data = []
    
    def initUI(self):
        """Initialize the driver updates UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)
        
        # Header
        header = HeaderLabel("Driver Updates")
        subheader = SubHeaderLabel("Scan, update and manage device drivers")
        
        layout.addWidget(header)
        layout.addWidget(subheader)
        
        # Status cards
        status_layout = QHBoxLayout()
        
        self.total_drivers_card = InfoCard(
            title="Total Drivers",
            icon_name="drivers",
            value="--",
            description="Installed drivers"
        )
        
        self.outdated_card = InfoCard(
            title="Outdated",
            icon_name="update",
            value="--",
            description="Drivers needing updates"
        )
        
        self.problematic_card = InfoCard(
            title="Problematic",
            icon_name="warning",
            value="--",
            description="Drivers with issues"
        )
        
        status_layout.addWidget(self.total_drivers_card)
        status_layout.addWidget(self.outdated_card)
        status_layout.addWidget(self.problematic_card)
        
        layout.addLayout(status_layout)
        
        # Driver table card
        drivers_card = Card()
        drivers_layout = QVBoxLayout(drivers_card)
        
        # Options group
        options_group = QGroupBox("Scan Options")
        options_layout = QHBoxLayout(options_group)
        
        # Category filter
        self.category_combo = QComboBox()
        self.category_combo.addItem("All Devices")
        self.category_combo.addItems([
            "Display Adapters", 
            "Network Adapters", 
            "Audio Devices",
            "Input Devices",
            "Storage Controllers",
            "USB Controllers"
        ])
        self.category_combo.currentIndexChanged.connect(self.filter_drivers)
        
        # Show options
        self.show_outdated = QCheckBox("Show outdated only")
        self.show_outdated.toggled.connect(self.filter_drivers)
        
        self.show_problematic = QCheckBox("Show problematic only")
        self.show_problematic.toggled.connect(self.filter_drivers)
        
        # Add to options layout
        options_layout.addWidget(QLabel("Category:"))
        options_layout.addWidget(self.category_combo)
        options_layout.addSpacing(20)
        options_layout.addWidget(self.show_outdated)
        options_layout.addSpacing(10)
        options_layout.addWidget(self.show_problematic)
        options_layout.addStretch()
        
        drivers_layout.addWidget(options_group)
        
        # Driver table
        self.drivers_table = QTableWidget()
        self.drivers_table.setColumnCount(5)
        self.drivers_table.setHorizontalHeaderLabels(["Device", "Current Version", "Status", "Available Update", ""])
        self.drivers_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.drivers_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.drivers_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Fixed)
        self.drivers_table.setColumnWidth(4, 100)  # Width for action column
        self.drivers_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.drivers_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.drivers_table.setAlternatingRowColors(True)
        
        drivers_layout.addWidget(self.drivers_table, 1)
        
        # Action buttons
        buttons_layout = QHBoxLayout()
        
        scan_btn = ActionButton("Scan for Updates", "search", "primary")
        scan_btn.clicked.connect(self.scan_drivers)
        
        update_all_btn = ActionButton("Update Selected", "update", "success")
        update_all_btn.clicked.connect(self.update_selected_drivers)
        
        backup_btn = ActionButton("Backup Drivers", "backup", "warning")
        backup_btn.clicked.connect(self.backup_drivers)
        
        buttons_layout.addStretch()
        buttons_layout.addWidget(scan_btn)
        buttons_layout.addWidget(update_all_btn)
        buttons_layout.addWidget(backup_btn)
        
        drivers_layout.addLayout(buttons_layout)
        
        layout.addWidget(drivers_card, 1)
        
        # Summary card
        summary_card = Card()
        summary_layout = QVBoxLayout(summary_card)
        
        summary_header = QLabel("Driver Health Report")
        summary_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        summary_layout.addWidget(summary_header)
        
        self.summary_text = QLabel(
            "Click 'Scan for Updates' to check the status of your device drivers."
        )
        self.summary_text.setWordWrap(True)
        self.summary_text.setStyleSheet(f"color: {COLORS['text']}; padding: 10px;")
        summary_layout.addWidget(self.summary_text)
        
        layout.addWidget(summary_card)
    
    def scan_drivers(self):
        """Scan system for driver status."""
        self.show_loading("Scanning for Driver Updates...", "This may take a while.")
        
        # Create background worker
        worker = BackgroundWorker(self.driver_manager.scan_drivers)
        worker.taskFinished.connect(self.on_scan_complete)
        worker.start()
    
    def on_scan_complete(self, driver_data):
        """Handle completion of driver scan."""
        self.hide_loading()
        
        # Store driver data
        self.driver_data = driver_data
        
        # Update status cards
        total_count = len(driver_data)
        outdated_count = sum(1 for driver in driver_data if driver['status'] == 'outdated')
        problematic_count = sum(1 for driver in driver_data if driver['status'] == 'problematic')
        
        self.total_drivers_card.update_value(str(total_count))
        self.outdated_card.update_value(str(outdated_count))
        self.problematic_card.update_value(str(problematic_count))
        
        # Apply any active filters
        self.filter_drivers()
        
        # Update summary
        self.update_summary()
    
    def filter_drivers(self):
        """Filter the drivers table based on selected options."""
        if not self.driver_data:
            return
        
        # Get filter criteria
        category = self.category_combo.currentText()
        show_outdated_only = self.show_outdated.isChecked()
        show_problematic_only = self.show_problematic.isChecked()
        
        # Apply filters
        filtered_data = self.driver_data
        
        if category != "All Devices":
            filtered_data = [driver for driver in filtered_data if driver['category'] == category]
        
        if show_outdated_only:
            filtered_data = [driver for driver in filtered_data if driver['status'] == 'outdated']
        
        if show_problematic_only:
            filtered_data = [driver for driver in filtered_data if driver['status'] == 'problematic']
        
        # Update table
        self.update_drivers_table(filtered_data)
    
    def update_drivers_table(self, driver_data):
        """Update the drivers table with the provided data."""
        self.drivers_table.setRowCount(0)  # Clear table
        
        for i, driver in enumerate(driver_data):
            self.drivers_table.insertRow(i)
            
            # Device name
            device_item = QTableWidgetItem(driver['name'])
            self.drivers_table.setItem(i, 0, device_item)
            
            # Current version
            version_item = QTableWidgetItem(driver['current_version'])
            self.drivers_table.setItem(i, 1, version_item)
            
            # Status
            status_text = driver['status'].capitalize()
            status_item = QTableWidgetItem(status_text)
            
            # Color-code status
            if driver['status'] == 'up-to-date':
                status_item.setForeground(QColor(COLORS['secondary']))
            elif driver['status'] == 'outdated':
                status_item.setForeground(QColor(COLORS['warning']))
            elif driver['status'] == 'problematic':
                status_item.setForeground(QColor(COLORS['error']))
            
            self.drivers_table.setItem(i, 2, status_item)
            
            # Available update
            update_version = driver.get('available_version', 'N/A')
            update_item = QTableWidgetItem(update_version)
            self.drivers_table.setItem(i, 3, update_item)
            
            # Action button (update or fix)
            if driver['status'] == 'outdated':
                update_btn = QPushButton("Update")
                update_btn.setStyleSheet(
                    f"background-color: {COLORS['primary']}; color: white; "
                    f"padding: 4px; border-radius: 4px;"
                )
                update_btn.clicked.connect(lambda checked, idx=i: self.update_driver(idx))
                self.drivers_table.setCellWidget(i, 4, update_btn)
            elif driver['status'] == 'problematic':
                fix_btn = QPushButton("Fix")
                fix_btn.setStyleSheet(
                    f"background-color: {COLORS['error']}; color: white; "
                    f"padding: 4px; border-radius: 4px;"
                )
                fix_btn.clicked.connect(lambda checked, idx=i: self.fix_driver(idx))
                self.drivers_table.setCellWidget(i, 4, fix_btn)
    
    def update_summary(self):
        """Update the driver health summary text."""
        if not self.driver_data:
            self.summary_text.setText("Click 'Scan for Updates' to check the status of your device drivers.")
            return
        
        total_count = len(self.driver_data)
        outdated_count = sum(1 for driver in self.driver_data if driver['status'] == 'outdated')
        problematic_count = sum(1 for driver in self.driver_data if driver['status'] == 'problematic')
        up_to_date_count = total_count - outdated_count - problematic_count
        
        message = (
            f"<span style='font-weight: bold;'>System Driver Health Report</span><br><br>"
            f"Total drivers scanned: {total_count}<br>"
            f"Up-to-date drivers: {up_to_date_count} ({up_to_date_count/total_count*100:.1f}%)<br>"
            f"Outdated drivers: {outdated_count}<br>"
            f"Problematic drivers: {problematic_count}<br><br>"
        )
        
        if outdated_count == 0 and problematic_count == 0:
            message += (
                f"<span style='color: {COLORS['secondary']}; font-weight: bold;'>"
                f"Great! All your drivers are up-to-date and functioning properly."
                f"</span>"
            )
        elif outdated_count > 0 and problematic_count == 0:
            message += (
                f"<span style='color: {COLORS['warning']}; font-weight: bold;'>"
                f"Action recommended: Update the outdated drivers to ensure optimal performance."
                f"</span>"
            )
        else:
            message += (
                f"<span style='color: {COLORS['error']}; font-weight: bold;'>"
                f"Attention required: There are problematic drivers that need to be fixed "
                f"to avoid system instability or hardware issues."
                f"</span>"
            )
        
        self.summary_text.setText(message)
    
    def update_driver(self, row_index):
        """Update the driver at the specified row."""
        # Get the driver name from the table
        driver_name = self.drivers_table.item(row_index, 0).text()
        
        # Find the corresponding driver in the data
        for driver in self.driver_data:
            if driver['name'] == driver_name:
                self.show_loading(f"Updating Driver: {driver_name}...", "This may take a while.")
                
                # Create background worker
                worker = BackgroundWorker(self.driver_manager.update_driver, driver)
                worker.taskFinished.connect(lambda result, name=driver_name: self.on_driver_update_complete(result, name))
                worker.start()
                break
    
    def on_driver_update_complete(self, result, driver_name):
        """Handle completion of driver update."""
        self.hide_loading()
        
        if result.get('success', False):
            QMessageBox.information(
                self,
                "Driver Update Complete",
                f"The driver for {driver_name} has been successfully updated."
            )
            
            # Refresh the driver list
            self.scan_drivers()
        else:
            QMessageBox.warning(
                self,
                "Driver Update Failed",
                f"Failed to update driver for {driver_name}.\n\n"
                f"Error: {result.get('error', 'Unknown error')}"
            )
    
    def fix_driver(self, row_index):
        """Attempt to fix the problematic driver at the specified row."""
        # Get the driver name from the table
        driver_name = self.drivers_table.item(row_index, 0).text()
        
        # Find the corresponding driver in the data
        for driver in self.driver_data:
            if driver['name'] == driver_name:
                self.show_loading(f"Fixing Driver: {driver_name}...", "This may take a while.")
                
                # Create background worker
                worker = BackgroundWorker(self.driver_manager.fix_driver, driver)
                worker.taskFinished.connect(lambda result, name=driver_name: self.on_driver_fix_complete(result, name))
                worker.start()
                break
    
    def on_driver_fix_complete(self, result, driver_name):
        """Handle completion of driver fix."""
        self.hide_loading()
        
        if result.get('success', False):
            QMessageBox.information(
                self,
                "Driver Fix Complete",
                f"The problematic driver for {driver_name} has been successfully fixed."
            )
            
            # Refresh the driver list
            self.scan_drivers()
        else:
            QMessageBox.warning(
                self,
                "Driver Fix Failed",
                f"Failed to fix driver for {driver_name}.\n\n"
                f"Error: {result.get('error', 'Unknown error')}\n\n"
                f"Recommendation: {result.get('recommendation', '')}"
            )
    
    def update_selected_drivers(self):
        """Update all selected drivers."""
        selected_rows = list(set(item.row() for item in self.drivers_table.selectedItems()))
        
        if not selected_rows:
            QMessageBox.information(
                self,
                "No Drivers Selected",
                "Please select one or more drivers to update."
            )
            return
        
        # Get driver objects for selected rows
        selected_drivers = []
        for row in selected_rows:
            driver_name = self.drivers_table.item(row, 0).text()
            
            for driver in self.driver_data:
                if driver['name'] == driver_name and driver['status'] == 'outdated':
                    selected_drivers.append(driver)
                    break
        
        if not selected_drivers:
            QMessageBox.information(
                self,
                "No Updates Required",
                "The selected drivers are already up to date."
            )
            return
        
        # Confirm update
        count = len(selected_drivers)
        reply = QMessageBox.question(
            self,
            "Confirm Driver Updates",
            f"Do you want to update {count} driver{'s' if count > 1 else ''}?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Yes:
            self.show_loading("Updating Selected Drivers...", "This may take a while.")
            
            # Create background worker
            worker = BackgroundWorker(self.driver_manager.update_multiple_drivers, selected_drivers)
            worker.taskFinished.connect(self.on_multiple_updates_complete)
            worker.start()
    
    def on_multiple_updates_complete(self, results):
        """Handle completion of multiple driver updates."""
        self.hide_loading()
        
        success_count = sum(1 for result in results if result.get('success', False))
        total_count = len(results)
        
        if success_count == total_count:
            QMessageBox.information(
                self,
                "Updates Complete",
                f"All {total_count} driver updates completed successfully."
            )
        else:
            QMessageBox.warning(
                self,
                "Updates Incomplete",
                f"{success_count} of {total_count} driver updates completed successfully.\n\n"
                f"Some drivers could not be updated. Please try updating them individually."
            )
        
        # Refresh the driver list
        self.scan_drivers()
    
    def backup_drivers(self):
        """Backup all system drivers."""
        reply = QMessageBox.question(
            self,
            "Confirm Driver Backup",
            "Do you want to create a backup of all system drivers? "
            "This can be used to restore drivers if an update causes issues.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Yes:
            self.show_loading("Backing Up System Drivers...", "This may take a while.")
            
            # Create background worker
            worker = BackgroundWorker(self.driver_manager.backup_drivers)
            worker.taskFinished.connect(self.on_backup_complete)
            worker.start()
    
    def on_backup_complete(self, result):
        """Handle completion of driver backup."""
        self.hide_loading()
        
        if result.get('success', False):
            QMessageBox.information(
                self,
                "Backup Complete",
                f"Driver backup completed successfully.\n\n"
                f"Backup location: {result.get('backup_path', 'Unknown')}"
            )
        else:
            QMessageBox.warning(
                self,
                "Backup Failed",
                f"Failed to backup drivers.\n\n"
                f"Error: {result.get('error', 'Unknown error')}"
            )
    
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
