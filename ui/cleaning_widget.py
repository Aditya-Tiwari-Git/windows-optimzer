"""
Cleaning widget for the Windows System Optimizer.
This module provides the interface for system cleaning operations
including browser cache, temp files, and registry cleanup.
"""

import os
import concurrent.futures
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QFrame, QCheckBox, QPushButton, QProgressBar,
                            QSpacerItem, QSizePolicy, QGroupBox, QScrollArea,
                            QTabWidget, QListWidget, QListWidgetItem, QMessageBox)
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QIcon

from ..services.cleaner import SystemCleaner
from .loading_screen import LoadingOverlay
from .styles import COLORS

class CleaningTask(QThread):
    """Background thread for executing cleaning operations."""
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    completed = pyqtSignal(bool, str)  # Success, message
    
    def __init__(self, tasks, parent=None):
        super().__init__(parent)
        self.tasks = tasks
        self.cleaner = SystemCleaner()
    
    def run(self):
        """Execute the cleaning tasks."""
        try:
            total_tasks = len(self.tasks)
            completed_tasks = 0
            
            for task in self.tasks:
                self.status_updated.emit(f"Processing: {task}")
                
                if task == "browser_chrome":
                    self.cleaner.clean_chrome_cache()
                elif task == "browser_edge":
                    self.cleaner.clean_edge_cache()
                elif task == "temp_files":
                    self.cleaner.clean_temp_files()
                elif task == "windows_temp":
                    self.cleaner.clean_windows_temp()
                elif task == "recycle_bin":
                    self.cleaner.empty_recycle_bin()
                elif task == "registry_scan":
                    self.cleaner.scan_registry()
                
                completed_tasks += 1
                progress = (completed_tasks / total_tasks) * 100
                self.progress_updated.emit(int(progress))
            
            self.completed.emit(True, "Cleaning completed successfully")
        
        except Exception as e:
            self.completed.emit(False, f"Error during cleaning: {str(e)}")


class CleaningCheckBox(QCheckBox):
    """Checkbox with description for cleaning tasks."""
    def __init__(self, text, description, parent=None):
        super().__init__(text, parent)
        self.setChecked(True)
        self.description = description


class CleaningWidget(QWidget):
    """Widget for system cleaning operations."""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.cleaner = SystemCleaner()
        self.thread_pool = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        
        # Set up main layout
        main_layout = QVBoxLayout(self)
        
        # Header
        header_label = QLabel("System Cleaning")
        header_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        main_layout.addWidget(header_label)
        
        description_label = QLabel(
            "Clean unnecessary files to free up disk space and improve system performance."
        )
        description_label.setWordWrap(True)
        main_layout.addWidget(description_label)
        
        # Tab widget for different cleaning types
        self.tab_widget = QTabWidget()
        
        # Browser cleaning tab
        self.browser_tab = self.create_browser_tab()
        self.tab_widget.addTab(self.browser_tab, "Browser Cleaning")
        
        # System cleaning tab
        self.system_tab = self.create_system_tab()
        self.tab_widget.addTab(self.system_tab, "System Cleaning")
        
        # Registry tab
        self.registry_tab = self.create_registry_tab()
        self.tab_widget.addTab(self.registry_tab, "Registry Cleaning")
        
        main_layout.addWidget(self.tab_widget)
        
        # Progress area
        progress_group = QGroupBox("Cleaning Progress")
        progress_layout = QVBoxLayout(progress_group)
        
        self.status_label = QLabel("Ready")
        progress_layout.addWidget(self.status_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        main_layout.addWidget(progress_group)
        
        # Action buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.scan_button = QPushButton("Scan Selected Items")
        self.scan_button.clicked.connect(self.start_scan)
        self.scan_button.setMinimumWidth(150)
        button_layout.addWidget(self.scan_button)
        
        self.clean_button = QPushButton("Clean Selected Items")
        self.clean_button.clicked.connect(self.start_cleaning)
        self.clean_button.setMinimumWidth(150)
        button_layout.addWidget(self.clean_button)
        
        main_layout.addLayout(button_layout)
        
        # Loading overlay
        self.loading_overlay = LoadingOverlay(self)
        self.loading_overlay.hide()
    
    def create_browser_tab(self):
        """Create the browser cleaning tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Browser cleaning options
        description = QLabel(
            "Select the browser caches you want to clean:"
        )
        layout.addWidget(description)
        
        # Chrome option
        self.chrome_checkbox = CleaningCheckBox(
            "Google Chrome",
            "Cleans temporary files, cache, cookies and history"
        )
        layout.addWidget(self.chrome_checkbox)
        
        chrome_details = QLabel(self.chrome_checkbox.description)
        chrome_details.setIndent(20)
        chrome_details.setStyleSheet("color: gray;")
        layout.addWidget(chrome_details)
        
        # Edge option
        self.edge_checkbox = CleaningCheckBox(
            "Microsoft Edge",
            "Cleans temporary files, cache, cookies and history"
        )
        layout.addWidget(self.edge_checkbox)
        
        edge_details = QLabel(self.edge_checkbox.description)
        edge_details.setIndent(20)
        edge_details.setStyleSheet("color: gray;")
        layout.addWidget(edge_details)
        
        # Stats
        stats_frame = QFrame()
        stats_layout = QVBoxLayout(stats_frame)
        
        self.browser_stats_label = QLabel("Estimated space to be freed: Calculating...")
        stats_layout.addWidget(self.browser_stats_label)
        
        layout.addWidget(stats_frame)
        layout.addStretch()
        
        return tab
    
    def create_system_tab(self):
        """Create the system cleaning tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        description = QLabel(
            "Select the system areas you want to clean:"
        )
        layout.addWidget(description)
        
        # Temp files option
        self.temp_files_checkbox = CleaningCheckBox(
            "Temporary Files",
            "Cleans user temporary files that are no longer needed"
        )
        layout.addWidget(self.temp_files_checkbox)
        
        temp_details = QLabel(self.temp_files_checkbox.description)
        temp_details.setIndent(20)
        temp_details.setStyleSheet("color: gray;")
        layout.addWidget(temp_details)
        
        # Windows temp option
        self.windows_temp_checkbox = CleaningCheckBox(
            "Windows Temporary Files",
            "Cleans Windows temporary files in the Windows directory"
        )
        layout.addWidget(self.windows_temp_checkbox)
        
        win_temp_details = QLabel(self.windows_temp_checkbox.description)
        win_temp_details.setIndent(20)
        win_temp_details.setStyleSheet("color: gray;")
        layout.addWidget(win_temp_details)
        
        # Recycle bin option
        self.recycle_bin_checkbox = CleaningCheckBox(
            "Empty Recycle Bin",
            "Permanently removes deleted files from the Recycle Bin"
        )
        layout.addWidget(self.recycle_bin_checkbox)
        
        recycle_details = QLabel(self.recycle_bin_checkbox.description)
        recycle_details.setIndent(20)
        recycle_details.setStyleSheet("color: gray;")
        layout.addWidget(recycle_details)
        
        # Stats
        stats_frame = QFrame()
        stats_layout = QVBoxLayout(stats_frame)
        
        self.system_stats_label = QLabel("Estimated space to be freed: Calculating...")
        stats_layout.addWidget(self.system_stats_label)
        
        layout.addWidget(stats_frame)
        layout.addStretch()
        
        return tab
    
    def create_registry_tab(self):
        """Create the registry cleaning tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        description = QLabel(
            "Registry scanning and cleaning:"
        )
        layout.addWidget(description)
        
        # Registry scan option
        self.registry_checkbox = CleaningCheckBox(
            "Scan Registry for Issues",
            "Scans for invalid or obsolete registry entries"
        )
        layout.addWidget(self.registry_checkbox)
        
        registry_details = QLabel(self.registry_checkbox.description)
        registry_details.setIndent(20)
        registry_details.setStyleSheet("color: gray;")
        layout.addWidget(registry_details)
        
        # Registry issues list
        self.registry_list = QListWidget()
        self.registry_list.setMinimumHeight(200)
        layout.addWidget(self.registry_list)
        
        # Warning label
        warning_label = QLabel(
            "Note: Registry cleaning should be used with caution. "
            "Always back up your registry before making changes."
        )
        warning_label.setStyleSheet("color: #FF9500;")  # Warning color
        warning_label.setWordWrap(True)
        layout.addWidget(warning_label)
        
        # Stats
        stats_frame = QFrame()
        stats_layout = QVBoxLayout(stats_frame)
        
        self.registry_stats_label = QLabel("Registry issues found: 0")
        stats_layout.addWidget(self.registry_stats_label)
        
        layout.addWidget(stats_frame)
        
        return tab
    
    def start_scan(self):
        """Start scanning for issues without cleaning."""
        self.loading_overlay.show()
        self.loading_overlay.set_message("Scanning for issues...")
        self.progress_bar.setValue(0)
        
        # Get selected tasks
        tasks = self.get_selected_tasks()
        
        # Start background task for scanning
        self.thread_pool.submit(self._background_scan, tasks)
    
    def _background_scan(self, tasks):
        """Background task for scanning."""
        try:
            # Calculate sizes and counts
            browser_size = 0
            system_size = 0
            registry_issues = []
            
            for task in tasks:
                if task == "browser_chrome":
                    chrome_size = self.cleaner.get_chrome_cache_size()
                    browser_size += chrome_size
                
                elif task == "browser_edge":
                    edge_size = self.cleaner.get_edge_cache_size()
                    browser_size += edge_size
                
                elif task == "temp_files":
                    temp_size = self.cleaner.get_temp_files_size()
                    system_size += temp_size
                
                elif task == "windows_temp":
                    win_temp_size = self.cleaner.get_windows_temp_size()
                    system_size += win_temp_size
                
                elif task == "recycle_bin":
                    recycle_size = self.cleaner.get_recycle_bin_size()
                    system_size += recycle_size
                
                elif task == "registry_scan":
                    registry_issues = self.cleaner.scan_registry(fix=False)
            
            # Update UI with results
            browser_size_mb = browser_size / (1024 * 1024)  # Convert to MB
            system_size_mb = system_size / (1024 * 1024)    # Convert to MB
            
            self.browser_stats_label.setText(f"Estimated space to be freed: {browser_size_mb:.2f} MB")
            self.system_stats_label.setText(f"Estimated space to be freed: {system_size_mb:.2f} MB")
            
            # Update registry list
            self.registry_list.clear()
            for issue in registry_issues:
                self.registry_list.addItem(issue)
            
            self.registry_stats_label.setText(f"Registry issues found: {len(registry_issues)}")
            
            # Complete
            self.progress_bar.setValue(100)
            self.status_label.setText("Scan completed")
            self.loading_overlay.hide()
            
            # Show summary message
            QMessageBox.information(
                self, 
                "Scan Complete",
                f"Scan completed successfully!\n\n"
                f"Browser cache: {browser_size_mb:.2f} MB\n"
                f"System temporary files: {system_size_mb:.2f} MB\n"
                f"Registry issues found: {len(registry_issues)}"
            )
        
        except Exception as e:
            self.loading_overlay.hide()
            self.status_label.setText(f"Error: {str(e)}")
            QMessageBox.critical(
                self, 
                "Error",
                f"An error occurred during scanning: {str(e)}"
            )
    
    def start_cleaning(self):
        """Start the cleaning process for selected items."""
        self.loading_overlay.show()
        self.loading_overlay.set_message("Cleaning in progress...")
        self.progress_bar.setValue(0)
        
        # Get selected tasks
        tasks = self.get_selected_tasks()
        
        if not tasks:
            self.loading_overlay.hide()
            QMessageBox.warning(
                self, 
                "No Selection",
                "Please select at least one item to clean."
            )
            return
        
        # Confirmation dialog
        confirm = QMessageBox.question(
            self,
            "Confirm Cleaning",
            "Are you sure you want to clean the selected items?\n"
            "This action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if confirm == QMessageBox.No:
            self.loading_overlay.hide()
            return
        
        # Create and start cleaning thread
        self.clean_thread = CleaningTask(tasks)
        self.clean_thread.progress_updated.connect(self.progress_bar.setValue)
        self.clean_thread.status_updated.connect(self.status_label.setText)
        self.clean_thread.completed.connect(self.cleaning_completed)
        self.clean_thread.start()
    
    def cleaning_completed(self, success, message):
        """Handle completion of cleaning process."""
        self.loading_overlay.hide()
        self.progress_bar.setValue(100 if success else 0)
        self.status_label.setText(message)
        
        if success:
            QMessageBox.information(self, "Cleaning Complete", message)
            
            # Reset stats after cleaning
            self.browser_stats_label.setText("Estimated space to be freed: 0 MB")
            self.system_stats_label.setText("Estimated space to be freed: 0 MB")
            self.registry_list.clear()
            self.registry_stats_label.setText("Registry issues found: 0")
        else:
            QMessageBox.critical(self, "Error", message)
    
    def get_selected_tasks(self):
        """Get list of selected cleaning tasks."""
        tasks = []
        
        if self.chrome_checkbox.isChecked():
            tasks.append("browser_chrome")
        
        if self.edge_checkbox.isChecked():
            tasks.append("browser_edge")
        
        if self.temp_files_checkbox.isChecked():
            tasks.append("temp_files")
        
        if self.windows_temp_checkbox.isChecked():
            tasks.append("windows_temp")
        
        if self.recycle_bin_checkbox.isChecked():
            tasks.append("recycle_bin")
        
        if self.registry_checkbox.isChecked():
            tasks.append("registry_scan")
        
        return tasks
