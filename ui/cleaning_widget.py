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

from services.cleaner import SystemCleaner
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
                elif task == "browser_firefox":
                    self.cleaner.clean_firefox_cache()
                elif task == "browser_opera":
                    self.cleaner.clean_opera_cache()
                elif task == "browser_brave":
                    self.cleaner.clean_brave_cache()
                elif task == "temp_files":
                    self.cleaner.clean_temp_files()
                elif task == "windows_temp":
                    self.cleaner.clean_windows_temp()
                elif task == "recycle_bin":
                    self.cleaner.empty_recycle_bin()
                
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
    
    # Add a signal to update theme
    theme_changed = pyqtSignal(bool)  # is_dark_mode
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.cleaner = SystemCleaner()
        self.thread_pool = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        
        # Store all primary-colored buttons for theme updates
        self.primary_buttons = []
        
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
        
        # Loading overlay
        self.loading_overlay = LoadingOverlay(self)
        self.loading_overlay.hide()
        
        # Connect theme change signal (will be connected externally from main window)
        self.theme_changed.connect(self.update_button_styling)
    
    def update_button_styling(self, is_dark_mode):
        """Update button styling when theme changes"""
        primary_color = COLORS['dark_primary'] if is_dark_mode else COLORS['primary']
        secondary_color = COLORS['dark_secondary'] if is_dark_mode else COLORS['secondary']
        
        # Update all primary buttons
        for button in self.primary_buttons:
            button.setStyleSheet(f"background-color: {primary_color}; color: white;")
        
        # Update registry fix button separately
        if hasattr(self, 'fix_button'):
            self.fix_button.setStyleSheet(f"background-color: {secondary_color}; color: white;")
    
    def create_browser_tab(self):
        """Create the browser cleaning tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Browser cleaning options
        description = QLabel(
            "Select the browser caches and cookies you want to clean:"
        )
        layout.addWidget(description)
        
        # Chrome option
        chrome_layout = QHBoxLayout()
        self.chrome_checkbox = CleaningCheckBox(
            "Google Chrome",
            "Cleans cache and cookies (browsing history is preserved)"
        )
        chrome_layout.addWidget(self.chrome_checkbox)
        
        # Add clean button for Chrome
        chrome_clean_btn = QPushButton("Clean")
        chrome_clean_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white;")
        chrome_clean_btn.setFixedWidth(80)
        chrome_clean_btn.clicked.connect(self.clean_chrome)
        chrome_layout.addWidget(chrome_clean_btn)
        self.primary_buttons.append(chrome_clean_btn)
        
        layout.addLayout(chrome_layout)
        
        chrome_details = QLabel(self.chrome_checkbox.description)
        chrome_details.setIndent(20)
        chrome_details.setStyleSheet("color: gray;")
        layout.addWidget(chrome_details)
        
        # Edge option
        edge_layout = QHBoxLayout()
        self.edge_checkbox = CleaningCheckBox(
            "Microsoft Edge",
            "Cleans cache and cookies (browsing history is preserved)"
        )
        edge_layout.addWidget(self.edge_checkbox)
        
        # Add clean button for Edge
        edge_clean_btn = QPushButton("Clean")
        edge_clean_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white;")
        edge_clean_btn.setFixedWidth(80)
        edge_clean_btn.clicked.connect(self.clean_edge)
        edge_layout.addWidget(edge_clean_btn)
        self.primary_buttons.append(edge_clean_btn)
        
        layout.addLayout(edge_layout)
        
        edge_details = QLabel(self.edge_checkbox.description)
        edge_details.setIndent(20)
        edge_details.setStyleSheet("color: gray;")
        layout.addWidget(edge_details)
        
        # Firefox option
        firefox_layout = QHBoxLayout()
        self.firefox_checkbox = CleaningCheckBox(
            "Mozilla Firefox",
            "Cleans cache and cookies (browsing history is preserved)"
        )
        firefox_layout.addWidget(self.firefox_checkbox)
        
        # Add clean button for Firefox
        firefox_clean_btn = QPushButton("Clean")
        firefox_clean_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white;")
        firefox_clean_btn.setFixedWidth(80)
        firefox_clean_btn.clicked.connect(self.clean_firefox)
        firefox_layout.addWidget(firefox_clean_btn)
        self.primary_buttons.append(firefox_clean_btn)
        
        layout.addLayout(firefox_layout)
        
        firefox_details = QLabel(self.firefox_checkbox.description)
        firefox_details.setIndent(20)
        firefox_details.setStyleSheet("color: gray;")
        layout.addWidget(firefox_details)
        
        # Opera option
        opera_layout = QHBoxLayout()
        self.opera_checkbox = CleaningCheckBox(
            "Opera",
            "Cleans cache and cookies (browsing history is preserved)"
        )
        opera_layout.addWidget(self.opera_checkbox)
        
        # Add clean button for Opera
        opera_clean_btn = QPushButton("Clean")
        opera_clean_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white;")
        opera_clean_btn.setFixedWidth(80)
        opera_clean_btn.clicked.connect(self.clean_opera)
        opera_layout.addWidget(opera_clean_btn)
        self.primary_buttons.append(opera_clean_btn)
        
        layout.addLayout(opera_layout)
        
        opera_details = QLabel(self.opera_checkbox.description)
        opera_details.setIndent(20)
        opera_details.setStyleSheet("color: gray;")
        layout.addWidget(opera_details)
        
        # Brave option
        brave_layout = QHBoxLayout()
        self.brave_checkbox = CleaningCheckBox(
            "Brave",
            "Cleans cache and cookies (browsing history is preserved)"
        )
        brave_layout.addWidget(self.brave_checkbox)
        
        # Add clean button for Brave
        brave_clean_btn = QPushButton("Clean")
        brave_clean_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white;")
        brave_clean_btn.setFixedWidth(80)
        brave_clean_btn.clicked.connect(self.clean_brave)
        brave_layout.addWidget(brave_clean_btn)
        self.primary_buttons.append(brave_clean_btn)
        
        layout.addLayout(brave_layout)
        
        brave_details = QLabel(self.brave_checkbox.description)
        brave_details.setIndent(20)
        brave_details.setStyleSheet("color: gray;")
        layout.addWidget(brave_details)
        
        # Remove estimated space label
        layout.addStretch()
        
        # Clean all browsers button at bottom
        clean_all_browsers_btn = QPushButton("Clean All Selected Browsers")
        clean_all_browsers_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white;")
        clean_all_browsers_btn.clicked.connect(self.clean_all_browsers)
        layout.addWidget(clean_all_browsers_btn)
        self.primary_buttons.append(clean_all_browsers_btn)
        
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
        temp_layout = QHBoxLayout()
        self.temp_files_checkbox = CleaningCheckBox(
            "Temporary Files",
            "Cleans user temporary files that are no longer needed"
        )
        temp_layout.addWidget(self.temp_files_checkbox)
        
        # Add clean button for Temp Files
        temp_clean_btn = QPushButton("Clean")
        temp_clean_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white;")
        temp_clean_btn.setFixedWidth(80)
        temp_clean_btn.clicked.connect(self.clean_temp_files)
        temp_layout.addWidget(temp_clean_btn)
        self.primary_buttons.append(temp_clean_btn)
        
        layout.addLayout(temp_layout)
        
        temp_details = QLabel(self.temp_files_checkbox.description)
        temp_details.setIndent(20)
        temp_details.setStyleSheet("color: gray;")
        layout.addWidget(temp_details)
        
        # Windows temp option
        win_temp_layout = QHBoxLayout()
        self.windows_temp_checkbox = CleaningCheckBox(
            "Windows Temporary Files",
            "Cleans Windows temporary files in the Windows directory"
        )
        win_temp_layout.addWidget(self.windows_temp_checkbox)
        
        # Add clean button for Windows Temp
        win_temp_clean_btn = QPushButton("Clean")
        win_temp_clean_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white;")
        win_temp_clean_btn.setFixedWidth(80)
        win_temp_clean_btn.clicked.connect(self.clean_windows_temp)
        win_temp_layout.addWidget(win_temp_clean_btn)
        self.primary_buttons.append(win_temp_clean_btn)
        
        layout.addLayout(win_temp_layout)
        
        win_temp_details = QLabel(self.windows_temp_checkbox.description)
        win_temp_details.setIndent(20)
        win_temp_details.setStyleSheet("color: gray;")
        layout.addWidget(win_temp_details)
        
        # Recycle bin option
        recycle_layout = QHBoxLayout()
        self.recycle_bin_checkbox = CleaningCheckBox(
            "Empty Recycle Bin",
            "Permanently removes deleted files from the Recycle Bin"
        )
        recycle_layout.addWidget(self.recycle_bin_checkbox)
        
        # Add clean button for Recycle Bin
        recycle_clean_btn = QPushButton("Clean")
        recycle_clean_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white;")
        recycle_clean_btn.setFixedWidth(80)
        recycle_clean_btn.clicked.connect(self.clean_recycle_bin)
        recycle_layout.addWidget(recycle_clean_btn)
        self.primary_buttons.append(recycle_clean_btn)
        
        layout.addLayout(recycle_layout)
        
        recycle_details = QLabel(self.recycle_bin_checkbox.description)
        recycle_details.setIndent(20)
        recycle_details.setStyleSheet("color: gray;")
        layout.addWidget(recycle_details)
        
        # Remove estimated space label
        layout.addStretch()
        
        # Clean all system items button at bottom
        clean_all_system_btn = QPushButton("Clean All Selected Items")
        clean_all_system_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white;")
        clean_all_system_btn.clicked.connect(self.clean_all_system)
        layout.addWidget(clean_all_system_btn)
        self.primary_buttons.append(clean_all_system_btn)
        
        return tab
    
    # Individual browser clean functions
    def clean_chrome(self):
        self._clean_specific_items(["browser_chrome"])
    
    def clean_edge(self):
        self._clean_specific_items(["browser_edge"])
    
    def clean_firefox(self):
        self._clean_specific_items(["browser_firefox"])
    
    def clean_opera(self):
        self._clean_specific_items(["browser_opera"])
    
    def clean_brave(self):
        self._clean_specific_items(["browser_brave"])
    
    # System clean functions
    def clean_temp_files(self):
        self._clean_specific_items(["temp_files"])
    
    def clean_windows_temp(self):
        self._clean_specific_items(["windows_temp"])
    
    def clean_recycle_bin(self):
        self._clean_specific_items(["recycle_bin"])
    
    # Group clean functions
    def clean_all_browsers(self):
        tasks = []
        if self.chrome_checkbox.isChecked():
            tasks.append("browser_chrome")
        if self.edge_checkbox.isChecked():
            tasks.append("browser_edge")
        if self.firefox_checkbox.isChecked():
            tasks.append("browser_firefox")
        if self.opera_checkbox.isChecked():
            tasks.append("browser_opera")
        if self.brave_checkbox.isChecked():
            tasks.append("browser_brave")
        
        if tasks:
            self._clean_specific_items(tasks)
        else:
            QMessageBox.warning(self, "No Selection", "Please select at least one browser to clean.")
    
    def clean_all_system(self):
        tasks = []
        if self.temp_files_checkbox.isChecked():
            tasks.append("temp_files")
        if self.windows_temp_checkbox.isChecked():
            tasks.append("windows_temp")
        if self.recycle_bin_checkbox.isChecked():
            tasks.append("recycle_bin")
        
        if tasks:
            self._clean_specific_items(tasks)
        else:
            QMessageBox.warning(self, "No Selection", "Please select at least one system item to clean.")
    
    def _clean_specific_items(self, tasks):
        """Start the cleaning process for specific items."""
        self.loading_overlay.show()
        self.loading_overlay.set_message("Cleaning in progress...")
        self.progress_bar.setValue(0)
        
        if not tasks:
            self.loading_overlay.hide()
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
        else:
            QMessageBox.critical(self, "Error", message)
    
    def get_selected_tasks(self):
        """Get list of selected cleaning tasks."""
        tasks = []
        
        # Browser tasks
        if self.chrome_checkbox.isChecked():
            tasks.append("browser_chrome")
        
        if self.edge_checkbox.isChecked():
            tasks.append("browser_edge")
        
        if self.firefox_checkbox.isChecked():
            tasks.append("browser_firefox")
        
        if self.opera_checkbox.isChecked():
            tasks.append("browser_opera")
        
        if self.brave_checkbox.isChecked():
            tasks.append("browser_brave")
        
        # System tasks
        if self.temp_files_checkbox.isChecked():
            tasks.append("temp_files")
        
        if self.windows_temp_checkbox.isChecked():
            tasks.append("windows_temp")
        
        if self.recycle_bin_checkbox.isChecked():
            tasks.append("recycle_bin")
        
        return tasks
