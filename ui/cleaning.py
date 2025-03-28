"""
System Cleaning widget for managing system cleanup operations.
"""

import os
import shutil
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QFrame, QProgressBar, QTabWidget,
                             QCheckBox, QListWidget, QListWidgetItem, QFileDialog,
                             QSpacerItem, QSizePolicy, QScrollArea, QGroupBox, QComboBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt, pyqtSignal, pyqtSlot, QThread, QSize

from ui.theme import COLORS
from ui.components import Card, HeaderLabel, SubHeaderLabel, ActionButton, InfoCard
from ui.loading_screen import LoadingScreen
from utils.cleaner import SystemCleaner
from utils.registry_tools import RegistryCleaner
from utils.background_worker import BackgroundWorker
from assets.icons import get_icon


class CleaningListItem(QFrame):
    """Custom list item widget for cleaning options."""
    
    def __init__(self, title, description, icon_name, parent=None):
        super().__init__(parent)
        self.title = title
        self.description = description
        
        self.setMinimumHeight(60)
        self.setObjectName("cleaning_list_item")
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 5, 10, 5)
        
        # Checkbox
        self.checkbox = QCheckBox()
        self.checkbox.setChecked(True)
        layout.addWidget(self.checkbox)
        
        # Icon
        icon_label = QLabel()
        icon = QIcon(get_icon(icon_name))
        pixmap = icon.pixmap(QSize(24, 24))
        icon_label.setPixmap(pixmap)
        layout.addWidget(icon_label)
        
        # Text content
        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(2)
        
        title_label = QLabel(title)
        title_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        
        desc_label = QLabel(description)
        desc_label.setFont(QFont("Segoe UI", 8))
        desc_label.setStyleSheet(f"color: {COLORS['text']}; opacity: 0.8;")
        
        text_layout.addWidget(title_label)
        text_layout.addWidget(desc_label)
        
        layout.addLayout(text_layout, 1)
        
        # Status indicator
        self.size_label = QLabel("--")
        self.size_label.setFont(QFont("Segoe UI", 9))
        layout.addWidget(self.size_label)
    
    def is_checked(self):
        """Return checkbox state."""
        return self.checkbox.isChecked()
    
    def set_size(self, size_text):
        """Update the size label."""
        self.size_label.setText(size_text)


class CleaningWidget(QWidget):
    """Widget for system cleaning operations."""
    
    def __init__(self):
        super().__init__()
        self.cleaner = SystemCleaner()
        self.registry_cleaner = RegistryCleaner()
        self.initUI()
        
        # Loading screen instance
        self.loading_screen = None
    
    def initUI(self):
        """Initialize the cleaning UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)
        
        # Header
        header = HeaderLabel("System Cleaning")
        subheader = SubHeaderLabel("Clean up system files and free up disk space")
        
        layout.addWidget(header)
        layout.addWidget(subheader)
        
        # Tabs for different cleaning operations
        tabs = QTabWidget()
        tabs.setObjectName("cleaning_tabs")
        
        # File Cleaner Tab
        file_cleaner_tab = QWidget()
        file_cleaner_layout = QVBoxLayout(file_cleaner_tab)
        file_cleaner_layout.setContentsMargins(15, 15, 15, 15)
        
        # Status cards
        status_layout = QHBoxLayout()
        
        self.total_space_card = InfoCard(
            title="Total Space",
            icon_name="disk",
            value="0 GB",
            description="Total disk space"
        )
        
        self.free_space_card = InfoCard(
            title="Free Space",
            icon_name="clean",
            value="0 GB",
            description="Available disk space"
        )
        
        self.recoverable_card = InfoCard(
            title="Recoverable",
            icon_name="recycle",
            value="0 MB",
            description="Potential space savings"
        )
        
        status_layout.addWidget(self.total_space_card)
        status_layout.addWidget(self.free_space_card)
        status_layout.addWidget(self.recoverable_card)
        
        file_cleaner_layout.addLayout(status_layout)
        file_cleaner_layout.addSpacing(15)
        
        # Cleaning options list
        options_card = Card()
        options_layout = QVBoxLayout(options_card)
        options_layout.setContentsMargins(0, 0, 0, 0)
        
        options_header = QLabel("Cleaning Options")
        options_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        options_header.setContentsMargins(15, 15, 15, 5)
        options_layout.addWidget(options_header)
        
        # Scrollable list of cleaning options
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setContentsMargins(5, 5, 5, 5)
        scroll_layout.setSpacing(5)
        
        # Create cleaning option items
        self.cleaning_items = []
        
        temp_files_item = CleaningListItem(
            "Temporary Files", 
            "System temporary files and folders",
            "temp"
        )
        
        windows_temp_item = CleaningListItem(
            "Windows Temp", 
            "Windows update and system temp files",
            "windows"
        )
        
        recyclebin_item = CleaningListItem(
            "Recycle Bin", 
            "Deleted files in the recycle bin",
            "recycle"
        )
        
        chrome_cache_item = CleaningListItem(
            "Chrome Cache", 
            "Google Chrome browser cache and cookies",
            "chrome"
        )
        
        edge_cache_item = CleaningListItem(
            "Edge Cache", 
            "Microsoft Edge browser cache and cookies",
            "edge"
        )
        
        downloads_item = CleaningListItem(
            "Downloads Cleanup", 
            "Old files in Downloads folder",
            "download"
        )
        
        thumbnails_item = CleaningListItem(
            "Thumbnails Cache", 
            "Windows Explorer thumbnails cache",
            "image"
        )
        
        self.cleaning_items = [
            temp_files_item,
            windows_temp_item,
            recyclebin_item,
            chrome_cache_item,
            edge_cache_item,
            downloads_item,
            thumbnails_item
        ]
        
        for item in self.cleaning_items:
            scroll_layout.addWidget(item)
        
        scroll_layout.addStretch()
        scroll_area.setWidget(scroll_content)
        options_layout.addWidget(scroll_area)
        
        file_cleaner_layout.addWidget(options_card, 1)
        file_cleaner_layout.addSpacing(10)
        
        # Action buttons
        buttons_layout = QHBoxLayout()
        
        analyze_btn = ActionButton("Analyze", "search", "primary")
        analyze_btn.clicked.connect(self.analyze_system)
        
        clean_btn = ActionButton("Clean", "clean", "success")
        clean_btn.clicked.connect(self.clean_system)
        
        buttons_layout.addStretch()
        buttons_layout.addWidget(analyze_btn)
        buttons_layout.addWidget(clean_btn)
        
        file_cleaner_layout.addLayout(buttons_layout)
        
        # Registry Cleaner Tab
        registry_tab = QWidget()
        registry_layout = QVBoxLayout(registry_tab)
        registry_layout.setContentsMargins(15, 15, 15, 15)
        
        # Registry information
        registry_info = QLabel(
            "The Registry Cleaner scans for and repairs issues in the Windows Registry "
            "such as invalid entries, broken shortcuts, and outdated software references."
        )
        registry_info.setWordWrap(True)
        registry_info.setStyleSheet("background-color: rgba(0, 122, 255, 0.1); padding: 10px; border-radius: 5px;")
        
        registry_layout.addWidget(registry_info)
        registry_layout.addSpacing(15)
        
        # Registry scan options
        registry_card = Card()
        registry_options_layout = QVBoxLayout(registry_card)
        
        registry_header = QLabel("Registry Scan Options")
        registry_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        registry_options_layout.addWidget(registry_header)
        
        # Registry options
        reg_options_group = QGroupBox("Select scan areas:")
        reg_options_layout = QVBoxLayout(reg_options_group)
        
        self.reg_options = []
        
        reg_software = QCheckBox("Software and App Paths")
        reg_software.setChecked(True)
        
        reg_file_extensions = QCheckBox("File Extensions")
        reg_file_extensions.setChecked(True)
        
        reg_startup = QCheckBox("Startup Programs")
        reg_startup.setChecked(True)
        
        reg_uninstall = QCheckBox("Uninstall Information")
        reg_uninstall.setChecked(True)
        
        reg_shared_dlls = QCheckBox("Shared DLLs")
        reg_shared_dlls.setChecked(True)
        
        self.reg_options = [
            reg_software,
            reg_file_extensions,
            reg_startup,
            reg_uninstall,
            reg_shared_dlls
        ]
        
        for option in self.reg_options:
            reg_options_layout.addWidget(option)
        
        registry_options_layout.addWidget(reg_options_group)
        
        # Registry backup option
        backup_group = QGroupBox("Backup")
        backup_layout = QVBoxLayout(backup_group)
        
        self.create_backup_cb = QCheckBox("Create Registry backup before cleaning")
        self.create_backup_cb.setChecked(True)
        
        backup_info = QLabel("Recommended: Always create a backup before making registry changes.")
        backup_info.setWordWrap(True)
        backup_info.setStyleSheet(f"color: {COLORS['warning']};")
        
        backup_layout.addWidget(self.create_backup_cb)
        backup_layout.addWidget(backup_info)
        
        registry_options_layout.addWidget(backup_group)
        registry_options_layout.addStretch()
        
        registry_layout.addWidget(registry_card, 1)
        
        # Registry results display
        self.registry_results = QLabel("No scan performed yet.")
        self.registry_results.setAlignment(Qt.AlignCenter)
        self.registry_results.setStyleSheet("padding: 20px;")
        registry_layout.addWidget(self.registry_results)
        
        # Registry action buttons
        registry_buttons_layout = QHBoxLayout()
        
        reg_scan_btn = ActionButton("Scan Registry", "search", "primary")
        reg_scan_btn.clicked.connect(self.scan_registry)
        
        reg_fix_btn = ActionButton("Fix Issues", "clean", "success")
        reg_fix_btn.clicked.connect(self.fix_registry)
        
        registry_buttons_layout.addStretch()
        registry_buttons_layout.addWidget(reg_scan_btn)
        registry_buttons_layout.addWidget(reg_fix_btn)
        
        registry_layout.addLayout(registry_buttons_layout)
        
        # Add tabs
        tabs.addTab(file_cleaner_tab, "File Cleaner")
        tabs.addTab(registry_tab, "Registry Cleaner")
        
        layout.addWidget(tabs, 1)
        
        # Initialize disk space information
        self.update_disk_info()
    
    def update_disk_info(self):
        """Update disk space information."""
        disk_info = self.cleaner.get_disk_info()
        
        self.total_space_card.update_value(f"{disk_info['total_gb']:.1f} GB")
        self.free_space_card.update_value(f"{disk_info['free_gb']:.1f} GB")
    
    def analyze_system(self):
        """Analyze system for cleanable files."""
        self.show_loading("Analyzing system files...", "This may take a moment.")
        
        # Get selected cleaning options
        selected_options = []
        for item in self.cleaning_items:
            if item.is_checked():
                selected_options.append(item.title)
        
        # Create background worker
        worker = BackgroundWorker(self.cleaner.analyze, selected_options)
        worker.taskFinished.connect(self.on_analysis_complete)
        worker.start()
    
    def on_analysis_complete(self, result):
        """Handle completion of system analysis."""
        self.hide_loading()
        
        total_bytes = 0
        
        # Update UI with analysis results
        for item_title, size_bytes in result.items():
            for item_widget in self.cleaning_items:
                if item_widget.title == item_title:
                    if size_bytes < 1024:
                        size_text = f"{size_bytes} B"
                    elif size_bytes < 1024 * 1024:
                        size_text = f"{size_bytes / 1024:.1f} KB"
                    else:
                        size_text = f"{size_bytes / (1024 * 1024):.1f} MB"
                    
                    item_widget.set_size(size_text)
                    total_bytes += size_bytes
        
        # Update recoverable space
        if total_bytes < 1024 * 1024:
            recoverable_text = f"{total_bytes / 1024:.1f} KB"
        else:
            recoverable_text = f"{total_bytes / (1024 * 1024):.1f} MB"
        
        self.recoverable_card.update_value(recoverable_text)
    
    def clean_system(self):
        """Clean system based on selected options."""
        self.show_loading("Cleaning system files...", "This may take a while.")
        
        # Get selected cleaning options
        selected_options = []
        for item in self.cleaning_items:
            if item.is_checked():
                selected_options.append(item.title)
        
        # Create background worker
        worker = BackgroundWorker(self.cleaner.clean, selected_options)
        worker.taskFinished.connect(self.on_cleaning_complete)
        worker.start()
    
    def on_cleaning_complete(self, result):
        """Handle completion of system cleaning."""
        self.hide_loading()
        
        # Reset size labels
        for item in self.cleaning_items:
            item.set_size("--")
        
        # Update recoverable space
        self.recoverable_card.update_value("0 MB")
        
        # Update disk info
        self.update_disk_info()
    
    def scan_registry(self):
        """Scan registry for issues."""
        self.show_loading("Scanning Windows Registry...", "Looking for issues...")
        
        # Get selected registry scan options
        selected_options = []
        for option in self.reg_options:
            if option.isChecked():
                selected_options.append(option.text())
        
        # Create background worker
        worker = BackgroundWorker(self.registry_cleaner.scan, selected_options)
        worker.taskFinished.connect(self.on_registry_scan_complete)
        worker.start()
    
    def on_registry_scan_complete(self, result):
        """Handle completion of registry scan."""
        self.hide_loading()
        
        issue_count = result.get('issue_count', 0)
        
        if issue_count > 0:
            self.registry_results.setText(
                f"Found {issue_count} issues in the registry that can be fixed."
            )
            self.registry_results.setStyleSheet(
                f"padding: 20px; background-color: rgba(255, 149, 0, 0.1); border-radius: 5px; color: {COLORS['warning']};"
            )
        else:
            self.registry_results.setText("No registry issues found.")
            self.registry_results.setStyleSheet(
                f"padding: 20px; background-color: rgba(0, 183, 74, 0.1); border-radius: 5px; color: {COLORS['secondary']};"
            )
    
    def fix_registry(self):
        """Fix registry issues."""
        self.show_loading("Fixing Registry Issues...", "This may take a moment.")
        
        # Check if backup is selected
        create_backup = self.create_backup_cb.isChecked()
        
        # Create background worker
        worker = BackgroundWorker(self.registry_cleaner.fix_issues, create_backup)
        worker.taskFinished.connect(self.on_registry_fix_complete)
        worker.start()
    
    def on_registry_fix_complete(self, result):
        """Handle completion of registry fixing."""
        self.hide_loading()
        
        fixed_count = result.get('fixed_count', 0)
        backup_path = result.get('backup_path', '')
        
        if fixed_count > 0:
            message = f"Fixed {fixed_count} registry issues."
            if backup_path:
                message += f"\nRegistry backup created at: {backup_path}"
            
            self.registry_results.setText(message)
            self.registry_results.setStyleSheet(
                f"padding: 20px; background-color: rgba(0, 183, 74, 0.1); border-radius: 5px; color: {COLORS['secondary']};"
            )
        else:
            self.registry_results.setText("No issues were fixed.")
            self.registry_results.setStyleSheet(
                f"padding: 20px; background-color: rgba(0, 122, 255, 0.1); border-radius: 5px; color: {COLORS['primary']};"
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
