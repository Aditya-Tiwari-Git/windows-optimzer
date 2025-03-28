"""
Settings widget for configuring application preferences.
"""

import os
import sys
import json
import logging
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QFrame, QCheckBox, QComboBox,
                             QGroupBox, QSpinBox, QFileDialog, QScrollArea,
                             QFormLayout, QRadioButton, QButtonGroup, QSlider,
                             QSpacerItem, QSizePolicy)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, pyqtSignal, QSettings, QSize

from ui.theme import COLORS, set_theme, DARK_MODE, LIGHT_MODE
from ui.components import Card, HeaderLabel, SubHeaderLabel, ActionButton
from assets.icons import get_icon

logger = logging.getLogger(__name__)


class SettingsWidget(QWidget):
    """Widget for application settings and preferences."""
    
    themeChanged = pyqtSignal(bool)
    
    def __init__(self):
        super().__init__()
        self.settings = QSettings("WindowsOptimizer", "Windows System Optimizer")
        self.initUI()
        self.load_settings()
    
    def initUI(self):
        """Initialize the settings UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)
        
        # Header
        header = HeaderLabel("Settings")
        subheader = SubHeaderLabel("Configure application preferences")
        
        layout.addWidget(header)
        layout.addWidget(subheader)
        
        # Create a scroll area for settings
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setContentsMargins(0, 0, 0, 0)
        scroll_layout.setSpacing(20)
        
        # Appearance settings
        appearance_card = Card()
        appearance_layout = QVBoxLayout(appearance_card)
        
        appearance_header = QLabel("Appearance")
        appearance_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        appearance_layout.addWidget(appearance_header)
        
        # Theme setting
        theme_layout = QHBoxLayout()
        theme_layout.setContentsMargins(10, 10, 10, 10)
        
        theme_label = QLabel("Theme:")
        theme_label.setFont(QFont("Segoe UI", 10))
        
        self.theme_toggle = QPushButton()
        self.theme_toggle.setCheckable(True)
        self.theme_toggle.setMinimumWidth(120)
        self.theme_toggle.clicked.connect(self.toggle_theme)
        
        theme_layout.addWidget(theme_label)
        theme_layout.addStretch()
        theme_layout.addWidget(self.theme_toggle)
        
        appearance_layout.addLayout(theme_layout)
        
        # Font size
        font_layout = QHBoxLayout()
        font_layout.setContentsMargins(10, 10, 10, 10)
        
        font_label = QLabel("Font Size:")
        font_label.setFont(QFont("Segoe UI", 10))
        
        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems(["Small", "Medium", "Large"])
        self.font_size_combo.setCurrentIndex(1)  # Default to Medium
        
        font_layout.addWidget(font_label)
        font_layout.addStretch()
        font_layout.addWidget(self.font_size_combo)
        
        appearance_layout.addLayout(font_layout)
        
        # Animation
        animation_layout = QHBoxLayout()
        animation_layout.setContentsMargins(10, 10, 10, 10)
        
        animation_label = QLabel("Enable Animations:")
        animation_label.setFont(QFont("Segoe UI", 10))
        
        self.animation_checkbox = QCheckBox()
        self.animation_checkbox.setChecked(True)
        
        animation_layout.addWidget(animation_label)
        animation_layout.addStretch()
        animation_layout.addWidget(self.animation_checkbox)
        
        appearance_layout.addLayout(animation_layout)
        
        scroll_layout.addWidget(appearance_card)
        
        # Behavior settings
        behavior_card = Card()
        behavior_layout = QVBoxLayout(behavior_card)
        
        behavior_header = QLabel("Behavior")
        behavior_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        behavior_layout.addWidget(behavior_header)
        
        # Start with Windows
        startup_layout = QHBoxLayout()
        startup_layout.setContentsMargins(10, 10, 10, 10)
        
        startup_label = QLabel("Start with Windows:")
        startup_label.setFont(QFont("Segoe UI", 10))
        
        self.startup_checkbox = QCheckBox()
        self.startup_checkbox.setChecked(False)
        self.startup_checkbox.toggled.connect(self.toggle_startup)
        
        startup_layout.addWidget(startup_label)
        startup_layout.addStretch()
        startup_layout.addWidget(self.startup_checkbox)
        
        behavior_layout.addLayout(startup_layout)
        
        # Minimize to tray
        tray_layout = QHBoxLayout()
        tray_layout.setContentsMargins(10, 10, 10, 10)
        
        tray_label = QLabel("Minimize to System Tray:")
        tray_label.setFont(QFont("Segoe UI", 10))
        
        self.tray_checkbox = QCheckBox()
        self.tray_checkbox.setChecked(True)
        
        tray_layout.addWidget(tray_label)
        tray_layout.addStretch()
        tray_layout.addWidget(self.tray_checkbox)
        
        behavior_layout.addLayout(tray_layout)
        
        # Background scan interval
        scan_layout = QHBoxLayout()
        scan_layout.setContentsMargins(10, 10, 10, 10)
        
        scan_label = QLabel("Background Scan Interval:")
        scan_label.setFont(QFont("Segoe UI", 10))
        
        self.scan_combo = QComboBox()
        self.scan_combo.addItems(["Never", "Daily", "Weekly", "Monthly"])
        self.scan_combo.setCurrentIndex(1)  # Default to Daily
        
        scan_layout.addWidget(scan_label)
        scan_layout.addStretch()
        scan_layout.addWidget(self.scan_combo)
        
        behavior_layout.addLayout(scan_layout)
        
        scroll_layout.addWidget(behavior_card)
        
        # Notifications settings
        notifications_card = Card()
        notifications_layout = QVBoxLayout(notifications_card)
        
        notifications_header = QLabel("Notifications")
        notifications_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        notifications_layout.addWidget(notifications_header)
        
        # Enable notifications
        enable_notif_layout = QHBoxLayout()
        enable_notif_layout.setContentsMargins(10, 10, 10, 10)
        
        enable_notif_label = QLabel("Enable Notifications:")
        enable_notif_label.setFont(QFont("Segoe UI", 10))
        
        self.enable_notif_checkbox = QCheckBox()
        self.enable_notif_checkbox.setChecked(True)
        
        enable_notif_layout.addWidget(enable_notif_label)
        enable_notif_layout.addStretch()
        enable_notif_layout.addWidget(self.enable_notif_checkbox)
        
        notifications_layout.addLayout(enable_notif_layout)
        
        # Notify of issues
        issues_notif_layout = QHBoxLayout()
        issues_notif_layout.setContentsMargins(10, 10, 10, 10)
        
        issues_notif_label = QLabel("Notify of System Issues:")
        issues_notif_label.setFont(QFont("Segoe UI", 10))
        
        self.issues_notif_checkbox = QCheckBox()
        self.issues_notif_checkbox.setChecked(True)
        
        issues_notif_layout.addWidget(issues_notif_label)
        issues_notif_layout.addStretch()
        issues_notif_layout.addWidget(self.issues_notif_checkbox)
        
        notifications_layout.addLayout(issues_notif_layout)
        
        # Notify of updates
        updates_notif_layout = QHBoxLayout()
        updates_notif_layout.setContentsMargins(10, 10, 10, 10)
        
        updates_notif_label = QLabel("Notify of Driver Updates:")
        updates_notif_label.setFont(QFont("Segoe UI", 10))
        
        self.updates_notif_checkbox = QCheckBox()
        self.updates_notif_checkbox.setChecked(True)
        
        updates_notif_layout.addWidget(updates_notif_label)
        updates_notif_layout.addStretch()
        updates_notif_layout.addWidget(self.updates_notif_checkbox)
        
        notifications_layout.addLayout(updates_notif_layout)
        
        scroll_layout.addWidget(notifications_card)
        
        # Advanced settings
        advanced_card = Card()
        advanced_layout = QVBoxLayout(advanced_card)
        
        advanced_header = QLabel("Advanced")
        advanced_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        advanced_layout.addWidget(advanced_header)
        
        # Log level
        log_layout = QHBoxLayout()
        log_layout.setContentsMargins(10, 10, 10, 10)
        
        log_label = QLabel("Logging Level:")
        log_label.setFont(QFont("Segoe UI", 10))
        
        self.log_combo = QComboBox()
        self.log_combo.addItems(["Error", "Warning", "Info", "Debug"])
        self.log_combo.setCurrentIndex(2)  # Default to Info
        self.log_combo.currentIndexChanged.connect(self.change_log_level)
        
        log_layout.addWidget(log_label)
        log_layout.addStretch()
        log_layout.addWidget(self.log_combo)
        
        advanced_layout.addLayout(log_layout)
        
        # Backup directory
        backup_layout = QHBoxLayout()
        backup_layout.setContentsMargins(10, 10, 10, 10)
        
        backup_label = QLabel("Backup Directory:")
        backup_label.setFont(QFont("Segoe UI", 10))
        
        self.backup_path = QLabel(os.path.join(os.path.expanduser("~"), "Documents", "WindowsOptimizer"))
        self.backup_path.setStyleSheet(f"color: {COLORS['primary']};")
        
        backup_browse_btn = QPushButton("Browse")
        backup_browse_btn.clicked.connect(self.browse_backup_dir)
        
        backup_layout.addWidget(backup_label)
        backup_layout.addWidget(self.backup_path, 1)
        backup_layout.addWidget(backup_browse_btn)
        
        advanced_layout.addLayout(backup_layout)
        
        scroll_layout.addWidget(advanced_card)
        
        # Add a reset and save button
        buttons_layout = QHBoxLayout()
        
        reset_btn = ActionButton("Reset to Default", "reset", "warning")
        reset_btn.clicked.connect(self.reset_settings)
        
        save_btn = ActionButton("Save Settings", "save", "success")
        save_btn.clicked.connect(self.save_settings)
        
        buttons_layout.addWidget(reset_btn)
        buttons_layout.addStretch()
        buttons_layout.addWidget(save_btn)
        
        scroll_layout.addLayout(buttons_layout)
        
        # Add spacer at the bottom
        scroll_layout.addItem(QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding))
        
        # Set up the scroll area
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)
    
    def load_settings(self):
        """Load settings from QSettings."""
        # Theme
        is_dark_mode = self.settings.value("appearance/dark_mode", True, type=bool)
        self.theme_toggle.setChecked(is_dark_mode)
        self.update_theme_button(is_dark_mode)
        
        # Font size
        font_size = self.settings.value("appearance/font_size", "Medium")
        index = self.font_size_combo.findText(font_size)
        if index >= 0:
            self.font_size_combo.setCurrentIndex(index)
        
        # Animations
        enable_animations = self.settings.value("appearance/animations", True, type=bool)
        self.animation_checkbox.setChecked(enable_animations)
        
        # Start with Windows
        start_with_windows = self.settings.value("behavior/start_with_windows", False, type=bool)
        self.startup_checkbox.setChecked(start_with_windows)
        
        # Minimize to tray
        minimize_to_tray = self.settings.value("behavior/minimize_to_tray", True, type=bool)
        self.tray_checkbox.setChecked(minimize_to_tray)
        
        # Scan interval
        scan_interval = self.settings.value("behavior/scan_interval", "Daily")
        index = self.scan_combo.findText(scan_interval)
        if index >= 0:
            self.scan_combo.setCurrentIndex(index)
        
        # Enable notifications
        enable_notifications = self.settings.value("notifications/enable", True, type=bool)
        self.enable_notif_checkbox.setChecked(enable_notifications)
        
        # System issues notifications
        issues_notifications = self.settings.value("notifications/system_issues", True, type=bool)
        self.issues_notif_checkbox.setChecked(issues_notifications)
        
        # Driver updates notifications
        updates_notifications = self.settings.value("notifications/driver_updates", True, type=bool)
        self.updates_notif_checkbox.setChecked(updates_notifications)
        
        # Log level
        log_level = self.settings.value("advanced/log_level", "Info")
        index = self.log_combo.findText(log_level)
        if index >= 0:
            self.log_combo.setCurrentIndex(index)
        
        # Backup directory
        backup_dir = self.settings.value("advanced/backup_directory", 
                                      os.path.join(os.path.expanduser("~"), "Documents", "WindowsOptimizer"))
        self.backup_path.setText(backup_dir)
    
    def save_settings(self):
        """Save settings to QSettings."""
        # Theme
        self.settings.setValue("appearance/dark_mode", self.theme_toggle.isChecked())
        
        # Font size
        self.settings.setValue("appearance/font_size", self.font_size_combo.currentText())
        
        # Animations
        self.settings.setValue("appearance/animations", self.animation_checkbox.isChecked())
        
        # Start with Windows
        self.settings.setValue("behavior/start_with_windows", self.startup_checkbox.isChecked())
        
        # Minimize to tray
        self.settings.setValue("behavior/minimize_to_tray", self.tray_checkbox.isChecked())
        
        # Scan interval
        self.settings.setValue("behavior/scan_interval", self.scan_combo.currentText())
        
        # Enable notifications
        self.settings.setValue("notifications/enable", self.enable_notif_checkbox.isChecked())
        
        # System issues notifications
        self.settings.setValue("notifications/system_issues", self.issues_notif_checkbox.isChecked())
        
        # Driver updates notifications
        self.settings.setValue("notifications/driver_updates", self.updates_notif_checkbox.isChecked())
        
        # Log level
        self.settings.setValue("advanced/log_level", self.log_combo.currentText())
        
        # Backup directory
        self.settings.setValue("advanced/backup_directory", self.backup_path.text())
        
        # Sync settings to disk
        self.settings.sync()
        
        # Call functions that need to apply settings immediately
        self.change_log_level(self.log_combo.currentIndex())
    
    def reset_settings(self):
        """Reset settings to default values."""
        # Theme
        self.theme_toggle.setChecked(True)
        self.update_theme_button(True)
        
        # Font size
        self.font_size_combo.setCurrentText("Medium")
        
        # Animations
        self.animation_checkbox.setChecked(True)
        
        # Start with Windows
        self.startup_checkbox.setChecked(False)
        
        # Minimize to tray
        self.tray_checkbox.setChecked(True)
        
        # Scan interval
        self.scan_combo.setCurrentText("Daily")
        
        # Enable notifications
        self.enable_notif_checkbox.setChecked(True)
        
        # System issues notifications
        self.issues_notif_checkbox.setChecked(True)
        
        # Driver updates notifications
        self.updates_notif_checkbox.setChecked(True)
        
        # Log level
        self.log_combo.setCurrentText("Info")
        
        # Backup directory
        default_backup_dir = os.path.join(os.path.expanduser("~"), "Documents", "WindowsOptimizer")
        self.backup_path.setText(default_backup_dir)
    
    def toggle_theme(self):
        """Toggle between light and dark theme."""
        is_dark_mode = self.theme_toggle.isChecked()
        self.update_theme_button(is_dark_mode)
        
        # Emit signal to notify app to change theme
        self.themeChanged.emit(is_dark_mode)
    
    def update_theme_button(self, is_dark_mode):
        """Update the theme button text and icon based on current mode."""
        if is_dark_mode:
            self.theme_toggle.setText("Dark Mode")
            self.theme_toggle.setIcon(QIcon(get_icon("moon")))
        else:
            self.theme_toggle.setText("Light Mode")
            self.theme_toggle.setIcon(QIcon(get_icon("sun")))
    
    def toggle_startup(self, checked):
        """Set application to start with Windows."""
        import winreg
        app_name = "Windows System Optimizer"
        app_path = sys.executable
        
        try:
            if checked:
                # Add to startup registry
                reg_key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Run",
                    0, winreg.KEY_SET_VALUE
                )
                winreg.SetValueEx(reg_key, app_name, 0, winreg.REG_SZ, app_path)
                winreg.CloseKey(reg_key)
                logger.info("Added application to startup registry")
            else:
                # Remove from startup registry
                reg_key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Run",
                    0, winreg.KEY_SET_VALUE
                )
                try:
                    winreg.DeleteValue(reg_key, app_name)
                except FileNotFoundError:
                    # Key doesn't exist, that's fine
                    pass
                winreg.CloseKey(reg_key)
                logger.info("Removed application from startup registry")
        except Exception as e:
            logger.error(f"Error setting startup registry: {str(e)}")
            # Revert checkbox if operation failed
            self.startup_checkbox.setChecked(not checked)
    
    def change_log_level(self, index):
        """Change the application's logging level."""
        log_levels = {
            "Error": logging.ERROR,
            "Warning": logging.WARNING,
            "Info": logging.INFO,
            "Debug": logging.DEBUG
        }
        
        selected_level = self.log_combo.currentText()
        level = log_levels.get(selected_level, logging.INFO)
        
        # Set root logger level
        logging.getLogger().setLevel(level)
        logger.info(f"Changed logging level to {selected_level}")
    
    def browse_backup_dir(self):
        """Open file dialog to select backup directory."""
        current_dir = self.backup_path.text()
        
        directory = QFileDialog.getExistingDirectory(
            self,
            "Select Backup Directory",
            current_dir,
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )
        
        if directory:
            self.backup_path.setText(directory)
