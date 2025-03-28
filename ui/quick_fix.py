"""
Quick Fix Tools widget for Microsoft Teams and Outlook issues.
"""

import os
import sys
import time
import logging
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QFrame, QProgressBar, QTabWidget, 
                             QCheckBox, QListWidget, QListWidgetItem, 
                             QSpacerItem, QSizePolicy, QGroupBox)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, pyqtSignal, pyqtSlot, QThread

from ui.theme import COLORS
from ui.components import Card, HeaderLabel, SubHeaderLabel, ActionButton, InfoCard
from ui.loading_screen import LoadingScreen
from utils.quick_fixes import MSTeamsFixer, OutlookFixer
from utils.background_worker import BackgroundWorker

logger = logging.getLogger(__name__)


class FixListItem(QFrame):
    """Custom list item widget for quick fixes."""
    
    def __init__(self, title, description, icon_name, parent=None):
        super().__init__(parent)
        self.title = title
        self.description = description
        
        self.setMinimumHeight(60)
        self.setObjectName("fix_list_item")
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 5, 10, 5)
        
        # Checkbox
        self.checkbox = QCheckBox()
        self.checkbox.setChecked(True)
        layout.addWidget(self.checkbox)
        
        # Icon
        icon_label = QLabel()
        icon = QIcon(get_icon(icon_name))
        pixmap = icon.pixmap(24, 24)
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
    
    def is_checked(self):
        """Return checkbox state."""
        return self.checkbox.isChecked()


class QuickFixWidget(QWidget):
    """Widget for quick fix tools for Microsoft Teams and Outlook."""
    
    def __init__(self):
        super().__init__()
        self.teams_fixer = MSTeamsFixer()
        self.outlook_fixer = OutlookFixer()
        self.initUI()
        
        # Loading screen instance
        self.loading_screen = None
    
    def initUI(self):
        """Initialize the quick fix UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)
        
        # Header
        header = HeaderLabel("Quick Fix Tools")
        subheader = SubHeaderLabel("Fix common issues with Microsoft Teams and Outlook")
        
        layout.addWidget(header)
        layout.addWidget(subheader)
        
        # Status cards
        status_layout = QHBoxLayout()
        
        self.teams_status_card = InfoCard(
            title="MS Teams",
            icon_name="teams",
            value="Checking...",
            description="Teams status"
        )
        
        self.outlook_status_card = InfoCard(
            title="Outlook",
            icon_name="outlook",
            value="Checking...",
            description="Outlook status"
        )
        
        self.cache_size_card = InfoCard(
            title="Cache Size",
            icon_name="cache",
            value="--",
            description="Total application cache"
        )
        
        status_layout.addWidget(self.teams_status_card)
        status_layout.addWidget(self.outlook_status_card)
        status_layout.addWidget(self.cache_size_card)
        
        layout.addLayout(status_layout)
        
        # Tabs for different applications
        tabs = QTabWidget()
        tabs.setObjectName("quickfix_tabs")
        
        # MS Teams Tab
        teams_tab = QWidget()
        teams_layout = QVBoxLayout(teams_tab)
        teams_layout.setContentsMargins(15, 15, 15, 15)
        
        # MS Teams Info
        teams_info = QLabel(
            "Microsoft Teams can encounter various issues that affect its performance "
            "or functionality. Below are common fixes that can resolve these issues."
        )
        teams_info.setWordWrap(True)
        teams_info.setStyleSheet("background-color: rgba(0, 122, 255, 0.1); padding: 10px; border-radius: 5px;")
        teams_layout.addWidget(teams_info)
        teams_layout.addSpacing(10)
        
        # Teams Fixes
        teams_fixes_card = Card()
        teams_fixes_layout = QVBoxLayout(teams_fixes_card)
        
        teams_fixes_header = QLabel("Microsoft Teams Fixes")
        teams_fixes_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        teams_fixes_layout.addWidget(teams_fixes_header)
        
        # Teams fixes list
        self.teams_fixes = []
        
        teams_cache = FixListItem(
            "Clear Cache",
            "Remove Teams cache files to fix loading and sign-in issues",
            "cache"
        )
        
        teams_restart = FixListItem(
            "End Background Processes",
            "Terminate all Teams processes that may be running in the background",
            "process"
        )
        
        teams_config = FixListItem(
            "Reset Configuration",
            "Reset Teams application configuration to default",
            "settings"
        )
        
        teams_cookies = FixListItem(
            "Clear Web Storage",
            "Clear cookies and site data used by Teams",
            "clean"
        )
        
        teams_network = FixListItem(
            "Reset Network Settings",
            "Reset network configuration for Teams connections",
            "network"
        )
        
        teams_gpu = FixListItem(
            "Disable GPU Hardware Acceleration",
            "Fix graphical glitches and crashes on some systems",
            "gpu"
        )
        
        self.teams_fixes = [
            teams_cache,
            teams_restart,
            teams_config,
            teams_cookies,
            teams_network,
            teams_gpu
        ]
        
        for fix in self.teams_fixes:
            teams_fixes_layout.addWidget(fix)
        
        teams_layout.addWidget(teams_fixes_card, 1)
        
        # Teams actions
        teams_actions = QHBoxLayout()
        
        teams_analyze_btn = ActionButton("Analyze Issues", "search", "primary")
        teams_analyze_btn.clicked.connect(self.analyze_teams)
        
        teams_fix_btn = ActionButton("Apply Selected Fixes", "tools", "success")
        teams_fix_btn.clicked.connect(self.fix_teams)
        
        teams_actions.addStretch()
        teams_actions.addWidget(teams_analyze_btn)
        teams_actions.addWidget(teams_fix_btn)
        
        teams_layout.addLayout(teams_actions)
        
        # Outlook Tab
        outlook_tab = QWidget()
        outlook_layout = QVBoxLayout(outlook_tab)
        outlook_layout.setContentsMargins(15, 15, 15, 15)
        
        # Outlook Info
        outlook_info = QLabel(
            "Microsoft Outlook issues can affect your email productivity and workflow. "
            "The tools below can help resolve common Outlook problems."
        )
        outlook_info.setWordWrap(True)
        outlook_info.setStyleSheet("background-color: rgba(0, 122, 255, 0.1); padding: 10px; border-radius: 5px;")
        outlook_layout.addWidget(outlook_info)
        outlook_layout.addSpacing(10)
        
        # Outlook Fixes
        outlook_fixes_card = Card()
        outlook_fixes_layout = QVBoxLayout(outlook_fixes_card)
        
        outlook_fixes_header = QLabel("Microsoft Outlook Fixes")
        outlook_fixes_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        outlook_fixes_layout.addWidget(outlook_fixes_header)
        
        # Outlook fixes list
        self.outlook_fixes = []
        
        outlook_profile = FixListItem(
            "Repair Outlook Profile",
            "Fix profile corruption and sign-in issues",
            "profile"
        )
        
        outlook_cache = FixListItem(
            "Clear Outlook Cache",
            "Remove OST/PST cache files to fix search and performance issues",
            "cache"
        )
        
        outlook_search = FixListItem(
            "Rebuild Search Index",
            "Fix search functionality issues in Outlook",
            "search"
        )
        
        outlook_autodiscover = FixListItem(
            "Reset AutoDiscover",
            "Fix connection issues with Exchange servers",
            "network"
        )
        
        outlook_rules = FixListItem(
            "Repair Rules",
            "Fix broken or non-functioning email rules",
            "rules"
        )
        
        outlook_addins = FixListItem(
            "Disable Add-ins",
            "Resolve crashes and slowness caused by problematic add-ins",
            "settings"
        )
        
        self.outlook_fixes = [
            outlook_profile,
            outlook_cache,
            outlook_search,
            outlook_autodiscover,
            outlook_rules,
            outlook_addins
        ]
        
        for fix in self.outlook_fixes:
            outlook_fixes_layout.addWidget(fix)
        
        outlook_layout.addWidget(outlook_fixes_card, 1)
        
        # Outlook actions
        outlook_actions = QHBoxLayout()
        
        outlook_analyze_btn = ActionButton("Analyze Issues", "search", "primary")
        outlook_analyze_btn.clicked.connect(self.analyze_outlook)
        
        outlook_fix_btn = ActionButton("Apply Selected Fixes", "tools", "success")
        outlook_fix_btn.clicked.connect(self.fix_outlook)
        
        outlook_actions.addStretch()
        outlook_actions.addWidget(outlook_analyze_btn)
        outlook_actions.addWidget(outlook_fix_btn)
        
        outlook_layout.addLayout(outlook_actions)
        
        # Add tabs
        tabs.addTab(teams_tab, "Microsoft Teams")
        tabs.addTab(outlook_tab, "Microsoft Outlook")
        
        layout.addWidget(tabs, 1)
        
        # Update status on load
        self.update_application_status()
    
    def update_application_status(self):
        """Update the status cards with application information."""
        # Create background worker for Teams status
        teams_worker = BackgroundWorker(self.teams_fixer.check_status)
        teams_worker.taskFinished.connect(self.update_teams_status)
        teams_worker.start()
        
        # Create background worker for Outlook status
        outlook_worker = BackgroundWorker(self.outlook_fixer.check_status)
        outlook_worker.taskFinished.connect(self.update_outlook_status)
        outlook_worker.start()
        
        # Update cache size
        cache_worker = BackgroundWorker(self.calculate_cache_size)
        cache_worker.taskFinished.connect(self.update_cache_size)
        cache_worker.start()
    
    def calculate_cache_size(self):
        """Calculate the total cache size for Teams and Outlook."""
        teams_cache = self.teams_fixer.get_cache_size()
        outlook_cache = self.outlook_fixer.get_cache_size()
        
        return {
            'teams_size': teams_cache,
            'outlook_size': outlook_cache,
            'total_size': teams_cache + outlook_cache
        }
    
    def update_teams_status(self, status):
        """Update Teams status card with results."""
        status_text = status.get('status', 'Unknown')
        status_color = COLORS['secondary'] if status_text == "Healthy" else COLORS['warning']
        
        self.teams_status_card.update_value(status_text, color=status_color)
    
    def update_outlook_status(self, status):
        """Update Outlook status card with results."""
        status_text = status.get('status', 'Unknown')
        status_color = COLORS['secondary'] if status_text == "Healthy" else COLORS['warning']
        
        self.outlook_status_card.update_value(status_text, color=status_color)
    
    def update_cache_size(self, cache_info):
        """Update cache size information."""
        total_size = cache_info.get('total_size', 0)
        
        if total_size < 1024 * 1024:  # Less than 1 MB
            size_text = f"{total_size / 1024:.1f} KB"
        else:
            size_text = f"{total_size / (1024 * 1024):.1f} MB"
        
        self.cache_size_card.update_value(size_text)
    
    def analyze_teams(self):
        """Analyze Microsoft Teams for issues."""
        self.show_loading("Analyzing Microsoft Teams...", "Checking for issues...")
        
        # Create background worker
        worker = BackgroundWorker(self.teams_fixer.analyze)
        worker.taskFinished.connect(self.on_teams_analysis_complete)
        worker.start()
    
    def on_teams_analysis_complete(self, issues):
        """Handle completion of Teams analysis."""
        self.hide_loading()
        
        # Update the Teams status
        if issues.get('found_issues', False):
            self.teams_status_card.update_value("Issues Found", color=COLORS['warning'])
            
            # Check/uncheck fixes based on analysis
            for fix in self.teams_fixes:
                if fix.title in issues.get('recommended_fixes', []):
                    fix.checkbox.setChecked(True)
                else:
                    fix.checkbox.setChecked(False)
        else:
            self.teams_status_card.update_value("Healthy", color=COLORS['secondary'])
    
    def fix_teams(self):
        """Apply selected Teams fixes."""
        # Get selected fixes
        selected_fixes = []
        for fix in self.teams_fixes:
            if fix.is_checked():
                selected_fixes.append(fix.title)
        
        if not selected_fixes:
            return
        
        self.show_loading("Fixing Microsoft Teams...", "Applying selected fixes...")
        
        # Create background worker
        worker = BackgroundWorker(self.teams_fixer.apply_fixes, selected_fixes)
        worker.taskFinished.connect(self.on_teams_fixes_complete)
        worker.start()
    
    def on_teams_fixes_complete(self, results):
        """Handle completion of Teams fixes."""
        self.hide_loading()
        
        # Update the Teams status
        if results.get('success', False):
            self.teams_status_card.update_value("Fixed", color=COLORS['secondary'])
        else:
            self.teams_status_card.update_value("Fix Failed", color=COLORS['error'])
        
        # Update cache size
        cache_worker = BackgroundWorker(self.calculate_cache_size)
        cache_worker.taskFinished.connect(self.update_cache_size)
        cache_worker.start()
    
    def analyze_outlook(self):
        """Analyze Microsoft Outlook for issues."""
        self.show_loading("Analyzing Microsoft Outlook...", "Checking for issues...")
        
        # Create background worker
        worker = BackgroundWorker(self.outlook_fixer.analyze)
        worker.taskFinished.connect(self.on_outlook_analysis_complete)
        worker.start()
    
    def on_outlook_analysis_complete(self, issues):
        """Handle completion of Outlook analysis."""
        self.hide_loading()
        
        # Update the Outlook status
        if issues.get('found_issues', False):
            self.outlook_status_card.update_value("Issues Found", color=COLORS['warning'])
            
            # Check/uncheck fixes based on analysis
            for fix in self.outlook_fixes:
                if fix.title in issues.get('recommended_fixes', []):
                    fix.checkbox.setChecked(True)
                else:
                    fix.checkbox.setChecked(False)
        else:
            self.outlook_status_card.update_value("Healthy", color=COLORS['secondary'])
    
    def fix_outlook(self):
        """Apply selected Outlook fixes."""
        # Get selected fixes
        selected_fixes = []
        for fix in self.outlook_fixes:
            if fix.is_checked():
                selected_fixes.append(fix.title)
        
        if not selected_fixes:
            return
        
        self.show_loading("Fixing Microsoft Outlook...", "Applying selected fixes...")
        
        # Create background worker
        worker = BackgroundWorker(self.outlook_fixer.apply_fixes, selected_fixes)
        worker.taskFinished.connect(self.on_outlook_fixes_complete)
        worker.start()
    
    def on_outlook_fixes_complete(self, results):
        """Handle completion of Outlook fixes."""
        self.hide_loading()
        
        # Update the Outlook status
        if results.get('success', False):
            self.outlook_status_card.update_value("Fixed", color=COLORS['secondary'])
        else:
            self.outlook_status_card.update_value("Fix Failed", color=COLORS['error'])
        
        # Update cache size
        cache_worker = BackgroundWorker(self.calculate_cache_size)
        cache_worker.taskFinished.connect(self.update_cache_size)
        cache_worker.start()
    
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


# Import at the bottom to avoid circular import
from assets.icons import get_icon
