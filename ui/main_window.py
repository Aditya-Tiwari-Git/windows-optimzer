"""
Main window for the Windows System Optimizer.
This module provides the main application window with navigation and global controls.
"""

import sys
import os
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QLabel, QStackedWidget, QSizePolicy,
                            QFrame, QSpacerItem, QMessageBox)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QIcon, QFont, QPixmap

from .dashboard_widget import DashboardWidget
from .cleaning_widget import CleaningWidget
from .network_widget import NetworkWidget
from .optimization_widget import OptimizationWidget
from .styles import COLORS, set_dark_mode, set_light_mode

class SidebarButton(QPushButton):
    """Custom sidebar button with icon and text."""
    def __init__(self, icon_name, text, parent=None):
        super().__init__(parent)
        self.setText(text)
        self.setIcon(QIcon(icon_name))
        self.setIconSize(QSize(24, 24))
        self.setMinimumHeight(48)
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        
class MainWindow(QMainWindow):
    """Main application window."""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Windows System Optimizer")
        self.setMinimumSize(1000, 650)
        self.dark_mode = False  # Default to light mode
        
        # Create central widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        self.main_layout = QHBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        # Create and setup the sidebar
        self.setup_sidebar()
        
        # Create and setup the content area
        self.setup_content_area()
        
        # Initialize with dashboard page
        self.show_dashboard()
        self.dashboard_btn.setChecked(True)
        
        # Apply initial styling
        self.apply_styles()
    
    def setup_sidebar(self):
        """Set up the sidebar with navigation buttons."""
        # Sidebar container
        self.sidebar = QFrame()
        self.sidebar.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.sidebar.setFixedWidth(220)
        
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(10, 20, 10, 20)
        sidebar_layout.setSpacing(10)
        
        # App title
        title_label = QLabel("Win Optimizer")
        title_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(title_label)
        
        sidebar_layout.addSpacing(30)
        
        # Navigation buttons
        self.dashboard_btn = SidebarButton("fa.home", "Dashboard")
        self.dashboard_btn.clicked.connect(self.show_dashboard)
        
        self.cleaning_btn = SidebarButton("fa.broom", "System Cleaning")
        self.cleaning_btn.clicked.connect(self.show_cleaning)
        
        self.network_btn = SidebarButton("fa.network-wired", "Network Diagnostics")
        self.network_btn.clicked.connect(self.show_network)
        
        self.optimization_btn = SidebarButton("fa.sliders-h", "Optimization")
        self.optimization_btn.clicked.connect(self.show_optimization)
        
        # Add buttons to layout
        for btn in [self.dashboard_btn, self.cleaning_btn, self.network_btn, self.optimization_btn]:
            sidebar_layout.addWidget(btn)
        
        # Add spacer to push settings to bottom
        sidebar_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        
        # Settings button
        self.toggle_theme_btn = SidebarButton("fa.adjust", "Toggle Dark Mode")
        self.toggle_theme_btn.clicked.connect(self.toggle_theme)
        sidebar_layout.addWidget(self.toggle_theme_btn)
        
        # About button
        self.about_btn = SidebarButton("fa.info-circle", "About")
        self.about_btn.clicked.connect(self.show_about)
        sidebar_layout.addWidget(self.about_btn)
        
        # Add sidebar to main layout
        self.main_layout.addWidget(self.sidebar)
    
    def setup_content_area(self):
        """Set up the main content area with stacked widgets."""
        # Content container
        self.content_container = QFrame()
        
        content_layout = QVBoxLayout(self.content_container)
        content_layout.setContentsMargins(20, 20, 20, 20)
        
        # Stack widget to manage different pages
        self.stack_widget = QStackedWidget()
        
        # Create page widgets
        self.dashboard_widget = DashboardWidget()
        self.cleaning_widget = CleaningWidget()
        self.network_widget = NetworkWidget()
        self.optimization_widget = OptimizationWidget()
        
        # Add widgets to stack
        self.stack_widget.addWidget(self.dashboard_widget)
        self.stack_widget.addWidget(self.cleaning_widget)
        self.stack_widget.addWidget(self.network_widget)
        self.stack_widget.addWidget(self.optimization_widget)
        
        content_layout.addWidget(self.stack_widget)
        
        # Add content container to main layout
        self.main_layout.addWidget(self.content_container)
    
    def show_dashboard(self):
        """Switch to dashboard page."""
        self.stack_widget.setCurrentIndex(0)
        self._update_button_states(self.dashboard_btn)
    
    def show_cleaning(self):
        """Switch to cleaning page."""
        self.stack_widget.setCurrentIndex(1)
        self._update_button_states(self.cleaning_btn)
    
    def show_network(self):
        """Switch to network page."""
        self.stack_widget.setCurrentIndex(2)
        self._update_button_states(self.network_btn)
    
    def show_optimization(self):
        """Switch to optimization page."""
        self.stack_widget.setCurrentIndex(3)
        self._update_button_states(self.optimization_btn)
    
    def _update_button_states(self, active_button):
        """Update the state of sidebar buttons to show active page."""
        for btn in [self.dashboard_btn, self.cleaning_btn, self.network_btn, self.optimization_btn]:
            btn.setChecked(btn == active_button)
    
    def toggle_theme(self):
        """Toggle between light and dark mode."""
        self.dark_mode = not self.dark_mode
        self.apply_styles()
        
        # Update button text
        if self.dark_mode:
            self.toggle_theme_btn.setText("Toggle Light Mode")
        else:
            self.toggle_theme_btn.setText("Toggle Dark Mode")
    
    def apply_styles(self):
        """Apply the current theme styles to all components."""
        if self.dark_mode:
            set_dark_mode(self)
        else:
            set_light_mode(self)
    
    def show_about(self):
        """Show the about dialog."""
        about_box = QMessageBox(self)
        about_box.setWindowTitle("About Windows System Optimizer")
        about_box.setTextFormat(Qt.RichText)
        
        about_text = """
        <h3>Windows System Optimizer</h3>
        <p>Version 1.0.0</p>
        <p>A modern Windows system optimization utility with cleaning, 
        network diagnostics, and performance monitoring features.</p>
        <p>&copy; 2023 WinOptimizer</p>
        """
        
        about_box.setText(about_text)
        about_box.setStandardButtons(QMessageBox.Ok)
        about_box.exec_()
