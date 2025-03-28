"""
Reusable UI components for the application.
"""

from PyQt5.QtWidgets import (QFrame, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
                             QSizePolicy)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QSize, pyqtSignal

from ui.theme import COLORS
from assets.icons import get_icon


class Card(QFrame):
    """Card component with shadow and rounded corners."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setProperty("class", "Card")
        
        # Set minimum size and size policy
        self.setMinimumSize(100, 80)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        
        # Set up shadow and border
        self.setStyleSheet(
            "border-radius: 8px; "
            "background-color: white; "
            "border: 1px solid #E0E0E0;"
        )


class HeaderLabel(QLabel):
    """Main header label with consistent styling."""
    
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setProperty("class", "header")
        self.setFont(QFont("Segoe UI", 18, QFont.Bold))


class SubHeaderLabel(QLabel):
    """Sub-header label with consistent styling."""
    
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setProperty("class", "subheader")
        self.setFont(QFont("Segoe UI", 12))
        self.setStyleSheet(f"color: {COLORS['text']}; opacity: 0.8;")


class ActionButton(QPushButton):
    """Styled action button with icon and text."""
    
    def __init__(self, text, icon_name=None, button_type="primary", parent=None):
        super().__init__(text, parent)
        
        # Set button properties
        self.setProperty("class", button_type)
        
        # Add icon if specified
        if icon_name:
            self.setIcon(QIcon(get_icon(icon_name)))
            self.setIconSize(QSize(18, 18))
        
        # Set fixed height and minimum width
        self.setFixedHeight(36)
        self.setMinimumWidth(120)
        
        # Add padding
        self.setStyleSheet(f"padding-left: 16px; padding-right: 16px;")


class SidebarButton(QPushButton):
    """Styled button for sidebar navigation."""
    
    def __init__(self, text, icon_name=None, parent=None):
        super().__init__(text, parent)
        self.setObjectName("sidebar_button")
        
        # Set fixed height
        self.setFixedHeight(40)
        
        # Add icon if specified
        if icon_name:
            self.setIcon(QIcon(get_icon(icon_name)))
            self.setIconSize(QSize(20, 20))
        
        # Set alignment and style
        self.setStyleSheet(
            "QPushButton {"
            f"    color: white;"
            f"    background-color: transparent;"
            f"    text-align: left;"
            f"    border: none;"
            f"    border-radius: 4px;"
            f"    padding-left: 10px;"
            f"    font-size: 12px;"
            "}"
            "QPushButton:hover {"
            f"    background-color: rgba(255, 255, 255, 0.1);"
            "}"
            "QPushButton:pressed {"
            f"    background-color: rgba(255, 255, 255, 0.2);"
            "}"
            "QPushButton[active=true] {"
            f"    background-color: {COLORS['primary']};"
            f"    color: white;"
            "}"
        )
        
        self.setProperty("active", False)
    
    def set_active(self, active):
        """Set the active state of the button."""
        self.setProperty("active", active)
        self.style().unpolish(self)
        self.style().polish(self)
        self.update()


class InfoCard(QFrame):
    """Information card with icon, title, and value."""
    
    def __init__(self, title, icon_name, value, description, parent=None):
        super().__init__(parent)
        self.setObjectName("info_card")
        
        # Set up card appearance
        self.setStyleSheet(
            "border-radius: 8px; "
            "background-color: white; "
            "border: 1px solid #E0E0E0;"
            "padding: 10px;"
        )
        
        # Set minimum size
        self.setMinimumSize(180, 100)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        
        # Create layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(5)
        
        # Header with icon and title
        header_layout = QHBoxLayout()
        header_layout.setSpacing(10)
        
        # Icon
        icon_label = QLabel()
        icon = QIcon(get_icon(icon_name))
        pixmap = icon.pixmap(QSize(20, 20))
        icon_label.setPixmap(pixmap)
        header_layout.addWidget(icon_label)
        
        # Title
        title_label = QLabel(title)
        title_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        
        layout.addLayout(header_layout)
        
        # Value
        self.value_label = QLabel(value)
        self.value_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        self.value_label.setAlignment(Qt.AlignLeft)
        layout.addWidget(self.value_label)
        
        # Description
        desc_label = QLabel(description)
        desc_label.setFont(QFont("Segoe UI", 9))
        desc_label.setStyleSheet(f"color: {COLORS['text']}; opacity: 0.7;")
        layout.addWidget(desc_label)
        
        layout.addStretch()
    
    def update_value(self, value, color=None):
        """Update the value displayed in the card."""
        self.value_label.setText(value)
        
        if color:
            self.value_label.setStyleSheet(f"color: {color};")
        else:
            self.value_label.setStyleSheet("")
