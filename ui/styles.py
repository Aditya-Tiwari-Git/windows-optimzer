"""
Styling module for the Windows System Optimizer.
This module provides consistent styling and theming for the application.
"""

from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QPalette, QColor, QFont

# Color scheme
COLORS = {
    "primary": "#007AFF",       # Modern blue
    "secondary": "#00B74A",     # Success green
    "background": "#F5F7FA",    # Light grey
    "text": "#2C3E50",          # Dark blue-grey
    "warning": "#FF9500",       # Orange
    "error": "#FF3B30",         # Red
    
    # Dark mode variants
    "dark_primary": "#0A84FF",
    "dark_secondary": "#30D158",
    "dark_background": "#1C1C1E",
    "dark_card": "#2C2C2E",
    "dark_text": "#FFFFFF",
}

def set_light_mode(window):
    """Apply light mode styling to the application."""
    app = QApplication.instance()
    
    # Set light palette
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(COLORS["background"]))
    palette.setColor(QPalette.WindowText, QColor(COLORS["text"]))
    palette.setColor(QPalette.Base, QColor("#FFFFFF"))
    palette.setColor(QPalette.AlternateBase, QColor(COLORS["background"]))
    palette.setColor(QPalette.ToolTipBase, QColor("#FFFFFF"))
    palette.setColor(QPalette.ToolTipText, QColor(COLORS["text"]))
    palette.setColor(QPalette.Text, QColor(COLORS["text"]))
    palette.setColor(QPalette.Button, QColor(COLORS["background"]))
    palette.setColor(QPalette.ButtonText, QColor(COLORS["text"]))
    palette.setColor(QPalette.Link, QColor(COLORS["primary"]))
    palette.setColor(QPalette.Highlight, QColor(COLORS["primary"]))
    palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    
    app.setPalette(palette)
    
    # Set light stylesheet
    app.setStyleSheet("""
        QMainWindow, QWidget {
            font-family: 'Segoe UI', 'Roboto', sans-serif;
            font-size: 10pt;
        }
        
        QPushButton {
            background-color: #FFFFFF;
            border: 1px solid #CCCCCC;
            border-radius: 4px;
            padding: 6px 12px;
            color: #2C3E50;
        }
        
        QPushButton:hover {
            background-color: #F0F0F0;
        }
        
        QPushButton:pressed {
            background-color: #E0E0E0;
        }
        
        QPushButton:checked {
            background-color: #007AFF;
            color: white;
        }
        
        QProgressBar {
            border: 1px solid #CCCCCC;
            border-radius: 4px;
            text-align: center;
            height: 20px;
        }
        
        QProgressBar::chunk {
            background-color: #007AFF;
            width: 20px;
        }
        
        QTabWidget::pane {
            border: 1px solid #CCCCCC;
            border-top: 0px;
        }
        
        QTabBar::tab {
            background-color: #F5F7FA;
            border: 1px solid #CCCCCC;
            border-bottom: none;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
            padding: 8px 12px;
            margin-right: 2px;
        }
        
        QTabBar::tab:selected {
            background-color: #FFFFFF;
            border-bottom: none;
        }
        
        QTabBar::tab:!selected {
            margin-top: 2px;
        }
        
        QGroupBox {
            font-weight: bold;
            border: 1px solid #CCCCCC;
            border-radius: 4px;
            margin-top: 1.5ex;
            padding: 10px;
        }
        
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 0 5px;
        }
        
        QFrame#metricCard {
            background-color: white;
            border: 1px solid #CCCCCC;
            border-radius: 4px;
            padding: 10px;
        }
    """)


def set_dark_mode(window):
    """Apply dark mode styling to the application."""
    app = QApplication.instance()
    
    # Set dark palette
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(COLORS["dark_background"]))
    palette.setColor(QPalette.WindowText, QColor(COLORS["dark_text"]))
    palette.setColor(QPalette.Base, QColor(COLORS["dark_card"]))
    palette.setColor(QPalette.AlternateBase, QColor("#3A3A3C"))
    palette.setColor(QPalette.ToolTipBase, QColor(COLORS["dark_card"]))
    palette.setColor(QPalette.ToolTipText, QColor(COLORS["dark_text"]))
    palette.setColor(QPalette.Text, QColor(COLORS["dark_text"]))
    palette.setColor(QPalette.Button, QColor(COLORS["dark_card"]))
    palette.setColor(QPalette.ButtonText, QColor(COLORS["dark_text"]))
    palette.setColor(QPalette.Link, QColor(COLORS["dark_primary"]))
    palette.setColor(QPalette.Highlight, QColor(COLORS["dark_primary"]))
    palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    
    app.setPalette(palette)
    
    # Set dark stylesheet
    app.setStyleSheet("""
        QMainWindow, QWidget {
            font-family: 'Segoe UI', 'Roboto', sans-serif;
            font-size: 10pt;
            background-color: #1C1C1E;
            color: #FFFFFF;
        }
        
        QPushButton {
            background-color: #3A3A3C;
            border: 1px solid #555555;
            border-radius: 4px;
            padding: 6px 12px;
            color: #FFFFFF;
        }
        
        QPushButton:hover {
            background-color: #48484A;
        }
        
        QPushButton:pressed {
            background-color: #545456;
        }
        
        QPushButton:checked {
            background-color: #0A84FF;
            color: white;
        }
        
        QProgressBar {
            border: 1px solid #555555;
            border-radius: 4px;
            text-align: center;
            color: #FFFFFF;
            height: 20px;
        }
        
        QProgressBar::chunk {
            background-color: #0A84FF;
            width: 20px;
        }
        
        QTabWidget::pane {
            border: 1px solid #555555;
            border-top: 0px;
        }
        
        QTabBar::tab {
            background-color: #2C2C2E;
            border: 1px solid #555555;
            border-bottom: none;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
            padding: 8px 12px;
            margin-right: 2px;
        }
        
        QTabBar::tab:selected {
            background-color: #3A3A3C;
            border-bottom: none;
        }
        
        QTabBar::tab:!selected {
            margin-top: 2px;
        }
        
        QGroupBox {
            font-weight: bold;
            border: 1px solid #555555;
            border-radius: 4px;
            margin-top: 1.5ex;
            padding: 10px;
        }
        
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 0 5px;
        }
        
        QLineEdit, QTextEdit, QComboBox, QSpinBox {
            background-color: #2C2C2E;
            color: #FFFFFF;
            border: 1px solid #555555;
            border-radius: 4px;
            padding: 4px;
        }
        
        QFrame#metricCard {
            background-color: #2C2C2E;
            border: 1px solid #555555;
            border-radius: 4px;
            padding: 10px;
        }
    """)
