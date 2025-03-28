"""
Theme management and styling for the application.
"""

# Define theme colors
COLORS = {
    'primary': '#007AFF',     # Modern blue
    'secondary': '#00B74A',   # Success green
    'background': '#F5F7FA',  # Light grey
    'text': '#2C3E50',        # Dark blue-grey
    'warning': '#FF9500',     # Orange
    'error': '#FF3B30',       # Red
    'card': '#FFFFFF',        # White for cards 
    'sidebar': '#394867',     # Darker blue for sidebar
    'highlight': '#E0E7FF',   # Light blue highlight
    'border': '#E0E0E0',      # Light grey for borders
}

# Dark mode colors
DARK_COLORS = {
    'primary': '#0A84FF',     # Modern blue (slightly adjusted for dark mode)
    'secondary': '#30D158',   # Success green (slightly adjusted for dark mode)
    'background': '#1C1C1E',  # Dark background
    'text': '#F5F7FA',        # Light text
    'warning': '#FF9F0A',     # Orange (slightly adjusted for dark mode)
    'error': '#FF453A',       # Red (slightly adjusted for dark mode)
    'card': '#2C2C2E',        # Dark card background
    'sidebar': '#252A34',     # Dark sidebar
    'highlight': '#2D3748',   # Dark blue highlight
    'border': '#3A3A3C',      # Dark border
}

# Theme identifiers
LIGHT_MODE = "light"
DARK_MODE = "dark"

# Light mode stylesheet
LIGHT_STYLESHEET = f"""
    QWidget {{
        font-family: "Segoe UI", "Roboto", sans-serif;
        font-size: 10pt;
        color: {COLORS['text']};
        background-color: {COLORS['background']};
    }}
    
    QMainWindow, QDialog {{
        background-color: {COLORS['background']};
    }}
    
    /* Sidebar styling */
    #sidebar {{
        background-color: {COLORS['sidebar']};
        color: white;
        border: none;
        padding: 0px;
    }}
    
    #app_title {{
        color: white;
        font-weight: bold;
        font-size: 16px;
    }}
    
    /* Content area styling */
    #content_container {{
        background-color: {COLORS['background']};
        border: none;
    }}
    
    /* Card styling */
    .Card {{
        background-color: {COLORS['card']};
        border-radius: 8px;
        border: 1px solid {COLORS['border']};
    }}
    
    /* Button styling */
    QPushButton {{
        background-color: {COLORS['primary']};
        color: white;
        border: none;
        border-radius: 4px;
        padding: 8px 16px;
        font-weight: bold;
    }}
    
    QPushButton:hover {{
        background-color: #0069D9;
    }}
    
    QPushButton:pressed {{
        background-color: #0051A8;
    }}
    
    QPushButton:disabled {{
        background-color: #CCCCCC;
        color: #666666;
    }}
    
    /* Success button */
    QPushButton[class="success"] {{
        background-color: {COLORS['secondary']};
    }}
    
    QPushButton[class="success"]:hover {{
        background-color: #00A23E;
    }}
    
    QPushButton[class="success"]:pressed {{
        background-color: #008732;
    }}
    
    /* Warning button */
    QPushButton[class="warning"] {{
        background-color: {COLORS['warning']};
    }}
    
    QPushButton[class="warning"]:hover {{
        background-color: #E68600;
    }}
    
    QPushButton[class="warning"]:pressed {{
        background-color: #CC7700;
    }}
    
    /* Error button */
    QPushButton[class="error"] {{
        background-color: {COLORS['error']};
    }}
    
    QPushButton[class="error"]:hover {{
        background-color: #E62E2E;
    }}
    
    QPushButton[class="error"]:pressed {{
        background-color: #CC2929;
    }}
    
    /* Text inputs */
    QLineEdit, QTextEdit {{
        background-color: white;
        color: {COLORS['text']};
        border: 1px solid {COLORS['border']};
        border-radius: 4px;
        padding: 5px;
    }}
    
    QLineEdit:focus, QTextEdit:focus {{
        border: 1px solid {COLORS['primary']};
    }}
    
    /* Headers */
    QLabel[class="header"] {{
        font-size: 20px;
        font-weight: bold;
        color: {COLORS['text']};
    }}
    
    QLabel[class="subheader"] {{
        font-size: 14px;
        color: {COLORS['text']};
        opacity: 0.8;
    }}
    
    /* Tabs */
    QTabWidget::pane {{
        border: 1px solid {COLORS['border']};
        border-radius: 4px;
        background-color: {COLORS['card']};
    }}
    
    QTabBar::tab {{
        background-color: {COLORS['background']};
        color: {COLORS['text']};
        border: 1px solid {COLORS['border']};
        border-bottom: none;
        border-top-left-radius: 4px;
        border-top-right-radius: 4px;
        padding: 8px 16px;
        margin-right: 2px;
    }}
    
    QTabBar::tab:selected {{
        background-color: {COLORS['card']};
        border-bottom: none;
    }}
    
    QTabBar::tab:!selected {{
        margin-top: 3px;
    }}
    
    /* Progress bar */
    QProgressBar {{
        border: 1px solid {COLORS['border']};
        border-radius: 4px;
        text-align: center;
    }}
    
    QProgressBar::chunk {{
        background-color: {COLORS['primary']};
    }}
    
    /* Combo box */
    QComboBox {{
        border: 1px solid {COLORS['border']};
        border-radius: 4px;
        padding: 5px;
        background-color: white;
    }}
    
    QComboBox::drop-down {{
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 20px;
        border-left: 1px solid {COLORS['border']};
    }}
    
    /* Table widget */
    QTableWidget {{
        background-color: {COLORS['card']};
        alternate-background-color: {COLORS['highlight']};
        border: 1px solid {COLORS['border']};
    }}
    
    QTableWidget::item:selected {{
        background-color: {COLORS['primary']};
        color: white;
    }}
    
    QHeaderView::section {{
        background-color: {COLORS['sidebar']};
        color: white;
        padding: 5px;
        border: 1px solid {COLORS['border']};
    }}
    
    /* Scroll area */
    QScrollArea {{
        border: none;
        background-color: transparent;
    }}
    
    /* Scroll bars */
    QScrollBar:vertical {{
        border: none;
        background: {COLORS['background']};
        width: 10px;
        margin: 0px;
    }}
    
    QScrollBar::handle:vertical {{
        background: {COLORS['border']};
        min-height: 20px;
        border-radius: 5px;
    }}
    
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
    }}
    
    QScrollBar:horizontal {{
        border: none;
        background: {COLORS['background']};
        height: 10px;
        margin: 0px;
    }}
    
    QScrollBar::handle:horizontal {{
        background: {COLORS['border']};
        min-width: 20px;
        border-radius: 5px;
    }}
    
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0px;
    }}
    
    /* Check boxes */
    QCheckBox::indicator {{
        width: 18px;
        height: 18px;
        border-radius: 3px;
        border: 1px solid {COLORS['border']};
    }}
    
    QCheckBox::indicator:checked {{
        background-color: {COLORS['primary']};
        border: 1px solid {COLORS['primary']};
    }}
"""

# Dark mode stylesheet
DARK_STYLESHEET = f"""
    QWidget {{
        font-family: "Segoe UI", "Roboto", sans-serif;
        font-size: 10pt;
        color: {DARK_COLORS['text']};
        background-color: {DARK_COLORS['background']};
    }}
    
    QMainWindow, QDialog {{
        background-color: {DARK_COLORS['background']};
    }}
    
    /* Sidebar styling */
    #sidebar {{
        background-color: {DARK_COLORS['sidebar']};
        color: white;
        border: none;
        padding: 0px;
    }}
    
    #app_title {{
        color: white;
        font-weight: bold;
        font-size: 16px;
    }}
    
    /* Content area styling */
    #content_container {{
        background-color: {DARK_COLORS['background']};
        border: none;
    }}
    
    /* Card styling */
    .Card {{
        background-color: {DARK_COLORS['card']};
        border-radius: 8px;
        border: 1px solid {DARK_COLORS['border']};
    }}
    
    /* Button styling */
    QPushButton {{
        background-color: {DARK_COLORS['primary']};
        color: white;
        border: none;
        border-radius: 4px;
        padding: 8px 16px;
        font-weight: bold;
    }}
    
    QPushButton:hover {{
        background-color: #2490FF;
    }}
    
    QPushButton:pressed {{
        background-color: #006DD6;
    }}
    
    QPushButton:disabled {{
        background-color: #555555;
        color: #999999;
    }}
    
    /* Success button */
    QPushButton[class="success"] {{
        background-color: {DARK_COLORS['secondary']};
    }}
    
    QPushButton[class="success"]:hover {{
        background-color: #40DC68;
    }}
    
    QPushButton[class="success"]:pressed {{
        background-color: #20C148;
    }}
    
    /* Warning button */
    QPushButton[class="warning"] {{
        background-color: {DARK_COLORS['warning']};
    }}
    
    QPushButton[class="warning"]:hover {{
        background-color: #FFAF2A;
    }}
    
    QPushButton[class="warning"]:pressed {{
        background-color: #E68F00;
    }}
    
    /* Error button */
    QPushButton[class="error"] {{
        background-color: {DARK_COLORS['error']};
    }}
    
    QPushButton[class="error"]:hover {{
        background-color: #FF655A;
    }}
    
    QPushButton[class="error"]:pressed {{
        background-color: #E6352A;
    }}
    
    /* Text inputs */
    QLineEdit, QTextEdit {{
        background-color: {DARK_COLORS['card']};
        color: {DARK_COLORS['text']};
        border: 1px solid {DARK_COLORS['border']};
        border-radius: 4px;
        padding: 5px;
    }}
    
    QLineEdit:focus, QTextEdit:focus {{
        border: 1px solid {DARK_COLORS['primary']};
    }}
    
    /* Headers */
    QLabel[class="header"] {{
        font-size: 20px;
        font-weight: bold;
        color: {DARK_COLORS['text']};
    }}
    
    QLabel[class="subheader"] {{
        font-size: 14px;
        color: {DARK_COLORS['text']};
        opacity: 0.8;
    }}
    
    /* Tabs */
    QTabWidget::pane {{
        border: 1px solid {DARK_COLORS['border']};
        border-radius: 4px;
        background-color: {DARK_COLORS['card']};
    }}
    
    QTabBar::tab {{
        background-color: {DARK_COLORS['background']};
        color: {DARK_COLORS['text']};
        border: 1px solid {DARK_COLORS['border']};
        border-bottom: none;
        border-top-left-radius: 4px;
        border-top-right-radius: 4px;
        padding: 8px 16px;
        margin-right: 2px;
    }}
    
    QTabBar::tab:selected {{
        background-color: {DARK_COLORS['card']};
        border-bottom: none;
    }}
    
    QTabBar::tab:!selected {{
        margin-top: 3px;
    }}
    
    /* Progress bar */
    QProgressBar {{
        border: 1px solid {DARK_COLORS['border']};
        border-radius: 4px;
        text-align: center;
    }}
    
    QProgressBar::chunk {{
        background-color: {DARK_COLORS['primary']};
    }}
    
    /* Combo box */
    QComboBox {{
        border: 1px solid {DARK_COLORS['border']};
        border-radius: 4px;
        padding: 5px;
        background-color: {DARK_COLORS['card']};
    }}
    
    QComboBox::drop-down {{
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 20px;
        border-left: 1px solid {DARK_COLORS['border']};
    }}
    
    /* Table widget */
    QTableWidget {{
        background-color: {DARK_COLORS['card']};
        alternate-background-color: {DARK_COLORS['highlight']};
        border: 1px solid {DARK_COLORS['border']};
    }}
    
    QTableWidget::item:selected {{
        background-color: {DARK_COLORS['primary']};
        color: white;
    }}
    
    QHeaderView::section {{
        background-color: {DARK_COLORS['sidebar']};
        color: white;
        padding: 5px;
        border: 1px solid {DARK_COLORS['border']};
    }}
    
    /* Scroll area */
    QScrollArea {{
        border: none;
        background-color: transparent;
    }}
    
    /* Scroll bars */
    QScrollBar:vertical {{
        border: none;
        background: {DARK_COLORS['background']};
        width: 10px;
        margin: 0px;
    }}
    
    QScrollBar::handle:vertical {{
        background: {DARK_COLORS['border']};
        min-height: 20px;
        border-radius: 5px;
    }}
    
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
    }}
    
    QScrollBar:horizontal {{
        border: none;
        background: {DARK_COLORS['background']};
        height: 10px;
        margin: 0px;
    }}
    
    QScrollBar::handle:horizontal {{
        background: {DARK_COLORS['border']};
        min-width: 20px;
        border-radius: 5px;
    }}
    
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0px;
    }}
    
    /* Check boxes */
    QCheckBox::indicator {{
        width: 18px;
        height: 18px;
        border-radius: 3px;
        border: 1px solid {DARK_COLORS['border']};
    }}
    
    QCheckBox::indicator:checked {{
        background-color: {DARK_COLORS['primary']};
        border: 1px solid {DARK_COLORS['primary']};
    }}
"""


def set_theme(app, theme_mode):
    """Set the application theme.
    
    Args:
        app: QApplication instance
        theme_mode: either LIGHT_MODE or DARK_MODE
    """
    global COLORS
    
    if theme_mode == DARK_MODE:
        app.setStyleSheet(DARK_STYLESHEET)
        COLORS = DARK_COLORS
    else:
        app.setStyleSheet(LIGHT_STYLESHEET)
        COLORS = COLORS
    
    return COLORS


def toggle_theme(app, current_theme):
    """Toggle between light and dark theme.
    
    Args:
        app: QApplication instance
        current_theme: current theme mode
    
    Returns:
        New theme mode
    """
    new_theme = LIGHT_MODE if current_theme == DARK_MODE else DARK_MODE
    set_theme(app, new_theme)
    return new_theme
