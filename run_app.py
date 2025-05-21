#!/usr/bin/env python3
"""
Launcher script for Windows System Optimizer
This script ensures the imports work correctly by setting up the Python path.
"""

import os
import sys
import logging
import traceback

# Determine if the application is running in a PyInstaller bundle
is_frozen = getattr(sys, 'frozen', False)

# Get the application base directory
if is_frozen:
    # PyInstaller creates a temporary folder and stores path in _MEIPASS
    base_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(sys.executable)))
else:
    # Normal execution
    base_dir = os.path.abspath(os.path.dirname(__file__))

# Add the base directory to the Python path
sys.path.insert(0, base_dir)

def setup_logging():
    """Setup application logging"""
    # Setup log directory
    log_dir = os.path.join(os.path.expanduser("~"), "WinOptimizer", "logs")
    os.makedirs(log_dir, exist_ok=True)
    
    # Log file path
    log_file = os.path.join(log_dir, "winoptimizer.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger("WinOptimizer")
    logger.info(f"Application base directory: {base_dir}")
    logger.info(f"Running in frozen mode: {is_frozen}")
    return logger

def handle_exception(exc_type, exc_value, exc_traceback):
    """Global exception handler to log unhandled exceptions"""
    logger = logging.getLogger("WinOptimizer")
    logger.error("Unhandled exception", exc_info=(exc_type, exc_value, exc_traceback))
    sys.__excepthook__(exc_type, exc_value, exc_traceback)

def main():
    """Main application entry point"""
    logger = setup_logging()
    logger.info("Starting Windows System Optimizer")
    
    # Set up global exception handler
    sys.excepthook = handle_exception
    
    try:
        # Import the necessary PyQt modules
        from PyQt5.QtWidgets import QApplication
        from PyQt5.QtCore import Qt
        from PyQt5.QtGui import QIcon
        
        # Enable high DPI scaling
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
        
        app = QApplication(sys.argv)
        app.setApplicationName("Windows System Optimizer")
        app.setOrganizationName("WinOptimizer")
        
        # Set application icon
        icon_path = os.path.join(base_dir, "assets", "app_icon.png")
        if os.path.exists(icon_path):
            app_icon = QIcon(icon_path)
            app.setWindowIcon(app_icon)
            logger.info(f"Set application icon from {icon_path}")
        else:
            logger.warning(f"Application icon not found at {icon_path}")
            # Fallback to looking in the root directory
            icon_path = os.path.join(base_dir, "app_icon.png")
            if os.path.exists(icon_path):
                app_icon = QIcon(icon_path)
                app.setWindowIcon(app_icon)
                logger.info(f"Set application icon from fallback path {icon_path}")
            else:
                logger.warning(f"Application icon not found at fallback path {icon_path}")
        
        # Import main window - ensure the module paths are correctly set
        try:
            from ui.main_window import MainWindow
            logger.info("Successfully imported ui.main_window module")
        except ImportError as e:
            logger.error(f"Failed to import ui.main_window: {str(e)}")
            # Try an alternative import method for PyInstaller
            if is_frozen:
                logger.info("Attempting alternative import method for PyInstaller environment")
                sys.path.append(os.path.join(base_dir, "ui"))
                sys.path.append(os.path.join(base_dir, "utils"))
                sys.path.append(os.path.join(base_dir, "services"))
                
                # Print available modules for debugging
                logger.info(f"Python path: {sys.path}")
                if os.path.exists(os.path.join(base_dir, "ui")):
                    logger.info(f"UI directory contents: {os.listdir(os.path.join(base_dir, 'ui'))}")
                
                from ui.main_window import MainWindow
            else:
                raise
        
        # Create and show the main window
        main_window = MainWindow()
        
        # Also set the icon for the main window
        if 'app_icon' in locals():
            main_window.setWindowIcon(app_icon)
            
        main_window.show()
        
        # Start the application
        logger.info("Application started successfully")
        return app.exec_()
    
    except ImportError as e:
        logger.error(f"Import error: {str(e)}")
        print(f"Error importing modules: {str(e)}")
        print("Check if all required packages are installed by running: pip install -r requirements.txt")
        traceback.print_exc()
        return 1
    
    except Exception as e:
        logger.error(f"Application error: {str(e)}")
        print(f"Error: {str(e)}")
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main()) 