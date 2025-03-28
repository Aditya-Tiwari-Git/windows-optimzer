#!/usr/bin/env python3
"""
Windows System Optimizer - Main Application Entry Point
A modern Windows system optimization utility with cleaning, network diagnostics, 
and performance monitoring features.
"""

import sys
import os
import logging
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt
from ui.main_window import MainWindow

# Set up logging
def setup_logging():
    log_dir = os.path.join(os.path.expanduser("~"), "WinOptimizer", "logs")
    os.makedirs(log_dir, exist_ok=True)
    
    log_file = os.path.join(log_dir, "winoptimizer.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger("WinOptimizer")

if __name__ == "__main__":
    # Set up application
    logger = setup_logging()
    logger.info("Starting Windows System Optimizer")
    
    # Enable high DPI scaling
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    app.setApplicationName("Windows System Optimizer")
    app.setOrganizationName("WinOptimizer")
    
    # Create and show the main window
    main_window = MainWindow()
    main_window.show()
    
    # Start the application
    try:
        sys.exit(app.exec_())
    except Exception as e:
        logger.error(f"Application crashed: {str(e)}")
        sys.exit(1)
