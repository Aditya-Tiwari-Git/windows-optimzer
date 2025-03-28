"""
Loading screen overlay for the Windows System Optimizer.
This module provides a semi-transparent overlay with a spinning loader
to indicate background operations in progress.
"""

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QProgressBar
from PyQt5.QtCore import Qt, QTimer, QSize
from PyQt5.QtGui import QColor, QPainter, QPainterPath, QFont

class LoadingOverlay(QWidget):
    """Semi-transparent overlay with spinning loader and message."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        
        # Set up the overlay
        self.setWindowFlags(Qt.Widget | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        
        # Create layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignCenter)
        
        # Message label
        self.message_label = QLabel("Processing...")
        self.message_label.setAlignment(Qt.AlignCenter)
        self.message_label.setFont(QFont("Segoe UI", 12))
        self.message_label.setStyleSheet("color: white; background-color: transparent;")
        
        # Add to layout with spacers to center
        layout.addStretch(1)
        layout.addWidget(self.message_label)
        layout.addStretch(1)
        
        # Animation properties
        self.angle = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.rotate)
        self.timer.start(30)  # Update every 30ms for smooth animation
        
        # Initial position and size
        self.resize(parent.size())
        self.show()
    
    def set_message(self, message):
        """Update the displayed message."""
        self.message_label.setText(message)
    
    def rotate(self):
        """Update the rotation angle for the spinner."""
        self.angle = (self.angle + 5) % 360
        self.update()
    
    def paintEvent(self, event):
        """Paint the overlay and spinner."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Semi-transparent background
        painter.fillRect(self.rect(), QColor(0, 0, 0, 180))
        
        # Draw the spinner
        painter.save()
        
        center = self.rect().center()
        painter.translate(center)
        painter.rotate(self.angle)
        
        # Spinner parameters
        radius = 30
        width = 5
        color = QColor("#007AFF")  # Primary blue
        
        # Create a path for the arc
        path = QPainterPath()
        path.addEllipse(-radius, -radius, radius * 2, radius * 2)
        
        # Draw the spinner
        pen = painter.pen()
        pen.setWidth(width)
        pen.setColor(color)
        painter.setPen(pen)
        painter.drawPath(path)
        
        painter.restore()
    
    def showEvent(self, event):
        """Handle show event to adjust size to parent."""
        self.resize(self.parent.size())
        super().showEvent(event)
    
    def resizeEvent(self, event):
        """Handle resize event from parent."""
        if self.parent:
            self.resize(self.parent.size())
