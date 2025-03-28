"""
System Optimization widget for system performance tuning and monitoring.
"""

import os
import time
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QFrame, QProgressBar, QTabWidget,
                             QCheckBox, QListWidget, QListWidgetItem, QFileDialog,
                             QSpacerItem, QSizePolicy, QScrollArea, QGroupBox,
                             QSlider, QComboBox, QGridLayout)
from PyQt5.QtGui import QIcon, QFont, QPalette, QColor
from PyQt5.QtCore import Qt, pyqtSignal, pyqtSlot, QThread, QTimer

from ui.theme import COLORS
from ui.components import Card, HeaderLabel, SubHeaderLabel, ActionButton, InfoCard
from ui.loading_screen import LoadingScreen
from utils.optimizer import SystemOptimizer
from utils.background_worker import BackgroundWorker


class OptimizationListItem(QFrame):
    """Custom list item widget for optimization options."""
    
    toggled = pyqtSignal(str, bool)
    
    def __init__(self, title, description, icon_name, option_id, default_state=True, parent=None):
        super().__init__(parent)
        self.title = title
        self.description = description
        self.option_id = option_id
        
        self.setMinimumHeight(60)
        self.setObjectName("optimization_list_item")
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 5, 10, 5)
        
        # Checkbox
        self.checkbox = QCheckBox()
        self.checkbox.setChecked(default_state)
        self.checkbox.toggled.connect(self.on_toggle)
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
    
    def on_toggle(self, checked):
        """Emit signal when checkbox is toggled."""
        self.toggled.emit(self.option_id, checked)
    
    def is_checked(self):
        """Return checkbox state."""
        return self.checkbox.isChecked()


class SliderOption(QFrame):
    """Custom slider option widget for performance settings."""
    
    valueChanged = pyqtSignal(str, int)
    
    def __init__(self, title, description, option_id, min_val, max_val, default_val, parent=None):
        super().__init__(parent)
        self.title = title
        self.description = description
        self.option_id = option_id
        
        self.setMinimumHeight(80)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 5, 10, 5)
        
        # Title and description
        title_label = QLabel(title)
        title_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        
        desc_label = QLabel(description)
        desc_label.setFont(QFont("Segoe UI", 8))
        desc_label.setStyleSheet(f"color: {COLORS['text']}; opacity: 0.8;")
        
        layout.addWidget(title_label)
        layout.addWidget(desc_label)
        
        # Slider and value
        slider_layout = QHBoxLayout()
        
        self.slider = QSlider(Qt.Horizontal)
        self.slider.setMinimum(min_val)
        self.slider.setMaximum(max_val)
        self.slider.setValue(default_val)
        self.slider.setTickPosition(QSlider.TicksBelow)
        self.slider.setTickInterval((max_val - min_val) // 5)
        self.slider.valueChanged.connect(self.on_value_changed)
        
        self.value_label = QLabel(f"{default_val}")
        self.value_label.setMinimumWidth(30)
        self.value_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        
        slider_layout.addWidget(self.slider, 1)
        slider_layout.addWidget(self.value_label)
        
        layout.addLayout(slider_layout)
    
    def on_value_changed(self, value):
        """Update label and emit signal when slider value changes."""
        self.value_label.setText(f"{value}")
        self.valueChanged.emit(self.option_id, value)
    
    def get_value(self):
        """Return current slider value."""
        return self.slider.value()


class OptimizationWidget(QWidget):
    """Widget for system optimization features."""
    
    def __init__(self):
        super().__init__()
        self.optimizer = SystemOptimizer()
        self.initUI()
        
        # Loading screen instance
        self.loading_screen = None
        
        # Timer for auto-refresh of running apps
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.update_running_apps)
        self.refresh_timer.start(5000)  # Refresh every 5 seconds
    
    def initUI(self):
        """Initialize the optimization UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)
        
        # Header
        header = HeaderLabel("System Optimization")
        subheader = SubHeaderLabel("Improve system performance and resource usage")
        
        layout.addWidget(header)
        layout.addWidget(subheader)
        
        # Tabs for different optimization tools
        tabs = QTabWidget()
        tabs.setObjectName("optimization_tabs")
        
        # Performance Optimization Tab
        perf_tab = QWidget()
        perf_layout = QVBoxLayout(perf_tab)
        perf_layout.setContentsMargins(15, 15, 15, 15)
        
        # Status cards
        status_layout = QHBoxLayout()
        
        self.cpu_usage_card = InfoCard(
            title="CPU Usage",
            icon_name="cpu",
            value="0%",
            description="Current utilization"
        )
        
        self.memory_usage_card = InfoCard(
            title="Memory Usage",
            icon_name="memory",
            value="0 / 0 GB",
            description="RAM utilization"
        )
        
        self.startup_time_card = InfoCard(
            title="Startup Time",
            icon_name="time",
            value="--",
            description="System boot time"
        )
        
        status_layout.addWidget(self.cpu_usage_card)
        status_layout.addWidget(self.memory_usage_card)
        status_layout.addWidget(self.startup_time_card)
        
        perf_layout.addLayout(status_layout)
        perf_layout.addSpacing(15)
        
        # Performance settings
        settings_card = Card()
        settings_layout = QVBoxLayout(settings_card)
        
        settings_header = QLabel("Performance Settings")
        settings_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        settings_layout.addWidget(settings_header)
        
        # Optimize options
        options_scroll = QScrollArea()
        options_scroll.setWidgetResizable(True)
        options_scroll.setFrameShape(QFrame.NoFrame)
        
        options_content = QWidget()
        options_layout = QVBoxLayout(options_content)
        options_layout.setContentsMargins(5, 5, 5, 5)
        options_layout.setSpacing(5)
        
        self.optimization_items = []
        
        # Create a list of optimization options
        visual_effects = OptimizationListItem(
            "Disable Visual Effects", 
            "Reduces animations and visual effects for better performance",
            "visual",
            "visual_effects",
            True
        )
        
        background_apps = OptimizationListItem(
            "Limit Background Apps", 
            "Restricts apps running in the background",
            "apps",
            "background_apps",
            True
        )
        
        startup_items = OptimizationListItem(
            "Optimize Startup Programs", 
            "Disable unnecessary startup items",
            "startup",
            "startup_items",
            True
        )
        
        system_services = OptimizationListItem(
            "Optimize System Services", 
            "Tune system services for better performance",
            "services",
            "system_services",
            True
        )
        
        power_settings = OptimizationListItem(
            "Performance Power Settings", 
            "Configure power settings for maximum performance",
            "power",
            "power_settings",
            True
        )
        
        disk_optimization = OptimizationListItem(
            "Disk Performance", 
            "Optimize disk caching and file system",
            "disk",
            "disk_optimization",
            True
        )
        
        network_optimization = OptimizationListItem(
            "Network Optimization", 
            "Tune TCP/IP parameters for better network performance",
            "network",
            "network_optimization",
            True
        )
        
        self.optimization_items = [
            visual_effects,
            background_apps,
            startup_items,
            system_services,
            power_settings,
            disk_optimization,
            network_optimization
        ]
        
        for item in self.optimization_items:
            item.toggled.connect(self.on_option_toggled)
            options_layout.addWidget(item)
        
        options_layout.addStretch()
        options_scroll.setWidget(options_content)
        settings_layout.addWidget(options_scroll, 1)
        
        # Performance mode
        mode_group = QGroupBox("Performance Mode")
        mode_layout = QVBoxLayout(mode_group)
        
        self.perf_mode_combo = QComboBox()
        self.perf_mode_combo.addItems(["Balanced", "Performance", "Maximum Performance", "Power Saver"])
        self.perf_mode_combo.setCurrentIndex(1)  # Default to Performance
        self.perf_mode_combo.currentIndexChanged.connect(self.on_mode_changed)
        
        mode_description = QLabel(
            "Performance mode adjusts multiple system settings at once to optimize "
            "for your selected preference between speed and energy efficiency."
        )
        mode_description.setWordWrap(True)
        mode_description.setStyleSheet(f"color: {COLORS['text']}; opacity: 0.8;")
        
        mode_layout.addWidget(self.perf_mode_combo)
        mode_layout.addWidget(mode_description)
        
        settings_layout.addWidget(mode_group)
        
        perf_layout.addWidget(settings_card, 1)
        perf_layout.addSpacing(10)
        
        # Action buttons
        perf_buttons_layout = QHBoxLayout()
        
        perf_analyze_btn = ActionButton("Analyze Performance", "search", "primary")
        perf_analyze_btn.clicked.connect(self.analyze_performance)
        
        perf_optimize_btn = ActionButton("Apply Optimizations", "bolt", "success")
        perf_optimize_btn.clicked.connect(self.apply_optimizations)
        
        perf_reset_btn = ActionButton("Reset to Default", "reset", "warning")
        perf_reset_btn.clicked.connect(self.reset_optimization_settings)
        
        perf_buttons_layout.addStretch()
        perf_buttons_layout.addWidget(perf_analyze_btn)
        perf_buttons_layout.addWidget(perf_optimize_btn)
        perf_buttons_layout.addWidget(perf_reset_btn)
        
        perf_layout.addLayout(perf_buttons_layout)
        
        # Process Management Tab
        process_tab = QWidget()
        process_layout = QVBoxLayout(process_tab)
        process_layout.setContentsMargins(15, 15, 15, 15)
        
        # Running apps
        apps_card = Card()
        apps_layout = QVBoxLayout(apps_card)
        
        apps_header = QLabel("Running Applications")
        apps_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        apps_layout.addWidget(apps_header)
        
        # Running apps table
        self.running_apps_list = QListWidget()
        self.running_apps_list.setAlternatingRowColors(True)
        self.running_apps_list.setSelectionMode(QListWidget.ExtendedSelection)
        apps_layout.addWidget(self.running_apps_list, 1)
        
        # App management buttons
        apps_buttons_layout = QHBoxLayout()
        
        refresh_apps_btn = ActionButton("Refresh", "refresh", "primary")
        refresh_apps_btn.clicked.connect(self.update_running_apps)
        
        end_task_btn = ActionButton("End Selected", "stop", "error")
        end_task_btn.clicked.connect(self.end_selected_tasks)
        
        apps_buttons_layout.addWidget(refresh_apps_btn)
        apps_buttons_layout.addStretch()
        apps_buttons_layout.addWidget(end_task_btn)
        
        apps_layout.addLayout(apps_buttons_layout)
        
        process_layout.addWidget(apps_card, 1)
        
        # Memory optimization
        memory_group = QGroupBox("Memory Optimization")
        memory_layout = QVBoxLayout(memory_group)
        
        memory_description = QLabel(
            "Free up unused memory to improve system performance. "
            "This can help when the system becomes sluggish after extended use."
        )
        memory_description.setWordWrap(True)
        
        self.memory_usage_progress = QProgressBar()
        self.memory_usage_progress.setRange(0, 100)
        self.memory_usage_progress.setValue(50)  # Placeholder
        
        optimize_memory_btn = ActionButton("Optimize Memory", "memory", "primary")
        optimize_memory_btn.clicked.connect(self.optimize_memory)
        
        memory_layout.addWidget(memory_description)
        memory_layout.addWidget(self.memory_usage_progress)
        memory_layout.addWidget(optimize_memory_btn, 0, Qt.AlignRight)
        
        process_layout.addWidget(memory_group)
        
        # Add tabs
        tabs.addTab(perf_tab, "Performance Optimization")
        tabs.addTab(process_tab, "Process Management")
        
        layout.addWidget(tabs, 1)
        
        # Update the status cards and running apps list
        self.update_system_stats()
        self.update_running_apps()
    
    def update_system_stats(self):
        """Update system statistics displayed in the cards."""
        # Get system information
        cpu_percent = self.optimizer.get_cpu_usage()
        memory_info = self.optimizer.get_memory_info()
        startup_time = self.optimizer.get_startup_time()
        
        # Update UI elements
        self.cpu_usage_card.update_value(f"{cpu_percent}%")
        
        self.memory_usage_card.update_value(
            f"{memory_info['used_gb']:.1f} / {memory_info['total_gb']:.1f} GB"
        )
        
        self.startup_time_card.update_value(startup_time)
        
        # Update memory progress bar
        self.memory_usage_progress.setValue(int(memory_info['percent']))
        
        # Set color based on usage
        if memory_info['percent'] < 60:
            self.memory_usage_progress.setStyleSheet(
                f"QProgressBar::chunk {{ background-color: {COLORS['secondary']}; }}"
            )
        elif memory_info['percent'] < 80:
            self.memory_usage_progress.setStyleSheet(
                f"QProgressBar::chunk {{ background-color: {COLORS['warning']}; }}"
            )
        else:
            self.memory_usage_progress.setStyleSheet(
                f"QProgressBar::chunk {{ background-color: {COLORS['error']}; }}"
            )
    
    def update_running_apps(self):
        """Update the list of running applications."""
        running_apps = self.optimizer.get_running_apps()
        
        # Store current selection
        selected_items = [self.running_apps_list.item(i).text().split(' (')[0] 
                          for i in range(self.running_apps_list.count()) 
                          if self.running_apps_list.item(i).isSelected()]
        
        # Clear and repopulate list
        self.running_apps_list.clear()
        
        for app in running_apps:
            memory_usage = app['memory_mb']
            list_text = f"{app['name']} (PID: {app['pid']}, Memory: {memory_usage:.1f} MB)"
            item = QListWidgetItem(list_text)
            
            # Highlight high memory usage
            if memory_usage > 500:
                item.setForeground(QColor(COLORS['warning']))
            
            self.running_apps_list.addItem(item)
            
            # Restore selection
            if app['name'] in selected_items:
                item.setSelected(True)
        
        # Also update system stats while we're at it
        self.update_system_stats()
    
    def on_option_toggled(self, option_id, checked):
        """Handle when an optimization option is toggled."""
        # This would typically update a configuration or setting
        # For now, just print to verify the callback works
        print(f"Option '{option_id}' set to: {checked}")
    
    def on_mode_changed(self, index):
        """Handle when performance mode is changed."""
        modes = ["Balanced", "Performance", "Maximum Performance", "Power Saver"]
        selected_mode = modes[index]
        
        # Update optimization options based on selected mode
        if selected_mode == "Balanced":
            # Default settings
            self.update_optimization_preset({
                "visual_effects": True,
                "background_apps": True,
                "startup_items": True,
                "system_services": True,
                "power_settings": False,
                "disk_optimization": True,
                "network_optimization": True
            })
        elif selected_mode == "Performance":
            # Performance-focused settings
            self.update_optimization_preset({
                "visual_effects": False,
                "background_apps": False,
                "startup_items": False,
                "system_services": True,
                "power_settings": True,
                "disk_optimization": True,
                "network_optimization": True
            })
        elif selected_mode == "Maximum Performance":
            # Maximum performance, all optimizations on
            self.update_optimization_preset({
                "visual_effects": False,
                "background_apps": False,
                "startup_items": False,
                "system_services": False,
                "power_settings": True,
                "disk_optimization": True,
                "network_optimization": True
            })
        elif selected_mode == "Power Saver":
            # Power-saving settings
            self.update_optimization_preset({
                "visual_effects": True,
                "background_apps": True,
                "startup_items": True,
                "system_services": True,
                "power_settings": False,
                "disk_optimization": False,
                "network_optimization": False
            })
    
    def update_optimization_preset(self, settings):
        """Update checkboxes based on a preset configuration."""
        for item in self.optimization_items:
            if item.option_id in settings:
                item.checkbox.setChecked(settings[item.option_id])
    
    def analyze_performance(self):
        """Analyze system performance."""
        self.show_loading("Analyzing System Performance...", "This may take a moment.")
        
        # Create background worker
        worker = BackgroundWorker(self.optimizer.analyze_performance)
        worker.taskFinished.connect(self.on_analysis_complete)
        worker.start()
    
    def on_analysis_complete(self, results):
        """Handle completion of performance analysis."""
        self.hide_loading()
        
        # Update the UI with analysis results
        self.update_system_stats()
        
        # Update optimization settings based on analysis recommendations
        if 'recommendations' in results:
            self.update_optimization_preset(results['recommendations'])
        
        # Select best performance mode based on analysis
        if 'recommended_mode' in results:
            modes = ["Balanced", "Performance", "Maximum Performance", "Power Saver"]
            if results['recommended_mode'] in modes:
                self.perf_mode_combo.setCurrentText(results['recommended_mode'])
    
    def apply_optimizations(self):
        """Apply selected optimization settings."""
        self.show_loading("Applying System Optimizations...", "This may take a while.")
        
        # Gather all selected optimization options
        optimizations = {}
        for item in self.optimization_items:
            optimizations[item.option_id] = item.is_checked()
        
        # Add performance mode
        optimizations['performance_mode'] = self.perf_mode_combo.currentText()
        
        # Create background worker
        worker = BackgroundWorker(self.optimizer.apply_optimizations, optimizations)
        worker.taskFinished.connect(self.on_optimization_complete)
        worker.start()
    
    def on_optimization_complete(self, results):
        """Handle completion of optimization application."""
        self.hide_loading()
        
        # Update the UI with optimization results
        self.update_system_stats()
        
        # If successful, update running apps list
        self.update_running_apps()
    
    def reset_optimization_settings(self):
        """Reset optimization settings to defaults."""
        # Set to balanced mode
        self.perf_mode_combo.setCurrentText("Balanced")
        
        # Reset all checkboxes to default
        default_settings = {
            "visual_effects": True,
            "background_apps": True,
            "startup_items": True,
            "system_services": True,
            "power_settings": False,
            "disk_optimization": True,
            "network_optimization": True
        }
        
        self.update_optimization_preset(default_settings)
    
    def optimize_memory(self):
        """Perform memory optimization."""
        self.show_loading("Optimizing System Memory...", "Freeing up unused memory...")
        
        # Create background worker
        worker = BackgroundWorker(self.optimizer.optimize_memory)
        worker.taskFinished.connect(self.on_memory_optimization_complete)
        worker.start()
    
    def on_memory_optimization_complete(self, results):
        """Handle completion of memory optimization."""
        self.hide_loading()
        
        # Update the UI with memory optimization results
        self.update_system_stats()
        self.update_running_apps()
    
    def end_selected_tasks(self):
        """End selected tasks in the running apps list."""
        selected_items = self.running_apps_list.selectedItems()
        if not selected_items:
            return
        
        # Extract PIDs from selected items
        pids = []
        for item in selected_items:
            # Format is: "App Name (PID: 1234, Memory: 123.4 MB)"
            try:
                pid_part = item.text().split("PID: ")[1].split(",")[0]
                pids.append(int(pid_part))
            except:
                continue
        
        if not pids:
            return
        
        self.show_loading("Ending Selected Processes...", "This may take a moment.")
        
        # Create background worker
        worker = BackgroundWorker(self.optimizer.end_processes, pids)
        worker.taskFinished.connect(self.on_end_tasks_complete)
        worker.start()
    
    def on_end_tasks_complete(self, results):
        """Handle completion of task termination."""
        self.hide_loading()
        
        # Update the running apps list
        self.update_running_apps()
    
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
    
    def showEvent(self, event):
        """When widget becomes visible, start the update timer."""
        super().showEvent(event)
        self.refresh_timer.start(5000)
        self.update_system_stats()
        self.update_running_apps()
    
    def hideEvent(self, event):
        """When widget is hidden, stop the update timer."""
        super().hideEvent(event)
        self.refresh_timer.stop()


# Import at the bottom to avoid circular import
from assets.icons import get_icon
