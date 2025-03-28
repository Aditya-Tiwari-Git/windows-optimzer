"""
Background worker for running operations in separate threads.
"""

import os
import sys
import time
import logging
import traceback
from PyQt5.QtCore import QObject, pyqtSignal, QThread, QTimer

logger = logging.getLogger(__name__)


class WorkerSignals(QObject):
    """Signals for the background worker."""
    
    taskFinished = pyqtSignal(object)
    taskFailed = pyqtSignal(object)
    taskUpdate = pyqtSignal(object)


class BackgroundWorker(QThread):
    """Background worker for running operations in a separate thread.
    
    This worker executes a function in a separate thread to prevent freezing the UI.
    It provides signals to report progress and completion.
    """
    
    def __init__(self, fn, *args, **kwargs):
        """Initialize the background worker.
        
        Args:
            fn: Function to execute in the background
            *args: Arguments to pass to the function
            **kwargs: Keyword arguments to pass to the function
        """
        super().__init__()
        
        # Store the function and arguments
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        
        # Set up signals
        self.signals = WorkerSignals()
        
        # Connect signals to our methods
        self.taskFinished = self.signals.taskFinished
        self.taskFailed = self.signals.taskFailed
        self.taskUpdate = self.signals.taskUpdate
    
    def run(self):
        """Run the worker thread."""
        # Execute the function and return the result
        try:
            # Special handling for generator functions that yield progress updates
            result = self.fn(*self.args, **self.kwargs)
            
            # Check if result is a generator (has yields)
            if hasattr(result, '__iter__') and not isinstance(result, (dict, list, tuple, str, bytes)):
                try:
                    # Process generator results
                    last_result = None
                    for update in result:
                        # Send progress update
                        self.signals.taskUpdate.emit(update)
                        last_result = update
                    
                    # Use last yielded value as the result
                    self.signals.taskFinished.emit(last_result)
                except Exception as e:
                    # Generator function raised an exception
                    logger.error(f"Error in generator worker: {str(e)}")
                    self.signals.taskFailed.emit(str(e))
            else:
                # Regular function, just return the result
                self.signals.taskFinished.emit(result)
        
        except Exception as e:
            # Log the full exception traceback
            logger.error(f"Error in background worker: {str(e)}")
            logger.error(traceback.format_exc())
            
            # Emit the error
            self.signals.taskFailed.emit(str(e))
    
    def __del__(self):
        """Clean up the worker."""
        # Make sure the thread is properly closed
        self.wait()


class TimedTaskRunner(QObject):
    """Runner for scheduled or repeating tasks."""
    
    taskCompleted = pyqtSignal(object)
    
    def __init__(self, interval=60000, parent=None):
        """Initialize the timed task runner.
        
        Args:
            interval: Interval in milliseconds between task executions
            parent: Parent object
        """
        super().__init__(parent)
        
        # Set up the timer
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.run_task)
        self.interval = interval
        
        # Store task function and arguments
        self.task_fn = None
        self.task_args = []
        self.task_kwargs = {}
        
        # Worker thread
        self.worker = None
    
    def set_task(self, fn, *args, **kwargs):
        """Set the task to run.
        
        Args:
            fn: Function to execute
            *args: Arguments to pass to the function
            **kwargs: Keyword arguments to pass to the function
        """
        self.task_fn = fn
        self.task_args = args
        self.task_kwargs = kwargs
    
    def start(self, run_immediately=False):
        """Start the timed task runner.
        
        Args:
            run_immediately: Whether to run the task immediately or wait for the first interval
        """
        if self.task_fn:
            if run_immediately:
                self.run_task()
            
            self.timer.start(self.interval)
    
    def stop(self):
        """Stop the timed task runner."""
        self.timer.stop()
        
        # Stop worker if running
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
    
    def set_interval(self, interval):
        """Set the interval between task executions.
        
        Args:
            interval: Interval in milliseconds
        """
        self.interval = interval
        
        # Restart timer if it's running
        if self.timer.isActive():
            self.timer.stop()
            self.timer.start(self.interval)
    
    def run_task(self):
        """Run the task in a background thread."""
        if not self.task_fn:
            return
        
        # Don't start a new task if the previous one is still running
        if self.worker and self.worker.isRunning():
            return
        
        # Create a new worker thread
        self.worker = BackgroundWorker(self.task_fn, *self.task_args, **self.task_kwargs)
        self.worker.taskFinished.connect(self.on_task_complete)
        self.worker.taskFailed.connect(self.on_task_failed)
        self.worker.start()
    
    def on_task_complete(self, result):
        """Handle task completion.
        
        Args:
            result: Result from the task
        """
        self.taskCompleted.emit(result)
    
    def on_task_failed(self, error):
        """Handle task failure.
        
        Args:
            error: Error message from the task
        """
        logger.error(f"Timed task failed: {error}")
