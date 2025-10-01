#!/usr/bin/env python3
"""
Word Appendix Manager - Main Entry Point
Author: Senior Python Developer
Version: 1.0.0

Professional desktop application for adding PDF appendices to Word documents.
"""

import sys
import os
from pathlib import Path
from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtCore import Qt, QDir
from PySide6.QtGui import QIcon

# Add src directory to Python path for imports
src_dir = Path(__file__).parent
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from utils.logger import setup_logging
from utils.exceptions import ApplicationError
from utils.theme_manager import ThemeManager
from config.settings import AppSettings
from gui.main_window import MainWindow
from gui.controller import AppController


class WordAppendixManager:
    """Main application class for Word Appendix Manager."""
    
    def __init__(self):
        self.app = None
        self.main_window = None
        self.controller = None
        self.settings = None
        self.theme_manager = None
        
    def initialize_application(self):
        """Initialize the Qt application with proper settings."""
        # Set DPI scaling policy BEFORE creating QApplication
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )

        # Create QApplication
        self.app = QApplication(sys.argv)
        self.app.setApplicationName("Word Appendix Manager")
        self.app.setApplicationVersion("1.0.0")
        self.app.setOrganizationName("PointOne")
    
        # Set application icon
        icon_path = self._get_resource_path("icons/app_icon.ico")
        if icon_path and os.path.exists(icon_path):
            self.app.setWindowIcon(QIcon(icon_path))
    
        return self.app

    
    def initialize_components(self):
        """Initialize application components."""
        try:
            # Setup logging
            setup_logging()
            
            # Load application settings
            self.settings = AppSettings()
            
            # Initialize theme manager
            self.theme_manager = ThemeManager(self.settings)
            
            # Apply saved theme or default
            self.theme_manager.load_saved_theme()
            
        except Exception as e:
            self._show_error(f"Failed to initialize application components: {str(e)}")
            return False
        
        return True
    
    def create_main_window(self):
        """Create and configure the main application window."""
        try:
            self.main_window = MainWindow(settings=self.settings, theme_manager=self.theme_manager)
            
            # Create controller to handle business logic
            self.controller = AppController(self.main_window, self.settings)
            
            self.main_window.show()
            
            # Center window on screen
            self._center_window()
            
        except Exception as e:
            self._show_error(f"Failed to create main window: {str(e)}")
            return False
            
        return True
    
    def _center_window(self):
        """Center the main window on the screen."""
        if self.main_window:
            screen = self.app.primaryScreen()
            screen_geometry = screen.availableGeometry()
            window_geometry = self.main_window.frameGeometry()
            
            center_point = screen_geometry.center()
            window_geometry.moveCenter(center_point)
            self.main_window.move(window_geometry.topLeft())
    
    def _get_resource_path(self, relative_path):
        """Get absolute path to resource file."""
        base_dir = Path(__file__).parent.parent
        resource_path = base_dir / "resources" / relative_path
        return str(resource_path) if resource_path.exists() else None
    
    def _show_error(self, message):
        """Show error message to user."""
        if self.app:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setWindowTitle("Application Error")
            msg_box.setText("Application Initialization Error")
            msg_box.setDetailedText(message)
            msg_box.exec()
        else:
            print(f"CRITICAL ERROR: {message}")
    
    def run(self):
        """Run the application."""
        try:
            # Initialize Qt application
            if not self.initialize_application():
                return 1
            
            # Initialize components
            if not self.initialize_components():
                return 1
            
            # Create main window
            if not self.create_main_window():
                return 1
            
            # Show startup message
            print("🚀 Word Appendix Manager v1.0 - Started Successfully!")
            print("📄 Ready to process Word documents with PDF appendices")
            print("=" * 60)
            
            # Run event loop
            return self.app.exec()
            
        except KeyboardInterrupt:
            print("\nApplication interrupted by user")
            return 0
        except Exception as e:
            self._show_error(f"Unexpected application error: {str(e)}")
            return 1
        finally:
            # Cleanup
            if self.controller:
                self.controller.shutdown()


def main():
    """Application entry point."""
    print("=" * 60)
    print("🎯 WORD APPENDIX MANAGER v1.0")
    print("📋 Professional PDF Appendix Tool for Word Documents")
    print("⚡ Built with PySide6 & Python")
    print("=" * 60)
    
    app_manager = WordAppendixManager()
    exit_code = app_manager.run()
    
    print("👋 Word Appendix Manager - Goodbye!")
    sys.exit(exit_code)


if __name__ == "__main__":
    main()