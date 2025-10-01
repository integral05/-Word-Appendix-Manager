"""
Theme Manager
Handles theme loading and switching for the application.
"""

import os
from pathlib import Path
from typing import Optional
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QObject, Signal

from utils.logger import get_logger, LoggerMixin


class ThemeManager(QObject, LoggerMixin):
    """Manages application themes and styling."""
    
    # Signal emitted when theme changes
    theme_changed = Signal(str)  # theme name
    
    # Available themes
    THEMES = {
        "light": "main_style.qss",
        "dark": "dark_theme.qss",
        "system": None  # Will detect system theme
    }
    
    def __init__(self, settings=None):
        super().__init__()
        self.settings = settings
        self.current_theme = "light"
        self.resource_dir = self._get_resource_directory()
        self.custom_widgets_path = self.resource_dir / "custom_widgets.qss"
        
        # Cache loaded stylesheets
        self._stylesheet_cache = {}
        
        self.logger.info("Theme manager initialized")
    
    def _get_resource_directory(self) -> Path:
        """Get the resources directory path."""
        # Get the src directory (where main.py is)
        src_dir = Path(__file__).parent.parent
        
        # Go up one level to project root, then into resources/styles
        resource_dir = src_dir.parent / "resources" / "styles"
        
        if not resource_dir.exists():
            self.logger.warning(f"Resources directory not found: {resource_dir}")
            # Try alternative path (for packaged applications)
            import sys
            if hasattr(sys, '_MEIPASS'):
                resource_dir = Path(sys._MEIPASS) / "resources" / "styles"
        
        return resource_dir
    
    def load_theme(self, theme_name: str) -> bool:
        """Load and apply a theme."""
        if theme_name not in self.THEMES:
            self.logger.error(f"Unknown theme: {theme_name}")
            return False
        
        try:
            # Handle system theme
            if theme_name == "system":
                theme_name = self._detect_system_theme()
            
            # Load the theme stylesheet
            stylesheet = self._load_stylesheet(theme_name)
            
            if stylesheet:
                # Apply to application
                app = QApplication.instance()
                if app:
                    app.setStyleSheet(stylesheet)
                    self.current_theme = theme_name
                    self.theme_changed.emit(theme_name)
                    self.logger.info(f"Applied theme: {theme_name}")
                    return True
            
            return False
            
        except Exception as e:
            self.logger.error(f"Failed to load theme {theme_name}: {e}")
            return False
    
    def _load_stylesheet(self, theme_name: str) -> Optional[str]:
        """Load stylesheet content for a theme."""
        # Check cache first
        if theme_name in self._stylesheet_cache:
            return self._stylesheet_cache[theme_name]
        
        try:
            # Load custom widgets stylesheet (shared across themes)
            custom_widgets_css = ""
            if self.custom_widgets_path.exists():
                with open(self.custom_widgets_path, 'r', encoding='utf-8') as f:
                    custom_widgets_css = f.read()
            else:
                self.logger.warning(f"Custom widgets stylesheet not found: {self.custom_widgets_path}")
            
            # Load theme-specific stylesheet
            theme_file = self.THEMES.get(theme_name)
            theme_css = ""
            
            if theme_file:
                theme_path = self.resource_dir / theme_file
                if theme_path.exists():
                    with open(theme_path, 'r', encoding='utf-8') as f:
                        theme_css = f.read()
                else:
                    self.logger.error(f"Theme file not found: {theme_path}")
                    return None
            
            # Combine stylesheets: theme-specific first (for CSS variable definitions), then custom widgets
            combined_stylesheet = theme_css + "\n\n" + custom_widgets_css
            
            # Cache the result
            self._stylesheet_cache[theme_name] = combined_stylesheet
            
            return combined_stylesheet
            
        except Exception as e:
            self.logger.error(f"Error loading stylesheet for {theme_name}: {e}")
            return None
    
    def _detect_system_theme(self) -> str:
        """Detect the system theme preference."""
        try:
            from PySide6.QtGui import QPalette
            from PySide6.QtWidgets import QApplication
            
            app = QApplication.instance()
            if app:
                palette = app.palette()
                # Check if window background is dark
                window_color = palette.color(QPalette.Window)
                luminance = (0.299 * window_color.red() + 
                           0.587 * window_color.green() + 
                           0.114 * window_color.blue())
                
                # If luminance is low, it's a dark theme
                if luminance < 128:
                    self.logger.info("Detected dark system theme")
                    return "dark"
            
            self.logger.info("Detected light system theme")
            return "light"
            
        except Exception as e:
            self.logger.warning(f"Could not detect system theme: {e}")
            return "light"
    
    def get_current_theme(self) -> str:
        """Get the currently active theme name."""
        return self.current_theme
    
    def get_available_themes(self) -> list:
        """Get list of available theme names."""
        return list(self.THEMES.keys())
    
    def refresh_theme(self):
        """Refresh the current theme (reload from file)."""
        # Clear cache for current theme
        if self.current_theme in self._stylesheet_cache:
            del self._stylesheet_cache[self.current_theme]
        
        # Reload theme
        self.load_theme(self.current_theme)
    
    def apply_custom_stylesheet(self, widget, custom_css: str):
        """Apply custom CSS to a specific widget."""
        try:
            widget.setStyleSheet(custom_css)
            self.logger.debug(f"Applied custom stylesheet to {widget.__class__.__name__}")
        except Exception as e:
            self.logger.error(f"Failed to apply custom stylesheet: {e}")
    
    def get_theme_color(self, color_name: str) -> str:
        """Get a color value from the current theme."""
        # Define color mappings for themes
        theme_colors = {
            "light": {
                "primary": "#2196f3",
                "secondary": "#f8f9fa",
                "success": "#4caf50",
                "warning": "#ff9800",
                "error": "#f44336",
                "info": "#2196f3",
                "text_primary": "#212529",
                "text_secondary": "#6c757d",
                "bg_primary": "#ffffff",
                "bg_secondary": "#f8f9fa",
                "border": "#d0d0d0"
            },
            "dark": {
                "primary": "#007acc",
                "secondary": "#252526",
                "success": "#4ec9b0",
                "warning": "#ce9178",
                "error": "#f48771",
                "info": "#007acc",
                "text_primary": "#cccccc",
                "text_secondary": "#999999",
                "bg_primary": "#1e1e1e",
                "bg_secondary": "#252526",
                "border": "#3e3e42"
            }
        }
        
        current_colors = theme_colors.get(self.current_theme, theme_colors["light"])
        return current_colors.get(color_name, "#000000")
    
    def create_inline_style(self, widget_type: str, properties: dict) -> str:
        """Create an inline stylesheet string."""
        style_parts = []
        for prop, value in properties.items():
            # Convert Python-style property names to CSS
            css_prop = prop.replace("_", "-")
            style_parts.append(f"{css_prop}: {value}")
        
        return f"{widget_type} {{ {'; '.join(style_parts)}; }}"
    
    def save_theme_preference(self, theme_name: str):
        """Save theme preference to settings."""
        if self.settings:
            self.settings.ui_theme = theme_name
            self.settings.save_settings()
            self.logger.info(f"Saved theme preference: {theme_name}")
    
    def load_saved_theme(self):
        """Load theme from saved settings."""
        if self.settings:
            saved_theme = self.settings.ui_theme
            self.load_theme(saved_theme)
        else:
            self.load_theme("light")


# Convenience function for getting theme manager instance
_theme_manager_instance = None

def get_theme_manager(settings=None) -> ThemeManager:
    """Get or create the global theme manager instance."""
    global _theme_manager_instance
    if _theme_manager_instance is None:
        _theme_manager_instance = ThemeManager(settings)
    return _theme_manager_instance