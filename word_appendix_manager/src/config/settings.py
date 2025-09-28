"""
Application Settings Manager
Handles configuration, user preferences, and persistent settings.
"""

import json
import os
from pathlib import Path
from typing import Dict, Any, Optional
from PySide6.QtCore import QSettings, QStandardPaths


class AppSettings:
    """Manages application settings and user preferences."""
    
    def __init__(self):
        self.qt_settings = QSettings("PointOne", "WordAppendixManager")
        self.config_dir = Path(QStandardPaths.writableLocation(QStandardPaths.AppConfigLocation))
        self.config_file = self.config_dir / "app_config.json"
        
        # Ensure config directory exists
        self.config_dir.mkdir(parents=True, exist_ok=True)
        
        # Default settings
        self.defaults = {
            "ui": {
                "theme": "light",
                "window_width": 1200,
                "window_height": 800,
                "window_maximized": False,
                "last_directory": str(Path.home()),
                "show_pdf_preview": True,
                "drag_drop_enabled": True
            },
            "document": {
                "auto_backup": True,
                "backup_directory": str(Path.home() / "Documents" / "WordAppendixManager_Backups"),
                "continue_page_numbering": True,
                "appendix_numbering_style": "alphabetical",
                "heading_style": {
                    "font_name": "Arial",
                    "font_size": 14,
                    "bold": True,
                    "italic": False
                }
            },
            "pdf": {
                "max_file_size_mb": 200,
                "max_pages_warning": 2000,
                "scale_to_fit": True,
                "preserve_orientation": True,
                "compression_level": "medium"
            },
            "advanced": {
                "enable_logging": True,
                "log_level": "INFO",
                "temp_directory": str(Path.home() / "tmp" / "WordAppendixManager"),
                "cleanup_temp_files": True,
                "max_undo_levels": 10
            }
        }
        
        # Load existing settings
        self.settings = self._load_settings()
    
    def _load_settings(self) -> Dict[str, Any]:
        """Load settings from file, falling back to defaults."""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                    return self._merge_settings(self.defaults, loaded_settings)
            except (json.JSONDecodeError, IOError) as e:
                print(f"Warning: Could not load settings file: {e}. Using defaults.")
        
        return self.defaults.copy()
    
    def _merge_settings(self, defaults: Dict[str, Any], loaded: Dict[str, Any]) -> Dict[str, Any]:
        """Recursively merge loaded settings with defaults."""
        result = defaults.copy()
        
        for key, value in loaded.items():
            if key in result:
                if isinstance(result[key], dict) and isinstance(value, dict):
                    result[key] = self._merge_settings(result[key], value)
                else:
                    result[key] = value
            else:
                result[key] = value
        
        return result
    
    def save_settings(self):
        """Save current settings to file."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=2, ensure_ascii=False)
        except IOError as e:
            print(f"Warning: Could not save settings: {e}")
    
    def get(self, key_path: str, default=None) -> Any:
        """Get a setting value using dot notation."""
        keys = key_path.split('.')
        value = self.settings
        
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default
    
    def set(self, key_path: str, value: Any):
        """Set a setting value using dot notation."""
        keys = key_path.split('.')
        current = self.settings
        
        for key in keys[:-1]:
            if key not in current:
                current[key] = {}
            current = current[key]
        
        current[keys[-1]] = value
    
    def reset_to_defaults(self):
        """Reset all settings to default values."""
        self.settings = self.defaults.copy()
        self.save_settings()
    
    # Convenience properties
    @property
    def ui_theme(self) -> str:
        return self.get('ui.theme', 'light')
    
    @ui_theme.setter
    def ui_theme(self, value: str):
        self.set('ui.theme', value)
    
    @property
    def window_geometry(self) -> tuple:
        return (
            self.get('ui.window_width', 1200),
            self.get('ui.window_height', 800),
            self.get('ui.window_maximized', False)
        )
    
    @window_geometry.setter
    def window_geometry(self, value: tuple):
        width, height, maximized = value
        self.set('ui.window_width', width)
        self.set('ui.window_height', height)
        self.set('ui.window_maximized', maximized)
    
    @property
    def last_directory(self) -> str:
        return self.get('ui.last_directory', str(Path.home()))
    
    @last_directory.setter
    def last_directory(self, value: str):
        self.set('ui.last_directory', value)
    
    @property
    def auto_backup_enabled(self) -> bool:
        return self.get('document.auto_backup', True)
    
    @property
    def backup_directory(self) -> str:
        return self.get('document.backup_directory', 
                       str(Path.home() / "Documents" / "WordAppendixManager_Backups"))
    
    @property
    def max_pdf_size_mb(self) -> int:
        return self.get('pdf.max_file_size_mb', 200)
    
    @property
    def appendix_numbering_style(self) -> str:
        return self.get('document.appendix_numbering_style', 'alphabetical')
    
    @property
    def heading_style(self) -> Dict[str, Any]:
        return self.get('document.heading_style', {
            "font_name": "Arial",
            "font_size": 14,
            "bold": True,
            "italic": False
        })
    
    def __del__(self):
        """Save settings when object is destroyed."""
        try:
            self.save_settings()
        except:
            pass 
