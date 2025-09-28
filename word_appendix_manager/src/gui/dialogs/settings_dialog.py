"""
Settings Dialog
Comprehensive settings dialog for configuring application preferences.
"""

import sys
from pathlib import Path

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent.parent))

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QFormLayout, QGroupBox, QLineEdit, QSpinBox, QComboBox,
    QPushButton, QCheckBox, QLabel, QFileDialog, QMessageBox
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from typing import Dict, Any

from utils.logger import get_logger, LoggerMixin


class SettingsDialog(QDialog, LoggerMixin):
    """Comprehensive settings dialog."""
    
    settings_changed = Signal()
    
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.temp_settings = {}  # Store temporary changes
        
        self.setup_ui()
        self.load_settings()
        self.setup_connections()
        
        self.logger.info("Opened settings dialog")
    
    def setup_ui(self):
        """Set up the settings dialog UI."""
        self.setWindowTitle("Application Settings")
        self.setModal(True)
        self.resize(600, 500)
        
        layout = QVBoxLayout(self)
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)
        
        # Create tabs
        self.create_general_tab()
        self.create_document_tab()
        self.create_pdf_tab()
        self.create_advanced_tab()
        
        # Dialog buttons
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        self.defaults_btn = QPushButton("Restore Defaults")
        buttons_layout.addWidget(self.defaults_btn)
        
        self.cancel_btn = QPushButton("Cancel")
        buttons_layout.addWidget(self.cancel_btn)
        
        self.apply_btn = QPushButton("Apply")
        buttons_layout.addWidget(self.apply_btn)
        
        self.ok_btn = QPushButton("OK")
        self.ok_btn.setDefault(True)
        self.ok_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196f3;
                color: white;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976d2;
            }
        """)
        buttons_layout.addWidget(self.ok_btn)
        
        layout.addLayout(buttons_layout)
    
    def create_general_tab(self):
        """Create the general settings tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # UI Settings Group
        ui_group = QGroupBox("User Interface")
        ui_layout = QFormLayout(ui_group)
        
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Light", "Dark", "System Default"])
        ui_layout.addRow("Theme:", self.theme_combo)
        
        self.show_preview_checkbox = QCheckBox("Show PDF preview panel")
        ui_layout.addRow("Preview:", self.show_preview_checkbox)
        
        self.drag_drop_checkbox = QCheckBox("Enable drag and drop")
        ui_layout.addRow("Drag & Drop:", self.drag_drop_checkbox)
        
        layout.addWidget(ui_group)
        
        # File Settings Group
        file_group = QGroupBox("File Locations")
        file_layout = QFormLayout(file_group)
        
        # Last directory
        last_dir_layout = QHBoxLayout()
        self.last_directory_edit = QLineEdit()
        self.last_directory_edit.setReadOnly(True)
        self.browse_last_dir_btn = QPushButton("Browse...")
        self.browse_last_dir_btn.setMaximumWidth(80)
        last_dir_layout.addWidget(self.last_directory_edit)
        last_dir_layout.addWidget(self.browse_last_dir_btn)
        file_layout.addRow("Default Directory:", last_dir_layout)
        
        # Backup directory
        backup_dir_layout = QHBoxLayout()
        self.backup_directory_edit = QLineEdit()
        self.backup_directory_edit.setReadOnly(True)
        self.browse_backup_dir_btn = QPushButton("Browse...")
        self.browse_backup_dir_btn.setMaximumWidth(80)
        backup_dir_layout.addWidget(self.backup_directory_edit)
        backup_dir_layout.addWidget(self.browse_backup_dir_btn)
        file_layout.addRow("Backup Directory:", backup_dir_layout)
        
        layout.addWidget(file_group)
        
        # Behavior Settings Group
        behavior_group = QGroupBox("Application Behavior")
        behavior_layout = QFormLayout(behavior_group)
        
        self.auto_backup_checkbox = QCheckBox("Automatically create backups")
        behavior_layout.addRow("Backups:", self.auto_backup_checkbox)
        
        self.cleanup_temp_checkbox = QCheckBox("Clean up temporary files on exit")
        behavior_layout.addRow("Cleanup:", self.cleanup_temp_checkbox)
        
        layout.addWidget(behavior_group)
        
        layout.addStretch()
        self.tab_widget.addTab(tab, "General")
    
    def create_document_tab(self):
        """Create the document settings tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Page Numbering Group
        numbering_group = QGroupBox("Appendix Numbering")
        numbering_layout = QFormLayout(numbering_group)
        
        self.numbering_style_combo = QComboBox()
        self.numbering_style_combo.addItems([
            "Alphabetical (A, B, C...)",
            "Numeric (1, 2, 3...)"
        ])
        numbering_layout.addRow("Default Style:", self.numbering_style_combo)
        
        self.continue_numbering_checkbox = QCheckBox("Continue page numbering from main document")
        numbering_layout.addRow("Page Numbers:", self.continue_numbering_checkbox)
        
        layout.addWidget(numbering_group)
        
        # Heading Style Group
        heading_group = QGroupBox("Appendix Heading Style")
        heading_layout = QFormLayout(heading_group)
        
        # Font selection
        font_layout = QHBoxLayout()
        self.font_label = QLabel("Arial, 14pt")
        self.font_btn = QPushButton("Change Font...")
        self.font_btn.setMaximumWidth(100)
        font_layout.addWidget(self.font_label)
        font_layout.addWidget(self.font_btn)
        heading_layout.addRow("Font:", font_layout)
        
        self.bold_checkbox = QCheckBox("Bold")
        heading_layout.addRow("Style:", self.bold_checkbox)
        
        self.italic_checkbox = QCheckBox("Italic")
        heading_layout.addRow("", self.italic_checkbox)
        
        layout.addWidget(heading_group)
        
        # Document Processing Group
        processing_group = QGroupBox("Document Processing")
        processing_layout = QFormLayout(processing_group)
        
        self.max_undo_spinbox = QSpinBox()
        self.max_undo_spinbox.setMinimum(1)
        self.max_undo_spinbox.setMaximum(50)
        processing_layout.addRow("Max Undo Levels:", self.max_undo_spinbox)
        
        layout.addWidget(processing_group)
        
        layout.addStretch()
        self.tab_widget.addTab(tab, "Document")
    
    def create_pdf_tab(self):
        """Create the PDF settings tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # File Size Limits Group
        limits_group = QGroupBox("File Size Limits")
        limits_layout = QFormLayout(limits_group)
        
        self.max_file_size_spinbox = QSpinBox()
        self.max_file_size_spinbox.setMinimum(1)
        self.max_file_size_spinbox.setMaximum(1000)
        self.max_file_size_spinbox.setSuffix(" MB")
        limits_layout.addRow("Max PDF Size:", self.max_file_size_spinbox)
        
        self.max_pages_spinbox = QSpinBox()
        self.max_pages_spinbox.setMinimum(1)
        self.max_pages_spinbox.setMaximum(10000)
        limits_layout.addRow("Max Pages Warning:", self.max_pages_spinbox)
        
        layout.addWidget(limits_group)
        
        # Processing Options Group
        pdf_processing_group = QGroupBox("PDF Processing")
        pdf_processing_layout = QFormLayout(pdf_processing_group)
        
        self.scale_to_fit_checkbox = QCheckBox("Scale PDFs to fit Word page size")
        pdf_processing_layout.addRow("Scaling:", self.scale_to_fit_checkbox)
        
        self.preserve_orientation_checkbox = QCheckBox("Preserve original orientation")
        pdf_processing_layout.addRow("Orientation:", self.preserve_orientation_checkbox)
        
        self.compression_combo = QComboBox()
        self.compression_combo.addItems(["Low", "Medium", "High"])
        pdf_processing_layout.addRow("Compression Level:", self.compression_combo)
        
        layout.addWidget(pdf_processing_group)
        
        layout.addStretch()
        self.tab_widget.addTab(tab, "PDF")
    
    def create_advanced_tab(self):
        """Create the advanced settings tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Logging Group
        logging_group = QGroupBox("Logging & Debugging")
        logging_layout = QFormLayout(logging_group)
        
        self.enable_logging_checkbox = QCheckBox("Enable application logging")
        logging_layout.addRow("Logging:", self.enable_logging_checkbox)
        
        self.log_level_combo = QComboBox()
        self.log_level_combo.addItems(["DEBUG", "INFO", "WARNING", "ERROR"])
        logging_layout.addRow("Log Level:", self.log_level_combo)
        
        layout.addWidget(logging_group)
        
        # Performance Group
        performance_group = QGroupBox("Performance")
        performance_layout = QFormLayout(performance_group)
        
        # Temp directory
        temp_dir_layout = QHBoxLayout()
        self.temp_directory_edit = QLineEdit()
        self.temp_directory_edit.setReadOnly(True)
        self.browse_temp_dir_btn = QPushButton("Browse...")
        self.browse_temp_dir_btn.setMaximumWidth(80)
        temp_dir_layout.addWidget(self.temp_directory_edit)
        temp_dir_layout.addWidget(self.browse_temp_dir_btn)
        performance_layout.addRow("Temp Directory:", temp_dir_layout)
        
        layout.addWidget(performance_group)
        
        layout.addStretch()
        self.tab_widget.addTab(tab, "Advanced")
    
    def setup_connections(self):
        """Set up signal-slot connections."""
        # File browsers
        self.browse_last_dir_btn.clicked.connect(self.browse_last_directory)
        self.browse_backup_dir_btn.clicked.connect(self.browse_backup_directory)
        self.browse_temp_dir_btn.clicked.connect(self.browse_temp_directory)
        
        # Font selection
        self.font_btn.clicked.connect(self.select_font)
        
        # Dialog buttons
        self.defaults_btn.clicked.connect(self.restore_defaults)
        self.cancel_btn.clicked.connect(self.reject)
        self.apply_btn.clicked.connect(self.apply_settings)
        self.ok_btn.clicked.connect(self.accept_settings)
    
    def load_settings(self):
        """Load current settings into the dialog."""
        # General tab
        theme = self.settings.ui_theme
        theme_index = {"light": 0, "dark": 1, "system": 2}.get(theme, 0)
        self.theme_combo.setCurrentIndex(theme_index)
        
        self.show_preview_checkbox.setChecked(self.settings.get('ui.show_pdf_preview', True))
        self.drag_drop_checkbox.setChecked(self.settings.get('ui.drag_drop_enabled', True))
        
        self.last_directory_edit.setText(self.settings.last_directory)
        self.backup_directory_edit.setText(self.settings.backup_directory)
        
        self.auto_backup_checkbox.setChecked(self.settings.auto_backup_enabled)
        self.cleanup_temp_checkbox.setChecked(self.settings.get('advanced.cleanup_temp_files', True))
        
        # Document tab
        numbering_style = self.settings.appendix_numbering_style
        numbering_index = 0 if numbering_style == "alphabetical" else 1
        self.numbering_style_combo.setCurrentIndex(numbering_index)
        
        self.continue_numbering_checkbox.setChecked(self.settings.get('document.continue_page_numbering', True))
        
        # Load font settings
        heading_style = self.settings.heading_style
        font_name = heading_style.get('font_name', 'Arial')
        font_size = heading_style.get('font_size', 14)
        self.font_label.setText(f"{font_name}, {font_size}pt")
        
        self.bold_checkbox.setChecked(heading_style.get('bold', True))
        self.italic_checkbox.setChecked(heading_style.get('italic', False))
        
        self.max_undo_spinbox.setValue(self.settings.get('advanced.max_undo_levels', 10))
        
        # PDF tab
        self.max_file_size_spinbox.setValue(self.settings.max_pdf_size_mb)
        self.max_pages_spinbox.setValue(self.settings.get('pdf.max_pages_warning', 2000))
        
        self.scale_to_fit_checkbox.setChecked(self.settings.get('pdf.scale_to_fit', True))
        self.preserve_orientation_checkbox.setChecked(self.settings.get('pdf.preserve_orientation', True))
        
        compression_level = self.settings.get('pdf.compression_level', 'medium')
        compression_index = {"low": 0, "medium": 1, "high": 2}.get(compression_level, 1)
        self.compression_combo.setCurrentIndex(compression_index)
        
        # Advanced tab
        self.enable_logging_checkbox.setChecked(self.settings.get('advanced.enable_logging', True))
        
        log_level = self.settings.get('advanced.log_level', 'INFO')
        try:
            log_level_index = ["DEBUG", "INFO", "WARNING", "ERROR"].index(log_level)
            self.log_level_combo.setCurrentIndex(log_level_index)
        except ValueError:
            self.log_level_combo.setCurrentIndex(1)  # Default to INFO
        
        self.temp_directory_edit.setText(self.settings.get('advanced.temp_directory', ''))
    
    def browse_last_directory(self):
        """Browse for last directory."""
        directory = QFileDialog.getExistingDirectory(
            self, "Select Default Directory", self.last_directory_edit.text()
        )
        if directory:
            self.last_directory_edit.setText(directory)
    
    def browse_backup_directory(self):
        """Browse for backup directory."""
        directory = QFileDialog.getExistingDirectory(
            self, "Select Backup Directory", self.backup_directory_edit.text()
        )
        if directory:
            self.backup_directory_edit.setText(directory)
    
    def browse_temp_directory(self):
        """Browse for temp directory."""
        directory = QFileDialog.getExistingDirectory(
            self, "Select Temp Directory", self.temp_directory_edit.text()
        )
        if directory:
            self.temp_directory_edit.setText(directory)
    
    def select_font(self):
        """Select font for appendix headings."""
        from PySide6.QtWidgets import QFontDialog
        
        heading_style = self.settings.heading_style
        current_font = QFont(
            heading_style.get('font_name', 'Arial'),
            heading_style.get('font_size', 14)
        )
        current_font.setBold(heading_style.get('bold', True))
        current_font.setItalic(heading_style.get('italic', False))
        
        font, ok = QFontDialog.getFont(current_font, self)
        if ok:
            self.font_label.setText(f"{font.family()}, {font.pointSize()}pt")
            
            # Store font info for later
            self.temp_settings['font_family'] = font.family()
            self.temp_settings['font_size'] = font.pointSize()
            self.bold_checkbox.setChecked(font.bold())
            self.italic_checkbox.setChecked(font.italic())
    
    def restore_defaults(self):
        """Restore all settings to defaults."""
        reply = QMessageBox.question(
            self, "Restore Defaults", 
            "Are you sure you want to restore all settings to their default values?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.settings.reset_to_defaults()
            self.load_settings()
            QMessageBox.information(self, "Defaults Restored", "All settings have been reset to defaults.")
    
    def apply_settings(self):
        """Apply settings without closing dialog."""
        self.save_settings()
        self.settings_changed.emit()
    
    def accept_settings(self):
        """Accept and apply settings, then close dialog."""
        self.save_settings()
        self.settings_changed.emit()
        self.accept()
    
    def save_settings(self):
        """Save settings from dialog to settings object."""
        # General settings
        theme_map = {0: "light", 1: "dark", 2: "system"}
        self.settings.ui_theme = theme_map[self.theme_combo.currentIndex()]
        
        self.settings.set('ui.show_pdf_preview', self.show_preview_checkbox.isChecked())
        self.settings.set('ui.drag_drop_enabled', self.drag_drop_checkbox.isChecked())
        
        self.settings.last_directory = self.last_directory_edit.text()
        self.settings.set('document.backup_directory', self.backup_directory_edit.text())
        
        self.settings.set('document.auto_backup', self.auto_backup_checkbox.isChecked())
        self.settings.set('advanced.cleanup_temp_files', self.cleanup_temp_checkbox.isChecked())
        
        # Document settings
        numbering_style = "alphabetical" if self.numbering_style_combo.currentIndex() == 0 else "numeric"
        self.settings.set('document.appendix_numbering_style', numbering_style)
        
        self.settings.set('document.continue_page_numbering', self.continue_numbering_checkbox.isChecked())
        
        # Font settings
        font_family = self.temp_settings.get('font_family', 'Arial')
        font_size = self.temp_settings.get('font_size', 14)
        
        heading_style = {
            'font_name': font_family,
            'font_size': font_size,
            'bold': self.bold_checkbox.isChecked(),
            'italic': self.italic_checkbox.isChecked()
        }
        self.settings.set('document.heading_style', heading_style)
        
        self.settings.set('advanced.max_undo_levels', self.max_undo_spinbox.value())
        
        # PDF settings
        self.settings.set('pdf.max_file_size_mb', self.max_file_size_spinbox.value())
        self.settings.set('pdf.max_pages_warning', self.max_pages_spinbox.value())
        
        self.settings.set('pdf.scale_to_fit', self.scale_to_fit_checkbox.isChecked())
        self.settings.set('pdf.preserve_orientation', self.preserve_orientation_checkbox.isChecked())
        
        compression_map = {0: "low", 1: "medium", 2: "high"}
        compression_level = compression_map[self.compression_combo.currentIndex()]
        self.settings.set('pdf.compression_level', compression_level)
        
        # Advanced settings
        self.settings.set('advanced.enable_logging', self.enable_logging_checkbox.isChecked())
        
        log_levels = ["DEBUG", "INFO", "WARNING", "ERROR"]
        log_level = log_levels[self.log_level_combo.currentIndex()]
        self.settings.set('advanced.log_level', log_level)
        
        temp_dir = self.temp_directory_edit.text()
        if temp_dir:
            self.settings.set('advanced.temp_directory', temp_dir)
        
        # Save to file
        self.settings.save_settings()
        self.logger.info("Settings saved successfully")
    
    def closeEvent(self, event):
        """Handle dialog close event."""
        self.logger.info("Closed settings dialog")
        event.accept() 
