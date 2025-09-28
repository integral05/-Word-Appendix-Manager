"""
Appendix Edit Dialog
Dialog for editing appendix properties and settings.
"""

import sys
from pathlib import Path

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent.parent))

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QSpinBox, QComboBox, QPushButton, 
    QLabel, QGroupBox, QCheckBox, QTextEdit, QFileDialog,
    QMessageBox
)
from PySide6.QtCore import Qt
from typing import Dict, Any

from utils.logger import get_logger, LoggerMixin


class AppendixEditDialog(QDialog, LoggerMixin):
    """Dialog for editing appendix properties."""
    
    def __init__(self, appendix_data: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.appendix_data = appendix_data.copy()
        self.original_data = appendix_data.copy()
        
        self.setup_ui()
        self.populate_fields()
        self.setup_connections()
        
        self.logger.info(f"Opened appendix edit dialog for: {appendix_data.get('title', 'Unknown')}")
    
    def setup_ui(self):
        """Set up the dialog UI."""
        self.setWindowTitle("Edit Appendix")
        self.setModal(True)
        self.resize(500, 600)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        
        # Basic Information Group
        basic_group = QGroupBox("Basic Information")
        basic_layout = QFormLayout(basic_group)
        
        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("e.g., Appendix A - Financial Data")
        basic_layout.addRow("Title:", self.title_edit)
        
        # File path display and change
        file_layout = QHBoxLayout()
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.setMaximumWidth(80)
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(self.browse_btn)
        basic_layout.addRow("PDF File:", file_layout)
        
        layout.addWidget(basic_group)
        
        # File Information Group
        info_group = QGroupBox("File Information")
        info_layout = QFormLayout(info_group)
        
        self.pages_label = QLabel()
        info_layout.addRow("Pages:", self.pages_label)
        
        self.size_label = QLabel()
        info_layout.addRow("File Size:", self.size_label)
        
        self.orientation_label = QLabel()
        info_layout.addRow("Orientation:", self.orientation_label)
        
        layout.addWidget(info_group)
        
        # Display Options Group
        display_group = QGroupBox("Display Options")
        display_layout = QFormLayout(display_group)
        
        self.numbering_combo = QComboBox()
        self.numbering_combo.addItems([
            "Auto (follows global setting)",
            "Alphabetical (A, B, C...)",
            "Numeric (1, 2, 3...)",
            "Custom"
        ])
        display_layout.addRow("Numbering Style:", self.numbering_combo)
        
        self.custom_number_edit = QLineEdit()
        self.custom_number_edit.setPlaceholderText("e.g., Appendix I, Section A")
        self.custom_number_edit.setEnabled(False)
        display_layout.addRow("Custom Number:", self.custom_number_edit)
        
        layout.addWidget(display_group)
        
        # Processing Options Group
        processing_group = QGroupBox("Processing Options")
        processing_layout = QVBoxLayout(processing_group)
        
        self.scale_checkbox = QCheckBox("Scale PDF to fit Word page size")
        self.scale_checkbox.setChecked(True)
        processing_layout.addWidget(self.scale_checkbox)
        
        self.preserve_orientation_checkbox = QCheckBox("Preserve original orientation")
        self.preserve_orientation_checkbox.setChecked(True)
        processing_layout.addWidget(self.preserve_orientation_checkbox)
        
        self.optimize_checkbox = QCheckBox("Optimize PDF for smaller file size")
        processing_layout.addWidget(self.optimize_checkbox)
        
        layout.addWidget(processing_group)
        
        # Notes/Comments
        notes_group = QGroupBox("Notes")
        notes_layout = QVBoxLayout(notes_group)
        
        self.notes_edit = QTextEdit()
        self.notes_edit.setMaximumHeight(80)
        self.notes_edit.setPlaceholderText("Optional notes about this appendix...")
        notes_layout.addWidget(self.notes_edit)
        
        layout.addWidget(notes_group)
        
        # Warnings/Errors display
        if self.appendix_data.get('warnings') or not self.appendix_data.get('valid', True):
            warnings_group = QGroupBox("Warnings & Issues")
            warnings_layout = QVBoxLayout(warnings_group)
            
            self.warnings_text = QTextEdit()
            self.warnings_text.setMaximumHeight(60)
            self.warnings_text.setReadOnly(True)
            self.warnings_text.setStyleSheet("background-color: #fff3cd; border: 1px solid #ffc107;")
            warnings_layout.addWidget(self.warnings_text)
            
            layout.addWidget(warnings_group)
        
        # Dialog buttons
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        self.reset_btn = QPushButton("Reset")
        self.reset_btn.setToolTip("Reset to original values")
        buttons_layout.addWidget(self.reset_btn)
        
        self.cancel_btn = QPushButton("Cancel")
        buttons_layout.addWidget(self.cancel_btn)
        
        self.ok_btn = QPushButton("OK")
        self.ok_btn.setDefault(True)
        self.ok_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196f3;
                color: white;
                border: none;
                padding: 8px 24px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976d2;
            }
        """)
        buttons_layout.addWidget(self.ok_btn)
        
        layout.addLayout(buttons_layout)
    
    def setup_connections(self):
        """Set up signal-slot connections."""
        self.browse_btn.clicked.connect(self.browse_file)
        self.numbering_combo.currentTextChanged.connect(self.on_numbering_changed)
        self.title_edit.textChanged.connect(self.validate_input)
        
        self.reset_btn.clicked.connect(self.reset_fields)
        self.cancel_btn.clicked.connect(self.reject)
        self.ok_btn.clicked.connect(self.accept_changes)
    
    def populate_fields(self):
        """Populate dialog fields with appendix data."""
        # Basic information
        self.title_edit.setText(self.appendix_data.get('title', ''))
        self.file_path_edit.setText(self.appendix_data.get('path', ''))
        
        # File information
        self.pages_label.setText(str(self.appendix_data.get('page_count', 0)))
        size_mb = self.appendix_data.get('size_mb', 0)
        self.size_label.setText(f"{size_mb:.1f} MB")
        orientation = self.appendix_data.get('orientation', 'Unknown')
        self.orientation_label.setText(orientation.title())
        
        # Display options
        numbering_style = self.appendix_data.get('numbering_style', 'auto')
        if numbering_style == 'alphabetical':
            self.numbering_combo.setCurrentIndex(1)
        elif numbering_style == 'numeric':
            self.numbering_combo.setCurrentIndex(2)
        elif numbering_style == 'custom':
            self.numbering_combo.setCurrentIndex(3)
            self.custom_number_edit.setText(self.appendix_data.get('custom_number', ''))
        else:
            self.numbering_combo.setCurrentIndex(0)
        
        # Processing options
        self.scale_checkbox.setChecked(self.appendix_data.get('scale_to_fit', True))
        self.preserve_orientation_checkbox.setChecked(self.appendix_data.get('preserve_orientation', True))
        self.optimize_checkbox.setChecked(self.appendix_data.get('optimize_pdf', False))
        
        # Notes
        self.notes_edit.setPlainText(self.appendix_data.get('notes', ''))
        
        # Warnings
        if hasattr(self, 'warnings_text'):
            warnings = self.appendix_data.get('warnings', [])
            if warnings:
                self.warnings_text.setPlainText('\n'.join(f"• {warning}" for warning in warnings))
            elif not self.appendix_data.get('valid', True):
                self.warnings_text.setPlainText("• This PDF file has validation errors")
    
    def on_numbering_changed(self, text: str):
        """Handle numbering style change."""
        is_custom = "Custom" in text
        self.custom_number_edit.setEnabled(is_custom)
        if not is_custom:
            self.custom_number_edit.clear()
    
    def validate_input(self):
        """Validate user input and enable/disable OK button."""
        title = self.title_edit.text().strip()
        file_path = self.file_path_edit.text().strip()
        
        # Check if title is not empty
        title_valid = bool(title)
        
        # Check if file exists
        file_valid = bool(file_path) and Path(file_path).exists()
        
        # Enable OK button only if both are valid
        self.ok_btn.setEnabled(title_valid and file_valid)
        
        # Update title preview
        if title_valid:
            self.title_edit.setStyleSheet("")
        else:
            self.title_edit.setStyleSheet("border: 1px solid #dc3545;")
    
    def browse_file(self):
        """Browse for a new PDF file."""
        current_path = self.file_path_edit.text()
        initial_dir = str(Path(current_path).parent) if current_path else str(Path.home())
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select PDF File",
            initial_dir,
            "PDF Files (*.pdf);;All Files (*)"
        )
        
        if file_path:
            self.file_path_edit.setText(file_path)
            
            # Update file information
            try:
                from core.pdf_handler import PDFHandler
                pdf_handler = PDFHandler()
                pdf_info = pdf_handler.validate_pdf_file(file_path)
                
                self.pages_label.setText(str(pdf_info.get('page_count', 0)))
                size_mb = pdf_info.get('size_mb', 0)
                self.size_label.setText(f"{size_mb:.1f} MB")
                orientation = pdf_info.get('orientation', 'Unknown')
                self.orientation_label.setText(orientation.title())
                
                # Update appendix data
                self.appendix_data['path'] = file_path
                self.appendix_data['page_count'] = pdf_info.get('page_count', 0)
                self.appendix_data['size_mb'] = size_mb
                self.appendix_data['orientation'] = orientation
                self.appendix_data['valid'] = pdf_info.get('valid', False)
                self.appendix_data['warnings'] = pdf_info.get('warnings', [])
                
                self.logger.info(f"Updated PDF file to: {Path(file_path).name}")
                
            except Exception as e:
                self.logger.error(f"Failed to analyze new PDF: {e}")
                # Show error but don't prevent user from continuing
                self.pages_label.setText("Unknown")
                self.size_label.setText("Unknown")
                self.orientation_label.setText("Unknown")
            
            self.validate_input()
    
    def reset_fields(self):
        """Reset all fields to original values."""
        self.appendix_data = self.original_data.copy()
        self.populate_fields()
        self.validate_input()
        self.logger.info("Reset appendix edit dialog to original values")
    
    def accept_changes(self):
        """Accept changes and update appendix data."""
        if not self.ok_btn.isEnabled():
            return
        
        # Update appendix data with form values
        self.appendix_data['title'] = self.title_edit.text().strip()
        self.appendix_data['path'] = self.file_path_edit.text().strip()
        
        # Numbering style
        numbering_text = self.numbering_combo.currentText()
        if "Alphabetical" in numbering_text:
            self.appendix_data['numbering_style'] = 'alphabetical'
        elif "Numeric" in numbering_text:
            self.appendix_data['numbering_style'] = 'numeric'
        elif "Custom" in numbering_text:
            self.appendix_data['numbering_style'] = 'custom'
            self.appendix_data['custom_number'] = self.custom_number_edit.text().strip()
        else:
            self.appendix_data['numbering_style'] = 'auto'
        
        # Processing options
        self.appendix_data['scale_to_fit'] = self.scale_checkbox.isChecked()
        self.appendix_data['preserve_orientation'] = self.preserve_orientation_checkbox.isChecked()
        self.appendix_data['optimize_pdf'] = self.optimize_checkbox.isChecked()
        
        # Notes
        self.appendix_data['notes'] = self.notes_edit.toPlainText().strip()
        
        self.logger.info(f"Accepted changes for appendix: {self.appendix_data['title']}")
        self.accept()
    
    def get_appendix_data(self) -> Dict[str, Any]:
        """Get the updated appendix data."""
        return self.appendix_data.copy()
    
    def closeEvent(self, event):
        """Handle dialog close event."""
        self.logger.info("Closed appendix edit dialog")
        event.accept() 
