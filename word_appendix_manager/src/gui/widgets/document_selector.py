"""
Document Selector Widget
Provides interface for selecting Word documents (open documents or files).
"""

import sys
from pathlib import Path

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent.parent))

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QComboBox, 
    QPushButton, QLabel, QButtonGroup, QRadioButton, QFileDialog
)
from PySide6.QtCore import Qt, Signal, QTimer
from typing import Dict, List, Any, Optional

from utils.logger import get_logger, LoggerMixin


class DocumentSelectorWidget(QWidget, LoggerMixin):
    """Widget for selecting Word documents."""
    
    # Signals
    document_selected = Signal(dict)  # Document info dict
    refresh_requested = Signal()
    
    def __init__(self, settings=None, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.open_documents = []
        self.selected_document = None
        
        # Auto-refresh timer
        self.refresh_timer = QTimer()
        self.refresh_timer.timeout.connect(self.auto_refresh)
        self.refresh_timer.setSingleShot(False)
        
        self.setup_ui()
        self.refresh()
        
        # Start auto-refresh every 5 seconds for open documents
        self.refresh_timer.start(5000)
    
    def setup_ui(self):
        """Set up the user interface."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Selection method radio buttons
        method_layout = QHBoxLayout()
        
        self.method_group = QButtonGroup(self)
        
        self.open_docs_radio = QRadioButton("Open Documents")
        self.open_docs_radio.setChecked(True)
        self.open_docs_radio.setToolTip("Select from currently open Word documents")
        self.method_group.addButton(self.open_docs_radio, 0)
        method_layout.addWidget(self.open_docs_radio)
        
        self.file_radio = QRadioButton("Browse File")
        self.file_radio.setToolTip("Browse and select a Word document file")
        self.method_group.addButton(self.file_radio, 1)
        method_layout.addWidget(self.file_radio)
        
        method_layout.addStretch()
        layout.addLayout(method_layout)
        
        # Document selection area
        selection_layout = QHBoxLayout()
        
        # Document dropdown/display
        self.document_combo = QComboBox()
        self.document_combo.setMinimumWidth(300)
        self.document_combo.setToolTip("Select a Word document")
        selection_layout.addWidget(self.document_combo)
        
        # Browse button (for file mode)
        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.setVisible(False)
        self.browse_btn.setToolTip("Browse for Word document file")
        selection_layout.addWidget(self.browse_btn)
        
        # Refresh button (for open documents mode)
        self.refresh_btn = QPushButton("🔄")
        self.refresh_btn.setMaximumWidth(30)
        self.refresh_btn.setToolTip("Refresh list of open documents")
        selection_layout.addWidget(self.refresh_btn)
        
        layout.addLayout(selection_layout)
        
        # Status label
        self.status_label = QLabel("Checking for open documents...")
        self.status_label.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(self.status_label)
        
        # Connect signals
        self.setup_connections()
    
    def setup_connections(self):
        """Set up signal-slot connections."""
        self.method_group.buttonClicked.connect(self.on_method_changed)
        self.document_combo.currentTextChanged.connect(self.on_document_changed)
        self.browse_btn.clicked.connect(self.browse_document)
        self.refresh_btn.clicked.connect(self.refresh)
    
    def on_method_changed(self, button):
        """Handle selection method change."""
        is_file_mode = self.file_radio.isChecked()
        
        # Show/hide appropriate controls
        self.browse_btn.setVisible(is_file_mode)
        self.refresh_btn.setVisible(not is_file_mode)
        
        if is_file_mode:
            # Switch to file selection mode
            self.document_combo.clear()
            self.document_combo.addItem("Click 'Browse...' to select a file")
            self.document_combo.setEnabled(False)
            self.status_label.setText("File mode - click Browse to select a document")
        else:
            # Switch to open documents mode
            self.document_combo.setEnabled(True)
            self.refresh()
        
        self.logger.info(f"Selection method changed to: {'file' if is_file_mode else 'open_documents'}")
    
    def refresh(self):
        """Refresh the list of available documents."""
        if self.file_radio.isChecked():
            return  # No refresh needed in file mode
        
        self.status_label.setText("Refreshing...")
        
        # Import here to avoid circular imports
        try:
            from core.word_manager import WordManager
            
            # Get list of open documents
            word_manager = WordManager(self.settings)
            self.open_documents = word_manager.get_open_documents()
            word_manager.close()
            
            # Update combo box
            self.update_document_combo()
            
            if self.open_documents:
                self.status_label.setText(f"Found {len(self.open_documents)} open document(s)")
            else:
                self.status_label.setText("No open Word documents found")
            
            self.logger.info(f"Refreshed documents: {len(self.open_documents)} found")
            
        except Exception as e:
            self.logger.error(f"Failed to refresh documents: {e}")
            self.status_label.setText(f"Error: {str(e)}")
            self.open_documents = []
            self.update_document_combo()
    
    def auto_refresh(self):
        """Automatically refresh open documents (but don't change selection)."""
        if self.open_docs_radio.isChecked():
            # Save current selection
            current_text = self.document_combo.currentText()
            
            # Refresh list
            self.refresh()
            
            # Restore selection if possible
            index = self.document_combo.findText(current_text)
            if index >= 0:
                self.document_combo.setCurrentIndex(index)
    
    def update_document_combo(self):
        """Update the document combo box with available documents."""
        self.document_combo.clear()
        
        if not self.open_documents:
            self.document_combo.addItem("No open documents found")
            self.document_combo.setEnabled(False)
            return
        
        self.document_combo.setEnabled(True)
        
        # Add placeholder
        self.document_combo.addItem("Select a document...")
        
        # Add open documents
        for doc in self.open_documents:
            display_name = doc['name']
            if not doc['saved']:
                display_name += " (unsaved)"
            
            self.document_combo.addItem(display_name)
            
            # Store document info as user data
            self.document_combo.setItemData(
                self.document_combo.count() - 1, 
                doc, 
                Qt.UserRole
            )
    
    def browse_document(self):
        """Browse for a Word document file."""
        file_path, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Word Document",
            self.settings.last_directory if self.settings else "",
            "Word Documents (*.docx *.doc);;All Files (*)"
        )
        
        if file_path:
            file_path = file_path[0]  # Take first selected file
            # Create document info for the selected file
            file_path = Path(file_path)
            doc_info = {
                'name': file_path.name,
                'full_name': str(file_path),
                'path': str(file_path),
                'saved': True,
                'index': -1,  # File mode indicator
                'source': 'file'
            }
            
            # Update combo and select this document
            self.document_combo.clear()
            self.document_combo.addItem(f"📄 {file_path.name}")
            self.document_combo.setItemData(0, doc_info, Qt.UserRole)
            self.document_combo.setCurrentIndex(0)
            
            # Update status
            self.status_label.setText(f"Selected file: {file_path.name}")
            
            # Update last directory setting
            if self.settings:
                self.settings.last_directory = str(file_path.parent)
            
            self.logger.info(f"Selected document file: {file_path}")
    
    def on_document_changed(self, text: str):
        """Handle document selection change."""
        if not text or text.startswith("Select a document") or text.startswith("No open"):
            self.selected_document = None
            return
        
        # Get document info from combo box data
        current_index = self.document_combo.currentIndex()
        doc_info = self.document_combo.itemData(current_index, Qt.UserRole)
        
        if doc_info:
            self.selected_document = doc_info
            self.document_selected.emit(doc_info)
            
            self.logger.info(f"Document selected: {doc_info['name']}")
    
    def clear_selection(self):
        """Clear the current document selection."""
        self.selected_document = None
        
        if self.open_docs_radio.isChecked():
            if self.document_combo.count() > 0:
                self.document_combo.setCurrentIndex(0)  # Select placeholder
        else:
            self.document_combo.clear()
            self.document_combo.addItem("Click 'Browse...' to select a file")
    
    def get_selected_document(self) -> Optional[Dict[str, Any]]:
        """Get the currently selected document info."""
        return self.selected_document
    
    def is_document_selected(self) -> bool:
        """Check if a document is currently selected."""
        return self.selected_document is not None
    
    def setEnabled(self, enabled: bool):
        """Enable or disable the widget."""
        super().setEnabled(enabled)
        
        # Stop auto-refresh when disabled
        if enabled:
            if not self.refresh_timer.isActive():
                self.refresh_timer.start(5000)
        else:
            self.refresh_timer.stop() 
