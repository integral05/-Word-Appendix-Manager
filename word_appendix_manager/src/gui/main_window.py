"""
Main Window
The primary user interface for the Word Appendix Manager application.
"""

import sys
from pathlib import Path

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent))

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QLabel, QComboBox, QPushButton, QTextEdit, QProgressBar, 
    QStatusBar, QMenuBar, QMenu, QGroupBox, QFrame, QMessageBox
)
from PySide6.QtCore import Qt, QSize, Signal, QTimer
from PySide6.QtGui import QAction, QIcon
from typing import Optional, Dict, Any, List

from utils.logger import get_logger, LoggerMixin
from gui.widgets.document_selector import DocumentSelectorWidget
from gui.widgets.appendix_list_widget import AppendixListWidget
from gui.widgets.drag_drop_area import DragDropArea
from gui.widgets.pdf_preview_widget import PDFPreviewWidget


class MainWindow(QMainWindow, LoggerMixin):
    """Main application window."""
    
    # Signals
    document_selected = Signal(dict)     # Document info
    pdf_files_added = Signal(list)       # List of PDF file paths
    appendix_removed = Signal(int)       # Index of appendix to remove
    process_requested = Signal()         # Request to process appendices
    
    def __init__(self, settings=None, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.current_document = None
        self.appendix_data = []
        
        # Initialize UI
        self.setup_ui()
        self.setup_menus()
        self.setup_status_bar()
        self.setup_connections()
        self.apply_settings()
        
        self.logger.info("Main window initialized")
    
    def setup_ui(self):
        """Set up the main user interface."""
        self.setWindowTitle("Word Appendix Manager v1.0")
        self.setMinimumSize(1000, 700)
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Document selection section
        doc_section = self.create_document_section()
        main_layout.addWidget(doc_section)
        
        # Main content splitter
        content_splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(content_splitter)
        
        # Left panel - Appendix management
        left_panel = self.create_left_panel()
        content_splitter.addWidget(left_panel)
        
        # Right panel - Preview and details
        right_panel = self.create_right_panel()
        content_splitter.addWidget(right_panel)
        
        # Set splitter proportions (60% left, 40% right)
        content_splitter.setSizes([600, 400])
        
        # Progress section
        progress_section = self.create_progress_section()
        main_layout.addWidget(progress_section)
        
        # Action buttons
        button_section = self.create_button_section()
        main_layout.addWidget(button_section)
    
    def create_document_section(self) -> QWidget:
        """Create the document selection section."""
        group = QGroupBox("Document Selection")
        layout = QVBoxLayout(group)
        
        # Document selector widget
        self.document_selector = DocumentSelectorWidget(self.settings)
        layout.addWidget(self.document_selector)
        
        # Document info display
        self.doc_info_label = QLabel("No document selected")
        self.doc_info_label.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(self.doc_info_label)
        
        return group
    
    def create_left_panel(self) -> QWidget:
        """Create the left panel with appendix management."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # PDF file input section
        input_group = QGroupBox("Add PDF Files")
        input_layout = QVBoxLayout(input_group)
        
        # Drag and drop area
        self.drag_drop_area = DragDropArea()
        self.drag_drop_area.setMinimumHeight(100)
        input_layout.addWidget(self.drag_drop_area)
        
        # Browse button
        browse_btn = QPushButton("Browse PDF Files...")
        browse_btn.clicked.connect(self.browse_pdf_files)
        input_layout.addWidget(browse_btn)
        
        layout.addWidget(input_group)
        
        # Appendix list section
        list_group = QGroupBox("Appendices")
        list_layout = QVBoxLayout(list_group)
        
        # Appendix list widget
        self.appendix_list = AppendixListWidget()
        self.appendix_list.setMinimumHeight(200)
        list_layout.addWidget(self.appendix_list)
        
        # List control buttons
        list_buttons_layout = QHBoxLayout()
        
        self.move_up_btn = QPushButton("↑ Move Up")
        self.move_up_btn.setEnabled(False)
        list_buttons_layout.addWidget(self.move_up_btn)
        
        self.move_down_btn = QPushButton("↓ Move Down")
        self.move_down_btn.setEnabled(False)
        list_buttons_layout.addWidget(self.move_down_btn)
        
        self.remove_btn = QPushButton("🗑 Remove")
        self.remove_btn.setEnabled(False)
        list_buttons_layout.addWidget(self.remove_btn)
        
        list_layout.addLayout(list_buttons_layout)
        
        layout.addWidget(list_group)
        
        return panel
    
    def create_right_panel(self) -> QWidget:
        """Create the right panel with preview and details."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # PDF preview section
        preview_group = QGroupBox("PDF Preview")
        preview_layout = QVBoxLayout(preview_group)
        
        self.pdf_preview = PDFPreviewWidget()
        preview_layout.addWidget(self.pdf_preview)
        
        layout.addWidget(preview_group)
        
        # Details section
        details_group = QGroupBox("Details")
        details_layout = QVBoxLayout(details_group)
        
        self.details_text = QTextEdit()
        self.details_text.setReadOnly(True)
        self.details_text.setMaximumHeight(150)
        self.details_text.setHtml("<i>Select an appendix to view details</i>")
        details_layout.addWidget(self.details_text)
        
        layout.addWidget(details_group)
        
        # Settings section
        settings_group = QGroupBox("Settings")
        settings_layout = QVBoxLayout(settings_group)
        
        # Numbering style
        numbering_layout = QHBoxLayout()
        numbering_layout.addWidget(QLabel("Appendix Numbering:"))
        self.numbering_combo = QComboBox()
        self.numbering_combo.addItems(["Alphabetical (A, B, C...)", "Numeric (1, 2, 3...)"])
        numbering_layout.addWidget(self.numbering_combo)
        settings_layout.addLayout(numbering_layout)
        
        # Auto backup checkbox
        self.backup_checkbox = QPushButton("✓ Auto-create backup")
        self.backup_checkbox.setCheckable(True)
        self.backup_checkbox.setChecked(True)
        settings_layout.addWidget(self.backup_checkbox)
        
        layout.addWidget(settings_group)
        
        return panel
    
    def create_progress_section(self) -> QWidget:
        """Create the progress display section."""
        frame = QFrame()
        frame.setFrameStyle(QFrame.Box)
        frame.setVisible(False)  # Hidden by default
        
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(10, 5, 10, 5)
        
        self.progress_label = QLabel("Ready")
        layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        self.progress_frame = frame
        return frame
    
    def create_button_section(self) -> QWidget:
        """Create the main action buttons section."""
        frame = QFrame()
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(0, 5, 0, 0)
        
        # Add stretch to push buttons to the right
        layout.addStretch()
        
        # Preview button
        self.preview_btn = QPushButton("👁 Preview Changes")
        self.preview_btn.setEnabled(False)
        self.preview_btn.setMinimumHeight(35)
        layout.addWidget(self.preview_btn)
        
        # Process button
        self.process_btn = QPushButton("▶ Process Appendices")
        self.process_btn.setEnabled(False)
        self.process_btn.setMinimumHeight(35)
        self.process_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        layout.addWidget(self.process_btn)
        
        return frame
    
    def setup_menus(self):
        """Set up the application menu bar."""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("&File")
        
        # New project
        new_action = QAction("&New Project", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_project)
        file_menu.addAction(new_action)
        
        # Open project
        open_action = QAction("&Open Project...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_project)
        file_menu.addAction(open_action)
        
        # Save project
        save_action = QAction("&Save Project", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_project)
        file_menu.addAction(save_action)
        
        file_menu.addSeparator()
        
        # Exit
        exit_action = QAction("E&xit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Edit menu
        edit_menu = menubar.addMenu("&Edit")
        
        # Settings
        settings_action = QAction("&Settings...", self)
        settings_action.triggered.connect(self.show_settings)
        edit_menu.addAction(settings_action)
        
        # Help menu
        help_menu = menubar.addMenu("&Help")
        
        # About
        about_action = QAction("&About", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def setup_status_bar(self):
        """Set up the status bar."""
        status_bar = self.statusBar()
        
        # Main status label
        self.status_label = QLabel("Ready")
        status_bar.addWidget(self.status_label)
        
        # Document status
        status_bar.addPermanentWidget(QLabel(" | "))
        self.doc_status_label = QLabel("No document")
        status_bar.addPermanentWidget(self.doc_status_label)
        
        # Appendix count
        status_bar.addPermanentWidget(QLabel(" | "))
        self.appendix_count_label = QLabel("0 appendices")
        status_bar.addPermanentWidget(self.appendix_count_label)
    
    def setup_connections(self):
        """Set up signal-slot connections."""
        # Document selector connections
        self.document_selector.document_selected.connect(self.on_document_selected)
        
        # Drag and drop connections
        self.drag_drop_area.files_dropped.connect(self.add_pdf_files)
        
        # Appendix list connections
        self.appendix_list.selection_changed.connect(self.on_appendix_selected)
        self.appendix_list.item_double_clicked.connect(self.edit_appendix)
        
        # Button connections
        self.move_up_btn.clicked.connect(self.move_appendix_up)
        self.move_down_btn.clicked.connect(self.move_appendix_down)
        self.remove_btn.clicked.connect(self.remove_appendix)
        self.preview_btn.clicked.connect(self.preview_changes)
        self.process_btn.clicked.connect(self.process_appendices)
        
        # Settings connections
        self.numbering_combo.currentTextChanged.connect(self.on_numbering_changed)
        self.backup_checkbox.toggled.connect(self.on_backup_setting_changed)
    
    def apply_settings(self):
        """Apply settings to the UI."""
        if not self.settings:
            return
        
        # Window geometry
        width, height, maximized = self.settings.window_geometry
        self.resize(width, height)
        if maximized:
            self.showMaximized()
        
        # Numbering style
        numbering_style = self.settings.appendix_numbering_style
        if numbering_style == "numeric":
            self.numbering_combo.setCurrentIndex(1)
        else:
            self.numbering_combo.setCurrentIndex(0)
        
        # Auto backup
        self.backup_checkbox.setChecked(self.settings.auto_backup_enabled)
    
    def browse_pdf_files(self):
        """Open file browser to select PDF files."""
        from PySide6.QtWidgets import QFileDialog
        
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PDF Files",
            self.settings.last_directory if self.settings else "",
            "PDF Files (*.pdf);;All Files (*)"
        )
        
        if file_paths:
            self.add_pdf_files(file_paths)
            
            # Update last directory
            if self.settings:
                self.settings.last_directory = str(Path(file_paths[0]).parent)
    
    def add_pdf_files(self, file_paths: list):
        """Add PDF files as appendices."""
        self.pdf_files_added.emit(file_paths)
        self.logger.info(f"Added {len(file_paths)} PDF files")
    
    def on_document_selected(self, document_info):
        """Handle document selection."""
        self.current_document = document_info
        self.document_selected.emit(document_info)
        
        # Update UI
        doc_name = document_info.get('name', 'Unknown')
        self.doc_info_label.setText(f"Selected: {doc_name}")
        self.doc_status_label.setText(doc_name)
        
        # Enable controls
        self.enable_controls(True)
        
        self.logger.info(f"Document selected: {doc_name}")
    
    def on_appendix_selected(self, index: int):
        """Handle appendix selection."""
        if 0 <= index < len(self.appendix_data):
            appendix = self.appendix_data[index]
            
            # Update preview
            self.pdf_preview.load_pdf(appendix.get('path', ''))
            
            # Update details
            self.update_appendix_details(appendix)
            
            # Enable/disable buttons
            self.move_up_btn.setEnabled(index > 0)
            self.move_down_btn.setEnabled(index < len(self.appendix_data) - 1)
            self.remove_btn.setEnabled(True)
        else:
            self.move_up_btn.setEnabled(False)
            self.move_down_btn.setEnabled(False)
            self.remove_btn.setEnabled(False)
    
    def update_appendix_details(self, appendix: Dict[str, Any]):
        """Update the details panel with appendix information."""
        details_html = f"""
        <h3>{appendix.get('title', 'Unknown Appendix')}</h3>
        <p><b>File:</b> {Path(appendix.get('path', '')).name}</p>
        <p><b>Pages:</b> {appendix.get('page_count', 0)}</p>
        <p><b>Size:</b> {appendix.get('size_mb', 0):.1f} MB</p>
        <p><b>Orientation:</b> {appendix.get('orientation', 'Unknown')}</p>
        """
        
        if appendix.get('warnings'):
            details_html += "<p><b>Warnings:</b></p><ul>"
            for warning in appendix['warnings']:
                details_html += f"<li style='color: orange;'>{warning}</li>"
            details_html += "</ul>"
        
        self.details_text.setHtml(details_html)
    
    def move_appendix_up(self):
        """Move selected appendix up in the list."""
        current_index = self.appendix_list.get_selected_index()
        if current_index > 0:
            # Swap items in data
            self.appendix_data[current_index], self.appendix_data[current_index - 1] = \
                self.appendix_data[current_index - 1], self.appendix_data[current_index]
            
            # Update UI
            self.refresh_appendix_list()
            self.appendix_list.select_item(current_index - 1)
            
            self.logger.info(f"Moved appendix from {current_index} to {current_index - 1}")
    
    def move_appendix_down(self):
        """Move selected appendix down in the list."""
        current_index = self.appendix_list.get_selected_index()
        if current_index < len(self.appendix_data) - 1:
            # Swap items in data
            self.appendix_data[current_index], self.appendix_data[current_index + 1] = \
                self.appendix_data[current_index + 1], self.appendix_data[current_index]
            
            # Update UI
            self.refresh_appendix_list()
            self.appendix_list.select_item(current_index + 1)
            
            self.logger.info(f"Moved appendix from {current_index} to {current_index + 1}")
    
    def remove_appendix(self):
        """Remove selected appendix from the list."""
        current_index = self.appendix_list.get_selected_index()
        if 0 <= current_index < len(self.appendix_data):
            appendix_name = self.appendix_data[current_index].get('title', 'Unknown')
            
            # Confirm removal
            reply = QMessageBox.question(
                self,
                "Remove Appendix",
                f"Are you sure you want to remove '{appendix_name}' from the appendices list?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                # Remove from data
                removed_appendix = self.appendix_data.pop(current_index)
                
                # Update UI
                self.refresh_appendix_list()
                self.update_appendix_count()
                
                # Clear preview and details if this was the selected item
                self.pdf_preview.clear()
                self.details_text.setHtml("<i>Select an appendix to view details</i>")
                
                self.logger.info(f"Removed appendix: {appendix_name}")
                self.appendix_removed.emit(current_index)
    
    def edit_appendix(self, index: int):
        """Edit appendix properties."""
        if 0 <= index < len(self.appendix_data):
            from gui.dialogs import AppendixEditDialog
            
            appendix = self.appendix_data[index]
            dialog = AppendixEditDialog(appendix, self)
            
            if dialog.exec() == dialog.Accepted:
                updated_appendix = dialog.get_appendix_data()
                self.appendix_data[index] = updated_appendix
                self.refresh_appendix_list()
                
                self.logger.info(f"Edited appendix: {updated_appendix.get('title', 'Unknown')}")
    
    def preview_changes(self):
        """Preview the changes that will be made to the document."""
        from gui.dialogs import PreviewDialog
        
        if not self.current_document or not self.appendix_data:
            QMessageBox.information(
                self,
                "Preview",
                "Please select a document and add some appendices first."
            )
            return
        
        dialog = PreviewDialog(self.current_document, self.appendix_data, self)
        dialog.exec()
    
    def process_appendices(self):
        """Process the appendices and add them to the document."""
        if not self.current_document:
            QMessageBox.warning(self, "No Document", "Please select a Word document first.")
            return
        
        if not self.appendix_data:
            QMessageBox.warning(self, "No Appendices", "Please add some PDF files as appendices first.")
            return
        
        # Confirm processing
        reply = QMessageBox.question(
            self,
            "Process Appendices",
            f"This will add {len(self.appendix_data)} appendices to the selected document.\n\n"
            f"{'A backup will be created automatically.' if self.backup_checkbox.isChecked() else 'No backup will be created.'}\n\n"
            "Do you want to continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Yes:
            self.process_requested.emit()
    
    def on_numbering_changed(self, text: str):
        """Handle numbering style change."""
        if self.settings:
            style = "numeric" if "Numeric" in text else "alphabetical"
            self.settings.set('document.appendix_numbering_style', style)
            self.refresh_appendix_list()
    
    def on_backup_setting_changed(self, enabled: bool):
        """Handle backup setting change."""
        if self.settings:
            self.settings.set('document.auto_backup', enabled)
    
    def refresh_appendix_list(self):
        """Refresh the appendix list display."""
        self.appendix_list.update_appendices(self.appendix_data, self.get_numbering_style())
        self.update_appendix_count()
        self.update_process_button()
    
    def update_appendix_count(self):
        """Update the appendix count in the status bar."""
        count = len(self.appendix_data)
        self.appendix_count_label.setText(f"{count} appendix{'es' if count != 1 else ''}")
    
    def update_process_button(self):
        """Update the state of the process button."""
        can_process = (
            self.current_document is not None and 
            len(self.appendix_data) > 0
        )
        self.process_btn.setEnabled(can_process)
        self.preview_btn.setEnabled(can_process)
    
    def get_numbering_style(self) -> str:
        """Get the current numbering style."""
        return "numeric" if self.numbering_combo.currentIndex() == 1 else "alphabetical"
    
    def enable_controls(self, enabled: bool):
        """Enable or disable main controls."""
        self.drag_drop_area.setEnabled(enabled)
        self.appendix_list.setEnabled(enabled)
        
        if not enabled:
            self.process_btn.setEnabled(False)
            self.preview_btn.setEnabled(False)
        else:
            self.update_process_button()
    
    def show_progress(self, message: str, progress: Optional[int] = None):
        """Show progress information."""
        self.progress_label.setText(message)
        self.progress_frame.setVisible(True)
        
        if progress is not None:
            self.progress_bar.setValue(progress)
            self.progress_bar.setVisible(True)
        else:
            self.progress_bar.setVisible(False)
        
        # Update status bar
        self.status_label.setText(message)
    
    def hide_progress(self):
        """Hide progress information."""
        self.progress_frame.setVisible(False)
        self.status_label.setText("Ready")
    
    def show_message(self, title: str, message: str, msg_type: str = "info"):
        """Show a message to the user."""
        if msg_type == "error":
            QMessageBox.critical(self, title, message)
        elif msg_type == "warning":
            QMessageBox.warning(self, title, message)
        else:
            QMessageBox.information(self, title, message)
    
    # Menu action handlers
    def new_project(self):
        """Create a new project."""
        reply = QMessageBox.question(
            self,
            "New Project",
            "This will clear the current project. Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.clear_project()
    
    def open_project(self):
        """Open an existing project file."""
        from PySide6.QtWidgets import QFileDialog
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Project",
            self.settings.last_directory if self.settings else "",
            "Word Appendix Project (*.wap);;All Files (*)"
        )
        
        if file_path:
            self.load_project(file_path)
    
    def save_project(self):
        """Save the current project."""
        from PySide6.QtWidgets import QFileDialog
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Project",
            self.settings.last_directory if self.settings else "",
            "Word Appendix Project (*.wap);;All Files (*)"
        )
        
        if file_path:
            self.save_project_to_file(file_path)
    
    def show_settings(self):
        """Show the settings dialog."""
        from gui.dialogs import SettingsDialog
        
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec() == dialog.Accepted:
            self.apply_settings()
            self.logger.info("Settings updated")
    
    def show_about(self):
        """Show the about dialog."""
        QMessageBox.about(
            self,
            "About Word Appendix Manager",
            "Word Appendix Manager v1.0\n\n"
            "A professional tool for adding PDF appendices to Word documents.\n\n"
            "Features:\n"
            "• Live Word document integration\n"
            "• Drag & drop PDF support\n"
            "• Automatic page numbering\n"
            "• Cross-platform compatibility\n\n"
            "Built with PySide6 and Python 3.8+"
        )
    
    def clear_project(self):
        """Clear the current project."""
        self.current_document = None
        self.appendix_data.clear()
        
        # Reset UI
        self.document_selector.clear_selection()
        self.appendix_list.update_appendices([], self.get_numbering_style())
        self.pdf_preview.clear()
        self.details_text.setHtml("<i>Select an appendix to view details</i>")
        self.doc_info_label.setText("No document selected")
        self.doc_status_label.setText("No document")
        
        # Disable controls
        self.enable_controls(False)
        self.update_appendix_count()
        
        self.logger.info("Project cleared")
    
    def load_project(self, file_path: str):
        """Load project from file."""
        try:
            import json
            with open(file_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            
            # Load project data
            self.current_document = project_data.get('document')
            self.appendix_data = project_data.get('appendices', [])
            
            # Update UI
            self.refresh_appendix_list()
            
            self.logger.info(f"Project loaded: {file_path}")
            
        except Exception as e:
            self.logger.error(f"Failed to load project: {e}")
            QMessageBox.critical(self, "Load Error", f"Failed to load project:\n{str(e)}")
    
    def save_project_to_file(self, file_path: str):
        """Save project to file."""
        try:
            import json
            project_data = {
                'version': '1.0',
                'document': self.current_document,
                'appendices': self.appendix_data,
                'settings': {
                    'numbering_style': self.get_numbering_style(),
                    'auto_backup': self.backup_checkbox.isChecked()
                }
            }
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, indent=2, ensure_ascii=False)
            
            self.logger.info(f"Project saved: {file_path}")
            
        except Exception as e:
            self.logger.error(f"Failed to save project: {e}")
            QMessageBox.critical(self, "Save Error", f"Failed to save project:\n{str(e)}")
    
    def closeEvent(self, event):
        """Handle window close event."""
        # Save window geometry
        if self.settings:
            self.settings.window_geometry = (
                self.width(),
                self.height(), 
                self.isMaximized()
            )
            self.settings.save_settings()
        
        self.logger.info("Main window closing")
        event.accept()
    
    # Public methods for external control (used by controller)
    def set_appendix_data(self, appendices: list):
        """Set the appendix data from external source."""
        self.appendix_data = appendices
        self.refresh_appendix_list()
    
    def get_appendix_data(self) -> list:
        """Get the current appendix data."""
        return self.appendix_data.copy()
    
    def get_current_document(self) -> Optional[Dict]:
        """Get the current document information."""
        return self.current_document
    
    def add_appendix(self, appendix_data: Dict[str, Any]):
        """Add a single appendix to the list."""
        self.appendix_data.append(appendix_data)
        self.refresh_appendix_list()