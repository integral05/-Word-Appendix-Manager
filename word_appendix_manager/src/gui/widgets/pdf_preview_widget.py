"""
PDF Preview Widget
Provides a preview of PDF files with thumbnail display.
"""

import sys
from pathlib import Path

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent.parent))

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QPushButton, QScrollArea, QFrame, QSpinBox
)
from PySide6.QtCore import Qt, Signal, QThread
from PySide6.QtGui import QPixmap, QFont
from typing import Optional

from utils.logger import get_logger, LoggerMixin


class ThumbnailWorker(QThread):
    """Worker thread for generating PDF thumbnails."""
    
    thumbnail_ready = Signal(bytes)  # Thumbnail image data
    thumbnail_failed = Signal(str)   # Error message
    
    def __init__(self, pdf_path: str, page_number: int = 0):
        super().__init__()
        self.pdf_path = pdf_path
        self.page_number = page_number
    
    def run(self):
        """Generate thumbnail in background thread."""
        try:
            from core.pdf_handler import PDFHandler
            
            pdf_handler = PDFHandler()
            thumbnail_data = pdf_handler.get_pdf_thumbnail(self.pdf_path, self.page_number)
            
            if thumbnail_data:
                self.thumbnail_ready.emit(thumbnail_data)
            else:
                self.thumbnail_failed.emit("Could not generate thumbnail")
                
        except Exception as e:
            self.thumbnail_failed.emit(str(e))


class PDFPreviewWidget(QWidget, LoggerMixin):
    """Widget for previewing PDF files."""
    
    # Signals
    page_changed = Signal(int)  # New page number
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_pdf_path = None
        self.current_page = 0
        self.total_pages = 0
        self.thumbnail_worker = None
        
        self.setup_ui()
        self.show_empty_state()
    
    def setup_ui(self):
        """Set up the preview widget UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)
        
        # Header with file info
        self.header_frame = QFrame()
        self.header_frame.setFrameStyle(QFrame.Box)
        self.header_frame.setStyleSheet("""
            QFrame {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        header_layout = QVBoxLayout(self.header_frame)
        header_layout.setContentsMargins(8, 4, 8, 4)
        
        # File name
        self.filename_label = QLabel("No PDF selected")
        font = QFont()
        font.setBold(True)
        self.filename_label.setFont(font)
        self.filename_label.setWordWrap(True)
        header_layout.addWidget(self.filename_label)
        
        # File details
        self.details_label = QLabel("")
        self.details_label.setStyleSheet("color: #666; font-size: 11px;")
        header_layout.addWidget(self.details_label)
        
        layout.addWidget(self.header_frame)
        
        # Page navigation
        self.nav_frame = QFrame()
        nav_layout = QHBoxLayout(self.nav_frame)
        nav_layout.setContentsMargins(0, 0, 0, 0)
        
        self.prev_btn = QPushButton("◀")
        self.prev_btn.setMaximumWidth(30)
        self.prev_btn.setEnabled(False)
        self.prev_btn.clicked.connect(self.previous_page)
        nav_layout.addWidget(self.prev_btn)
        
        self.page_label = QLabel("Page:")
        nav_layout.addWidget(self.page_label)
        
        self.page_spinbox = QSpinBox()
        self.page_spinbox.setMinimum(1)
        self.page_spinbox.setMaximum(1)
        self.page_spinbox.setEnabled(False)
        self.page_spinbox.valueChanged.connect(self.on_page_changed)
        nav_layout.addWidget(self.page_spinbox)
        
        self.total_pages_label = QLabel("of 0")
        nav_layout.addWidget(self.total_pages_label)
        
        self.next_btn = QPushButton("▶")
        self.next_btn.setMaximumWidth(30)
        self.next_btn.setEnabled(False)
        self.next_btn.clicked.connect(self.next_page)
        nav_layout.addWidget(self.next_btn)
        
        nav_layout.addStretch()
        
        layout.addWidget(self.nav_frame)
        
        # Preview area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setAlignment(Qt.AlignCenter)
        self.scroll_area.setStyleSheet("""
            QScrollArea {
                border: 1px solid #ccc;
                background-color: white;
            }
        """)
        
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumSize(200, 250)
        self.preview_label.setStyleSheet("""
            QLabel {
                background-color: white;
                border: none;
                margin: 10px;
            }
        """)
        
        self.scroll_area.setWidget(self.preview_label)
        layout.addWidget(self.scroll_area)
        
        # Initially hide navigation controls
        self.nav_frame.setVisible(False)
    
    def show_empty_state(self):
        """Show empty state when no PDF is loaded."""
        self.preview_label.clear()
        self.preview_label.setText("📄\n\nNo PDF selected\n\nSelect an appendix to preview")
        self.preview_label.setStyleSheet("""
            QLabel {
                color: #999;
                font-size: 14pt;
                background-color: #fafafa;
                border: 2px dashed #ddd;
                border-radius: 8px;
                margin: 20px;
                padding: 40px;
            }
        """)
        
        self.filename_label.setText("No PDF selected")
        self.details_label.setText("")
        self.nav_frame.setVisible(False)
        
        self.current_pdf_path = None
        self.current_page = 0
        self.total_pages = 0
    
    def load_pdf(self, pdf_path: str, page_number: int = 0):
        """Load a PDF file for preview."""
        if not pdf_path or not Path(pdf_path).exists():
            self.show_empty_state()
            return
        
        self.current_pdf_path = pdf_path
        self.current_page = page_number
        
        # Update header info
        pdf_file = Path(pdf_path)
        self.filename_label.setText(pdf_file.name)
        
        # Get PDF info
        try:
            from core.pdf_handler import PDFHandler
            pdf_handler = PDFHandler()
            pdf_info = pdf_handler.validate_pdf_file(pdf_path)
            
            self.total_pages = pdf_info.get('page_count', 0)
            size_mb = pdf_info.get('size_mb', 0)
            orientation = pdf_info.get('orientation', 'Unknown')
            
            self.details_label.setText(
                f"{self.total_pages} pages • {size_mb:.1f} MB • {orientation.title()}"
            )
            
            # Update navigation controls
            if self.total_pages > 1:
                self.nav_frame.setVisible(True)
                self.page_spinbox.setMaximum(self.total_pages)
                self.page_spinbox.setValue(page_number + 1)
                self.total_pages_label.setText(f"of {self.total_pages}")
                self.update_navigation_buttons()
            else:
                self.nav_frame.setVisible(False)
            
            # Load thumbnail
            self.load_thumbnail()
            
            self.logger.info(f"Loaded PDF preview: {pdf_file.name}")
            
        except Exception as e:
            self.logger.error(f"Failed to load PDF info: {e}")
            self.details_label.setText(f"Error: {str(e)}")
            self.show_error_state(str(e))
    
    def load_thumbnail(self):
        """Load thumbnail for current page."""
        if not self.current_pdf_path:
            return
        
        # Show loading state
        self.preview_label.clear()
        self.preview_label.setText("Loading preview...")
        self.preview_label.setStyleSheet("""
            QLabel {
                color: #666;
                font-size: 12pt;
                background-color: white;
                border: 1px solid #eee;
                border-radius: 4px;
                margin: 10px;
                padding: 20px;
            }
        """)
        
        # Cancel previous thumbnail generation
        if self.thumbnail_worker and self.thumbnail_worker.isRunning():
            self.thumbnail_worker.terminate()
            self.thumbnail_worker.wait()
        
        # Start thumbnail generation
        self.thumbnail_worker = ThumbnailWorker(self.current_pdf_path, self.current_page)
        self.thumbnail_worker.thumbnail_ready.connect(self.on_thumbnail_ready)
        self.thumbnail_worker.thumbnail_failed.connect(self.on_thumbnail_failed)
        self.thumbnail_worker.start()
    
    def on_thumbnail_ready(self, thumbnail_data: bytes):
        """Handle thumbnail generation completion."""
        try:
            # Create pixmap from thumbnail data
            pixmap = QPixmap()
            pixmap.loadFromData(thumbnail_data)
            
            if not pixmap.isNull():
                # Compute available size (subtract margins)
                margins = self.scroll_area.contentsMargins()
                preview_size = self.scroll_area.size()
                available_width = preview_size.width() - (margins.left() + margins.right())
                available_height = preview_size.height() - (margins.top() + margins.bottom())
                
                # Scale pixmap to fit preview area while maintaining aspect ratio
                scaled_pixmap = pixmap.scaled(
                    max(1, available_width - 40),
                    max(1, available_height - 40),
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                )
                
                self.preview_label.setPixmap(scaled_pixmap)
                self.preview_label.setStyleSheet("""
                    QLabel {
                        background-color: white;
                        border: 1px solid #ddd;
                        border-radius: 4px;
                        margin: 10px;
                        padding: 10px;
                    }
                """)
            else:
                self.show_error_state("Invalid thumbnail data")
                
        except Exception as e:
            self.logger.error(f"Failed to display thumbnail: {e}")
            self.show_error_state(str(e))

    
    def on_thumbnail_failed(self, error_message: str):
        """Handle thumbnail generation failure."""
        self.logger.error(f"Thumbnail generation failed: {error_message}")
        self.show_error_state(error_message)
    
    def show_error_state(self, error_message: str):
        """Show error state in preview."""
        self.preview_label.clear()
        self.preview_label.setText(f"❌\n\nPreview Error\n\n{error_message}")
        self.preview_label.setStyleSheet("""
            QLabel {
                color: #d32f2f;
                font-size: 12pt;
                background-color: #ffebee;
                border: 2px solid #f44336;
                border-radius: 8px;
                margin: 20px;
                padding: 30px;
            }
        """)
    
    def previous_page(self):
        """Go to previous page."""
        if self.current_page > 0:
            self.current_page -= 1
            self.page_spinbox.setValue(self.current_page + 1)
            self.load_thumbnail()
            self.update_navigation_buttons()
            self.page_changed.emit(self.current_page)
    
    def next_page(self):
        """Go to next page."""
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.page_spinbox.setValue(self.current_page + 1)
            self.load_thumbnail()
            self.update_navigation_buttons()
            self.page_changed.emit(self.current_page)
    
    def on_page_changed(self, page_number: int):
        """Handle page number change from spinbox."""
        new_page = page_number - 1  # Convert to 0-indexed
        if new_page != self.current_page and 0 <= new_page < self.total_pages:
            self.current_page = new_page
            self.load_thumbnail()
            self.update_navigation_buttons()
            self.page_changed.emit(self.current_page)
    
    def update_navigation_buttons(self):
        """Update the state of navigation buttons."""
        self.prev_btn.setEnabled(self.current_page > 0)
        self.next_btn.setEnabled(self.current_page < self.total_pages - 1)
        self.page_spinbox.setEnabled(self.total_pages > 0)
    
    def clear(self):
        """Clear the preview."""
        self.show_empty_state()
    
    def get_current_page(self) -> int:
        """Get the current page number (0-indexed)."""
        return self.current_page
    
    def get_total_pages(self) -> int:
        """Get the total number of pages."""
        return self.total_pages
    
    def get_current_pdf_path(self) -> Optional[str]:
        """Get the path of the currently loaded PDF."""
        return self.current_pdf_path 
