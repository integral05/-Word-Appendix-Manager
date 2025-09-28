"""
Drag and Drop Area Widget
Provides a user-friendly drag and drop interface for PDF files.
"""

import sys
from pathlib import Path

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent.parent))

from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QFrame
from PySide6.QtCore import Qt, Signal, QMimeData, QUrl, QTimer
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QPaintEvent, QPainter
from typing import List

from utils.logger import get_logger, LoggerMixin


class DragDropArea(QFrame, LoggerMixin):
    """Drag and drop area for PDF files."""
    
    # Signals
    files_dropped = Signal(list)  # List of file paths
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.is_dragging = False
        self.setup_ui()
    
    def setup_ui(self):
        """Set up the drag and drop area UI."""
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.Box)
        self.setLineWidth(2)
        
        # Layout
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(10)
        
        # Icon/visual indicator
        icon_label = QLabel("📁")
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet("""
            QLabel {
                font-size: 48px;
                color: #666;
                margin: 10px;
            }
        """)
        layout.addWidget(icon_label)
        
        # Primary text
        primary_label = QLabel("Drop PDF files here")
        primary_label.setAlignment(Qt.AlignCenter)
        primary_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                font-weight: bold;
                color: #333;
                margin: 5px;
            }
        """)
        layout.addWidget(primary_label)
        
        # Secondary text
        secondary_label = QLabel("or click Browse to select files")
        secondary_label.setAlignment(Qt.AlignCenter)
        secondary_label.setStyleSheet("""
            QLabel {
                font-size: 12px;
                color: #666;
                font-style: italic;
                margin: 5px;
            }
        """)
        layout.addWidget(secondary_label)
        
        # Supported formats text
        formats_label = QLabel("Supported: PDF files only")
        formats_label.setAlignment(Qt.AlignCenter)
        formats_label.setStyleSheet("""
            QLabel {
                font-size: 10px;
                color: #999;
                margin: 5px;
            }
        """)
        layout.addWidget(formats_label)
        
        # Set initial styling
        self.set_normal_style()
    
    def set_normal_style(self):
        """Set normal (non-dragging) style."""
        self.setStyleSheet("""
            QFrame {
                border: 2px dashed #ccc;
                border-radius: 8px;
                background-color: #fafafa;
            }
            QFrame:hover {
                border-color: #2196f3;
                background-color: #f0f8ff;
            }
        """)
    
    def set_drag_style(self):
        """Set dragging style."""
        self.setStyleSheet("""
            QFrame {
                border: 2px dashed #2196f3;
                border-radius: 8px;
                background-color: #e3f2fd;
                color: #1976d2;
            }
        """)
    
    def set_error_style(self):
        """Set error style for invalid drops."""
        self.setStyleSheet("""
            QFrame {
                border: 2px dashed #f44336;
                border-radius: 8px;
                background-color: #ffebee;
                color: #d32f2f;
            }
        """)
        
        # Reset to normal after a short delay
        QTimer.singleShot(1000, self.set_normal_style)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter event."""
        if self.has_valid_urls(event.mimeData()):
            self.is_dragging = True
            self.set_drag_style()
            event.acceptProposedAction()
            self.logger.debug("Drag enter: valid files detected")
        else:
            event.ignore()
            self.logger.debug("Drag enter: no valid files")
    
    def dragMoveEvent(self, event):
        """Handle drag move event."""
        if self.has_valid_urls(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def dragLeaveEvent(self, event):
        """Handle drag leave event."""
        self.is_dragging = False
        self.set_normal_style()
        self.logger.debug("Drag leave")
    
    def dropEvent(self, event: QDropEvent):
        """Handle drop event."""
        self.is_dragging = False
        
        mime_data = event.mimeData()
        if mime_data.hasUrls():
            file_paths = self.extract_pdf_paths(mime_data)
            
            if file_paths:
                self.set_normal_style()
                self.files_dropped.emit(file_paths)
                event.acceptProposedAction()
                self.logger.info(f"Files dropped: {len(file_paths)} PDF files")
            else:
                self.set_error_style()
                event.ignore()
                self.logger.warning("Drop event: no valid PDF files found")
        else:
            self.set_error_style()
            event.ignore()
            self.logger.warning("Drop event: no URLs in mime data")
    
    def has_valid_urls(self, mime_data: QMimeData) -> bool:
        """Check if mime data contains valid file URLs."""
        if not mime_data.hasUrls():
            return False
        
        urls = mime_data.urls()
        for url in urls:
            if url.isLocalFile():
                file_path = Path(url.toLocalFile())
                if file_path.is_file() and file_path.suffix.lower() == '.pdf':
                    return True
        
        return False
    
    def extract_pdf_paths(self, mime_data: QMimeData) -> List[str]:
        """Extract PDF file paths from mime data."""
        pdf_paths = []
        
        if mime_data.hasUrls():
            for url in mime_data.urls():
                if url.isLocalFile():
                    file_path = Path(url.toLocalFile())
                    if file_path.is_file() and file_path.suffix.lower() == '.pdf':
                        pdf_paths.append(str(file_path))
        
        return pdf_paths
    
    def mousePressEvent(self, event):
        """Handle mouse press for click-to-browse functionality."""
        if event.button() == Qt.LeftButton:
            self.logger.debug("Drag drop area clicked")
        super().mousePressEvent(event)
    
    def setEnabled(self, enabled: bool):
        """Enable or disable the drag drop area."""
        super().setEnabled(enabled)
        
        if not enabled:
            self.setStyleSheet("""
                QFrame {
                    border: 2px dashed #ddd;
                    border-radius: 8px;
                    background-color: #f5f5f5;
                    color: #bbb;
                }
            """)
            self.setAcceptDrops(False)
        else:
            self.set_normal_style()
            self.setAcceptDrops(True)
    
    def show_message(self, message: str, message_type: str = "info"):
        """Show a temporary message on the drag drop area."""
        # Create a temporary label to show messages
        if not hasattr(self, 'message_label'):
            self.message_label = QLabel(self)
            self.message_label.setAlignment(Qt.AlignCenter)
            self.message_label.hide()
        
        # Style based on message type
        if message_type == "error":
            color = "#d32f2f"
            bg_color = "rgba(255, 235, 238, 220)"
        elif message_type == "warning":
            color = "#f57c00"
            bg_color = "rgba(255, 248, 225, 220)"
        elif message_type == "success":
            color = "#388e3c"
            bg_color = "rgba(232, 245, 233, 220)"
        else:  # info
            color = "#1976d2"
            bg_color = "rgba(227, 242, 253, 220)"
        
        self.message_label.setStyleSheet(f"""
            QLabel {{
                background-color: {bg_color};
                border: 1px solid {color};
                border-radius: 4px;
                padding: 8px;
                font-weight: bold;
                color: {color};
            }}
        """)
        
        # Set message and show
        self.message_label.setText(message)
        self.message_label.adjustSize()
        
        # Center the label
        label_x = (self.width() - self.message_label.width()) // 2
        label_y = (self.height() - self.message_label.height()) // 2
        self.message_label.move(label_x, label_y)
        self.message_label.show()
        
        # Hide after 3 seconds
        QTimer.singleShot(3000, self.message_label.hide)