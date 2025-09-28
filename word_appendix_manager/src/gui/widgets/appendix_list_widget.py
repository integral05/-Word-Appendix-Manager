"""
Appendix List Widget
Custom widget for displaying and managing the list of appendices.
"""

import sys
from pathlib import Path

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent.parent))

from PySide6.QtWidgets import (
    QListWidget, QListWidgetItem, QWidget, QVBoxLayout, 
    QHBoxLayout, QLabel, QFrame
)
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtGui import QFont
from typing import List, Dict, Any, Optional

from utils.logger import get_logger, LoggerMixin


class AppendixItemWidget(QWidget):
    """Custom widget for displaying appendix information in the list."""
    
    def __init__(self, appendix_data: Dict[str, Any], numbering: str, parent=None):
        super().__init__(parent)
        self.appendix_data = appendix_data
        self.numbering = numbering
        self.setup_ui()
    
    def setup_ui(self):
        """Set up the item widget UI."""
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(8)
        
        # Status indicator (colored dot)
        status_frame = QFrame()
        status_frame.setFixedSize(12, 12)
        status_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {self.get_status_color()};
                border-radius: 6px;
                border: 1px solid #ccc;
            }}
        """)
        layout.addWidget(status_frame)
        
        # Main content
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(2)
        
        # Title line
        title_layout = QHBoxLayout()
        title_layout.setContentsMargins(0, 0, 0, 0)
        
        # Appendix number/letter
        number_label = QLabel(f"{self.numbering}")
        number_font = QFont()
        number_font.setBold(True)
        number_font.setPointSize(10)
        number_label.setFont(number_font)
        number_label.setStyleSheet("color: #2c5aa0; font-weight: bold;")
        number_label.setMinimumWidth(30)
        title_layout.addWidget(number_label)
        
        # File name
        filename = Path(self.appendix_data.get('path', '')).name
        name_label = QLabel(filename)
        name_font = QFont()
        name_font.setPointSize(9)
        name_label.setFont(name_font)
        name_label.setToolTip(self.appendix_data.get('path', ''))
        title_layout.addWidget(name_label, 1)
        
        # Page count badge
        page_count = self.appendix_data.get('page_count', 0)
        pages_label = QLabel(f"{page_count} page{'s' if page_count != 1 else ''}")
        pages_label.setStyleSheet("""
            QLabel {
                background-color: #e3f2fd;
                border: 1px solid #bbdefb;
                border-radius: 10px;
                padding: 2px 8px;
                font-size: 8pt;
                color: #1976d2;
            }
        """)
        title_layout.addWidget(pages_label)
        
        content_layout.addLayout(title_layout)
        
        # Details line
        details_layout = QHBoxLayout()
        details_layout.setContentsMargins(30, 0, 0, 0)  # Indent to align with filename
        
        # Size info
        size_mb = self.appendix_data.get('size_mb', 0)
        size_label = QLabel(f"{size_mb:.1f} MB")
        size_label.setStyleSheet("color: #666; font-size: 8pt;")
        details_layout.addWidget(size_label)
        
        # Orientation
        orientation = self.appendix_data.get('orientation', 'Unknown')
        orientation_label = QLabel(f"• {orientation.title()}")
        orientation_label.setStyleSheet("color: #666; font-size: 8pt;")
        details_layout.addWidget(orientation_label)
        
        # Warnings indicator
        if self.appendix_data.get('warnings'):
            warning_label = QLabel(f"• ⚠ {len(self.appendix_data['warnings'])} warning(s)")
            warning_label.setStyleSheet("color: #f57c00; font-size: 8pt;")
            warning_label.setToolTip("\n".join(self.appendix_data['warnings']))
            details_layout.addWidget(warning_label)
        
        details_layout.addStretch()
        content_layout.addLayout(details_layout)
        
        layout.addLayout(content_layout, 1)
    
    def get_status_color(self) -> str:
        """Get the status color for the appendix."""
        if self.appendix_data.get('warnings'):
            return "#ff9800"  # Orange for warnings
        elif self.appendix_data.get('valid', True):
            return "#4caf50"  # Green for valid
        else:
            return "#f44336"  # Red for errors


class AppendixListWidget(QListWidget, LoggerMixin):
    """Custom list widget for displaying appendices."""
    
    # Signals
    selection_changed = Signal(int)  # Index of selected item
    item_double_clicked = Signal(int)  # Index of double-clicked item
    items_reordered = Signal(int, int)  # From index, to index
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.appendix_data = []
        self.numbering_style = "alphabetical"
        
        self.setup_ui()
        self.setup_connections()
    
    def setup_ui(self):
        """Set up the list widget UI."""
        self.setAlternatingRowColors(True)
        self.setSelectionMode(QListWidget.SingleSelection)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setDragDropMode(QListWidget.InternalMove)
        
        # Styling
        self.setStyleSheet("""
            QListWidget {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: white;
                outline: none;
            }
            QListWidget::item {
                border-bottom: 1px solid #f0f0f0;
                padding: 0px;
                margin: 0px;
            }
            QListWidget::item:selected {
                background-color: #e3f2fd;
                border-left: 3px solid #2196f3;
            }
            QListWidget::item:hover {
                background-color: #f5f5f5;
            }
        """)
        
        # Set minimum item height
        self.setMinimumHeight(200)
        
        # Empty state
        self.show_empty_state()
    
    def setup_connections(self):
        """Set up signal-slot connections."""
        self.itemSelectionChanged.connect(self.on_selection_changed)
        self.itemDoubleClicked.connect(self.on_item_double_clicked)
    
    def update_appendices(self, appendices: List[Dict[str, Any]], numbering_style: str = "alphabetical"):
        """Update the list with new appendix data."""
        self.appendix_data = appendices.copy()
        self.numbering_style = numbering_style
        
        # Clear existing items
        self.clear()
        
        if not appendices:
            self.show_empty_state()
            return
        
        # Add appendix items
        for i, appendix in enumerate(appendices):
            numbering = self.get_numbering(i)
            
            # Create list item
            item = QListWidgetItem()
            item.setSizeHint(QSize(0, 60))  # Set item height
            
            # Create custom widget
            item_widget = AppendixItemWidget(appendix, numbering, self)
            
            # Add to list
            self.addItem(item)
            self.setItemWidget(item, item_widget)
        
        self.logger.info(f"Updated appendix list: {len(appendices)} items")
    
    def get_numbering(self, index: int) -> str:
        """Get the numbering string for an appendix at the given index."""
        if self.numbering_style == "numeric":
            return f"{index + 1}."
        else:  # alphabetical
            if index < 26:
                return f"Appendix {chr(ord('A') + index)}"
            else:
                # For more than 26 appendices, use AA, AB, AC...
                first_letter = chr(ord('A') + (index // 26) - 1)
                second_letter = chr(ord('A') + (index % 26))
                return f"Appendix {first_letter}{second_letter}"
    
    def show_empty_state(self):
        """Show empty state message."""
        item = QListWidgetItem()
        item.setSizeHint(QSize(0, 80))
        
        empty_widget = QWidget()
        layout = QVBoxLayout(empty_widget)
        layout.setAlignment(Qt.AlignCenter)
        
        # Empty state icon/text
        empty_label = QLabel("📄 No appendices added yet")
        empty_label.setAlignment(Qt.AlignCenter)
        empty_label.setStyleSheet("""
            QLabel {
                color: #999;
                font-size: 14pt;
                font-style: italic;
                padding: 20px;
            }
        """)
        layout.addWidget(empty_label)
        
        hint_label = QLabel("Drag & drop PDF files here or use the Browse button")
        hint_label.setAlignment(Qt.AlignCenter)
        hint_label.setStyleSheet("""
            QLabel {
                color: #bbb;
                font-size: 10pt;
                font-style: italic;
            }
        """)
        layout.addWidget(hint_label)
        
        self.addItem(item)
        self.setItemWidget(item, empty_widget)
        
        # Make the empty state item non-selectable
        item.setFlags(item.flags() & ~Qt.ItemIsSelectable)
    
    def on_selection_changed(self):
        """Handle selection change."""
        current_row = self.currentRow()
        if current_row >= 0 and current_row < len(self.appendix_data):
            self.selection_changed.emit(current_row)
    
    def on_item_double_clicked(self, item: QListWidgetItem):
        """Handle item double click."""
        row = self.row(item)
        if row >= 0 and row < len(self.appendix_data):
            self.item_double_clicked.emit(row)
    
    def get_selected_index(self) -> int:
        """Get the index of the currently selected item."""
        return self.currentRow()
    
    def select_item(self, index: int):
        """Select an item by index."""
        if 0 <= index < self.count() and index < len(self.appendix_data):
            self.setCurrentRow(index)
    
    def get_appendix_count(self) -> int:
        """Get the number of appendices."""
        return len(self.appendix_data)
    
    def is_empty(self) -> bool:
        """Check if the list is empty."""
        return len(self.appendix_data) == 0
    
    def get_total_pages(self) -> int:
        """Get the total number of pages across all appendices."""
        return sum(appendix.get('page_count', 0) for appendix in self.appendix_data)
    
    def get_total_size_mb(self) -> float:
        """Get the total size in MB across all appendices."""
        return sum(appendix.get('size_mb', 0) for appendix in self.appendix_data)
    
    def dropEvent(self, event):
        """Handle drop events for reordering."""
        # Get source and destination indices
        source_index = self.currentRow()
        
        # Call parent implementation to handle the drop
        super().dropEvent(event)
        
        # Get new index after drop
        dest_index = self.currentRow()
        
        if source_index != dest_index and source_index >= 0 and dest_index >= 0:
            # Move item in our data structure
            if 0 <= source_index < len(self.appendix_data) and 0 <= dest_index < len(self.appendix_data):
                # Move the data
                item = self.appendix_data.pop(source_index)
                self.appendix_data.insert(dest_index, item)
                
                # Refresh the display
                self.update_appendices(self.appendix_data, self.numbering_style)
                self.select_item(dest_index)
                
                # Emit reorder signal
                self.items_reordered.emit(source_index, dest_index)
                
                self.logger.info(f"Reordered appendix from {source_index} to {dest_index}") 
