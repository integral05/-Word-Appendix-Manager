"""
Preview Dialog
Shows a preview of changes that will be made to the document.
"""

import sys
from pathlib import Path

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent.parent))

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTreeWidget, QTreeWidgetItem,
    QPushButton, QLabel, QGroupBox, QTextEdit, QSplitter, QFrame, QWidget
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont
from typing import Dict, Any, List

from utils.logger import get_logger, LoggerMixin


class PreviewDialog(QDialog, LoggerMixin):
    """Dialog for previewing document changes before processing."""
    
    def __init__(self, document_info: Dict[str, Any], appendices_data: List[Dict[str, Any]], parent=None):
        super().__init__(parent)
        self.document_info = document_info
        self.appendices_data = appendices_data
        
        self.setup_ui()
        self.populate_preview()
        
        self.logger.info("Opened preview changes dialog")
    
    def setup_ui(self):
        """Set up the dialog UI."""
        self.setWindowTitle("Preview Changes")
        self.setModal(True)
        self.resize(800, 600)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        
        # Header
        header_layout = QVBoxLayout()
        
        title_label = QLabel("Preview Document Changes")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c5aa0; margin-bottom: 5px;")
        header_layout.addWidget(title_label)
        
        subtitle_label = QLabel("Review the changes that will be made to your document")
        subtitle_label.setStyleSheet("color: #666; font-style: italic; margin-bottom: 10px;")
        header_layout.addWidget(subtitle_label)
        
        layout.addLayout(header_layout)
        
        # Main content splitter
        splitter = QSplitter(Qt.Horizontal)
        
        # Left panel - Document structure
        left_panel = self.create_structure_panel()
        splitter.addWidget(left_panel)
        
        # Right panel - Details and summary
        right_panel = self.create_details_panel()
        splitter.addWidget(right_panel)
        
        # Set initial sizes (60% left, 40% right)
        splitter.setSizes([480, 320])
        layout.addWidget(splitter)
        
        # Dialog buttons
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        self.close_btn = QPushButton("Close")
        self.close_btn.setMinimumWidth(100)
        buttons_layout.addWidget(self.close_btn)
        
        self.proceed_btn = QPushButton("Proceed with Changes")
        self.proceed_btn.setDefault(True)
        self.proceed_btn.setMinimumWidth(150)
        self.proceed_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        buttons_layout.addWidget(self.proceed_btn)
        
        layout.addLayout(buttons_layout)
        
        # Connect signals
        self.close_btn.clicked.connect(self.reject)
        self.proceed_btn.clicked.connect(self.accept)
    
    def create_structure_panel(self) -> QGroupBox:
        """Create the document structure preview panel."""
        group = QGroupBox("Document Structure Preview")
        layout = QVBoxLayout(group)
        
        # Description
        desc_label = QLabel("This shows how your document will be structured after processing:")
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(desc_label)
        
        # Tree widget to show document structure
        self.structure_tree = QTreeWidget()
        self.structure_tree.setHeaderLabel("Document Contents")
        self.structure_tree.setAlternatingRowColors(True)
        self.structure_tree.setRootIsDecorated(True)
        layout.addWidget(self.structure_tree)
        
        return group
    
    def create_details_panel(self) -> QWidget:
        """Create the details and summary panel."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Document info group
        doc_group = QGroupBox("Document Information")
        doc_layout = QVBoxLayout(doc_group)
        
        self.doc_info_text = QTextEdit()
        self.doc_info_text.setMaximumHeight(120)
        self.doc_info_text.setReadOnly(True)
        doc_layout.addWidget(self.doc_info_text)
        
        layout.addWidget(doc_group)
        
        # Processing summary group
        summary_group = QGroupBox("Processing Summary")
        summary_layout = QVBoxLayout(summary_group)
        
        self.summary_text = QTextEdit()
        self.summary_text.setMaximumHeight(150)
        self.summary_text.setReadOnly(True)
        summary_layout.addWidget(self.summary_text)
        
        layout.addWidget(summary_group)
        
        # Impact analysis group
        impact_group = QGroupBox("Impact Analysis")
        impact_layout = QVBoxLayout(impact_group)
        
        self.impact_text = QTextEdit()
        self.impact_text.setReadOnly(True)
        impact_layout.addWidget(self.impact_text)
        
        layout.addWidget(impact_group)
        
        return panel
    
    def populate_preview(self):
        """Populate the preview with document and appendix data."""
        self.populate_structure_tree()
        self.populate_document_info()
        self.populate_processing_summary()
        self.populate_impact_analysis()
    
    def populate_structure_tree(self):
        """Populate the document structure tree."""
        self.structure_tree.clear()
        
        # Main document item
        doc_name = self.document_info.get('name', 'Document')
        doc_item = QTreeWidgetItem([f"📄 {doc_name}"])
        doc_item.setExpanded(True)
        self.structure_tree.addTopLevelItem(doc_item)
        
        # Original content
        original_item = QTreeWidgetItem(["📝 Original Content"])
        original_item.setToolTip(0, "The existing content of your document")
        doc_item.addChild(original_item)
        
        # Add each appendix
        if self.appendices_data:
            appendices_item = QTreeWidgetItem(["📚 Appendices"])
            appendices_item.setExpanded(True)
            appendices_item.setToolTip(0, f"{len(self.appendices_data)} appendices will be added")
            doc_item.addChild(appendices_item)
            
            for i, appendix in enumerate(self.appendices_data):
                title = appendix.get('title', f'Appendix {i+1}')
                page_count = appendix.get('page_count', 0)
                file_name = Path(appendix.get('path', '')).name
                
                appendix_item = QTreeWidgetItem([f"📄 {title}"])
                appendix_item.setToolTip(0, f"Source: {file_name}\nPages: {page_count}")
                
                # Add page items for visual representation
                if page_count > 0:
                    if page_count <= 5:  # Show individual pages for small appendices
                        for page in range(page_count):
                            page_item = QTreeWidgetItem([f"📃 Page {page + 1}"])
                            appendix_item.addChild(page_item)
                    else:  # Show page range for large appendices
                        pages_item = QTreeWidgetItem([f"📃 {page_count} Pages"])
                        appendix_item.addChild(pages_item)
                
                appendices_item.addChild(appendix_item)
        
        # Expand all items initially
        self.structure_tree.expandAll()
    
    def populate_document_info(self):
        """Populate document information."""
        doc_name = self.document_info.get('name', 'Unknown')
        doc_path = self.document_info.get('path', 'Unknown')
        doc_source = self.document_info.get('source', 'unknown')
        
        doc_html = f"""
        <h3>📄 {doc_name}</h3>
        <p><b>Location:</b> {doc_path}</p>
        <p><b>Source:</b> {'Open Document' if doc_source != 'file' else 'File'}</p>
        """
        
        if doc_source != 'file':
            doc_html += "<p><b>Status:</b> Document is currently open in Word</p>"
        
        self.doc_info_text.setHtml(doc_html)
    
    def populate_processing_summary(self):
        """Populate processing summary."""
        appendix_count = len(self.appendices_data)
        total_pages = sum(a.get('page_count', 0) for a in self.appendices_data)
        total_size_mb = sum(a.get('size_mb', 0) for a in self.appendices_data)
        
        # Check for warnings
        warnings_count = sum(len(a.get('warnings', [])) for a in self.appendices_data)
        invalid_count = sum(1 for a in self.appendices_data if not a.get('valid', True))
        
        summary_html = f"""
        <h4>📊 Processing Overview</h4>
        <ul>
            <li><b>Appendices to add:</b> {appendix_count}</li>
            <li><b>Total pages to insert:</b> {total_pages}</li>
            <li><b>Total file size:</b> {total_size_mb:.1f} MB</li>
        </ul>
        """
        
        if warnings_count > 0 or invalid_count > 0:
            summary_html += f"""
            <h4>⚠️ Issues Detected</h4>
            <ul>
                <li style="color: orange;"><b>Warnings:</b> {warnings_count}</li>
                <li style="color: red;"><b>Invalid files:</b> {invalid_count}</li>
            </ul>
            """
        
        self.summary_text.setHtml(summary_html)
    
    def populate_impact_analysis(self):
        """Populate impact analysis."""
        total_pages = sum(a.get('page_count', 0) for a in self.appendices_data)
        
        impact_html = """
        <h4>🔍 What Will Happen</h4>
        <ol>
            <li><b>Backup Creation:</b> A backup of your original document will be created</li>
            <li><b>Page Insertion:</b> Blank pages with headings will be added to your Word document</li>
        """
        
        impact_html += f"<li><b>Content Replacement:</b> The {total_pages} blank pages will be replaced with PDF content</li>"
        impact_html += """
            <li><b>Final Export:</b> A combined PDF will be created with all appendices</li>
        </ol>
        
        <h4>📋 Processing Steps</h4>
        <ol>
        """
        
        for i, appendix in enumerate(self.appendices_data, 1):
            title = appendix.get('title', f'Appendix {i}')
            page_count = appendix.get('page_count', 0)
            impact_html += f"<li>Add <b>{title}</b> ({page_count} pages)</li>"
        
        impact_html += """
        </ol>
        
        <h4>⏱️ Estimated Time</h4>
        <p>Processing should take approximately <b>30-60 seconds</b> depending on file sizes.</p>
        
        <h4>💾 Output</h4>
        <ul>
            <li>Updated Word document with appendices</li>
            <li>Combined PDF file (optional)</li>
            <li>Backup of original document</li>
        </ul>
        """
        
        # Add warnings if any
        all_warnings = []
        for appendix in self.appendices_data:
            warnings = appendix.get('warnings', [])
            if warnings:
                title = appendix.get('title', 'Unknown')
                for warning in warnings:
                    all_warnings.append(f"{title}: {warning}")
        
        if all_warnings:
            impact_html += """
            <h4 style="color: orange;">⚠️ Warnings to Review</h4>
            <ul style="color: orange;">
            """
            for warning in all_warnings[:5]:  # Show first 5 warnings
                impact_html += f"<li>{warning}</li>"
            
            if len(all_warnings) > 5:
                impact_html += f"<li>... and {len(all_warnings) - 5} more warnings</li>"
            
            impact_html += "</ul>"
        
        self.impact_text.setHtml(impact_html)
    
    def get_user_decision(self) -> bool:
        """Get whether user wants to proceed."""
        return self.exec() == QDialog.Accepted
    
    def closeEvent(self, event):
        """Handle dialog close event."""
        self.logger.info("Closed preview changes dialog")
        event.accept() 
