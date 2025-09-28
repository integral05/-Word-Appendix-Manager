"""
Application Constants
Defines constants used throughout the Word Appendix Manager application.
"""

# Application Information
APP_NAME = "Word Appendix Manager"
APP_VERSION = "1.0.0"
APP_AUTHOR = "PointOne Development Team"
APP_DESCRIPTION = "Professional tool for adding PDF appendices to Word documents"

# File Extensions
SUPPORTED_WORD_EXTENSIONS = ['.docx', '.doc']
SUPPORTED_PDF_EXTENSIONS = ['.pdf']
PROJECT_EXTENSION = '.wap'

# File Size Limits (in MB)
DEFAULT_MAX_PDF_SIZE_MB = 200
DEFAULT_MAX_PAGES_WARNING = 2000

# UI Constants
DEFAULT_WINDOW_WIDTH = 1200
DEFAULT_WINDOW_HEIGHT = 800
MIN_WINDOW_WIDTH = 1000
MIN_WINDOW_HEIGHT = 700

# Processing Constants
MAX_APPENDICES_LIMIT = 50
DEFAULT_THUMBNAIL_SIZE = 200

# Word Document Settings
DEFAULT_HEADING_FONT = "Arial"
DEFAULT_HEADING_SIZE = 14
DEFAULT_HEADING_BOLD = True
DEFAULT_HEADING_ITALIC = False

# Colors
COLORS = {
    'primary': '#2196f3',
    'success': '#4caf50',
    'warning': '#ff9800',
    'error': '#f44336',
    'background': '#f8f9fa',
    'text_primary': '#495057'
}

# File Dialog Filters
FILE_DIALOGS = {
    'word_documents': "Word Documents (*.docx *.doc);;All Files (*)",
    'pdf_files': "PDF Files (*.pdf);;All Files (*)",
    'project_files': "Word Appendix Project (*.wap);;All Files (*)"
}