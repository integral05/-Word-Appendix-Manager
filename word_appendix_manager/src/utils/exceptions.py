"""
Custom Exception Classes
Defines specific exceptions for different error scenarios in the application.
"""

from typing import Optional, Any


class ApplicationError(Exception):
    """Base exception class for all application errors."""
    
    def __init__(self, message: str, error_code: Optional[str] = None, details: Optional[dict] = None):
        super().__init__(message)
        self.message = message
        self.error_code = error_code or self.__class__.__name__
        self.details = details or {}
    
    def __str__(self):
        return f"[{self.error_code}] {self.message}"
    
    def to_dict(self) -> dict:
        """Convert exception to dictionary for logging/serialization."""
        return {
            "error_type": self.__class__.__name__,
            "error_code": self.error_code,
            "message": self.message,
            "details": self.details
        }


class WordDocumentError(ApplicationError):
    """Raised when there are issues with Word document operations."""
    
    def __init__(self, message: str, document_path: Optional[str] = None, operation: Optional[str] = None):
        details = {}
        if document_path:
            details["document_path"] = document_path
        if operation:
            details["operation"] = operation
        
        super().__init__(message, "WORD_ERROR", details)


class WordNotAvailableError(WordDocumentError):
    """Raised when Microsoft Word is not available or accessible."""
    
    def __init__(self, message: str = "Microsoft Word is not available or accessible"):
        super().__init__(message, operation="word_availability_check")


class WordDocumentNotFoundError(WordDocumentError):
    """Raised when a Word document cannot be found or opened."""
    
    def __init__(self, document_path: str, message: Optional[str] = None):
        if message is None:
            message = f"Word document not found or cannot be opened: {document_path}"
        super().__init__(message, document_path, "open_document")


class WordDocumentLockError(WordDocumentError):
    """Raised when a Word document is locked or in use by another process."""
    
    def __init__(self, document_path: str, message: Optional[str] = None):
        if message is None:
            message = f"Word document is locked or in use: {document_path}"
        super().__init__(message, document_path, "access_document")


class PDFError(ApplicationError):
    """Base class for PDF-related errors."""
    
    def __init__(self, message: str, pdf_path: Optional[str] = None, operation: Optional[str] = None):
        details = {}
        if pdf_path:
            details["pdf_path"] = pdf_path
        if operation:
            details["operation"] = operation
        
        super().__init__(message, "PDF_ERROR", details)


class PDFNotFoundError(PDFError):
    """Raised when a PDF file cannot be found."""
    
    def __init__(self, pdf_path: str):
        message = f"PDF file not found: {pdf_path}"
        super().__init__(message, pdf_path, "access_file")


class PDFCorruptedError(PDFError):
    """Raised when a PDF file is corrupted or unreadable."""
    
    def __init__(self, pdf_path: str, details: Optional[str] = None):
        message = f"PDF file is corrupted or unreadable: {pdf_path}"
        if details:
            message += f" ({details})"
        super().__init__(message, pdf_path, "read_file")


class PDFTooLargeError(PDFError):
    """Raised when a PDF file exceeds size limits."""
    
    def __init__(self, pdf_path: str, file_size_mb: float, max_size_mb: float):
        message = f"PDF file is too large: {file_size_mb:.1f}MB (max: {max_size_mb}MB)"
        details = {
            "file_size_mb": file_size_mb,
            "max_size_mb": max_size_mb
        }
        super().__init__(message, pdf_path, "size_check")
        self.details.update(details)


class PDFProcessingError(PDFError):
    """Raised when PDF processing operations fail."""
    
    def __init__(self, message: str, pdf_path: Optional[str] = None, operation: Optional[str] = None):
        super().__init__(message, pdf_path, operation)


class AppendixError(ApplicationError):
    """Raised when there are issues with appendix operations."""
    
    def __init__(self, message: str, appendix_name: Optional[str] = None, operation: Optional[str] = None):
        details = {}
        if appendix_name:
            details["appendix_name"] = appendix_name
        if operation:
            details["operation"] = operation
        
        super().__init__(message, "APPENDIX_ERROR", details)


class FileSystemError(ApplicationError):
    """Raised when there are file system operation errors."""
    
    def __init__(self, message: str, file_path: Optional[str] = None, operation: Optional[str] = None):
        details = {}
        if file_path:
            details["file_path"] = file_path
        if operation:
            details["operation"] = operation
        
        super().__init__(message, "FILESYSTEM_ERROR", details)


class BackupError(FileSystemError):
    """Raised when backup operations fail."""
    
    def __init__(self, message: str, source_path: Optional[str] = None, backup_path: Optional[str] = None):
        details = {}
        if source_path:
            details["source_path"] = source_path
        if backup_path:
            details["backup_path"] = backup_path
        
        super().__init__(message, operation="backup")
        self.details.update(details)


class ValidationError(ApplicationError):
    """Raised when data validation fails."""
    
    def __init__(self, message: str, field_name: Optional[str] = None, field_value: Optional[Any] = None):
        details = {}
        if field_name:
            details["field_name"] = field_name
        if field_value is not None:
            details["field_value"] = str(field_value)
        
        super().__init__(message, "VALIDATION_ERROR", details)


class UserCancelledError(ApplicationError):
    """Raised when user cancels an operation."""
    
    def __init__(self, message: str = "Operation cancelled by user", operation: Optional[str] = None):
        details = {}
        if operation:
            details["operation"] = operation
        
        super().__init__(message, "USER_CANCELLED", details)