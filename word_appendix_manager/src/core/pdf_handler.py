"""
PDF Handler
Manages PDF operations including page counting, merging, and processing.
"""

import os
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple, BinaryIO
import tempfile
import shutil

# Add parent directory to path for imports
sys.path.append(str(Path(__file__).parent.parent))

from utils.logger import get_logger, LoggerMixin
from utils.exceptions import (
    PDFError, PDFNotFoundError, PDFCorruptedError, 
    FileSystemError
)

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    from pypdf import PdfReader, PdfWriter, PdfMerger
    PYPDF_AVAILABLE = True
except ImportError:
    try:
        from PyPDF2 import PdfReader, PdfWriter, PdfMerger
        PYPDF_AVAILABLE = True
    except ImportError:
        PYPDF_AVAILABLE = False


class PDFHandler(LoggerMixin):
    """Handles all PDF-related operations."""
    
    def __init__(self, settings=None):
        self.settings = settings
        self.temp_dir = None
        self._temp_files = []
        
        # Check available backends
        if not PYMUPDF_AVAILABLE and not PYPDF_AVAILABLE:
            raise PDFError("No PDF processing library available. Install PyMuPDF or pypdf.")
        
        self.logger.info(f"PDFHandler initialized - PyMuPDF: {PYMUPDF_AVAILABLE}, pypdf: {PYPDF_AVAILABLE}")
    
    def validate_pdf_file(self, pdf_path: str) -> Dict[str, Any]:
        """Validate a PDF file and return information about it."""
        pdf_path = Path(pdf_path)
        
        if not pdf_path.exists():
            raise PDFNotFoundError(str(pdf_path))
        
        try:
            # Check file size
            file_size = pdf_path.stat().st_size
            file_size_mb = file_size / (1024 * 1024)
            
            max_size_mb = self.settings.max_pdf_size_mb if self.settings else 200
            if file_size_mb > max_size_mb:
                self.logger.warning(f"PDF file exceeds size limit: {file_size_mb:.1f}MB > {max_size_mb}MB")
            
            # Get PDF information
            pdf_info = self._get_pdf_info(pdf_path)
            
            validation_result = {
                "path": str(pdf_path),
                "name": pdf_path.name,
                "size_bytes": file_size,
                "size_mb": file_size_mb,
                "valid": True,
                "error": None,
                **pdf_info
            }
            
            # Check for warnings
            warnings = []
            if file_size_mb > 50:
                warnings.append(f"Large file size: {file_size_mb:.1f}MB")
            if pdf_info.get("page_count", 0) > 1000:
                warnings.append(f"Many pages: {pdf_info['page_count']}")
            if pdf_info.get("encrypted", False):
                warnings.append("PDF is encrypted")
            
            validation_result["warnings"] = warnings
            
            self.logger.info(f"PDF validated: {pdf_path.name} ({pdf_info.get('page_count', 0)} pages)")
            return validation_result
            
        except (PDFNotFoundError):
            raise
        except Exception as e:
            self.logger.error(f"PDF validation failed for {pdf_path}: {e}")
            raise PDFCorruptedError(str(pdf_path), str(e))
    
    def _get_pdf_info(self, pdf_path: Path) -> Dict[str, Any]:
        """Get detailed information about a PDF file."""
        try:
            if PYMUPDF_AVAILABLE:
                return self._get_pdf_info_pymupdf(pdf_path)
            elif PYPDF_AVAILABLE:
                return self._get_pdf_info_pypdf(pdf_path)
            else:
                raise PDFError("No PDF library available")
        except Exception as e:
            raise PDFCorruptedError(str(pdf_path), str(e))
    
    def _get_pdf_info_pymupdf(self, pdf_path: Path) -> Dict[str, Any]:
        """Get PDF info using PyMuPDF."""
        doc = fitz.open(str(pdf_path))
        try:
            info = {
                "page_count": doc.page_count,
                "encrypted": doc.needs_pass,
                "title": doc.metadata.get("title", ""),
                "author": doc.metadata.get("author", ""),
                "subject": doc.metadata.get("subject", ""),
                "creator": doc.metadata.get("creator", ""),
                "producer": doc.metadata.get("producer", ""),
                "creation_date": doc.metadata.get("creationDate", ""),
                "modification_date": doc.metadata.get("modDate", "")
            }
            
            # Get page dimensions (first page)
            if doc.page_count > 0:
                page = doc[0]
                rect = page.rect
                info["page_width"] = rect.width
                info["page_height"] = rect.height
                info["orientation"] = "landscape" if rect.width > rect.height else "portrait"
            
            return info
        finally:
            doc.close()
    
    def _get_pdf_info_pypdf(self, pdf_path: Path) -> Dict[str, Any]:
        """Get PDF info using pypdf."""
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            
            info = {
                "page_count": len(reader.pages),
                "encrypted": reader.is_encrypted,
                "title": "",
                "author": "",
                "subject": "",
                "creator": "",
                "producer": "",
                "creation_date": "",
                "modification_date": ""
            }
            
            # Try to get metadata
            try:
                if reader.metadata:
                    info.update({
                        "title": reader.metadata.get("/Title", ""),
                        "author": reader.metadata.get("/Author", ""),
                        "subject": reader.metadata.get("/Subject", ""),
                        "creator": reader.metadata.get("/Creator", ""),
                        "producer": reader.metadata.get("/Producer", ""),
                        "creation_date": str(reader.metadata.get("/CreationDate", "")),
                        "modification_date": str(reader.metadata.get("/ModDate", ""))
                    })
            except:
                pass
            
            # Get page dimensions (first page)
            if len(reader.pages) > 0:
                page = reader.pages[0]
                mediabox = page.mediabox
                width = float(mediabox.width)
                height = float(mediabox.height)
                info["page_width"] = width
                info["page_height"] = height
                info["orientation"] = "landscape" if width > height else "portrait"
            
            return info
    
    def get_pdf_thumbnail(self, pdf_path: str, page_number: int = 0) -> Optional[bytes]:
        """Generate a thumbnail image of a PDF page."""
        if not PYMUPDF_AVAILABLE:
            self.logger.warning("Thumbnail generation requires PyMuPDF")
            return None
        
        try:
            doc = fitz.open(pdf_path)
            try:
                if page_number >= doc.page_count:
                    page_number = 0
                
                page = doc[page_number]
                
                # Create thumbnail with max dimension of 200px
                mat = fitz.Matrix(1, 1)
                pix = page.get_pixmap(matrix=mat)
                
                # Scale down if necessary
                max_dimension = 200
                if pix.width > max_dimension or pix.height > max_dimension:
                    scale = min(max_dimension / pix.width, max_dimension / pix.height)
                    mat = fitz.Matrix(scale, scale)
                    pix = page.get_pixmap(matrix=mat)
                
                # Convert to PNG bytes
                thumbnail_bytes = pix.tobytes("png")
                
                self.logger.debug(f"Generated thumbnail for {Path(pdf_path).name}")
                return thumbnail_bytes
                
            finally:
                doc.close()
                
        except Exception as e:
            self.logger.error(f"Failed to generate thumbnail for {pdf_path}: {e}")
            return None
    
    def merge_pdfs(self, pdf_list: List[str], output_path: str) -> bool:
        """Merge multiple PDF files into one."""
        if not pdf_list:
            raise PDFError("No PDF files to merge")
        
        try:
            if PYPDF_AVAILABLE:
                return self._merge_pdfs_pypdf(pdf_list, output_path)
            elif PYMUPDF_AVAILABLE:
                return self._merge_pdfs_pymupdf(pdf_list, output_path)
            else:
                raise PDFError("No PDF library available for merging")
        
        except Exception as e:
            self.logger.error(f"PDF merge failed: {e}")
            raise PDFError(f"Failed to merge PDFs: {str(e)}")
    
    def _merge_pdfs_pypdf(self, pdf_list: List[str], output_path: str) -> bool:
        """Merge PDFs using pypdf."""
        merger = PdfMerger()
        try:
            for pdf_path in pdf_list:
                if not Path(pdf_path).exists():
                    raise PDFNotFoundError(pdf_path)
                merger.append(pdf_path)
            
            with open(output_path, 'wb') as output_file:
                merger.write(output_file)
            
            self.logger.info(f"Merged {len(pdf_list)} PDFs to {output_path}")
            return True
            
        finally:
            merger.close()
    
    def _merge_pdfs_pymupdf(self, pdf_list: List[str], output_path: str) -> bool:
        """Merge PDFs using PyMuPDF."""
        output_doc = fitz.open()
        try:
            for pdf_path in pdf_list:
                if not Path(pdf_path).exists():
                    raise PDFNotFoundError(pdf_path)
                
                input_doc = fitz.open(pdf_path)
                try:
                    output_doc.insert_pdf(input_doc)
                finally:
                    input_doc.close()
            
            output_doc.save(output_path)
            self.logger.info(f"Merged {len(pdf_list)} PDFs to {output_path}")
            return True
            
        finally:
            output_doc.close()
    
    def create_temp_directory(self) -> str:
        """Create a temporary directory for PDF operations."""
        if not self.temp_dir:
            base_temp_dir = self.settings.get('advanced.temp_directory') if self.settings else None
            if base_temp_dir:
                Path(base_temp_dir).mkdir(parents=True, exist_ok=True)
            
            self.temp_dir = tempfile.mkdtemp(
                prefix="word_appendix_",
                dir=base_temp_dir
            )
            self.logger.debug(f"Created temp directory: {self.temp_dir}")
        
        return self.temp_dir
    
    def add_temp_file(self, file_path: str):
        """Track a temporary file for cleanup."""
        self._temp_files.append(file_path)
    
    def cleanup_temp_files(self):
        """Clean up temporary files and directories."""
        try:
            # Clean up tracked temp files
            for temp_file in self._temp_files:
                try:
                    if Path(temp_file).exists():
                        os.remove(temp_file)
                except Exception as e:
                    self.logger.warning(f"Could not remove temp file {temp_file}: {e}")
            
            # Clean up temp directory
            if self.temp_dir and Path(self.temp_dir).exists():
                shutil.rmtree(self.temp_dir, ignore_errors=True)
                self.logger.debug(f"Cleaned up temp directory: {self.temp_dir}")
            
            self._temp_files.clear()
            self.temp_dir = None
            
        except Exception as e:
            self.logger.error(f"Error during temp file cleanup: {e}")
    
    def get_system_info(self) -> Dict[str, Any]:
        """Get information about available PDF processing capabilities."""
        return {
            "pymupdf_available": PYMUPDF_AVAILABLE,
            "pypdf_available": PYPDF_AVAILABLE,
            "thumbnail_support": PYMUPDF_AVAILABLE,
            "preferred_backend": "PyMuPDF" if PYMUPDF_AVAILABLE else "pypdf" if PYPDF_AVAILABLE else "None"
        }
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup_temp_files() 
