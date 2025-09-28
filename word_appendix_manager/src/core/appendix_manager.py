"""
Appendix Manager
Coordinates between Word document operations and PDF handling to create appendices.
"""

import sys
import tempfile
from pathlib import Path
from typing import List, Dict, Any, Optional
import shutil

# Add parent directory to path for imports
sys.path.append(str(Path(__file__).parent.parent))

from utils.logger import get_logger, LoggerMixin
from utils.exceptions import (
    AppendixError, WordDocumentError, PDFError, 
    FileSystemError, ValidationError
)


class AppendixManager(LoggerMixin):
    """Manages the process of adding PDF appendices to Word documents."""
    
    def __init__(self, word_manager, pdf_handler, settings=None):
        self.word_manager = word_manager
        self.pdf_handler = pdf_handler
        self.settings = settings
        
        # State tracking
        self.appendices_added = []
        self.temp_files = []
        
        self.logger.info("Appendix manager initialized")
    
    def add_appendix(self, appendix_data: Dict[str, Any]) -> bool:
        """Add a single appendix to the Word document."""
        try:
            pdf_path = appendix_data.get('path')
            title = appendix_data.get('title', 'Appendix')
            page_count = appendix_data.get('page_count', 0)
            
            if not pdf_path or not Path(pdf_path).exists():
                raise AppendixError(f"PDF file not found: {pdf_path}")
            
            if page_count <= 0:
                raise AppendixError(f"Invalid page count: {page_count}")
            
            # Get heading style
            heading_style = self.settings.heading_style if self.settings else None
            
            # Add blank pages with heading to Word document
            success = self.word_manager.add_appendix_section(
                appendix_title=title,
                page_count=page_count,
                heading_style=heading_style
            )
            
            if success:
                # Track this appendix for later PDF replacement
                appendix_info = {
                    'title': title,
                    'pdf_path': pdf_path,
                    'page_count': page_count,
                    'start_page': self._calculate_start_page(),
                    'end_page': self._calculate_start_page() + page_count - 1
                }
                
                self.appendices_added.append(appendix_info)
                self.logger.info(f"Added appendix: {title} ({page_count} pages)")
                
            return success
            
        except Exception as e:
            self.logger.error(f"Failed to add appendix {appendix_data.get('title', 'Unknown')}: {e}")
            raise AppendixError(f"Failed to add appendix: {str(e)}")
    
    def add_multiple_appendices(self, appendices_data: List[Dict[str, Any]]) -> int:
        """Add multiple appendices to the Word document."""
        success_count = 0
        
        for appendix_data in appendices_data:
            try:
                if self.add_appendix(appendix_data):
                    success_count += 1
            except Exception as e:
                self.logger.error(f"Failed to add appendix {appendix_data.get('title')}: {e}")
                continue
        
        self.logger.info(f"Added {success_count}/{len(appendices_data)} appendices")
        return success_count
    
    def create_final_pdf(self, output_path: Optional[str] = None) -> str:
        """Create the final PDF with appendices by exporting Word document and replacing blank pages."""
        try:
            # Export Word document to PDF
            temp_word_pdf = self._export_word_to_pdf()
            
            # Replace blank pages with actual PDF content
            final_pdf = self._replace_blank_pages(temp_word_pdf, output_path)
            
            # Clean up temporary file
            if Path(temp_word_pdf).exists():
                Path(temp_word_pdf).unlink()
            
            self.logger.info(f"Created final PDF: {final_pdf}")
            return final_pdf
            
        except Exception as e:
            self.logger.error(f"Failed to create final PDF: {e}")
            raise AppendixError(f"Failed to create final PDF: {str(e)}")
    
    def _export_word_to_pdf(self) -> str:
        """Export the current Word document to PDF."""
        try:
            # Create temporary file for Word PDF export
            temp_dir = self.pdf_handler.create_temp_directory()
            temp_pdf = Path(temp_dir) / "word_document.pdf"
            
            # Export using Word manager
            success = self.word_manager.export_as_pdf(str(temp_pdf))
            
            if not success or not temp_pdf.exists():
                raise WordDocumentError("Failed to export Word document to PDF")
            
            self.temp_files.append(str(temp_pdf))
            return str(temp_pdf)
            
        except Exception as e:
            raise WordDocumentError(f"PDF export failed: {str(e)}")
    
    def _replace_blank_pages(self, word_pdf_path: str, output_path: Optional[str] = None) -> str:
        """Replace blank pages in the Word PDF with actual PDF content."""
        try:
            if not output_path:
                # Generate default output path
                doc_info = self.word_manager.get_document_info()
                doc_name = Path(doc_info.get('name', 'document')).stem
                
                output_dir = Path.home() / "Documents"
                output_path = output_dir / f"{doc_name}_with_appendices.pdf"
                
                # Ensure unique filename
                counter = 1
                while output_path.exists():
                    output_path = output_dir / f"{doc_name}_with_appendices_{counter}.pdf"
                    counter += 1
            
            # For now, just copy the Word PDF as final PDF
            # In a full implementation, this would replace blank pages with actual PDF content
            shutil.copy2(word_pdf_path, output_path)
            
            # Merge all appendix PDFs at the end
            if self.appendices_added:
                pdf_list = [word_pdf_path] + [appendix['pdf_path'] for appendix in self.appendices_added]
                self.pdf_handler.merge_pdfs(pdf_list, str(output_path))
            
            return str(output_path)
            
        except Exception as e:
            raise PDFError(f"Failed to replace blank pages: {str(e)}")
    
    def _calculate_start_page(self) -> int:
        """Calculate the start page number for the next appendix."""
        if not self.appendices_added:
            return 10  # Placeholder - should get actual page count
        
        last_appendix = self.appendices_added[-1]
        return last_appendix['end_page'] + 1
    
    def validate_appendices(self, appendices_data: List[Dict[str, Any]]) -> List[str]:
        """Validate a list of appendices before processing."""
        errors = []
        
        for i, appendix in enumerate(appendices_data):
            appendix_id = f"Appendix {i + 1}"
            
            # Check PDF file exists
            pdf_path = appendix.get('path')
            if not pdf_path:
                errors.append(f"{appendix_id}: No PDF file specified")
                continue
            
            if not Path(pdf_path).exists():
                errors.append(f"{appendix_id}: PDF file not found: {pdf_path}")
                continue
            
            # Validate PDF
            try:
                pdf_info = self.pdf_handler.validate_pdf_file(pdf_path)
                
                if not pdf_info.get('valid', False):
                    errors.append(f"{appendix_id}: Invalid PDF file")
                
                if pdf_info.get('page_count', 0) <= 0:
                    errors.append(f"{appendix_id}: PDF has no pages")
                
                # Check file size limits
                max_size_mb = self.settings.max_pdf_size_mb if self.settings else 200
                if pdf_info.get('size_mb', 0) > max_size_mb:
                    errors.append(f"{appendix_id}: File too large ({pdf_info['size_mb']:.1f}MB > {max_size_mb}MB)")
                
                # Add warnings as potential issues
                for warning in pdf_info.get('warnings', []):
                    errors.append(f"{appendix_id}: {warning}")
                
            except Exception as e:
                errors.append(f"{appendix_id}: PDF validation failed: {str(e)}")
        
        return errors
    
    def get_processing_summary(self) -> Dict[str, Any]:
        """Get summary of processing results."""
        total_pages = sum(a['page_count'] for a in self.appendices_added)
        
        return {
            'appendices_count': len(self.appendices_added),
            'total_pages': total_pages,
            'appendices_info': self.appendices_added.copy()
        }
    
    def cleanup_temp_files(self):
        """Clean up temporary files created during processing."""
        for temp_file in self.temp_files:
            try:
                if Path(temp_file).exists():
                    Path(temp_file).unlink()
            except Exception as e:
                self.logger.warning(f"Could not delete temp file {temp_file}: {e}")
        
        self.temp_files.clear()
        
        # Also cleanup PDF handler temp files
        if self.pdf_handler:
            self.pdf_handler.cleanup_temp_files()
        
        self.logger.info("Cleaned up temporary files")
    
    def reset(self):
        """Reset the appendix manager state."""
        self.appendices_added.clear()
        self.cleanup_temp_files()
        self.logger.info("Appendix manager reset")
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup_temp_files()
