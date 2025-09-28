"""
Application Controller
Manages the communication between GUI and core business logic.
"""

import sys
from pathlib import Path

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent))

from PySide6.QtCore import QObject, QThread, Signal, QTimer
from PySide6.QtWidgets import QMessageBox, QProgressDialog, QApplication
from PySide6.QtCore import Qt
from typing import List, Dict, Any, Optional
from concurrent.futures import ThreadPoolExecutor

from utils.logger import get_logger, LoggerMixin
from utils.exceptions import (
    ApplicationError, WordDocumentError, PDFError, 
    ValidationError
)
from core.word_manager import WordManager
from core.pdf_handler import PDFHandler
from core.appendix_manager import AppendixManager


class ProcessingWorker(QObject):
    """Worker for handling long-running processing tasks in a separate thread."""
    
    # Signals
    progress_updated = Signal(str, int)  # message, percentage
    processing_completed = Signal(dict)  # result data
    processing_failed = Signal(str)      # error message
    
    def __init__(self, document_info, appendix_data, settings):
        super().__init__()
        self.document_info = document_info
        self.appendix_data = appendix_data
        self.settings = settings
        self.should_stop = False
    
    def process(self):
        """Main processing method that runs in the worker thread."""
        try:
            self.progress_updated.emit("Initializing...", 0)
            
            # Initialize managers
            word_manager = WordManager(self.settings)
            pdf_handler = PDFHandler(self.settings)
            appendix_manager = AppendixManager(word_manager, pdf_handler, self.settings)
            
            self.progress_updated.emit("Opening document...", 10)
            
            # Open/select document
            if self.document_info.get('source') == 'file':
                success = word_manager.open_document(self.document_info['path'])
            else:
                success = word_manager.select_open_document(self.document_info['index'])
            
            if not success:
                raise WordDocumentError("Failed to open document")
            
            self.progress_updated.emit("Creating backup...", 20)
            
            # Create backup if enabled
            backup_path = None
            if self.settings and self.settings.auto_backup_enabled:
                backup_path = word_manager.create_backup()
            
            self.progress_updated.emit("Processing appendices...", 30)
            
            # Process each appendix
            total_appendices = len(self.appendix_data)
            for i, appendix in enumerate(self.appendix_data):
                if self.should_stop:
                    break
                
                progress = 30 + (i * 50) // total_appendices
                self.progress_updated.emit(f"Adding {appendix.get('title', 'appendix')}...", progress)
                
                # Add appendix to document
                appendix_manager.add_appendix(appendix)
            
            if self.should_stop:
                return
            
            self.progress_updated.emit("Finalizing document...", 90)
            
            # Save document
            word_manager.save_document()
            
            self.progress_updated.emit("Exporting to PDF...", 95)
            
            # Export combined PDF (simplified for this implementation)
            try:
                output_path = appendix_manager.create_final_pdf()
            except Exception as e:
                # If PDF export fails, continue without it
                output_path = None
                print(f"Warning: PDF export failed: {e}")
            
            self.progress_updated.emit("Complete!", 100)
            
            # Prepare result
            result = {
                'success': True,
                'document_path': str(word_manager.document_path) if word_manager.document_path else '',
                'backup_path': backup_path,
                'output_pdf': output_path,
                'appendices_added': total_appendices
            }
            
            self.processing_completed.emit(result)
            
        except Exception as e:
            error_msg = f"Processing failed: {str(e)}"
            self.processing_failed.emit(error_msg)
    
    def stop(self):
        """Stop the processing."""
        self.should_stop = True


class AppController(QObject, LoggerMixin):
    """Main application controller."""
    
    def __init__(self, main_window, settings):
        super().__init__()
        self.main_window = main_window
        self.settings = settings
        
        # Core components
        self.word_manager = None
        self.pdf_handler = None
        self.appendix_manager = None
        
        # Processing
        self.processing_thread = None
        self.processing_worker = None
        self.progress_dialog = None
        
        # Data
        self.current_document = None
        self.appendix_data = []
        
        # Initialize components
        self.initialize_components()
        self.connect_signals()
        
        self.logger.info("Application controller initialized")
    
    def initialize_components(self):
        """Initialize core components."""
        try:
            self.pdf_handler = PDFHandler(self.settings)
            self.logger.info("Core components initialized")
        except Exception as e:
            self.logger.error(f"Failed to initialize components: {e}")
            self.main_window.show_message("Initialization Error", 
                                        f"Failed to initialize application components:\n{str(e)}", 
                                        "error")
    
    def connect_signals(self):
        """Connect main window signals to controller methods."""
        # Document selection
        self.main_window.document_selected.connect(self.on_document_selected)
        
        # PDF file handling
        self.main_window.pdf_files_added.connect(self.on_pdf_files_added)
        
        # Appendix management
        self.main_window.appendix_removed.connect(self.on_appendix_removed)
        
        # Processing
        self.main_window.process_requested.connect(self.on_process_requested)
    
    def on_document_selected(self, document_info):
        """Handle document selection."""
        try:
            # Initialize word manager if needed
            if not self.word_manager:
                self.word_manager = WordManager(self.settings)
            
            # Store document info
            self.current_document = document_info
            
            if self.current_document:
                self.logger.info(f"Document selected: {self.current_document.get('name', 'Unknown')}")
            
        except Exception as e:
            self.logger.error(f"Document selection failed: {e}")
            self.main_window.show_message("Document Selection Error", 
                                        f"Failed to select document:\n{str(e)}", 
                                        "error")
    
    def on_pdf_files_added(self, file_paths: List[str]):
        """Handle adding PDF files as appendices."""
        if not file_paths:
            return
        
        self.main_window.show_progress("Processing PDF files...", 0)
        
        # Process files in a separate thread to avoid blocking UI
        QTimer.singleShot(100, lambda: self._process_pdf_files(file_paths))
    
    def _process_pdf_files(self, file_paths: List[str]):
        """Process PDF files in background."""
        try:
            processed_files = []
            total_files = len(file_paths)
            
            for i, file_path in enumerate(file_paths):
                progress = (i * 100) // total_files
                self.main_window.show_progress(f"Processing {Path(file_path).name}...", progress)
                
                try:
                    # Validate PDF
                    pdf_info = self.pdf_handler.validate_pdf_file(file_path)
                    
                    # Create appendix data
                    appendix_data = {
                        'path': file_path,
                        'title': self._generate_appendix_title(len(self.appendix_data) + len(processed_files)),
                        'page_count': pdf_info.get('page_count', 0),
                        'size_mb': pdf_info.get('size_mb', 0),
                        'orientation': pdf_info.get('orientation', 'unknown'),
                        'valid': pdf_info.get('valid', False),
                        'warnings': pdf_info.get('warnings', []),
                        'pdf_info': pdf_info
                    }
                    
                    processed_files.append(appendix_data)
                    
                except Exception as e:
                    self.logger.error(f"Failed to process {file_path}: {e}")
                    # Add as invalid appendix
                    appendix_data = {
                        'path': file_path,
                        'title': self._generate_appendix_title(len(self.appendix_data) + len(processed_files)),
                        'page_count': 0,
                        'size_mb': 0,
                        'orientation': 'unknown',
                        'valid': False,
                        'warnings': [f"Processing error: {str(e)}"],
                        'pdf_info': {}
                    }
                    processed_files.append(appendix_data)
            
            # Add to appendix data
            self.appendix_data.extend(processed_files)
            
            # Update main window
            self.main_window.set_appendix_data(self.appendix_data)
            self.main_window.hide_progress()
            
            # Show summary
            valid_count = sum(1 for a in processed_files if a['valid'])
            if valid_count == len(processed_files):
                self.main_window.show_message("Files Added", 
                                            f"Successfully added {len(processed_files)} PDF files", 
                                            "info")
            else:
                invalid_count = len(processed_files) - valid_count
                self.main_window.show_message("Files Added with Warnings", 
                                            f"Added {len(processed_files)} files ({invalid_count} with warnings)", 
                                            "warning")
            
            self.logger.info(f"Added {len(processed_files)} PDF files ({valid_count} valid)")
            
        except Exception as e:
            self.main_window.hide_progress()
            self.logger.error(f"PDF processing failed: {e}")
            self.main_window.show_message("Processing Error", 
                                        f"Failed to process PDF files:\n{str(e)}", 
                                        "error")
    
    def _generate_appendix_title(self, index: int) -> str:
        """Generate appendix title based on numbering style."""
        numbering_style = self.main_window.get_numbering_style()
        
        if numbering_style == "numeric":
            return f"Appendix {index + 1}"
        else:  # alphabetical
            if index < 26:
                return f"Appendix {chr(ord('A') + index)}"
            else:
                # For more than 26 appendices
                first_letter = chr(ord('A') + (index // 26) - 1)
                second_letter = chr(ord('A') + (index % 26))
                return f"Appendix {first_letter}{second_letter}"
    
    def on_appendix_removed(self, index: int):
        """Handle appendix removal."""
        if 0 <= index < len(self.appendix_data):
            removed_appendix = self.appendix_data.pop(index)
            
            # Regenerate titles for remaining appendices
            for i, appendix in enumerate(self.appendix_data):
                appendix['title'] = self._generate_appendix_title(i)
            
            # Update main window
            self.main_window.set_appendix_data(self.appendix_data)
            
            self.logger.info(f"Removed appendix: {removed_appendix.get('title', 'Unknown')}")
    
    def on_process_requested(self):
        """Handle processing request."""
        if not self.current_document:
            self.main_window.show_message("No Document", "Please select a document first.", "warning")
            return
        
        if not self.appendix_data:
            self.main_window.show_message("No Appendices", "Please add some PDF files first.", "warning")
            return
        
        # Validate appendices
        errors = self._validate_appendices()
        if errors:
            error_text = "\n".join(errors[:5])  # Show first 5 errors
            if len(errors) > 5:
                error_text += f"\n... and {len(errors) - 5} more errors"
            
            reply = QMessageBox.question(
                self.main_window,
                "Validation Errors",
                f"Found {len(errors)} validation errors:\n\n{error_text}\n\nContinue anyway?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply != QMessageBox.Yes:
                return
        
        # Start processing
        self.start_processing()
    
    def _validate_appendices(self) -> List[str]:
        """Validate all appendices and return error messages."""
        errors = []
        
        for i, appendix in enumerate(self.appendix_data):
            # Check file existence
            if not Path(appendix['path']).exists():
                errors.append(f"{appendix['title']}: File not found")
            
            # Check page count
            if appendix.get('page_count', 0) <= 0:
                errors.append(f"{appendix['title']}: Invalid page count")
            
            # Check file size
            max_size = self.settings.max_pdf_size_mb if self.settings else 200
            if appendix.get('size_mb', 0) > max_size:
                errors.append(f"{appendix['title']}: File too large ({appendix['size_mb']:.1f}MB)")
            
            # Add warnings as errors if they're critical
            for warning in appendix.get('warnings', []):
                if 'corrupted' in warning.lower() or 'error' in warning.lower():
                    errors.append(f"{appendix['title']}: {warning}")
        
        return errors
    
    def start_processing(self):
        """Start the appendix processing in a separate thread."""
        try:
            # Create progress dialog
            self.progress_dialog = QProgressDialog(
                "Processing appendices...", "Cancel", 0, 100, self.main_window
            )
            self.progress_dialog.setWindowTitle("Processing")
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.setAutoClose(False)
            self.progress_dialog.setAutoReset(False)
            
            # Create worker and thread
            self.processing_worker = ProcessingWorker(
                self.current_document,
                self.appendix_data.copy(),
                self.settings
            )
            self.processing_thread = QThread()
            self.processing_worker.moveToThread(self.processing_thread)
            
            # Connect signals
            self.processing_thread.started.connect(self.processing_worker.process)
            self.processing_worker.progress_updated.connect(self.on_processing_progress)
            self.processing_worker.processing_completed.connect(self.on_processing_completed)
            self.processing_worker.processing_failed.connect(self.on_processing_failed)
            self.progress_dialog.canceled.connect(self.cancel_processing)
            
            # Clean up when done
            self.processing_worker.processing_completed.connect(self.processing_thread.quit)
            self.processing_worker.processing_failed.connect(self.processing_thread.quit)
            self.processing_thread.finished.connect(self.processing_worker.deleteLater)
            self.processing_thread.finished.connect(self.processing_thread.deleteLater)
            
            # Start processing
            self.processing_thread.start()
            self.progress_dialog.show()
            
            self.logger.info("Started appendix processing")
            
        except Exception as e:
            self.logger.error(f"Failed to start processing: {e}")
            self.main_window.show_message("Processing Error", 
                                        f"Failed to start processing:\n{str(e)}", 
                                        "error")
    
    def on_processing_progress(self, message: str, percentage: int):
        """Handle processing progress updates."""
        if self.progress_dialog:
            self.progress_dialog.setLabelText(message)
            self.progress_dialog.setValue(percentage)
        
        self.main_window.show_progress(message, percentage)
    
    def on_processing_completed(self, result: Dict[str, Any]):
        """Handle successful processing completion."""
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        
        self.main_window.hide_progress()
        
        # Show success message
        message = f"Successfully processed {result.get('appendices_added', 0)} appendices!\n\n"
        
        if result.get('backup_path'):
            message += f"Backup created: {Path(result['backup_path']).name}\n"
        
        if result.get('output_pdf'):
            message += f"Final PDF: {Path(result['output_pdf']).name}\n"
        
        message += f"\nDocument: {Path(result.get('document_path', '')).name}"
        
        QMessageBox.information(self.main_window, "Processing Complete", message)
        
        self.logger.info("Appendix processing completed successfully")
    
    def on_processing_failed(self, error_message: str):
        """Handle processing failure."""
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        
        self.main_window.hide_progress()
        
        self.main_window.show_message("Processing Failed", error_message, "error")
        self.logger.error(f"Appendix processing failed: {error_message}")
    
    def cancel_processing(self):
        """Cancel the current processing operation."""
        if self.processing_worker:
            self.processing_worker.stop()
        
        if self.processing_thread and self.processing_thread.isRunning():
            self.processing_thread.quit()
            self.processing_thread.wait(3000)  # Wait up to 3 seconds
        
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        
        self.main_window.hide_progress()
        self.main_window.show_message("Processing Cancelled", "Processing was cancelled by user.", "info")
        
        self.logger.info("Processing cancelled by user")
    
    def get_appendix_data(self) -> List[Dict[str, Any]]:
        """Get current appendix data."""
        return self.appendix_data.copy()
    
    def set_appendix_data(self, data: List[Dict[str, Any]]):
        """Set appendix data."""
        self.appendix_data = data.copy()
        self.main_window.set_appendix_data(self.appendix_data)
    
    def clear_appendices(self):
        """Clear all appendices."""
        self.appendix_data.clear()
        self.main_window.set_appendix_data(self.appendix_data)
    
    def shutdown(self):
        """Shutdown the controller and clean up resources."""
        # Cancel any ongoing processing
        try:
            if self.processing_thread and self.processing_thread.isRunning():
                self.cancel_processing()
        except RuntimeError:
            self.logger.warning("Processing thread already deleted")

        # Clean up managers
        if self.word_manager:
            self.word_manager.close()

        if self.pdf_handler:
            self.pdf_handler.cleanup_temp_files()

        self.logger.info("Application controller shutdown")
