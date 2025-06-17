"""
Word automation for DOCX to PDF conversion.
"""

import os
import time
from typing import Optional
from ..core.config import Config
from ..utils.logging_config import get_logger


class WordConverter:
    """Handles DOCX to PDF conversion using Microsoft Word automation."""
    
    def __init__(self):
        self.word_app = None
        self.is_connected = False
        self.logger = get_logger()
    
    def connect(self) -> bool:
        """
        Connect to Microsoft Word application.
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        if win32com is None:
            print("    ❌ pywin32 is not installed. Word automation is unavailable on this platform.")
            self.is_connected = False
            return False
        
        try:
            # Try to connect to existing Word instance first
            try:
                self.word_app = win32com.client.GetActiveObject("Word.Application")
                self.logger.debug("Connected to existing Word instance")
                self.is_connected = True
                return True
            except:
                pass
            
            # If no existing instance, create new one
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False  # Run in background
            self.logger.debug("Created new Word instance")
            self.is_connected = True
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to connect to Word: {e}", exc_info=True)
            self.is_connected = False
            return False
    
    def convert_to_pdf(self, docx_path: str, pdf_path: str) -> bool:
        """
        Convert DOCX file to PDF using the configured rendering engine.
        
        Args:
            docx_path: Path to input DOCX file
            pdf_path: Path to output PDF file
            
        Returns:
            bool: True if conversion successful, False otherwise
        """
        if win32com is None:
            print("    ❌ pywin32 is not installed. Cannot convert DOCX to PDF using Word automation.")
            return False
        
        if not self.is_connected:
            if not self.connect():
                return False
        
        doc: Optional[object] = None
        try:
            self.logger.info("Converting DOCX to PDF with bookmark generation")
            self.logger.debug(f"Input: {docx_path}")
            self.logger.debug(f"Output: {pdf_path}")
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
            
            # Open the document
            self.logger.debug(f"Opening document: {os.path.basename(docx_path)}")
            doc = self.word_app.Documents.Open(docx_path)
            
            # Wait a moment for document to fully load
            time.sleep(0.5)
            
            # Export to PDF with bookmark generation
            self.logger.debug(f"Exporting to PDF: {os.path.basename(pdf_path)}")
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=Config.WORD_EXPORT_FORMAT,  # PDF format
                OpenAfterExport=False,
                OptimizeFor=0,  # Print optimization - wdExportOptimizeForPrint
                Range=0,       # Export entire document - wdExportAllDocument
                Item=0,        # Export document content - wdExportDocumentContent
                CreateBookmarks=1,  # Create bookmarks from headings - wdExportCreateHeadingBookmarks
                DocStructureTags=True, # Create document structure tags for accessibility
                BitmapMissingFonts=True,
                UseISO19005_1=False # PDF/A compatibility, can be restrictive
            )
            
            self.logger.info(f"Successfully converted '{os.path.basename(docx_path)}' to PDF with bookmarks")
            return True
            
        except Exception as e:
            self.logger.error(f"Error converting DOCX to PDF: {e}", exc_info=True)
            return False
            
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                    self.logger.debug("Document closed")
                except Exception as e:
                    self.logger.warning(f"Error closing document: {e}")
    
    def disconnect(self) -> None:
        """Disconnect from Word application."""
        if self.word_app and self.is_connected:
            try:
                # Don't quit Word - might be used by other processes
                self.word_app = None
                self.is_connected = False
                self.logger.debug("Disconnected from Word")
            except Exception as e:
                self.logger.warning(f"Error disconnecting from Word: {e}")
    
    def __enter__(self):
        """Context manager entry."""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.disconnect()
