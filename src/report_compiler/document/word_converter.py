"""
Word automation for DOCX to PDF conversion.
"""

import os
import time
from typing import Optional
from ..core.config import Config
from ..utils.logging_config import get_logger

try:
    import win32com.client
except ImportError:
    win32com = None


class WordConverter:
    """Handles DOCX to PDF conversion using Microsoft Word automation."""
    
    def __init__(self):
        self.word_app = None
        self.is_connected = False
        self.logger = get_logger()

    def is_available(self) -> bool:
        """Checks if the Word application can be automated."""
        if win32com is None:
            return False
        # A lightweight check without launching the full app
        try:
            win32com.client.Dispatch("Word.Application")
            return True
        except Exception:
            return False
    
    def connect(self) -> bool:
        """
        Connect to Microsoft Word application.
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        if win32com is None:
            self.logger.error("pywin32 is not installed. Word automation is unavailable on this platform.")
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
    
    def update_fields_and_save_as_pdf(self, docx_path: str, pdf_path: str) -> bool:
        """
        Updates fields (like TOC) in a DOCX and then saves it as a PDF.
        
        Args:
            docx_path: Path to input DOCX file
            pdf_path: Path to output PDF file
            
        Returns:
            bool: True if conversion successful, False otherwise
        """
        if not self.is_connected:
            if not self.connect():
                return False
        
        doc: Optional[object] = None
        try:
            self.logger.debug(f"  > Opening document: {os.path.basename(docx_path)}")
            doc = self.word_app.Documents.Open(docx_path)
            
            self.logger.info("  > Updating document fields (e.g., Table of Contents)...")
            doc.Fields.Update()
            time.sleep(1) # Give Word a moment to update.

            self.logger.debug(f"  > Exporting to PDF: {os.path.basename(pdf_path)}")
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=Config.WORD_EXPORT_FORMAT,  # PDF format
                OpenAfterExport=False,
                OptimizeFor=0,  # Print optimization
                Range=0,       # Export entire document
                Item=0,        # Export document content
                CreateBookmarks=1,  # Create bookmarks from headings
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False
            )
            
            return True
            
        except Exception as e:
            self.logger.error(f"  > Error during Word conversion: {e}", exc_info=True)
            return False
            
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                    self.logger.debug("  > Document closed.")
                except Exception as e:
                    self.logger.warning(f"  > Error closing document: {e}")
    
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
