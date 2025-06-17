"""
Word automation for DOCX to PDF conversion.
"""

import os
import time
import win32com.client
from typing import Optional
from ..core.config import Config


class WordConverter:
    """Handles DOCX to PDF conversion using Microsoft Word automation."""
    
    def __init__(self):
        self.word_app = None
        self.is_connected = False
    
    def connect(self) -> bool:
        """
        Connect to Microsoft Word application.
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        try:
            # Try to connect to existing Word instance first
            try:
                self.word_app = win32com.client.GetActiveObject("Word.Application")
                print("    ✓ Connected to existing Word instance")
                self.is_connected = True
                return True
            except:
                pass
            
            # If no existing instance, create new one
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False  # Run in background
            print("    ✓ Created new Word instance")
            self.is_connected = True
            return True
            
        except Exception as e:
            print(f"    ❌ Failed to connect to Word: {e}")
            self.is_connected = False
            return False
    
    def convert_to_pdf(self, docx_path: str, pdf_path: str) -> bool:
        """
        Convert DOCX file to PDF using Word automation.
        
        Args:
            docx_path: Path to input DOCX file
            pdf_path: Path to output PDF file
            
        Returns:
            bool: True if conversion successful, False otherwise
        """
        if not self.is_connected:
            if not self.connect():
                return False
        
        doc = None
        try:
            print("    Converting DOCX to PDF...")
            print(f"    Input: {docx_path}")
            print(f"    Output: {pdf_path}")
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
            
            # Open the document
            print(f"    Opening document: {os.path.basename(docx_path)}")
            doc = self.word_app.Documents.Open(docx_path)
            
            # Wait a moment for document to fully load
            time.sleep(0.5)
              # Export to PDF
            print(f"    Exporting to PDF: {os.path.basename(pdf_path)}")
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=Config.WORD_EXPORT_FORMAT,  # PDF format
                OpenAfterExport=False,
                OptimizeFor=0,  # Print optimization
                BitmapMissingFonts=True,
                DocStructureTags=False,
                CreateBookmarks=False
            )
            
            print(f"    ✓ Successfully converted '{os.path.basename(docx_path)}' to PDF")
            return True
            
        except Exception as e:
            print(f"    ❌ Error converting to PDF: {e}")
            return False
            
        finally:
            # Close the document
            if doc:
                try:
                    doc.Close(SaveChanges=False)
                    print("    ✓ Document closed")
                except:
                    pass
    
    def disconnect(self) -> None:
        """Disconnect from Word application."""
        if self.word_app and self.is_connected:
            try:
                # Don't quit Word - might be used by other processes
                self.word_app = None
                self.is_connected = False
            except:
                pass
    
    def __enter__(self):
        """Context manager entry."""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.disconnect()
