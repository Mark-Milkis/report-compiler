"""
Main report compiler orchestrator class.
"""

import os
import time
from typing import Dict, List, Any, Optional
from pathlib import Path

from ..utils.file_manager import FileManager
from ..utils.validators import Validators
from ..document.placeholder_parser import PlaceholderParser
from ..document.docx_processor import DocxProcessor
from ..document.word_converter import WordConverter
from ..pdf.content_analyzer import ContentAnalyzer
from ..pdf.overlay_processor import OverlayProcessor
from ..pdf.merge_processor import MergeProcessor
from ..pdf.marker_remover import MarkerRemover
from ..core.config import Config


class ReportCompiler:
    """
    Main orchestrator class for report compilation.
    
    This class coordinates the entire process of converting DOCX documents
    with embedded PDF placeholders into final PDF reports.
    """
    
    def __init__(self, input_path: str, output_path: str, keep_temp: bool = False):
        """
        Initialize the report compiler.
        
        Args:
            input_path: Path to input DOCX file
            output_path: Path to output PDF file
            keep_temp: Whether to keep temporary files for debugging
        """
        self.input_path = input_path
        self.output_path = output_path
        self.keep_temp = keep_temp
          # Initialize components
        self.file_manager = FileManager(keep_temp)
        self.validators = Validators()
        self.placeholder_parser = PlaceholderParser()
        self.content_analyzer = ContentAnalyzer()
        self.docx_processor = DocxProcessor(input_path)
        self.word_converter = WordConverter()
        self.overlay_processor = OverlayProcessor()
        self.merge_processor = MergeProcessor()
        self.marker_remover = MarkerRemover()
        
        # Process state
        self.placeholders = {}
        self.base_directory = os.path.dirname(input_path)
        
        # File paths
        self.temp_docx_path = None
        self.temp_pdf_path = None
    
    def run(self) -> bool:
        """
        Run the complete report compilation process.
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            with self.file_manager:
                print("🔍 Input DOCX:", self.input_path)
                print("📄 Output PDF:", self.output_path)
                
                # Generate temporary file paths
                self.temp_docx_path = self.file_manager.generate_temp_path(
                    self.input_path, "modified_report")
                self.temp_pdf_path = self.file_manager.generate_temp_path(
                    self.output_path, "base")
                
                print("📋 Temp DOCX:", self.temp_docx_path)
                print("📑 Temp PDF:", self.temp_pdf_path)
                
                print("\\n=== Starting Report Compilation ===")
                
                # Step 1: Validate input
                if not self._validate_input():
                    return False
                
                # Step 2: Find and validate placeholders
                if not self._find_and_validate_placeholders():
                    return False
                
                # Step 3: Create modified DOCX
                if not self._create_modified_document():
                    return False
                
                # Step 4: Convert to PDF
                if not self._convert_to_pdf():
                    return False
                
                # Step 5: Process PDF insertions
                if not self._process_pdf_insertions():
                    return False
                
                print("\\n=== Report Compilation Complete ===")
                return True
                
        except Exception as e:
            print(f"\\n❌ Compilation failed: {e}")
            return False
    
    def _validate_input(self) -> bool:
        """Validate input file and output location."""
        print("Step 1: Validating input...")
        
        # Validate input DOCX
        docx_result = self.validators.validate_docx_path(self.input_path)
        if not docx_result['valid']:
            print(f"❌ {docx_result['error_message']}")
            return False
        
        print(f"✓ Input DOCX validated ({docx_result['file_size_mb']:.1f} MB)")
        
        # Validate output location
        output_result = self.validators.validate_output_path(self.output_path)
        if not output_result['valid']:
            print(f"❌ {output_result['error_message']}")
            return False
        
        if output_result['file_exists']:
            print(f"⚠️ Output file exists and will be overwritten")
        
        print("✓ Output location validated")
        return True
    
    def _find_and_validate_placeholders(self) -> bool:
        """Find and validate all PDF placeholders."""
        print("Step 2: Analyzing document for PDF placeholders...")
        
        # Find placeholders
        print("🔍 PHASE 1: Scanning document for placeholders...")
        self.placeholders = self.placeholder_parser.find_all_placeholders(self.input_path)
        
        if self.placeholders['total'] == 0:
            print("❌ No PDF placeholders found in document")
            print("   Use [[OVERLAY: path.pdf]] for table-based overlays")
            print("   Use [[INSERT: path.pdf]] for paragraph-based merges")
            return False
        
        # Validate placeholders
        placeholder_list = self.placeholders['table'] + self.placeholders['paragraph']
        validation_result = self.validators.validate_placeholders(placeholder_list)
        
        if not validation_result['valid']:
            print("❌ Placeholder validation failed:")
            for error in validation_result['errors']:
                print(f"   • {error}")
            return False
        
        # Validate PDF paths
        print("🔍 VALIDATING PDF references...")
        valid_count = 0
        
        for i, placeholder in enumerate(placeholder_list, 1):
            pdf_result = self.validators.validate_pdf_path(
                placeholder['pdf_path_raw'], 
                self.base_directory
            )
            
            placeholder_type = "overlay" if placeholder['type'] == 'overlay' else "merge"
            print(f"   📋 Placeholder #{i} ({placeholder_type}):")
            print(f"      • Raw path: {placeholder['pdf_path_raw']}")
            
            if pdf_result['valid']:
                print(f"      • Resolved path: {pdf_result['resolved_path']}")
                print(f"      ✅ Valid PDF with {pdf_result['page_count']} page(s)")
                placeholder['resolved_path'] = pdf_result['resolved_path']
                placeholder['page_count'] = pdf_result['page_count']
                valid_count += 1
            else:
                print(f"      ❌ {pdf_result['error_message']}")
                return False
        
        print(f"✅ Validated {valid_count}/{len(placeholder_list)} placeholders")
          # Summary
        print("📊 PROCESSING SUMMARY:")
        print(f"   • Overlay insertions (table-based): {len(self.placeholders['table'])}")
        print(f"   • Merge insertions (paragraph-based): {len(self.placeholders['paragraph'])}")
        
        return True
    
    def _create_modified_document(self) -> bool:
        """Create modified DOCX with markers."""
        print("Step 3: Creating modified document...")
        
        print("🔧 PHASE 2: Modifying document...")
        
        # Use DocxProcessor to create modified document
        success = self.docx_processor.create_modified_document(
            self.placeholders, self.temp_docx_path)        
        if success:
            print(f"✓ Modified document created: {self.temp_docx_path}")
        else:
            print("❌ Failed to create modified document")
        
        return success
    
    def _convert_to_pdf(self) -> bool:
        """Convert modified DOCX to PDF."""
        print("Step 4: Converting modified document to PDF...")
        
        print("🔄 PHASE 3: Converting to PDF...")
        
        # Use WordConverter to convert to PDF
        success = self.word_converter.convert_to_pdf(
            self.temp_docx_path, self.temp_pdf_path)
          if success:
            print(f"✓ PDF conversion successful: {self.temp_pdf_path}")
        else:
            print("❌ PDF conversion failed")
        
        return success
    
    def _process_pdf_insertions(self) -> bool:
        """Process all PDF insertions (overlays and merges)."""
        print("Step 5: Processing PDF insertions...")
        
        print("🔧 PHASE 4: Processing PDF insertions...")
        
        # Get table metadata from DocxProcessor
        table_metadata = self.docx_processor.get_table_metadata()
        
        # Process overlays first
        overlay_placeholders = self.placeholders.get('table', [])
        if overlay_placeholders:
            print(f"   📦 Processing {len(overlay_placeholders)} overlay insertions...")
            success = self.overlay_processor.process_overlays(
                self.temp_pdf_path, overlay_placeholders, 
                self.output_path, table_metadata)
            if not success:
                print("❌ Overlay processing failed")
                return False
        
        # Process merges
        merge_placeholders = self.placeholders.get('paragraph', [])
        if merge_placeholders:
            print(f"   📄 Processing {len(merge_placeholders)} merge insertions...")
            # Use output from overlay processing or temp PDF if no overlays
            input_pdf = self.output_path if overlay_placeholders else self.temp_pdf_path
            success = self.merge_processor.process_merges(
                input_pdf, merge_placeholders, self.output_path)
            if not success:
                print("❌ Merge processing failed")
                return False
        
        # If no placeholders to process, just copy the base PDF
        if not overlay_placeholders and not merge_placeholders:
            import shutil
            shutil.copy2(self.temp_pdf_path, self.output_path)
            print("✓ Base PDF copied to output (no insertions to process)")
        
        print(f"✓ Output PDF created: {self.output_path}")
        return True
