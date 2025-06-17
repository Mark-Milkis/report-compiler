"""
Report compiler core orchestration module.

This module contains the main ReportCompiler class that orchestrates the entire
report compilation process, from input validation through PDF generation.
"""

import os
import sys
from typing import Dict, Any

from ..utils.file_manager import FileManager
from ..utils.validators import Validators
from ..document.placeholder_parser import PlaceholderParser
from ..pdf.content_analyzer import ContentAnalyzer
from ..document.docx_processor import DocxProcessor
from ..document.word_converter import WordConverter
from ..document.libreoffice_converter import LibreOfficeConverter
from ..pdf.overlay_processor import OverlayProcessor
from ..pdf.merge_processor import MergeProcessor
from ..pdf.marker_remover import MarkerRemover
from ..utils.logging_config import get_compiler_logger


class ReportCompiler:
    """Main orchestrator class for report compilation."""
    
    def __init__(self, input_path: str, output_path: str, keep_temp: bool = False):
        """
        Initialize the report compiler.
        
        Args:
            input_path: Path to input DOCX file
            output_path: Path for output PDF file
            keep_temp: Whether to keep temporary files for debugging
        """
        self.input_path = input_path
        self.output_path = output_path
        self.keep_temp = keep_temp
        self.logger = get_compiler_logger()
        
        # Initialize components
        self.file_manager = FileManager(keep_temp)
        self.validators = Validators()
        self.placeholder_parser = PlaceholderParser()
        self.content_analyzer = ContentAnalyzer()
        self.docx_processor = DocxProcessor(input_path)
        if sys.platform == 'win32':
            self.word_converter = WordConverter()
        else:
            self.word_converter = None
        self.libreoffice_converter = LibreOfficeConverter()
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
        """Run the complete report compilation process."""
        try:
            with self.file_manager:
                self.logger.info("🔍 Input DOCX: %s", self.input_path)
                self.logger.info("📄 Output PDF: %s", self.output_path)
                
                # Generate temporary file paths
                self.temp_docx_path = self.file_manager.generate_temp_path(
                    self.input_path, "modified_report")
                self.temp_pdf_path = self.file_manager.generate_temp_path(
                    self.output_path, "base")
                
                self.logger.debug("📋 Temp DOCX: %s", self.temp_docx_path)
                self.logger.debug("📑 Temp PDF: %s", self.temp_pdf_path)
                
                self.logger.info("=== Starting Report Compilation ===")
                
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
                
                self.logger.info("=== Report Compilation Complete ===")
                return True
                
        except Exception as e:
            self.logger.error("❌ Compilation failed: %s", e, exc_info=True)
            return False
    
    def _validate_input(self) -> bool:
        """Validate input file and output location."""
        self.logger.info("Step 1: Validating input...")
        
        # Validate input DOCX
        docx_result = self.validators.validate_docx_path(self.input_path)
        if not docx_result['valid']:
            self.logger.error("❌ %s", docx_result['error_message'])
            return False
        
        self.logger.info("✓ Input DOCX validated (%.1f MB)", docx_result['file_size_mb'])
        
        # Validate output location
        output_result = self.validators.validate_output_path(self.output_path)
        if not output_result['valid']:
            self.logger.error("❌ %s", output_result['error_message'])
            return False
        
        if output_result['file_exists']:
            self.logger.warning("⚠️ Output file exists and will be overwritten")
        
        self.logger.info("✓ Output location validated")
        return True
    
    def _find_and_validate_placeholders(self) -> bool:
        """Find and validate all PDF placeholders."""
        self.logger.info("Step 2: Analyzing document for PDF placeholders...")
        
        # Find placeholders
        self.logger.info("🔍 PHASE 1: Scanning document for placeholders...")
        self.placeholders = self.placeholder_parser.find_all_placeholders(self.input_path)
        
        if self.placeholders['total'] == 0:
            self.logger.error("❌ No PDF placeholders found in document")
            self.logger.info("   Use [[OVERLAY: path.pdf]] for table-based overlays")
            self.logger.info("   Use [[INSERT: path.pdf]] for paragraph-based merges")
            return False
        
        # Validate placeholders
        placeholder_list = self.placeholders['table'] + self.placeholders['paragraph']
        validation_result = self.validators.validate_placeholders(placeholder_list)
        
        if not validation_result['valid']:
            self.logger.error("❌ Placeholder validation failed:")
            for error in validation_result['errors']:
                self.logger.error("   • %s", error)
            return False
        
        # Validate PDF paths
        self.logger.info("🔍 VALIDATING PDF references...")
        valid_count = 0
        
        for i, placeholder in enumerate(placeholder_list, 1):
            pdf_result = self.validators.validate_pdf_path(
                placeholder['pdf_path_raw'], 
                self.base_directory
            )
            
            placeholder_type = "overlay" if placeholder['type'] == 'overlay' else "merge"
            self.logger.info("   📋 Placeholder #%d (%s):", i, placeholder_type)
            self.logger.info("      • Raw path: %s", placeholder['pdf_path_raw'])
            
            if pdf_result['valid']:
                self.logger.info("      • Resolved path: %s", pdf_result['resolved_path'])
                self.logger.info("      ✅ Valid PDF with %d page(s)", pdf_result['page_count'])
                placeholder['resolved_path'] = pdf_result['resolved_path']
                placeholder['page_count'] = pdf_result['page_count']
                valid_count += 1
            else:
                self.logger.error("      ❌ %s", pdf_result['error_message'])
                return False
        
        self.logger.info("✅ Validated %d/%d placeholders", valid_count, len(placeholder_list))
        
        # Summary
        self.logger.info("📊 PROCESSING SUMMARY:")
        self.logger.info("   • Overlay insertions (table-based): %d", len(self.placeholders['table']))
        self.logger.info("   • Merge insertions (paragraph-based): %d", len(self.placeholders['paragraph']))
        
        return True
    
    def _create_modified_document(self) -> bool:
        """Create modified DOCX with markers."""
        self.logger.info("Step 3: Creating modified document...")
        
        self.logger.info("🔧 PHASE 2: Modifying document...")
        
        # Use DocxProcessor to create modified document
        success = self.docx_processor.create_modified_document(
            self.placeholders, self.temp_docx_path)
        
        if success:
            self.logger.info("✓ Modified document created: %s", self.temp_docx_path)
        else:
            self.logger.error("❌ Failed to create modified document")
        
        return success
    
    def _convert_to_pdf(self) -> bool:
        """Convert modified DOCX to PDF."""
        self.logger.info("Step 4: Converting modified document to PDF...")
        
        self.logger.info("🔄 PHASE 3: Converting to PDF...")
        
        # Use WordConverter to convert to PDF
        success = self.word_converter.convert_to_pdf(
            self.temp_docx_path, self.temp_pdf_path)
        
        if success:
            self.logger.info("✓ PDF conversion successful: %s", self.temp_pdf_path)
        else:
            self.logger.error("❌ PDF conversion failed")
        
        return success
    
    def _process_pdf_insertions(self) -> bool:
        """Process all PDF insertions (overlays and merges)."""
        self.logger.info("Step 5: Processing PDF insertions...")
        
        self.logger.info("🔧 PHASE 4: Processing PDF insertions...")
        
        # Get table metadata from DocxProcessor
        table_metadata = self.docx_processor.get_table_metadata()
        
        # Process overlays first
        overlay_placeholders = self.placeholders.get('table', [])
        if overlay_placeholders:
            self.logger.info("   📦 Processing %d overlay insertions...", len(overlay_placeholders))
            success = self.overlay_processor.process_overlays(
                self.temp_pdf_path, overlay_placeholders, 
                table_metadata, self.output_path)
            if not success:
                self.logger.error("❌ Overlay processing failed")
                return False
        
        # Process merges
        merge_placeholders = self.placeholders.get('paragraph', [])
        if merge_placeholders:
            self.logger.info("   📄 Processing %d merge insertions...", len(merge_placeholders))
            # Use output from overlay processing or temp PDF if no overlays
            input_pdf = self.output_path if overlay_placeholders else self.temp_pdf_path
            success = self.merge_processor.process_merges(
                input_pdf, merge_placeholders, self.output_path)
            if not success:
                self.logger.error("❌ Merge processing failed")
                return False
        
        # If no placeholders to process, just copy the base PDF
        if not overlay_placeholders and not merge_placeholders:
            import shutil
            shutil.copy2(self.temp_pdf_path, self.output_path)
            self.logger.info("✓ Base PDF copied to output (no insertions to process)")
        
        self.logger.info("✓ Output PDF created: %s", self.output_path)
        return True
