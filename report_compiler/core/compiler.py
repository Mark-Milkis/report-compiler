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
from ..core.config import Config


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
        self.docx_processor = DocxProcessor()  # No longer needs input_path
        self.word_converter = WordConverter()
        self.libreoffice_converter = LibreOfficeConverter()
        self.overlay_processor = OverlayProcessor()
        self.merge_processor = MergeProcessor()
        self.marker_remover = MarkerRemover()
        
        # Process state
        self.placeholders = {}
        self.base_directory = os.path.dirname(input_path)
        self.table_metadata = {}

        # File paths
        self.temp_docx_path = None
        self.temp_pdf_path = None
        self.final_pdf_path = None
        self.overlay_pdf_path = None
        self.toc_pages = []
        self.content_map = {}

    def run(self) -> bool:
        """Run the complete report compilation process."""
        try:
            with self.file_manager:
                if not self._initialize(): return False
                if not self._validate_inputs(): return False
                if not self._modify_docx(): return False
                if not self._convert_to_pdf(): return False
                if not self._analyze_base_pdf(): return False
                if not self._process_pdf_overlays(): return False
                if not self._process_pdf_merges(): return False
                if not self._finalize_pdf(): return False
                self.logger.info("\n=== Report Compilation Successful ===")
                return True
        except Exception as e:
            self.logger.error("❌ A critical error occurred: %s", e, exc_info=True)
            return False

    def _initialize(self) -> bool:
        """[Stage 1/8: Initialization] Set up temporary file paths and initial state."""
        self.logger.info("[Stage 1/8: Initialization]")
        self.temp_docx_path = self.file_manager.generate_temp_path(self.input_path, "modified_report")
        self.temp_pdf_path = self.file_manager.generate_temp_path(self.output_path, "base")
        self.overlay_pdf_path = self.file_manager.generate_temp_path(self.output_path, "with_overlays")
        self.final_pdf_path = self.output_path
        self.logger.debug("  > Input DOCX: %s", self.input_path)
        self.logger.debug("  > Output PDF: %s", self.final_pdf_path)
        self.logger.debug("  > Temp DOCX: %s", self.temp_docx_path)
        self.logger.debug("  > Temp PDF (Base): %s", self.temp_pdf_path)
        self.logger.debug("  > Temp PDF (Overlaid): %s", self.overlay_pdf_path)
        self.logger.info("  > Environment initialized.")
        return True

    def _validate_inputs(self) -> bool:
        """[Stage 2/8: Input Validation] Validate input files and placeholders."""
        self.logger.info("[Stage 2/8: Input Validation]")
        
        # Validate input DOCX
        self.logger.debug("  > Validating source DOCX file...")
        docx_result = self.validators.validate_docx_path(self.input_path)
        if not docx_result['valid']:
            self.logger.error("  > ❌ %s", docx_result['error_message'])
            return False
        self.logger.info("  > Source DOCX is valid (%.1f MB).", docx_result['file_size_mb'])

        # Validate output location
        self.logger.debug("  > Validating output path...")
        output_result = self.validators.validate_output_path(self.output_path)
        if not output_result['valid']:
            self.logger.error("  > ❌ %s", output_result['error_message'])
            return False
        if output_result['file_exists']:
            self.logger.warning("  > ⚠️ Output file exists and will be overwritten.")
        self.logger.debug("  > Output path is valid.")

        # Find and validate placeholders
        self.logger.info("  > Scanning document for placeholders...")
        self.placeholders = self.placeholder_parser.find_all_placeholders(self.input_path)
        if self.placeholders['total'] == 0:
            self.logger.warning("  > ⚠️ No PDF placeholders found. The document will be converted as-is.")
            self.logger.info("  > Validation complete.")
            return True

        self.logger.info("  > Found %d placeholders. Validating paths and parameters...", self.placeholders['total'])
        placeholder_list = self.placeholders['table'] + self.placeholders['paragraph']
        validation_result = self.validators.validate_placeholders(placeholder_list, self.base_directory)
        if not validation_result['valid']:
            for error in validation_result['errors']:
                self.logger.error("  > ❌ %s", error)
            return False
        self.logger.info("  > All placeholders are valid.")
        self.logger.info("  > Validation complete.")
        return True

    def _modify_docx(self) -> bool:
        """[Stage 3/8: DOCX Modification] Create a modified DOCX with placeholders handled."""
        self.logger.info("[Stage 3/8: DOCX Modification]")
        if self.placeholders['total'] == 0:
            self.logger.info("  > No placeholders to process. Copying original document for conversion.")
            self.file_manager.copy_file(self.input_path, self.temp_docx_path)
            self.table_metadata = {}
            return True

        self.logger.info("  > Inserting markers into DOCX for %d placeholders...", self.placeholders['total'])
        table_metadata = self.docx_processor.create_modified_docx(
            self.input_path, self.placeholders, self.temp_docx_path
        )

        if table_metadata is None:
            self.logger.error("  > ❌ Failed to create modified DOCX.")
            return False

        self.table_metadata = table_metadata
        self.logger.info("  > Modified DOCX created successfully.")
        return True

    def _convert_to_pdf(self) -> bool:
        """[Stage 4/8: PDF Conversion] Convert the DOCX to a base PDF."""
        self.logger.info("[Stage 4/8: PDF Conversion]")
        use_libreoffice = False
        if self.word_converter.is_available():
            self.logger.info("  > Attempting conversion with MS Word...")
            success = self.word_converter.update_fields_and_save_as_pdf(
                self.temp_docx_path, self.temp_pdf_path
            )
            if not success:
                self.logger.warning("  > MS Word conversion failed. Falling back to LibreOffice.")
                use_libreoffice = True
            else:
                self.logger.info("  > ✓ Conversion with MS Word successful.")
        else:
            self.logger.info("  > MS Word not available. Using LibreOffice for conversion.")
            use_libreoffice = True

        if use_libreoffice:
            if not self.libreoffice_converter.is_available():
                self.logger.error("  > ❌ Neither MS Word nor LibreOffice is available for PDF conversion.")
                return False
            self.logger.info("  > Attempting conversion with LibreOffice...")
            success = self.libreoffice_converter.convert_to_pdf(
                self.temp_docx_path, os.path.dirname(self.temp_pdf_path)
            )
            if not success:
                self.logger.error("  > ❌ LibreOffice conversion failed.")
                return False
            expected_pdf = self.temp_docx_path.replace(".docx", ".pdf")
            self.file_manager.move_file(expected_pdf, self.temp_pdf_path)
            self.logger.info("  > ✓ Conversion with LibreOffice successful.")
        
        self.logger.info("  > Base PDF created successfully.")
        return True

    def _analyze_base_pdf(self) -> bool:
        """[Stage 5/8: PDF Analysis] Find TOC pages and placeholder marker locations."""
        self.logger.info("[Stage 5/8: PDF Analysis]")
        self.logger.debug("  > Analyzing PDF: %s", self.temp_pdf_path)
        self.toc_pages = self.content_analyzer.find_toc_pages(self.temp_pdf_path)
        if self.toc_pages:
            self.logger.info("  > Table of Contents found on pages: %s", [p + 1 for p in self.toc_pages])
        else:
            self.logger.warning("  > No Table of Contents found in the document.")

        if self.placeholders['total'] == 0:
            self.logger.info("  > No placeholders were in the document, skipping marker analysis.")
            return True

        self.logger.info("  > Searching for placeholder markers in the PDF...")
        self.content_map = self.content_analyzer.find_placeholder_markers(
            self.temp_pdf_path, self.placeholders, self.table_metadata
        )
        if not self.content_map:
            self.logger.error("  > ❌ Could not find any placeholder markers in the converted PDF.")
            return False
        
        self.logger.info("  > Found %d placeholder markers.", len(self.content_map))
        self.logger.info("  > PDF analysis complete.")
        return True

    def _process_pdf_overlays(self) -> bool:
        """[Stage 6/8: PDF Overlay Processing] Handle all table-based PDF overlays."""
        num_overlays = len(self.placeholders.get('table', []))
        self.logger.info("[Stage 6/8: PDF Overlay Processing]")
        if num_overlays == 0:
            self.logger.info("  > No overlays to process. Skipping.")
            self.file_manager.copy_file(self.temp_pdf_path, self.overlay_pdf_path)
            return True

        self.logger.info("  > Processing %d overlay(s)...", num_overlays)
        success = self.overlay_processor.process_overlays(
            self.temp_pdf_path, self.content_map, self.overlay_pdf_path
        )
        if not success:
            self.logger.error("  > ❌ Overlay processing failed.")
            return False
        self.logger.info("  > Overlay processing complete.")
        return True

    def _process_pdf_merges(self) -> bool:
        """[Stage 7/8: PDF Merge Processing] Handle all paragraph-based PDF merges."""
        num_merges = len(self.placeholders.get('paragraph', []))
        self.logger.info("[Stage 7/8: PDF Merge Processing]")
        if num_merges == 0:
            self.logger.info("  > No merges to process. Skipping.")
            source_pdf = self.overlay_pdf_path if os.path.exists(self.overlay_pdf_path) else self.temp_pdf_path
            if os.path.exists(source_pdf):
                self.file_manager.copy_file(source_pdf, self.final_pdf_path)
            return True

        self.logger.info("  > Processing %d merge(s)...", num_merges)
        source_pdf = self.overlay_pdf_path
        success = self.merge_processor.process_merges(
            source_pdf, self.content_map, self.toc_pages, self.final_pdf_path
        )
        if not success:
            self.logger.error("  > ❌ Merge processing failed.")
            return False
        self.logger.info("  > Merge processing complete.")
        return True

    def _finalize_pdf(self) -> bool:
        """[Stage 8/8: Finalization] Clean up placeholder markers from the final PDF."""
        self.logger.info("[Stage 8/8: Finalization]")
        
        pdf_to_clean = self.final_pdf_path
        if not os.path.exists(pdf_to_clean) and os.path.exists(self.overlay_pdf_path):
             pdf_to_clean = self.overlay_pdf_path
        elif not os.path.exists(pdf_to_clean):
            pdf_to_clean = self.temp_pdf_path

        if self.placeholders['total'] == 0:
            self.logger.info("  > No placeholders were used, skipping cleanup.")
            if pdf_to_clean != self.final_pdf_path:
                 self.file_manager.copy_file(pdf_to_clean, self.final_pdf_path)
            self.logger.info("  > Final report is ready: %s", self.final_pdf_path)
            return True

        self.logger.info("  > Removing placeholder markers from the final report...")
        cleaned_pdf_path = self.file_manager.generate_temp_path(self.output_path, "cleaned")
        
        success = self.marker_remover.remove_markers(
            pdf_to_clean,
            list(self.content_map.keys()),
            cleaned_pdf_path
        )
        
        if not success:
            self.logger.warning("  > ⚠️ Failed to remove placeholder markers. The final file may contain them.")
            self.file_manager.copy_file(pdf_to_clean, self.final_pdf_path)
            return True # Not a fatal error, just a warning

        self.file_manager.move_file(cleaned_pdf_path, self.final_pdf_path)
        self.logger.info("  > Cleanup complete.")
        self.logger.info("  > Final report is ready: %s", self.final_pdf_path)
        return True
