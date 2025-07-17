"""
Report compiler core orchestration module.

This module contains the main ReportCompiler class that orchestrates the entire
report compilation process, from input validation through PDF generation.
"""

import os
import sys
from typing import Dict, Any, Set

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
    
    def __init__(self, input_path: str, output_path: str, keep_temp: bool = False, recursion_level: int = 0, file_manager: FileManager = None):
        """
        Initialize the report compiler.
        
        Args:
            input_path: Path to input DOCX file.
            output_path: Path for output PDF file.
            keep_temp: Whether to keep temporary files for debugging.
            recursion_level: The current recursion depth.
            file_manager: An existing file manager instance for shared state.
        """
        self.input_path = os.path.abspath(input_path)
        self.output_path = os.path.abspath(output_path)
        self.keep_temp = keep_temp
        self.recursion_level = recursion_level
        self.logger = get_compiler_logger()
        
        # Initialize components
        self.file_manager = file_manager if file_manager else FileManager(keep_temp)
        self.validators = Validators()
        self.placeholder_parser = PlaceholderParser()
        self.content_analyzer = ContentAnalyzer()
        self.docx_processor = DocxProcessor()
        self.word_converter = WordConverter()
        self.libreoffice_converter = LibreOfficeConverter()
        self.overlay_processor = OverlayProcessor()
        self.merge_processor = MergeProcessor()
        self.marker_remover = MarkerRemover()
        
        # Process state
        self.placeholders = {}
        self.base_directory = os.path.dirname(self.input_path)
        self.table_metadata = {}

        # File paths
        self.temp_docx_path = None
        self.temp_pdf_path = None
        self.final_pdf_path = None
        self.overlay_pdf_path = None
        self.merged_pdf_path = None
        self.toc_pages = []
        self.content_map = {}

    def _log_prefix(self) -> str:
        """Provides a prefix for logging based on recursion depth."""
        return f"  " * self.recursion_level

    def run(self, processed_files: Set[str] = None) -> bool:
        """Run the complete report compilation process."""
        if processed_files is None:
            processed_files = set()

        # Cycle detection
        if self.input_path in processed_files:
            self.logger.error(f"{self._log_prefix()}❌ Circular reference detected. Aborting compilation for: {self.input_path}")
            return False
        processed_files.add(self.input_path)

        self.logger.info(f"--- Starting Compilation for: {os.path.basename(self.input_path)} (Depth: {self.recursion_level}) ---")

        try:
            # The 'with' statement for file_manager is only used by the top-level call
            if self.recursion_level == 0:
                with self.file_manager:
                    result = self._execute_pipeline(processed_files)
            else:
                result = self._execute_pipeline(processed_files)
            
            if result:
                self.logger.info(f"--- Finished Compilation for: {os.path.basename(self.input_path)} ---")
            else:
                self.logger.error(f"--- Failed Compilation for: {os.path.basename(self.input_path)} ---")
            
            return result

        except Exception as e:
            self.logger.error(f"{self._log_prefix()}❌ A critical error occurred: %s", e, exc_info=True)
            return False
        finally:
            # Remove from set so sibling branches in the recursion tree can refer to this file
            if self.input_path in processed_files:
                processed_files.remove(self.input_path)

    def _execute_pipeline(self, processed_files: Set[str]) -> bool:
        """The core sequence of compilation steps."""
        if not self._initialize(): return False
        if not self._validate_paths(): return False
        if not self._find_placeholders(): return False
        if not self._resolve_docx_inserts(processed_files): return False
        if not self._validate_placeholders(): return False
        if not self._modify_docx(): return False
        if not self._convert_to_pdf(): return False
        if not self._analyze_base_pdf(): return False
        if not self._process_pdf_overlays(): return False
        if not self._process_pdf_merges(): return False
        if not self._finalize_pdf(): return False
        if self.recursion_level == 0:
            self.logger.info(f"\n{self._log_prefix()}=== Report Compilation Successful ===")
        return True

    def _initialize(self) -> bool:
        """[Stage 1/11: Initialization] Set up temporary file paths and initial state."""
        self.logger.info(f"{self._log_prefix()}[Stage 1/11: Initialization]")
        self.temp_docx_path = self.file_manager.generate_temp_path(self.input_path, "modified_report")
        # Use a more specific name for the base PDF to avoid collisions in recursion
        base_pdf_suffix = f"base_{self.recursion_level}"
        self.temp_pdf_path = self.file_manager.generate_temp_path(self.output_path, base_pdf_suffix)
        self.overlay_pdf_path = self.file_manager.generate_temp_path(self.output_path, f"with_overlays_{self.recursion_level}")
        self.final_pdf_path = self.output_path
        self.logger.debug(f"{self._log_prefix()}  > Input DOCX: {self.input_path}")
        self.logger.debug(f"{self._log_prefix()}  > Output PDF: {self.final_pdf_path}")
        self.logger.debug(f"{self._log_prefix()}  > Temp DOCX: {self.temp_docx_path}")
        self.logger.debug(f"{self._log_prefix()}  > Temp PDF (Base): {self.temp_pdf_path}")
        self.logger.debug(f"{self._log_prefix()}  > Temp PDF (Overlaid): {self.overlay_pdf_path}")
        self.logger.info(f"{self._log_prefix()}  > Environment initialized.")
        return True

    def _validate_paths(self) -> bool:
        """[Stage 2/11: Path Validation] Validate input and output paths."""
        self.logger.info(f"{self._log_prefix()}[Stage 2/11: Path Validation]")
        
        self.logger.debug(f"{self._log_prefix()}  > Validating source DOCX file...")
        docx_result = self.validators.validate_docx_path(self.input_path)
        if not docx_result['valid']:
            self.logger.error(f"{self._log_prefix()}  > ❌ {docx_result['error_message']}")
            return False
        self.logger.info(f"{self._log_prefix()}  > Source DOCX is valid (%.1f MB).", docx_result['file_size_mb'])

        self.logger.debug(f"{self._log_prefix()}  > Validating output path...")
        output_result = self.validators.validate_output_path(self.output_path)
        if not output_result['valid']:
            self.logger.error(f"{self._log_prefix()}  > ❌ {output_result['error_message']}")
            return False
        if output_result['file_exists'] and self.recursion_level == 0:
            self.logger.warning(f"{self._log_prefix()}  > ⚠️ Output file exists and will be overwritten.")
        self.logger.debug(f"{self._log_prefix()}  > Output path is valid.")
        return True

    def _find_placeholders(self) -> bool:
        """[Stage 3/11: Find Placeholders] Scan document for placeholders."""
        self.logger.info(f"{self._log_prefix()}[Stage 3/11: Find Placeholders]")
        self.logger.info(f"{self._log_prefix()}  > Scanning document for placeholders...")
        self.placeholders = self.placeholder_parser.find_all_placeholders(self.input_path)
        if self.placeholders['total'] == 0:
            self.logger.info(f"{self._log_prefix()}  > No placeholders found.")
        else:
            self.logger.info(f"{self._log_prefix()}  > Found {self.placeholders['total']} placeholders.")
        return True

    def _resolve_docx_inserts(self, processed_files: Set[str]) -> bool:
        """[Stage 4/11: Recursive DOCX Resolution] Recursively compile DOCX inserts to PDF."""
        self.logger.info(f"{self._log_prefix()}[Stage 4/11: Recursive DOCX Resolution]")
        
        placeholder_list = self.placeholders.get('table', []) + self.placeholders.get('paragraph', [])
        docx_inserts = [p for p in placeholder_list if p.get('is_recursive_docx')]
        
        if not docx_inserts:
            self.logger.info(f"{self._log_prefix()}  > No recursive DOCX inserts found. Skipping.")
            return True

        self.logger.info(f"{self._log_prefix()}  > Found {len(docx_inserts)} DOCX inserts to resolve.")

        for placeholder in docx_inserts:
            # The path from the parser is relative to the document, so make it absolute.
            docx_path = os.path.abspath(os.path.join(self.base_directory, placeholder['file_path']))
            
            self.logger.info(f"{self._log_prefix()}  > Resolving insert: {os.path.basename(docx_path)}")

            # Generate a unique temp PDF path for the compiled DOCX.
            # The extension must be .pdf for the rest of the process.
            temp_output_pdf = self.file_manager.generate_temp_path(
                docx_path, f"recursive_{self.recursion_level}"
            ).replace(".docx", ".pdf")

            # Create and run a sub-compiler.
            sub_compiler = ReportCompiler(
                input_path=docx_path,
                output_path=temp_output_pdf,
                keep_temp=self.keep_temp,
                recursion_level=self.recursion_level + 1,
                file_manager=self.file_manager
            )

            if not sub_compiler.run(processed_files):
                self.logger.error(f"{self._log_prefix()}  > ❌ Failed to compile recursive DOCX: {docx_path}")
                return False

            # Update placeholder to point to the new PDF. The path must be absolute.
            self.logger.info(f"{self._log_prefix()}  > ✓ Successfully compiled {os.path.basename(docx_path)} to {os.path.basename(temp_output_pdf)}")
            placeholder['file_path'] = temp_output_pdf
            placeholder['is_recursive_docx'] = False
            placeholder['original_path'] = docx_path # Keep track for debugging

        self.logger.info(f"{self._log_prefix()}  > All DOCX inserts resolved.")
        return True

    def _validate_placeholders(self) -> bool:
        """[Stage 5/11: Validate Placeholders] Validate placeholder paths and parameters."""
        self.logger.info(f"{self._log_prefix()}[Stage 5/11: Validate Placeholders]")
        
        if self.placeholders['total'] == 0:
            self.logger.info(f"{self._log_prefix()}  > No placeholders to validate.")
            return True

        self.logger.info(f"{self._log_prefix()}  > Validating paths and parameters for {self.placeholders['total']} placeholders...")
        placeholder_list = self.placeholders['table'] + self.placeholders['paragraph']
        
        # After recursion, all file paths are absolute, so base_directory is not strictly needed
        # but the validator can handle both absolute and relative paths.
        validation_result = self.validators.validate_placeholders(placeholder_list, self.base_directory)
        
        if not validation_result['valid']:
            for error in validation_result['errors']:
                self.logger.error(f"{self._log_prefix()}  > ❌ {error}")
            return False
            
        self.logger.info(f"{self._log_prefix()}  > All placeholders are valid.")
        return True

    def _modify_docx(self) -> bool:
        """[Stage 6/11: DOCX Modification] Create a modified DOCX with placeholders handled."""
        self.logger.info(f"{self._log_prefix()}[Stage 6/11: DOCX Modification]")
        if self.placeholders['total'] == 0:
            self.logger.info(f"{self._log_prefix()}  > No placeholders to process. Copying original document for conversion.")
            self.file_manager.copy_file(self.input_path, self.temp_docx_path)
            self.table_metadata = {}
            return True

        self.logger.info(f"{self._log_prefix()}  > Inserting markers into DOCX for {self.placeholders['total']} placeholders...")
        table_metadata = self.docx_processor.create_modified_docx(
            self.input_path, self.placeholders, self.temp_docx_path
        )

        if table_metadata is None:
            self.logger.error(f"{self._log_prefix()}  > ❌ Failed to create modified DOCX.")
            return False

        self.table_metadata = table_metadata
        self.logger.info(f"{self._log_prefix()}  > Modified DOCX created successfully.")
        return True

    def _convert_to_pdf(self) -> bool:
        """[Stage 7/11: PDF Conversion] Convert the DOCX to a base PDF."""
        self.logger.info(f"{self._log_prefix()}[Stage 7/11: PDF Conversion]")
        use_libreoffice = False
        if self.word_converter.is_available():
            self.logger.info(f"{self._log_prefix()}  > Attempting conversion with MS Word...")
            success = self.word_converter.update_fields_and_save_as_pdf(
                self.temp_docx_path, self.temp_pdf_path
            )
            if not success:
                self.logger.warning(f"{self._log_prefix()}  > MS Word conversion failed. Falling back to LibreOffice.")
                use_libreoffice = True
            else:
                self.logger.info(f"{self._log_prefix()}  > ✓ Conversion with MS Word successful.")
        else:
            self.logger.info(f"{self._log_prefix()}  > MS Word not available. Using LibreOffice for conversion.")
            use_libreoffice = True

        if use_libreoffice:
            if not self.libreoffice_converter.is_available():
                self.logger.error(f"{self._log_prefix()}  > ❌ Neither MS Word nor LibreOffice is available for PDF conversion.")
                return False
            self.logger.info(f"{self._log_prefix()}  > Attempting conversion with LibreOffice...")
            success = self.libreoffice_converter.convert_to_pdf(
                self.temp_docx_path, os.path.dirname(self.temp_pdf_path)
            )
            if not success:
                self.logger.error(f"{self._log_prefix()}  > ❌ LibreOffice conversion failed.")
                return False
            # LibreOffice may create a file with a different name, so we find it and move it.
            expected_pdf = os.path.splitext(self.temp_docx_path)[0] + ".pdf"
            if os.path.exists(expected_pdf):
                 self.file_manager.move_file(expected_pdf, self.temp_pdf_path)
            else:
                self.logger.error(f"{self._log_prefix()}  > ❌ LibreOffice did not produce the expected PDF file: {expected_pdf}")
                return False
            self.logger.info(f"{self._log_prefix()}  > ✓ Conversion with LibreOffice successful.")
        
        self.logger.info(f"{self._log_prefix()}  > Base PDF created successfully.")
        return True

    def _analyze_base_pdf(self) -> bool:
        """[Stage 8/11: PDF Analysis] Find markers and content in the base PDF."""
        self.logger.info(f"{self._log_prefix()}[Stage 8/11: PDF Analysis]")
        if self.placeholders['total'] == 0:
            self.logger.info(f"{self._log_prefix()}  > No placeholders were processed. Skipping analysis.")
            # If there are no placeholders, the temp_pdf is the final document.
            self.file_manager.copy_file(self.temp_pdf_path, self.final_pdf_path)
            return True

        self.logger.info(f"{self._log_prefix()}  > Analyzing base PDF for content and markers...")
        analysis_result = self.content_analyzer.analyze(
            self.temp_pdf_path, self.placeholders, self.table_metadata
        )
        if not analysis_result:
            self.logger.error(f"{self._log_prefix()}  > ❌ PDF analysis failed.")
            return False
        
        self.content_map, self.toc_pages = analysis_result
        self.logger.info(f"{self._log_prefix()}  > PDF analysis complete.")
        self.logger.debug(f"{self._log_prefix()}  > Content Map: {self.content_map}")
        self.logger.debug(f"{self._log_prefix()}  > TOC Pages: {self.toc_pages}")
        return True

    def _process_pdf_overlays(self) -> bool:
        """[Stage 9/11: PDF Overlays] Apply overlays to the base PDF."""
        self.logger.info(f"{self._log_prefix()}[Stage 9/11: PDF Overlays]")
        
        table_placeholders = self.placeholders.get('table', [])
        if not table_placeholders:
            self.logger.info(f"{self._log_prefix()}  > No table placeholders to process as overlays. Skipping.")
            self.file_manager.copy_file(self.temp_pdf_path, self.overlay_pdf_path)
            return True

        self.logger.info(f"{self._log_prefix()}  > Processing overlays...")
        success = self.overlay_processor.process_overlays(
            base_pdf_path=self.temp_pdf_path,
            output_path=self.overlay_pdf_path,
            content_map=self.content_map
        )
        if not success:
            self.logger.error(f"{self._log_prefix()}  > ❌ Failed to process overlays.")
            return False
        
        self.logger.info(f"{self._log_prefix()}  > Overlays processed successfully.")
        return True

    def _process_pdf_merges(self) -> bool:
        """[Stage 10/11: PDF Merging] Merge inserted PDFs into the main document."""
        self.logger.info(f"{self._log_prefix()}[Stage 10/11: PDF Merging]")
        
        paragraph_placeholders = self.placeholders.get('paragraph', [])
        if not paragraph_placeholders:
            self.logger.info(f"{self._log_prefix()}  > No paragraph placeholders to merge. Skipping.")
            # The overlay_pdf_path is carried forward to the finalization step
            return True

        if not self.content_map:
            self.logger.error(f"{self._log_prefix()}  > ❌ Paragraph placeholders exist, but content map is empty. Analysis likely failed.")
            return False

        self.logger.info(f"{self._log_prefix()}  > Merging PDF inserts...")
        # The final PDF before marker removal is created here.
        self.merged_pdf_path = self.file_manager.generate_temp_path(self.output_path, "merged")
        
        success = self.merge_processor.process_merges(
            base_pdf_path=self.overlay_pdf_path,
            output_path=self.merged_pdf_path,
            content_map=self.content_map,
            toc_pages=self.toc_pages
        )
        if not success:
            self.logger.error(f"{self._log_prefix()}  > ❌ Failed to merge PDFs.")
            return False
        
        self.logger.info(f"{self._log_prefix()}  > PDF inserts merged successfully.")
        return True

    def _finalize_pdf(self) -> bool:
        """[Stage 11/11: Finalization] Remove markers and save the final PDF."""
        self.logger.info(f"{self._log_prefix()}[Stage 11/11: Finalization]")
        
        # If no placeholders were processed, the final PDF is already in place.
        if self.placeholders['total'] == 0:
            self.logger.info(f"{self._log_prefix()}  > No markers to remove. Final PDF is ready.")
            if self.recursion_level > 0:
                 self.file_manager.copy_file(self.temp_pdf_path, self.output_path)
            return True

        # Determine the source PDF for marker removal
        source_pdf_for_removal = self.merged_pdf_path if self.merged_pdf_path else self.overlay_pdf_path

        self.logger.info(f"{self._log_prefix()}  > Removing markers from the final PDF...")
        all_markers = list(self.content_map.keys())
        success = self.marker_remover.remove_markers(
            input_pdf_path=source_pdf_for_removal,
            markers=all_markers,
            output_pdf_path=self.final_pdf_path
        )
        if not success:
            self.logger.error(f"{self._log_prefix()}  > ❌ Failed to remove markers from the final PDF.")
            return False
        
        self.logger.info(f"{self._log_prefix()}  > ✓ Final PDF created at: {self.final_pdf_path}")
        return True
