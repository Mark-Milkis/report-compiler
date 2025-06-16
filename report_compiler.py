#!/usr/bin/env python3
"""
Python PDF Report Compiler using DOCX Modification and PDF Overlay

This script automates the creation of a final PDF report by combining a main Word document 
with multiple PDF appendices. It modifies the Word document to create blank placeholder pages,
converts it to PDF, and then overlays the appendix content onto these blank pages.
"""

import os
import re
import argparse
import win32com.client as win32
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_BREAK
import fitz  # PyMuPDF
import time


class ReportCompiler:
    """
    A class to compile Word documents with PDF appendices into a single PDF report.
    
    The process involves:
    1. Finding PDF insertion placeholders in the Word document
    2. Modifying the document to create blank pages with hidden markers
    3. Converting the modified document to PDF
    4. Overlaying the appendix PDFs onto the blank pages
    
    Supported placeholder types:
    - Table-based overlays: Single-cell tables containing [[INSERT: path.pdf]]
    - Paragraph-based merges: Regular paragraphs containing [[INSERT: path.pdf]]
    """
    
    def __init__(self, input_docx_path, final_pdf_path, keep_temp=False):
        """
        Initialize the ReportCompiler with input and output paths.
        
        Args:
            input_docx_path (str): Absolute path to the source .docx file
            final_pdf_path (str): Absolute path for the final output .pdf file
            keep_temp (bool): Whether to keep temporary files for debugging
        """
        self.input_docx_path = os.path.abspath(input_docx_path)
        self.final_pdf_path = os.path.abspath(final_pdf_path)
        self.keep_temp = keep_temp
          # Store table coordinate metadata for overlay positioning
        self.table_coordinates = {}  # Maps table index to position/size info
        
        # Create unique temporary file names using timestamp
        timestamp = str(int(time.time() * 1000))
        base_dir = os.path.dirname(self.input_docx_path)
        self.temp_docx_path = os.path.join(base_dir, f"~temp_modified_report_{timestamp}.docx")
        self.temp_pdf_path = os.path.join(base_dir, f"~temp_base_{timestamp}.pdf")
        
        # Compile regex for finding placeholders
        self.placeholder_regex = re.compile(r"\[\[INSERT:\s*(.*?)\s*\]\]")
        
        print(f"Input DOCX: {self.input_docx_path}")
        print(f"Output PDF: {self.final_pdf_path}")
        print(f"Temp DOCX: {self.temp_docx_path}")
        print(f"Temp PDF: {self.temp_pdf_path}")
    
    def run(self):
        """
        Main public method that executes the entire workflow.
        
        The workflow:
        1. Find placeholders and modify the DOCX
        2. Convert modified DOCX to PDF
        3. Overlay appendix PDFs onto the base PDF
        4. Clean up temporary files
        """
        try:
            print("\n=== Starting Report Compilation ===")
            
            # Step 1: Find placeholders and modify DOCX
            print("\nStep 1: Analyzing document for PDF placeholders...")
            modified_doc, placeholders = self._find_placeholders_and_modify_docx()
            
            if not placeholders:
                print("No PDF placeholders found. Converting original document to PDF...")
                self._convert_docx_to_pdf(self.input_docx_path, self.final_pdf_path)
                print(f"✓ Conversion complete: {self.final_pdf_path}")
                return
            
            print(f"Found {len(placeholders)} PDF placeholder(s)")
            
            # Step 2: Save modified document and convert to PDF
            print("\nStep 2: Saving modified document...")
            modified_doc.save(self.temp_docx_path)
            print(f"✓ Modified document saved: {self.temp_docx_path}")
            
            print("\nStep 3: Converting modified document to PDF...")
            self._convert_docx_to_pdf(self.temp_docx_path, self.temp_pdf_path)
            print(f"✓ Base PDF created: {self.temp_pdf_path}")
              # Step 3: Overlay PDFs (only overlay placeholders)
            print("\nStep 4: Overlaying appendix PDFs...")
            overlay_placeholders = [p for p in placeholders if p.get('type') == 'overlay']
            if overlay_placeholders:
                self._overlay_pdfs(overlay_placeholders)
                print(f"✓ Final PDF created: {self.final_pdf_path}")
            else:
                print("   • No overlay placeholders to process")
                # Copy base PDF to final PDF
                import shutil
                shutil.copy2(self.temp_pdf_path, self.final_pdf_path)
                print(f"✓ Final PDF copied: {self.final_pdf_path}")            
            print("\n=== Report Compilation Complete ===")
            
        finally:
            # Always clean up temporary files
            self._cleanup()
    
    def _find_placeholders_and_modify_docx(self):
        """
        Find PDF placeholders and sort them into two types: overlay and merge insertions.
        
        Overlay insertions: Placeholders found inside single-cell tables - use table size/position
        Merge insertions: Placeholders found in regular paragraphs - full page inserts
        
        Returns:
            tuple: (modified_doc, placeholders_list)
                - modified_doc: The modified docx.Document object
                - placeholders_list: List of dictionaries containing placeholder info
        """
        print("🔍 PHASE 1: Scanning document for INSERT placeholders...")
        
        # Step 1: Scan for table placeholders (overlay type)
        print("\n📋 Scanning tables for overlay placeholders...")
        table_placeholders = self._find_table_placeholders()
        
        # Step 2: Scan for paragraph placeholders (merge type)
        print("\n📄 Scanning paragraphs for merge placeholders...")
        paragraph_placeholders = self._find_paragraph_placeholders()
        
        # Step 3: Validate and count pages for all placeholders
        print(f"\n✅ VALIDATION: Found {len(table_placeholders)} table + {len(paragraph_placeholders)} paragraph placeholders")
        all_placeholders = table_placeholders + paragraph_placeholders
        validated_placeholders = self._validate_pdf_references(all_placeholders)
        
        # Step 4: Sort placeholders by type for processing
        overlay_placeholders = [p for p in validated_placeholders if p['type'] == 'overlay']
        merge_placeholders = [p for p in validated_placeholders if p['type'] == 'merge']
        
        print(f"\n📊 PROCESSING SUMMARY:")
        print(f"   • Overlay insertions (table-based): {len(overlay_placeholders)}")
        print(f"   • Merge insertions (paragraph-based): {len(merge_placeholders)}")
        
        if not validated_placeholders:
            print("⚠️  No valid placeholders found - document will be converted as-is")
            return Document(self.input_docx_path), []
        
        # Step 5: Modify document based on placeholder types
        print(f"\n🔧 PHASE 2: Modifying document...")
        modified_doc = self._create_modified_document(overlay_placeholders, merge_placeholders)
        
        return modified_doc, validated_placeholders
    
    def _find_table_placeholders(self):
        """
        Find PDF placeholders inside single-cell tables (overlay type).
        
        Single-cell tables containing INSERT placeholders are considered overlay inserts
        because they can be precisely positioned and sized.
        
        Returns:
            list: List of dictionaries containing table placeholder info
        """
        placeholders = []
        
        try:
            doc = Document(self.input_docx_path)
            
            for table_idx, table in enumerate(doc.tables):
                rows = len(table.rows)
                cols = len(table.columns)
                
                # Only consider single-cell tables for overlay inserts
                if rows == 1 and cols == 1:
                    cell = table.rows[0].cells[0]
                    cell_text = cell.text.strip()
                    
                    # Check if this cell contains an INSERT placeholder
                    match = self.placeholder_regex.search(cell_text)
                    if match:
                        pdf_path_raw = match.group(1).strip()
                        
                        print(f"   📋 Found table placeholder #{len(placeholders)+1}:")
                        print(f"      • Raw path: {pdf_path_raw}")
                        print(f"      • Table index: {table_idx}")
                        print(f"      • Table type: Single-cell (1x1)")
                        print(f"      • Cell text: '{cell_text}'")
                        
                        # Try to get table dimensions
                        dimensions = self._get_table_dimensions(table, table_idx)
                        
                        table_info = {
                            'type': 'overlay',
                            'pdf_path_raw': pdf_path_raw,
                            'table_index': table_idx,
                            'table_text': cell_text,
                            'source': f'table_{table_idx}',
                            'insert_method': 'table'
                        }
                        
                        if dimensions:
                            table_info.update(dimensions)
                            if 'width_inches' in dimensions and 'height_inches' in dimensions:
                                print(f"      • Dimensions: {dimensions['width_inches']:.2f}\" x {dimensions['height_inches']:.2f}\"")
                            else:
                                print(f"      • Dimensions: {dimensions}")
                        else:
                            print(f"      • ⚠️ Could not determine table dimensions")
                        
                        placeholders.append(table_info)
                
                else:
                    # Multi-cell tables: scan all cells but don't classify as overlay
                    # This prevents inline text within tables from being misclassified
                    has_insert = False
                    for row in table.rows:
                        for cell in row.cells:
                            if self.placeholder_regex.search(cell.text):
                                has_insert = True
                                break
                        if has_insert:
                            break
                    
                    if has_insert:
                        print(f"   ⚠️  Multi-cell table #{table_idx} ({rows}x{cols}) contains INSERT but skipped (not overlay type)")
        
        except Exception as e:
            print(f"   ❌ Error scanning for table placeholders: {e}")
        
        print(f"   ✅ Found {len(placeholders)} table placeholders")
        return placeholders
    
    def _get_table_dimensions(self, table, table_idx):
        """
        Extract dimensions and position information from a table element.
        Enhanced for Word-to-PDF coordinate mapping.
        
        Returns:
            dict: Dictionary with width/height information or None
        """
        try:
            # Method 1: Check table preferred width
            dimensions = {}
            
            if hasattr(table, '_tbl'):
                tbl_element = table._tbl
                tbl_pr = tbl_element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblPr')
                if tbl_pr is not None:
                    # Look for width information
                    tbl_w = tbl_pr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblW')
                    if tbl_w is not None:
                        width_type = tbl_w.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'unknown')
                        width_val = tbl_w.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', 'unknown')
                        
                        if width_type == 'dxa' and width_val != 'unknown':
                            # Convert from twentieths of a point to inches
                            width_inches = int(width_val) / 1440.0
                            dimensions['width_inches'] = width_inches
                            dimensions['width_type'] = width_type
                            dimensions['width_raw'] = width_val
                    
                    # Look for table positioning information
                    tbl_pos = tbl_pr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblpPr')
                    if tbl_pos is not None:
                        # Table has absolute positioning
                        h_anchor = tbl_pos.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}horzAnchor', 'margin')
                        v_anchor = tbl_pos.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}vertAnchor', 'margin')
                        tbl_px = tbl_pos.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblpX')
                        tbl_py = tbl_pos.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblpY')
                        
                        dimensions['has_absolute_position'] = True
                        dimensions['h_anchor'] = h_anchor
                        dimensions['v_anchor'] = v_anchor
                        if tbl_px:
                            dimensions['pos_x_twips'] = int(tbl_px)
                            dimensions['pos_x_inches'] = int(tbl_px) / 1440.0
                        if tbl_py:
                            dimensions['pos_y_twips'] = int(tbl_py)
                            dimensions['pos_y_inches'] = int(tbl_py) / 1440.0
                    
                    # Look for table margins/indentation
                    tbl_ind = tbl_pr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblInd')
                    if tbl_ind is not None:
                        ind_type = tbl_ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'dxa')
                        ind_val = tbl_ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w')
                        if ind_val and ind_type == 'dxa':
                            dimensions['indent_twips'] = int(ind_val)
                            dimensions['indent_inches'] = int(ind_val) / 1440.0
            
            # Method 2: Check column widths
            if table.columns:
                try:
                    first_col = table.columns[0]
                    if hasattr(first_col, 'width') and first_col.width:
                        col_width_inches = first_col.width.inches
                        dimensions['column_width_inches'] = col_width_inches
                except:
                    pass
            
            # Method 3: Try to get row height
            if table.rows:
                try:
                    first_row = table.rows[0]
                    if hasattr(first_row, 'height') and first_row.height:
                        row_height_inches = first_row.height.inches
                        dimensions['row_height_inches'] = row_height_inches
                except:
                    pass
              # Store coordinate metadata for precise overlay positioning
            if dimensions:
                self.table_coordinates[table_idx] = {
                    'dimensions': dimensions,
                    'word_table_index': table_idx,
                    'extraction_method': 'word_xml_analysis'
                }
                print(f"      📍 Stored coordinate metadata for table {table_idx}")
            
            return dimensions if dimensions else None
            
        except Exception as e:
            print(f"      ⚠️ Error getting table dimensions: {e}")
            return None
    
    def _find_paragraph_placeholders(self):
        """
        Find PDF placeholders in regular document paragraphs (merge type).
        
        Returns:
            list: List of dictionaries containing paragraph placeholder info
        """
        placeholders = []
        doc = Document(self.input_docx_path)
        
        for para_idx, paragraph in enumerate(doc.paragraphs):
            match = self.placeholder_regex.search(paragraph.text)
            if match:
                pdf_path_raw = match.group(1).strip()
                
                print(f"   📄 Found paragraph placeholder #{len(placeholders)+1}:")
                print(f"      • Raw path: {pdf_path_raw}")
                print(f"      • Paragraph index: {para_idx}")
                
                placeholders.append({
                    'type': 'merge',
                    'pdf_path_raw': pdf_path_raw,
                    'paragraph_index': para_idx,
                    'source': f'paragraph_{para_idx}'
                })
        
        print(f"   ✅ Found {len(placeholders)} paragraph placeholders")
        return placeholders
    
    def _validate_pdf_references(self, placeholders):
        """
        Validate PDF file existence and count pages for all placeholders.
        
        Args:
            placeholders (list): List of placeholder dictionaries
            
        Returns:
            list: List of validated placeholders with page counts
        """
        validated = []
        input_dir = os.path.dirname(self.input_docx_path)
        
        print(f"\n🔍 VALIDATING {len(placeholders)} PDF references...")
        
        for i, placeholder in enumerate(placeholders):
            pdf_path_raw = placeholder['pdf_path_raw']
            
            print(f"\n   📋 Placeholder #{i+1} ({placeholder['type']}):")
            print(f"      • Raw path: {pdf_path_raw}")
            
            # Resolve relative path
            if not os.path.isabs(pdf_path_raw):
                pdf_path = os.path.join(input_dir, pdf_path_raw)
            else:
                pdf_path = pdf_path_raw
            
            pdf_path = os.path.abspath(pdf_path)
            print(f"      • Resolved path: {pdf_path}")
            
            # Check if file exists
            if not os.path.exists(pdf_path):
                print(f"      ❌ ERROR: PDF file not found")
                continue
            
            # Get page count
            try:
                pdf_doc = fitz.open(pdf_path)
                page_count = pdf_doc.page_count
                pdf_doc.close()
                print(f"      ✅ Valid PDF with {page_count} page(s)")
                  # Add resolved info to placeholder
                placeholder['pdf_path'] = pdf_path
                placeholder['page_count'] = page_count
                placeholder['index'] = len(validated)
                
                validated.append(placeholder)
                
            except Exception as e:
                print(f"      ❌ ERROR: Cannot read PDF file: {e}")
                continue
        
        print(f"\n✅ Validated {len(validated)}/{len(placeholders)} placeholders")
        return validated
    
    def _create_modified_document(self, overlay_placeholders, merge_placeholders):
        """
        Create a modified version of the document with insertion markers.
        
        Args:
            overlay_placeholders (list): Table-based placeholders
            merge_placeholders (list): Paragraph-based placeholders
            
        Returns:
            Document: Modified document with markers
        """
        print("\n🔧 Creating modified document...")
        doc = Document(self.input_docx_path)
        
        # Process merge placeholders (paragraph-based)
        if merge_placeholders:
            print(f"\n📄 Processing {len(merge_placeholders)} merge placeholders...")
            self._process_merge_placeholders(doc, merge_placeholders)
        
        # Process overlay placeholders (table-based)
        if overlay_placeholders:
            print(f"\n📦 Processing {len(overlay_placeholders)} overlay placeholders...")
            self._process_overlay_placeholders(doc, overlay_placeholders)
        
        print("✅ Document modification complete")
        return doc
    
    def _process_merge_placeholders(self, doc, merge_placeholders):
        """
        Process merge (paragraph-based) placeholders by replacing with page breaks.
        
        Args:
            doc (Document): The document to modify
            merge_placeholders (list): List of merge placeholders
        """
        # Sort by paragraph index in reverse order to avoid index shifting
        sorted_placeholders = sorted(merge_placeholders, key=lambda x: x['paragraph_index'], reverse=True)
        
        for placeholder in sorted_placeholders:
            para_idx = placeholder['paragraph_index']
            page_count = placeholder['page_count']
            marker_text = f"%%MERGE_START_{placeholder['index']}%%"
            
            print(f"   📄 Processing merge placeholder #{placeholder['index']+1}:")
            print(f"      • Paragraph {para_idx}, {page_count} pages")
            print(f"      • Marker: {marker_text}")
            
            # Get the paragraph and clear it
            paragraph = doc.paragraphs[para_idx]
            paragraph.clear()
            
            # Add visible marker for the first page
            marker_run = paragraph.add_run(marker_text)
            marker_run.font.size = Pt(12)
            marker_run.font.color.rgb = RGBColor(255, 0, 0)  # Red
              # Add page breaks for each page in the PDF
            for i in range(page_count):
                if i > 0:  # Add page break before each page except the first
                    run = paragraph.add_run()
                    run.add_break(WD_BREAK.PAGE)
            
            # Add the marker to the placeholder for overlay processing            placeholder['marker'] = marker_text            
            print(f"      ✅ Added marker and {page_count} page breaks")
    
    def _process_overlay_placeholders(self, doc, overlay_placeholders):
        """
        Process table-based overlay placeholders by replacing table content with markers.
        
        Args:
            doc (Document): The document to modify
            overlay_placeholders (list): List of table-based overlay placeholders
        """
        print("   � Table-based overlay placeholder processing:")
        
        if not overlay_placeholders:
            return
        
        # All overlay placeholders should be table-based now
        self._process_table_placeholders(doc, overlay_placeholders)
    
    def _process_table_placeholders(self, doc, table_placeholders):
        """
        Process table-based overlay placeholders by replacing table content with markers.
        
        Args:
            doc (Document): The document to modify
            table_placeholders (list): List of table placeholders
        """
        # Sort by table index in reverse order to avoid index shifting issues
        sorted_placeholders = sorted(table_placeholders, key=lambda x: x['table_index'], reverse=True)
        
        for placeholder in sorted_placeholders:
            table_idx = placeholder['table_index']
            page_count = placeholder['page_count']
            marker_text = f"%%OVERLAY_START_{placeholder['index']}%%"
            
            print(f"      � Processing table placeholder #{placeholder['index']+1}:")
            print(f"         • Table {table_idx}, {page_count} pages")
            print(f"         • Marker: {marker_text}")
            
            try:
                # Get the table and modify its content
                table = doc.tables[table_idx]
                cell = table.rows[0].cells[0]
                  # Clear the cell and add our marker
                cell.text = marker_text
                
                # Add the marker to the placeholder for overlay processing
                placeholder['marker'] = marker_text
                
                # Optional: style the marker text
                if cell.paragraphs:
                    for paragraph in cell.paragraphs:
                        if paragraph.runs:
                            for run in paragraph.runs:
                                run.font.size = Pt(12)
                                # Make marker less visible
                                run.font.color.rgb = None  # Default color
                
                print(f"         ✅ Table {table_idx} updated with overlay marker")
                
            except Exception as e:
                print(f"         ❌ Error modifying table {table_idx}: {e}")
                print(f"         📝 Falling back to paragraph marker...")                # Fallback: add paragraph marker
                paragraph = doc.add_paragraph()
                marker_run = paragraph.add_run(marker_text)
                marker_run.font.size = Pt(12)
    
    def _convert_docx_to_pdf(self, input_path, output_path):
        """
        Convert a DOCX file to PDF using Microsoft Word automation.
        Uses tested conversion function with simplified parameters.
        
        Args:
            input_path (str): Path to the input DOCX file
            output_path (str): Path for the output PDF file
        """
        print(f"    Converting DOCX to PDF...")
        print(f"    Input: {input_path}")
        print(f"    Output: {output_path}")
        
        word = None
        doc = None
        
        try:
            # Get or create Word application
            try:
                word = win32.GetActiveObject("Word.Application")
                print("    ✓ Connected to existing Word instance")
            except:
                word = win32.Dispatch("Word.Application")
                word.Visible = False
                print("    ✓ Created new Word instance")
            
            # Open the document
            input_path = os.path.abspath(input_path)
            print(f"    Opening document: {input_path}")
            doc = word.Documents.Open(input_path)
            
            # Export parameters
            wdExportFormatPDF = 17
            wdExportOptimizeForPrint = 0
            wdExportCreateHeadingBookmarks = 1  # Create bookmarks from headings
            wdExportItem = 7  # Export entire document including markups
            
            output_path = os.path.abspath(output_path)
            print(f"    Exporting to PDF: {output_path}")
            
            # Export to PDF with minimal, tested parameters
            doc.ExportAsFixedFormat(
                OutputFileName=output_path,
                ExportFormat=wdExportFormatPDF,
                OpenAfterExport=False,
                OptimizeFor=wdExportOptimizeForPrint,
                Item=wdExportItem,
                CreateBookmarks=wdExportCreateHeadingBookmarks
            )
            
            print(f"    ✓ Successfully converted '{os.path.basename(input_path)}' to PDF")
            
        except Exception as e:
            print(f"    ⚠ ERROR during Word conversion: {e}")
            raise
            
        finally:
            # Clean up Word objects using tested approach
            try:
                if 'doc' in locals() and doc:
                    doc.Close(SaveChanges=False)
                    print("    ✓ Document closed")
            except:
                pass
            
            try:
                if word and word.Documents.Count == 0:
                    word.Quit()
                    print("    ✓ Word application closed")
            except:
                pass
    
    def _overlay_pdfs(self, placeholders):
        """
        Overlay appendix PDFs onto the base PDF using precise table positioning.
        
        Args:
            placeholders (list): List of placeholder dictionaries
        """
        print(f"    Opening base PDF: {self.temp_pdf_path}")
        base_pdf = fitz.open(self.temp_pdf_path)
        
        try:
            for placeholder in placeholders:
                marker = placeholder['marker']
                pdf_path = placeholder['pdf_path']
                page_count = placeholder['page_count']
                index = placeholder['index']
                
                print(f"    Processing appendix {index + 1}: {os.path.basename(pdf_path)}")
                print(f"      Searching for marker: {marker}")
                
                # Search for the marker in the PDF and get its position
                marker_info = self._find_marker_position(base_pdf, marker)
                
                if not marker_info:
                    print(f"      ⚠ WARNING: Marker not found in PDF, skipping appendix")
                    continue
                
                start_page_index = marker_info['page_index']
                marker_rect = marker_info['rect']
                
                print(f"      ✓ Found marker on page {start_page_index + 1}")
                print(f"      📐 Marker position: ({marker_rect.x0:.1f}, {marker_rect.y0:.1f}) to ({marker_rect.x1:.1f}, {marker_rect.y1:.1f})")
                
                # Calculate overlay rectangle based on table dimensions
                overlay_rect = self._calculate_overlay_rectangle(placeholder, marker_rect, base_pdf[start_page_index])
                
                print(f"      📋 Overlay rectangle: ({overlay_rect.x0:.1f}, {overlay_rect.y0:.1f}) to ({overlay_rect.x1:.1f}, {overlay_rect.y1:.1f})")
                print(f"      📏 Overlay size: {overlay_rect.width:.1f} x {overlay_rect.height:.1f} points")
                
                # Remove the marker text from the page
                page = base_pdf[start_page_index]
                page.add_redact_annot(marker_rect, fill=(1, 1, 1))  # White fill
                page.apply_redactions()
                print(f"      ✓ Removed marker text from page {start_page_index + 1}")

                # Open the appendix PDF
                print(f"      Opening appendix PDF: {pdf_path}")
                appendix_pdf = fitz.open(pdf_path)
                
                try:
                    # Overlay each page of the appendix
                    for i in range(page_count):
                        target_page_index = start_page_index + i
                        
                        if target_page_index >= base_pdf.page_count:
                            print(f"      ⚠ WARNING: Not enough pages in base PDF for appendix page {i + 1}")
                            break
                        
                        print(f"        Overlaying page {i + 1}/{page_count} -> base page {target_page_index + 1}")
                        
                        # Get the target page in the base PDF
                        target_page = base_pdf[target_page_index]
                        
                        # For the first page, use calculated overlay rectangle
                        # For subsequent pages, use full page overlay
                        if i == 0:
                            # Precise overlay within table boundaries
                            print(f"        📌 Precise overlay within table boundaries")
                            target_page.show_pdf_page(overlay_rect, appendix_pdf, i)
                        else:
                            # Full page overlay for additional pages
                            print(f"        📄 Full page overlay for continuation")
                            page_rect = target_page.rect
                            target_page.show_pdf_page(page_rect, appendix_pdf, i)
                    
                    print(f"      ✓ Appendix {index + 1} overlay complete")
                    
                finally:
                    appendix_pdf.close()
            
            # Save the final PDF
            print(f"    Saving final PDF: {self.final_pdf_path}")
            base_pdf.save(self.final_pdf_path)
            print("    ✓ Final PDF saved successfully")
            
        finally:
            base_pdf.close()
    
    def _find_marker_position(self, pdf_doc, marker):
        """
        Find the position of a marker in the PDF document.
        
        Args:
            pdf_doc: The PDF document to search
            marker (str): The marker text to find
            
        Returns:
            dict: Dictionary with page_index and rect, or None if not found
        """
        for page_num in range(pdf_doc.page_count):
            page = pdf_doc[page_num]
            text_instances = page.search_for(marker)
            
            if text_instances:
                # Return the first instance found
                return {
                    'page_index': page_num,
                    'rect': text_instances[0]                }
        
        return None
    
    def _calculate_overlay_rectangle(self, placeholder, marker_rect, page):
        """
        Calculate the overlay rectangle using Word-to-PDF coordinate mapping.
        
        Args:
            placeholder (dict): Placeholder information including table index
            marker_rect: Rectangle where the marker was found
            page: The PDF page object
            
        Returns:
            fitz.Rect: Rectangle for overlay positioning
        """
        # Get page dimensions
        page_rect = page.rect
        page_width = page_rect.width
        page_height = page_rect.height
        
        print(f"      📄 Page size: {page_width:.1f} x {page_height:.1f} points")
          # Get table index from placeholder
        table_index = placeholder.get('table_index')
        
        if table_index is not None and table_index in self.table_coordinates:
            print(f"      🎯 Using Word-to-PDF coordinate mapping for table {table_index}")
            
            # Use stored Word table coordinates
            table_coord_data = self.table_coordinates[table_index]
            dimensions = table_coord_data['dimensions']
            
            overlay_rect = self._convert_word_coords_to_pdf(dimensions, page_rect, marker_rect)
            
            if overlay_rect:
                print(f"      ✅ Converted Word coordinates to PDF coordinates")
                print(f"      📋 Overlay rectangle: ({overlay_rect.x0:.1f}, {overlay_rect.y0:.1f}) to ({overlay_rect.x1:.1f}, {overlay_rect.y1:.1f})")
                print(f"      📏 Overlay size: {overlay_rect.width:.1f} x {overlay_rect.height:.1f} points")
                return overlay_rect
            else:
                print(f"      ⚠️ Word-to-PDF coordinate conversion failed, falling back")
          # Fallback to table detection methods if coordinate mapping fails
        print(f"      🔄 Falling back to table detection methods...")
        
        # Try to detect table boundaries using PyMuPDF table detection
        table_rect = self._detect_table_boundaries(page, marker_rect)
        
        if table_rect:
            print(f"      📋 Detected table boundaries: ({table_rect.x0:.1f}, {table_rect.y0:.1f}) to ({table_rect.x1:.1f}, {table_rect.y1:.1f})")
            print(f"      📏 Detected table size: {table_rect.width:.1f} x {table_rect.height:.1f} points")
            
            # Use detected table boundaries as overlay rectangle
            overlay_rect = table_rect
            
            # Validate the detected rectangle makes sense
            if (overlay_rect.width > 50 and overlay_rect.height > 30 and 
                overlay_rect.width < page_width and overlay_rect.height < page_height):
                print(f"      ✅ Using detected table boundaries for overlay")
                return overlay_rect
            else:
                print(f"      ⚠️ Detected table boundaries seem invalid, falling back to dimension-based calculation")
        
        # Final fallback: Use table dimensions from Word document
        table_width_inches = placeholder.get('width_inches')
        table_height_inches = placeholder.get('row_height_inches') or placeholder.get('height_inches')
        
        if table_width_inches and table_height_inches:
            # Convert inches to points (1 inch = 72 points)
            table_width_points = table_width_inches * 72
            table_height_points = table_height_inches * 72
            
            print(f"      📋 Word table size: {table_width_inches:.2f}\" x {table_height_inches:.2f}\" ({table_width_points:.1f} x {table_height_points:.1f} points)")
            
            # Use marker position but apply smarter positioning logic
            marker_center_x = (marker_rect.x0 + marker_rect.x1) / 2
            marker_center_y = (marker_rect.y0 + marker_rect.y1) / 2
            
            # Estimate table position based on marker and typical table layouts
            # Assume marker is near the top-left of the table cell
            estimated_table_x0 = max(20, marker_rect.x0 - 20)  # Small offset from marker
            estimated_table_y0 = max(20, marker_rect.y0 - 20)  # Small offset from marker
            estimated_table_x1 = min(page_width - 20, estimated_table_x0 + table_width_points)
            estimated_table_y1 = min(page_height - 20, estimated_table_y0 + table_height_points)
            
            # Adjust if table goes beyond page boundaries
            if estimated_table_x1 > page_width - 20:
                estimated_table_x0 = page_width - 20 - table_width_points
                estimated_table_x1 = page_width - 20
            if estimated_table_y1 > page_height - 20:
                estimated_table_y0 = page_height - 20 - table_height_points
                estimated_table_y1 = page_height - 20
                
            overlay_rect = fitz.Rect(estimated_table_x0, estimated_table_y0, estimated_table_x1, estimated_table_y1)
            print(f"      📐 Estimated table position: ({overlay_rect.x0:.1f}, {overlay_rect.y0:.1f}) to ({overlay_rect.x1:.1f}, {overlay_rect.y1:.1f})")
            
        else:
            print(f"      ⚠️ Table dimensions not available, using marker-based estimation")
            # Fallback: Use marker position and estimate reasonable size
            default_width = min(400, page_width * 0.6)  # 400 points or 60% of page width
            default_height = min(300, page_height * 0.4)  # 300 points or 40% of page height
            
            marker_center_x = (marker_rect.x0 + marker_rect.x1) / 2
            marker_center_y = (marker_rect.y0 + marker_rect.y1) / 2
            
            overlay_x0 = max(20, marker_center_x - default_width / 2)
            overlay_y0 = max(20, marker_center_y - default_height / 2)
            overlay_x1 = min(page_width - 20, overlay_x0 + default_width)
            overlay_y1 = min(page_height - 20, overlay_y0 + default_height)
            
            overlay_rect = fitz.Rect(overlay_x0, overlay_y0, overlay_x1, overlay_y1)
            
        return overlay_rect
    
    def _convert_word_coords_to_pdf(self, word_dimensions, page_rect, marker_rect):
        """
        Convert Word table coordinates to PDF coordinates for precise overlay placement.
        
        Args:
            word_dimensions (dict): Table dimensions and position from Word
            page_rect: PDF page rectangle
            marker_rect: Where the marker was found in the PDF
            
        Returns:
            fitz.Rect: Converted PDF coordinates or None if conversion fails
        """
        try:
            print(f"      🔄 Converting Word coordinates to PDF coordinates...")
            
            # Get Word table dimensions
            width_inches = word_dimensions.get('width_inches') or word_dimensions.get('column_width_inches')
            height_inches = word_dimensions.get('row_height_inches')
            
            if not width_inches or not height_inches:
                print(f"      ❌ Missing required dimensions (width: {width_inches}, height: {height_inches})")
                return None
            
            # Convert to PDF points (1 inch = 72 points)
            width_points = width_inches * 72
            height_points = height_inches * 72
            
            print(f"      📏 Word table: {width_inches:.2f}\" x {height_inches:.2f}\" = {width_points:.1f} x {height_points:.1f} points")
            
            # Calculate position based on Word positioning information
            if word_dimensions.get('has_absolute_position'):
                # Table has absolute positioning in Word
                pos_x_inches = word_dimensions.get('pos_x_inches', 0)
                pos_y_inches = word_dimensions.get('pos_y_inches', 0)
                
                # Convert to PDF coordinates
                pdf_x = pos_x_inches * 72
                pdf_y = pos_y_inches * 72
                
                print(f"      📍 Absolute position: ({pos_x_inches:.2f}\", {pos_y_inches:.2f}\") = ({pdf_x:.1f}, {pdf_y:.1f}) points")
                
            else:
                # Use marker position with intelligent offset calculation
                # Estimate table position relative to marker
                
                # Check for table indentation
                indent_inches = word_dimensions.get('indent_inches', 0)
                indent_points = indent_inches * 72
                
                # Calculate left margin (typical Word margins)
                # Standard Word margins: 1" top/bottom, 1.25" left/right
                standard_left_margin = 1.25 * 72  # 90 points
                
                # Position table accounting for indentation
                pdf_x = standard_left_margin + indent_points
                  # For Y position, use marker position as reference
                # The marker should be at the top of the table, not center
                marker_top_y = marker_rect.y0
                
                # Adjust for typical Word table positioning
                # In Word-to-PDF conversion, tables often have some padding above
                table_top_padding = 10  # points
                pdf_y = max(table_top_padding, marker_top_y - table_top_padding)
                
                print(f"      📍 Calculated position: margin({standard_left_margin:.1f}) + indent({indent_points:.1f}) = x:{pdf_x:.1f}")
                print(f"      📍 Y position: marker_top({marker_top_y:.1f}) - padding({table_top_padding}) = y:{pdf_y:.1f}")
            
            # Create the overlay rectangle
            overlay_rect = fitz.Rect(
                max(0, pdf_x),
                max(0, pdf_y),
                min(page_rect.width, pdf_x + width_points),
                min(page_rect.height, pdf_y + height_points)
            )
            
            # Validate the rectangle
            if (overlay_rect.width > 50 and overlay_rect.height > 30 and 
                overlay_rect.x1 <= page_rect.width and overlay_rect.y1 <= page_rect.height):
                return overlay_rect
            else:
                print(f"      ⚠️ Converted coordinates are invalid: {overlay_rect}")
                return None
                
        except Exception as e:
            print(f"      ❌ Word-to-PDF coordinate conversion failed: {e}")
            return None
            
    def _detect_table_boundaries(self, page, marker_rect):
        """
        Detect table boundaries using PyMuPDF's table detection capabilities.
        
        Args:
            page: The PDF page object
            marker_rect: Rectangle where the marker was found
            
        Returns:
            fitz.Rect: Table cell boundaries or None if not detected
        """
        try:
            print(f"      🔍 Attempting table detection around marker...")
            
            # Use PyMuPDF's table detection
            table_finder = page.find_tables()
            tables = list(table_finder)  # Convert TableFinder to list
            
            if not tables:
                print(f"      📋 No tables detected on page")
                # Try alternative detection method
                return self._detect_table_by_text_analysis(page, marker_rect)
            
            print(f"      📋 Found {len(tables)} table(s) on page")
            
            # Find the table that contains the marker
            marker_center_x = (marker_rect.x0 + marker_rect.x1) / 2
            marker_center_y = (marker_rect.y0 + marker_rect.y1) / 2
            marker_point = fitz.Point(marker_center_x, marker_center_y)
            
            for i, table in enumerate(tables):
                table_rect = table.bbox
                print(f"      📋 Table {i+1}: ({table_rect.x0:.1f}, {table_rect.y0:.1f}) to ({table_rect.x1:.1f}, {table_rect.y1:.1f})")
                
                # Check if marker is within this table
                if table_rect.contains(marker_point):
                    print(f"      ✅ Marker found in table {i+1}")
                    
                    # Try to find the specific cell containing the marker
                    cell_rect = self._find_cell_in_table(table, marker_point)
                    if cell_rect:
                        print(f"      📋 Found containing cell: ({cell_rect.x0:.1f}, {cell_rect.y0:.1f}) to ({cell_rect.x1:.1f}, {cell_rect.y1:.1f})")
                        return cell_rect
                    else:
                        # Fallback: use entire table
                        print(f"      📋 Using entire table as overlay area")
                        return table_rect
            
            print(f"      ⚠️ Marker not found within any detected table")
            return None
            
        except Exception as e:
            print(f"      ❌ Table detection failed: {e}")
            return None
    
    def _detect_table_by_text_analysis(self, page, marker_rect):
        """
        Alternative table detection method using text analysis when PyMuPDF table detection fails.
        
        This method analyzes text positioning to identify table-like structures.
        
        Args:
            page: The PDF page object
            marker_rect: Rectangle where the marker was found
            
        Returns:
            fitz.Rect: Estimated table boundaries or None if not detected
        """
        try:
            print(f"      🔍 Attempting text-based table detection...")
            
            # Get all text blocks on the page
            text_dict = page.get_text("dict")
            
            # Look for text elements near the marker
            nearby_texts = []
            marker_center_x = (marker_rect.x0 + marker_rect.x1) / 2
            marker_center_y = (marker_rect.y0 + marker_rect.y1) / 2
            
            # Search within a reasonable distance from the marker
            search_distance = 200  # points
            
            for block in text_dict["blocks"]:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            span_rect = fitz.Rect(span["bbox"])
                            span_center_x = (span_rect.x0 + span_rect.x1) / 2
                            span_center_y = (span_rect.y0 + span_rect.y1) / 2
                            
                            # Calculate distance from marker
                            distance = ((span_center_x - marker_center_x) ** 2 + 
                                      (span_center_y - marker_center_y) ** 2) ** 0.5
                            
                            if distance <= search_distance:
                                nearby_texts.append({
                                    'rect': span_rect,
                                    'text': span["text"],
                                    'distance': distance
                                })
            
            if not nearby_texts:
                print(f"      📋 No text elements found near marker")
                return None
            
            # Find the bounds of nearby text to estimate table area
            min_x = min(text['rect'].x0 for text in nearby_texts)
            max_x = max(text['rect'].x1 for text in nearby_texts)
            min_y = min(text['rect'].y0 for text in nearby_texts)
            max_y = max(text['rect'].y1 for text in nearby_texts)
            
            # Look for whitespace patterns that might indicate table boundaries
            # Group texts by their vertical positions to detect rows
            y_positions = sorted(set(text['rect'].y0 for text in nearby_texts))
            row_groups = []
            
            for y_pos in y_positions:
                row_texts = [text for text in nearby_texts if abs(text['rect'].y0 - y_pos) < 5]  # 5 point tolerance
                if row_texts:
                    row_groups.append(row_texts)
            
            print(f"      📋 Found {len(row_groups)} text row(s) near marker")
            
            # If we have multiple rows, try to detect actual table structure
            if len(row_groups) > 1:
                # Look for consistent left margins (column alignment)
                all_left_margins = []
                for row in row_groups:
                    left_margins = [text['rect'].x0 for text in row]
                    all_left_margins.extend(left_margins)
                
                # Find common left margins (potential column boundaries)
                margin_tolerance = 10  # points
                unique_margins = []
                for margin in sorted(set(all_left_margins)):
                    # Check if this margin is close to an existing one
                    is_new = True
                    for existing in unique_margins:
                        if abs(margin - existing) < margin_tolerance:
                            is_new = False
                            break
                    if is_new:
                        unique_margins.append(margin)
                
                print(f"      📋 Detected {len(unique_margins)} potential column boundary(ies)")
                
                # Use the leftmost and rightmost margins for table bounds
                if len(unique_margins) >= 2:
                    min_x = min(unique_margins) - 10  # Small padding
                    max_x = max(unique_margins) + 100  # Estimate column width
            
            # Add padding to the detected bounds
            padding_x = 20
            padding_y = 20
            
            # Ensure we don't include the marker text in the final bounds
            # if the table should be clear (this helps for invisible table borders)
            if marker_rect.intersects(fitz.Rect(min_x, min_y, max_x, max_y)):
                # Expand to include some space around the marker
                min_x = min(min_x, marker_rect.x0 - padding_x)
                max_x = max(max_x, marker_rect.x1 + padding_x)
                min_y = min(min_y, marker_rect.y0 - padding_y)
                max_y = max(max_y, marker_rect.y1 + padding_y)
            
            estimated_table = fitz.Rect(
                max(0, min_x - padding_x),
                max(0, min_y - padding_y),
                min(page.rect.width, max_x + padding_x),
                min(page.rect.height, max_y + padding_y)
            )
            
            print(f"      📋 Text-based table estimate: ({estimated_table.x0:.1f}, {estimated_table.y0:.1f}) to ({estimated_table.x1:.1f}, {estimated_table.y1:.1f})")
            print(f"      📏 Estimated size: {estimated_table.width:.1f} x {estimated_table.height:.1f} points")
            
            # Validate the estimated bounds
            min_width, min_height = 100, 50
            max_width, max_height = page.rect.width * 0.9, page.rect.height * 0.9
            
            if (estimated_table.width >= min_width and estimated_table.height >= min_height and 
                estimated_table.width <= max_width and estimated_table.height <= max_height):
                return estimated_table
            else:
                print(f"      ⚠️ Text-based estimation produced invalid bounds (size check failed)")
                print(f"         Size: {estimated_table.width:.1f}x{estimated_table.height:.1f}, "
                      f"Required: {min_width}-{max_width:.0f} x {min_height}-{max_height:.0f}")
                return None
                
        except Exception as e:
            print(f"      ❌ Text-based table detection failed: {e}")
            return None
    
    def _find_cell_in_table(self, table, marker_point):
        """
        Find the specific cell within a table that contains the marker point.
        
        Args:
            table: PyMuPDF table object
            marker_point: Point where the marker is located
            
        Returns:
            fitz.Rect: Cell boundaries or None if not found
        """
        try:
            # Extract table data to get cell boundaries
            table_data = table.extract()
            
            if not table_data:
                print(f"      ⚠️ Could not extract table data")
                return table.bbox  # Fallback to entire table
            
            # Try to get cell rectangles from the table
            # PyMuPDF table objects may have different methods depending on version
            if hasattr(table, 'cells') and table.cells:
                print(f"      🔍 Analyzing {len(table.cells)} cells in table")
                
                for i, cell_rect in enumerate(table.cells):
                    if isinstance(cell_rect, (list, tuple)) and len(cell_rect) >= 4:
                        # Convert to fitz.Rect if needed
                        cell_bbox = fitz.Rect(cell_rect[0], cell_rect[1], cell_rect[2], cell_rect[3])
                    elif hasattr(cell_rect, 'bbox'):
                        cell_bbox = cell_rect.bbox
                    else:
                        cell_bbox = fitz.Rect(cell_rect)
                    
                    print(f"      📋 Cell {i}: ({cell_bbox.x0:.1f}, {cell_bbox.y0:.1f}) to ({cell_bbox.x1:.1f}, {cell_bbox.y1:.1f})")
                    
                    # Check if marker point is within this cell
                    if cell_bbox.contains(marker_point):
                        print(f"      ✅ Marker found in cell {i}")
                        return cell_bbox
                
                print(f"      ⚠️ Marker not found in any specific cell")
            else:
                print(f"      ⚠️ Table cells not accessible, using entire table")
            
            # Fallback: return entire table bounds
            return table.bbox
            
        except Exception as e:
            print(f"      ⚠️ Cell detection within table failed: {e}")
            return table.bbox  # Safe fallback
    
    def _cleanup(self):
        """
        Clean up temporary files created during the process.
        Skips cleanup if keep_temp flag is set.
        """
        if self.keep_temp:
            print("\nKeeping temporary files for debugging:")
            files_to_keep = [self.temp_docx_path, self.temp_pdf_path]
            for file_path in files_to_keep:
                if os.path.exists(file_path):
                    print(f"  ✓ Kept: {file_path}")
                else:
                    print(f"  - Not found: {os.path.basename(file_path)}")
            return
        
        print("\nCleaning up temporary files...")
        
        files_to_remove = [self.temp_docx_path, self.temp_pdf_path]
        
        for file_path in files_to_remove:
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    print(f"  ✓ Removed: {os.path.basename(file_path)}")
                except Exception as e:
                    print(f"  ⚠ Could not remove {os.path.basename(file_path)}: {e}")
            else:
                print(f"  - Not found: {os.path.basename(file_path)}")


def main():
    """
    Main execution function with command-line argument parsing.
    """
    parser = argparse.ArgumentParser(
        description="Python PDF Report Compiler - Combine Word documents with PDF appendices",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python report_compiler.py report.docx final_report.pdf
  python report_compiler.py "C:\\Reports\\my_report.docx" "C:\\Output\\final.pdf"

Placeholder format in Word document:
  [[INSERT: appendices/calculations.pdf]]
  [[INSERT: C:\\Shared\\analysis.pdf]]        """
    )
    
    parser.add_argument(
        'input_file',
        help='Path to the input Word document (.docx)'
    )
    
    parser.add_argument(
        'output_file', 
        help='Path for the output PDF file (.pdf)'
    )
    
    parser.add_argument(
        '--keep-temp',
        action='store_true',
        help='Keep temporary files for debugging purposes'
    )
    
    args = parser.parse_args()
    
    # Validate input file
    if not args.input_file.lower().endswith('.docx'):
        print("ERROR: Input file must be a .docx file")
        return 1
    
    if not os.path.exists(args.input_file):
        print(f"ERROR: Input file not found: {args.input_file}")
        return 1
    
    # Validate output file
    if not args.output_file.lower().endswith('.pdf'):
        print("ERROR: Output file must be a .pdf file")
        return 1
      # Create output directory if it doesn't exist
    output_dir = os.path.dirname(os.path.abspath(args.output_file))
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            print(f"Created output directory: {output_dir}")
        except Exception as e:
            print(f"ERROR: Could not create output directory: {e}")
            return 1
    
    try:
        # Create and run the compiler
        compiler = ReportCompiler(args.input_file, args.output_file, keep_temp=args.keep_temp)
        compiler.run()
        return 0
        
    except Exception as e:
        print(f"\nERROR: Compilation failed: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit(main())
