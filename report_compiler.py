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
import zipfile
import xml.etree.ElementTree as ET


class ReportCompiler:
    """
    A class to compile Word documents with PDF appendices into a single PDF report.
    
    The process involves:
    1. Finding PDF insertion placeholders in the Word document
    2. Modifying the document to create blank pages with hidden markers
    3. Converting the modified document to PDF
    4. Overlaying the appendix PDFs onto the blank pages    """
    
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
            
            # Step 3: Overlay PDFs
            print("\nStep 4: Overlaying appendix PDFs...")
            self._overlay_pdfs(placeholders)
            print(f"✓ Final PDF created: {self.final_pdf_path}")
            
            print("\n=== Report Compilation Complete ===")
            
        finally:
            # Always clean up temporary files
            self._cleanup()
    
    def _find_placeholders_and_modify_docx(self):
        """
        Find PDF placeholders in the DOCX and modify the document.
        This includes both regular paragraph placeholders and text box placeholders.
        
        Returns:
            tuple: (modified_doc, placeholders_list)
                - modified_doc: The modified docx.Document object
                - placeholders_list: List of dictionaries containing placeholder info
        """
        # First, check for text box placeholders (these have priority for sized insertion)
        textbox_placeholders = self._find_textbox_placeholders()
        
        # Open the source document
        doc = Document(self.input_docx_path)
        placeholders = []
        placeholder_index = 0
        
        # Get the directory of the input document for resolving relative paths
        input_dir = os.path.dirname(self.input_docx_path)
        
        # Iterate through all paragraphs
        for para_idx, paragraph in enumerate(doc.paragraphs):
            # Check if this paragraph contains a placeholder
            match = self.placeholder_regex.search(paragraph.text)
            if match:
                pdf_path_raw = match.group(1).strip()
                
                # Resolve relative path
                if not os.path.isabs(pdf_path_raw):
                    pdf_path = os.path.join(input_dir, pdf_path_raw)
                else:
                    pdf_path = pdf_path_raw
                
                pdf_path = os.path.abspath(pdf_path)
                
                print(f"  Found placeholder: {pdf_path_raw}")
                print(f"    Resolved to: {pdf_path}")
                
                # Validate that the PDF file exists
                if not os.path.exists(pdf_path):
                    print(f"    ⚠ WARNING: PDF file not found: {pdf_path}")
                    continue
                
                # Get page count of the appendix PDF
                try:
                    pdf_doc = fitz.open(pdf_path)
                    page_count = pdf_doc.page_count
                    pdf_doc.close()
                    print(f"    PDF has {page_count} page(s)")
                except Exception as e:
                    print(f"    ⚠ ERROR: Cannot read PDF file: {e}")
                    continue
                
                # Create unique marker text for this placeholder
                marker_text = f"%%APPENDIX_START_{placeholder_index}%%"                # Clear the placeholder text and add page breaks for the appendix
                paragraph.clear()
                
                # Add visible marker for the first page (will be removed during overlay)
                marker_run = paragraph.add_run(marker_text)
                marker_run.font.size = Pt(12)  # Visible size
                marker_run.font.color.rgb = RGBColor(255, 0, 0)  # Red color (visible)
                
                # Add page breaks for each page in the appendix
                for i in range(page_count):
                    if i > 0:  # Add page break before each page except the first
                        run = paragraph.add_run()
                        run.add_break(WD_BREAK.PAGE)
                
                # Store placeholder information
                placeholders.append({
                    'marker': marker_text,
                    'pdf_path': pdf_path,
                    'page_count': page_count,
                    'index': placeholder_index                })
                
                placeholder_index += 1
        
        return doc, placeholders
    
    def _find_textbox_placeholders(self):
        """
        Find PDF placeholders inside text boxes and extract text box dimensions.
        
        Returns:
            list: List of dictionaries containing textbox placeholder info with dimensions
        """
        print("  Scanning for text box placeholders...")
        textbox_placeholders = []
        
        try:
            with zipfile.ZipFile(self.input_docx_path, 'r') as docx_zip:
                if 'word/document.xml' not in docx_zip.namelist():
                    return textbox_placeholders
                
                with docx_zip.open('word/document.xml') as xml_file:
                    xml_content = xml_file.read().decode('utf-8')
                    
                    # Parse XML
                    root = ET.fromstring(xml_content)
                    
                    # Define namespaces
                    namespaces = {
                        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                        'v': 'urn:schemas-microsoft-com:vml',
                        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
                        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
                    }
                    
                    # Search for text boxes with INSERT placeholders
                    textbox_types = [
                        ('.//v:textbox', 'VML textbox'),
                        ('.//w:txbxContent', 'Word textbox content')
                    ]
                    
                    for xpath, description in textbox_types:
                        textboxes = root.findall(xpath, namespaces)
                        
                        for i, textbox in enumerate(textboxes):
                            # Get all text content from the text box
                            text_elements = textbox.findall('.//w:t', namespaces)
                            if text_elements:
                                all_text = ''.join([t.text for t in text_elements if t.text])
                                
                                # Check if this text box contains an INSERT placeholder
                                match = self.placeholder_regex.search(all_text)
                                if match:
                                    pdf_path_raw = match.group(1).strip()
                                    
                                    print(f"    Found textbox placeholder: {pdf_path_raw}")
                                    
                                    # Resolve relative path
                                    input_dir = os.path.dirname(self.input_docx_path)
                                    if not os.path.isabs(pdf_path_raw):
                                        pdf_path = os.path.join(input_dir, pdf_path_raw)
                                    else:
                                        pdf_path = pdf_path_raw
                                    
                                    pdf_path = os.path.abspath(pdf_path)
                                    
                                    # Validate that the PDF file exists
                                    if not os.path.exists(pdf_path):
                                        print(f"      ⚠ WARNING: PDF file not found: {pdf_path}")
                                        continue
                                    
                                    # Get page count of the appendix PDF
                                    try:
                                        pdf_doc = fitz.open(pdf_path)
                                        page_count = pdf_doc.page_count
                                        pdf_doc.close()
                                        print(f"      PDF has {page_count} page(s)")
                                    except Exception as e:
                                        print(f"      ⚠ ERROR: Cannot read PDF file: {e}")
                                        continue
                                    
                                    # Try to find dimensions for this text box
                                    dimensions = self._get_textbox_dimensions(textbox, namespaces)
                                    
                                    textbox_info = {
                                        'pdf_path': pdf_path,
                                        'pdf_path_raw': pdf_path_raw,
                                        'page_count': page_count,
                                        'textbox_type': description,
                                        'textbox_text': all_text
                                    }
                                    
                                    if dimensions:
                                        textbox_info.update(dimensions)
                                        print(f"      Textbox dimensions: {dimensions['width_inches']:.2f}\" x {dimensions['height_inches']:.2f}\"")
                                    else:
                                        print(f"      ⚠ Could not determine textbox dimensions")
                                    
                                    textbox_placeholders.append(textbox_info)
        
        except Exception as e:
            print(f"    ⚠ Error scanning for textbox placeholders: {e}")
        
        return textbox_placeholders
    
    def _get_textbox_dimensions(self, textbox_elem, namespaces):
        """
        Extract dimensions from a text box element.
        
        Returns:
            dict: Dictionary with width/height information or None
        """
        try:
            # Method 1: Look for parent drawing element with extent
            current = textbox_elem
            for _ in range(10):  # Limit search depth
                if current is None:
                    break
                    
                # Look for extent elements in current or parent elements
                extents = current.findall('.//a:ext', namespaces)
                if extents:
                    extent = extents[0]
                    cx = extent.get('cx')
                    cy = extent.get('cy')
                    if cx and cy:
                        width_inches = int(cx) / 914400
                        height_inches = int(cy) / 914400
                        return {
                            'width_inches': width_inches,
                            'height_inches': height_inches,
                            'width_emus': cx,
                            'height_emus': cy
                        }
                
                # Move to parent element
                current = current.getparent() if hasattr(current, 'getparent') else None
            
            # Method 2: Look for VML style attributes
            style = textbox_elem.get('style', '')
            if style:
                width_match = re.search(r'width:([^;]+)', style)
                height_match = re.search(r'height:([^;]+)', style)
                if width_match and height_match:
                    # Try to parse dimensions (this is approximate)
                    width_str = width_match.group(1).strip()
                    height_str = height_match.group(1).strip()
                    return {
                        'width_style': width_str,
                        'height_style': height_str,
                        'note': 'Dimensions from VML style (approximate)'
                    }
            
            return None
            
        except Exception as e:
            print(f"        Error extracting textbox dimensions: {e}")
            return None

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
        Overlay appendix PDFs onto the base PDF using visible marker detection.
        
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
                
                # Search for the marker in the PDF
                start_page_index = None
                
                for page_num in range(base_pdf.page_count):
                    page = base_pdf[page_num]
                    text_instances = page.search_for(marker)
                    
                    if text_instances:
                        start_page_index = page_num
                        print(f"      ✓ Found marker on page {page_num + 1}")
                        
                        # Remove the marker text from the page
                        for inst in text_instances:
                            # Add a white rectangle to cover the marker text
                            page.add_redact_annot(inst, fill=(1, 1, 1))  # White fill
                        page.apply_redactions()
                        print(f"      ✓ Removed marker text from page {page_num + 1}")
                        break
                
                if start_page_index is None:
                    print(f"      ⚠ WARNING: Marker not found in PDF, skipping appendix")
                    continue

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
                        
                        # Get the source page from the appendix
                        source_page = appendix_pdf[i]
                        
                        # Create a rectangle covering the full page
                        page_rect = target_page.rect
                        
                        # Overlay the appendix page onto the target page
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
