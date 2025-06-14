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
    """
    
    def __init__(self, input_docx_path, final_pdf_path):
        """
        Initialize the ReportCompiler with input and output paths.
        
        Args:
            input_docx_path (str): Absolute path to the source .docx file
            final_pdf_path (str): Absolute path for the final output .pdf file
        """
        self.input_docx_path = os.path.abspath(input_docx_path)
        self.final_pdf_path = os.path.abspath(final_pdf_path)
        
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
        
        Returns:
            tuple: (modified_doc, placeholders_list)
                - modified_doc: The modified docx.Document object
                - placeholders_list: List of dictionaries containing placeholder info
        """
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
                marker_text = f"%%APPENDIX_START_{placeholder_index}%%"
                
                # Insert hidden marker before clearing the paragraph
                # Add the marker as invisible text
                run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
                run.clear()
                marker_run = paragraph.add_run(marker_text)
                marker_run.font.size = Pt(1)  # Minimal size
                marker_run.font.color.rgb = RGBColor(255, 255, 255)  # White color (invisible)
                
                # Clear the placeholder text
                paragraph.clear()
                  # Add page breaks for each page in the appendix
                for i in range(page_count):
                    if i == 0:
                        # First page - add marker and page break
                        marker_run = paragraph.add_run(marker_text)
                        marker_run.font.size = Pt(1)
                        marker_run.font.color.rgb = RGBColor(255, 255, 255)
                    
                    if i < page_count - 1:  # Don't add page break after the last page
                        # Add page break
                        run = paragraph.add_run()
                        run.add_break()
                
                # Store placeholder information
                placeholders.append({
                    'marker': marker_text,
                    'pdf_path': pdf_path,
                    'page_count': page_count,
                    'index': placeholder_index                })
                
                placeholder_index += 1
        
        return doc, placeholders
    
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
        Overlay appendix PDFs onto the base PDF using the placeholder markers.
        
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
                
                # Search for the marker text in the base PDF
                start_page_index = None
                for page_num in range(base_pdf.page_count):
                    page = base_pdf[page_num]
                    text_instances = page.search_for(marker)
                    
                    if text_instances:
                        start_page_index = page_num
                        print(f"      ✓ Found marker on page {page_num + 1}")
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
        """
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
  [[INSERT: C:\\Shared\\analysis.pdf]]
        """
    )
    
    parser.add_argument(
        'input_file',
        help='Path to the input Word document (.docx)'
    )
    
    parser.add_argument(
        'output_file', 
        help='Path for the output PDF file (.pdf)'
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
        compiler = ReportCompiler(args.input_file, args.output_file)
        compiler.run()
        return 0
        
    except Exception as e:
        print(f"\nERROR: Compilation failed: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit(main())
