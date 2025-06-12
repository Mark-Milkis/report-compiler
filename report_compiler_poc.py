#!/usr/bin/env python3
"""
Proof of Concept: PDF Report Compiler

Simple script to demonstrate:
1. Converting DOCX to Markdown using pandoc
2. Processing PDF placeholders [[INSERT: path]]
3. Converting Markdown to PDF using pandoc with LaTeX

Usage: python report_compiler_poc.py input.docx
"""

import os
import re
import subprocess
import sys
from pathlib import Path
import argparse


def run_command(cmd, description):
    """Run a shell command and return success status."""
    print(f"\n{'='*50}")
    print(f"STEP: {description}")
    print(f"Command: {cmd}")
    print(f"{'='*50}")
    
    try:
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        
        if result.stdout:
            print("STDOUT:")
            print(result.stdout)
        
        if result.stderr:
            print("STDERR:")
            print(result.stderr)
        
        if result.returncode == 0:
            print(f"✓ SUCCESS: {description}")
            return True
        else:
            print(f"✗ FAILED: {description} (return code: {result.returncode})")
            return False
            
    except Exception as e:
        print(f"✗ ERROR: {description} - {e}")
        return False


def process_pdf_placeholders(content, base_dir):
    """Process [[INSERT: path]] placeholders and replace with LaTeX includepdf commands."""
    print(f"\n{'='*50}")
    print("PROCESSING PDF PLACEHOLDERS")
    print(f"{'='*50}")
      # Handle both escaped and non-escaped brackets
    placeholder_pattern = r'\\?\[\\?\[INSERT:\s*([^\]\\]+)\\?\]\\?\]'
    placeholders_found = re.findall(placeholder_pattern, content)
    
    print(f"Found {len(placeholders_found)} PDF placeholders:")
    for placeholder in placeholders_found:
        print(f"  - {placeholder}")
    
    def replace_placeholder(match):
        file_path = match.group(1).strip()
        
        # Handle relative paths
        if not os.path.isabs(file_path):
            full_path = os.path.join(base_dir, file_path)
        else:
            full_path = file_path
        
        # For LaTeX, use relative path from the markdown file location
        if not os.path.isabs(file_path):
            latex_path = file_path.replace('\\', '/')
        else:
            # Convert absolute to relative
            latex_path = os.path.relpath(full_path, base_dir).replace('\\', '/')
        
        # Check if file exists
        if os.path.exists(full_path):
            print(f"  ✓ Found: {full_path}")
        else:
            print(f"  ⚠ Missing: {full_path}")
        
        # Return LaTeX command with relative path
        return f'\\newpage\n\\includepdf[pages=-]{{{latex_path}}}'
    
    processed_content = re.sub(placeholder_pattern, replace_placeholder, content)
    return processed_content


def process_figure_tags(content):
    """
    Converts <figure><img>...<figcaption>...</figcaption></figure> blocks
    into LaTeX commands that preserve sizing information.
    """
    print(f"\n{'='*50}")
    print("PROCESSING FIGURE/FIGCAPTION TAGS WITH SIZING")
    print(f"{'='*50}")
    
    # Pattern to match <figure><img ...><figcaption>...</figcaption></figure>
    figure_pattern = re.compile(
        r'<figure>\s*'
        r'<img\s+src="([^"]*)"([^>]*?)>\s*'
        r'<figcaption>(.*?)</figcaption>\s*'
        r'</figure>',
        re.DOTALL | re.IGNORECASE
    )

    processed_figure_count = 0

    def replace_figure(match):
        nonlocal processed_figure_count
        img_src = match.group(1)
        img_attributes = match.group(2)
        figcaption_html = match.group(3)

        # Extract caption text and clean it
        caption_text = re.sub(r'<[^>]+>', '', figcaption_html)
        caption_text = caption_text.replace('\\n', ' ').strip()
        
        # Remove figure numbering from caption text to avoid duplication in LaTeX
        caption_text = re.sub(r'^Figure\s+\d+:\s*', '', caption_text)
        
        # Fix image path - convert backslashes to forward slashes
        img_src = img_src.replace('\\', '/')
        
        # Extract sizing information from img attributes
        width_latex = None
        height_latex = None
        
        # Check for style attribute with width/height
        style_match = re.search(r'style="([^"]*)"', img_attributes)
        if style_match:
            style_content = style_match.group(1)
            width_match = re.search(r'width:\s*([^;]+)', style_content)
            height_match = re.search(r'height:\s*([^;]+)', style_content)
            
            if width_match:
                width_latex = width_match.group(1).strip()
                print(f"    Found width in style: {width_latex}")
            if height_match:
                height_latex = height_match.group(1).strip()
                print(f"    Found height in style: {height_latex}")
        
        # Check for direct width/height attributes
        if not width_latex:
            width_match = re.search(r'width="([^"]*)"', img_attributes)
            if width_match:
                width_latex = width_match.group(1).strip()
                print(f"    Found width attribute: {width_latex}")
        
        if not height_latex:
            height_match = re.search(r'height="([^"]*)"', img_attributes)
            if height_match:
                height_latex = height_match.group(1).strip()
                print(f"    Found height attribute: {height_latex}")
        
        # Create LaTeX command with sizing if available
        if width_latex or height_latex:
            size_options = []
            if width_latex:
                size_options.append(f"width={width_latex}")
            if height_latex:
                size_options.append(f"height={height_latex}")
            
            size_str = ",".join(size_options)
            latex_command = f'\\includegraphics[{size_str}]{{{img_src}}}'
            
            # Return LaTeX figure environment with sizing
            result = f'\\begin{{figure}}[h!]\n\\centering\n{latex_command}\n\\caption{{{caption_text}}}\n\\end{{figure}}'
            print(f"  - Found figure for image '{img_src}' with sizing. Replacing with LaTeX figure command.")
        else:
            # No sizing info, use standard markdown image syntax
            result = f'![{caption_text}]({img_src})'
            print(f"  - Found figure for image '{img_src}' without sizing. Replacing with: {result}")
        
        processed_figure_count += 1
        return result

    content_after_processing = figure_pattern.sub(replace_figure, content)
    if processed_figure_count == 0:
        print("  - No <figure>...<figcaption> blocks found or processed.")
    else:
        print(f"✓ Processed {processed_figure_count} figure/figcaption blocks.")
    return content_after_processing


def fix_image_paths(content, base_dir):
    """Fix image paths to be relative to the markdown file location and preserve sizing information.
    Only processes img tags that are NOT inside figure blocks (those are handled by process_figure_tags)."""
    print(f"\n{'='*50}")
    print("FIXING STANDALONE IMAGE PATHS AND PRESERVING SIZING")
    print(f"{'='*50}")
    
    # Pattern to match image src attributes in HTML img tags that are NOT inside figure blocks
    # We'll use a more careful approach to avoid processing images already in figures
    img_pattern = r'<img src="([^"]*)"([^>]*)>'
    
    # First, let's find all figure blocks so we can exclude them
    figure_pattern = r'<figure>.*?</figure>'
    figure_blocks = re.findall(figure_pattern, content, re.DOTALL | re.IGNORECASE)
    
    # Find all img tags
    images_found = re.findall(img_pattern, content)
    
    # Filter out img tags that are inside figure blocks
    standalone_images = []
    for img_path, attributes in images_found:
        img_tag = f'<img src="{img_path}"{attributes}>'
        is_in_figure = any(img_tag in figure_block for figure_block in figure_blocks)
        if not is_in_figure:
            standalone_images.append((img_path, attributes))
    
    print(f"Found {len(standalone_images)} standalone image references (excluding those in figure blocks):")
    for img_path, attributes in standalone_images:
        print(f"  - {img_path}")
        if 'style=' in attributes or 'width=' in attributes or 'height=' in attributes:
            print(f"    Sizing info: {attributes.strip()}")
    
    def fix_image_tag(match):
        img_path = match.group(1)
        attributes = match.group(2)
        
        # Check if this img tag is inside a figure block - if so, skip it
        img_tag = f'<img src="{img_path}"{attributes}>'
        is_in_figure = any(img_tag in figure_block for figure_block in figure_blocks)
        if is_in_figure:
            return match.group(0)  # Return original unchanged
        
        # Convert backslashes to forward slashes
        img_path = img_path.replace('\\', '/')
        
        # If the path contains the base directory, make it relative
        base_dir_str = str(base_dir).replace('\\', '/')
        if base_dir_str in img_path:
            # Extract just the relative part
            relative_path = img_path.split(base_dir_str + '/')[-1]
            print(f"  ✓ Converting '{img_path}' to '{relative_path}'")
            final_path = relative_path
        else:
            # Already relative, just fix slashes
            print(f"  ✓ Fixing slashes in '{img_path}'")
            final_path = img_path
        
        # Extract alt text if present
        alt_match = re.search(r'alt="([^"]*)"', attributes)
        alt_text = alt_match.group(1) if alt_match else "Image"
        
        # Remove figure numbering from alt text to avoid duplication in LaTeX
        # LaTeX will automatically number figures
        alt_text = re.sub(r'^Figure\s+\d+:\s*', '', alt_text)
        
        # Extract sizing information
        width_latex = None
        height_latex = None
        
        # Check for style attribute with width/height
        style_match = re.search(r'style="([^"]*)"', attributes)
        if style_match:
            style_content = style_match.group(1)
            width_match = re.search(r'width:\s*([^;]+)', style_content)
            height_match = re.search(r'height:\s*([^;]+)', style_content)
            
            if width_match:
                width_latex = width_match.group(1).strip()
                print(f"    Found width: {width_latex}")
            if height_match:
                height_latex = height_match.group(1).strip()
                print(f"    Found height: {height_latex}")
        
        # Check for direct width/height attributes
        if not width_latex:
            width_match = re.search(r'width="([^"]*)"', attributes)
            if width_match:
                width_latex = width_match.group(1).strip()
                print(f"    Found width attribute: {width_latex}")
        
        if not height_latex:
            height_match = re.search(r'height="([^"]*)"', attributes)
            if height_match:
                height_latex = height_match.group(1).strip()
                print(f"    Found height attribute: {height_latex}")
        
        # If we have sizing information, use LaTeX includegraphics with sizing
        if width_latex or height_latex:
            size_options = []
            if width_latex:
                size_options.append(f"width={width_latex}")
            if height_latex:
                size_options.append(f"height={height_latex}")
            
            size_str = ",".join(size_options)
            latex_command = f'\\includegraphics[{size_str}]{{{final_path}}}'
            
            # Return raw LaTeX command within a figure environment for proper captioning
            return f'\\begin{{figure}}[h!]\n\\centering\n{latex_command}\n\\caption{{{alt_text}}}\n\\end{{figure}}'
        else:
            # No sizing info, use standard markdown image syntax
            return f'![{alt_text}]({final_path})'
    
    fixed_content = re.sub(img_pattern, fix_image_tag, content)
    return fixed_content


def get_template_path():
    """Get the path to the LaTeX template file."""
    # Look for template.tex in the same directory as this script
    script_dir = Path(__file__).parent
    template_path = script_dir / "template.tex"
    
    if not template_path.exists():
        raise FileNotFoundError(f"Template file not found: {template_path}")
    
    return template_path

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Proof of Concept: PDF Report Compiler')
    parser.add_argument('input_file', help='Input DOCX file to process')
    parser.add_argument('--output', '-o', help='Output PDF file name (optional)')
    
    args = parser.parse_args()
    
    input_file = Path(args.input_file)
    if not input_file.exists():
        print(f"Error: Input file '{input_file}' not found!")
        sys.exit(1)
      # Set up file names
    base_name = input_file.stem
    base_dir = input_file.parent
    md_file = base_dir / f"{base_name}_temp.md"
    processed_md_file = base_dir / f"{base_name}_processed.md"
    template_file = get_template_path()  # Use external template file
    output_file = Path(args.output) if args.output else base_dir / f"{base_name}_compiled.pdf"
    
    print(f"Input DOCX: {input_file}")
    print(f"Working directory: {base_dir}")
    print(f"Template: {template_file}")
    print(f"Output PDF: {output_file}")
    
    # Step 1: Template is already available as external file
    print(f"\nUsing LaTeX template: {template_file}")
    if not template_file.exists():
        print(f"✗ Template file not found: {template_file}")
        sys.exit(1)
    print("✓ Template found")
    
    # Step 2: Convert DOCX to Markdown
    pandoc_cmd = f'pandoc "{input_file}" -t markdown --extract-media="{base_dir}" -o "{md_file}"'
    if not run_command(pandoc_cmd, "Converting DOCX to Markdown"):
        sys.exit(1)
    
    # Step 3: Read and process the markdown file
    print(f"\nReading markdown file: {md_file}")
    try:
        with open(md_file, 'r', encoding='utf-8') as f:
            content = f.read()
        print(f"✓ Read {len(content)} characters from markdown file")
    except Exception as e:
        print(f"✗ Error reading markdown file: {e}")
        sys.exit(1)
      # Step 4: Process <figure> and <figcaption> tags FIRST (before fixing individual img tags)
    content = process_figure_tags(content)

    # Step 5: Fix remaining individual image paths (those not in figure blocks)
    content = fix_image_paths(content, base_dir)

    # Step 5: Process PDF placeholders
    processed_content = process_pdf_placeholders(content, base_dir)
    
    # Step 6: Write processed markdown
    print(f"\nWriting processed markdown: {processed_md_file}")
    try:
        with open(processed_md_file, 'w', encoding='utf-8') as f:
            f.write(processed_content)
        print("✓ Processed markdown written")
    except Exception as e:
        print(f"✗ Error writing processed markdown: {e}")
        sys.exit(1)    # Step 7: Convert processed markdown to PDF (change to the directory where files are)
    original_dir = os.getcwd()
    os.chdir(base_dir)
    
    # Use relative paths for processed markdown and output, absolute path for template
    rel_processed_md = processed_md_file.name
    abs_template = str(template_file)  # Use absolute path for template
    rel_output = output_file.name
    
    pdf_cmd = f'pandoc "{rel_processed_md}" --template="{abs_template}" --pdf-engine=pdflatex -o "{rel_output}"'
    
    success = run_command(pdf_cmd, "Converting Markdown to PDF")
    
    # Change back to original directory
    os.chdir(original_dir)
    
    if not success:
        print("\n⚠ PDF conversion failed. This might be due to:")
        print("  - Missing LaTeX installation")
        print("  - Missing image files")
        print("  - Missing PDF files referenced in placeholders")
        print("  - LaTeX compilation errors")
        sys.exit(1)
    
    print(f"\n{'='*50}")
    print("SUCCESS! Report compilation completed.")
    print(f"Output PDF: {output_file}")
    print(f"{'='*50}")
    
    # Clean up temporary files
    # if md_file.exists():
    #     md_file.unlink()
    # if processed_md_file.exists():
    #     processed_md_file.unlink()
    # if template_file.exists():
    #     template_file.unlink()
    
    # print("Temporary files cleaned up.")
