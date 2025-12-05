#!/usr/bin/env python3
"""
Fix PDF page rotations - processes all PDFs in a directory.
Detects pages with rotation metadata and corrects them to 0 degrees.
"""

import pikepdf
import sys
import os
from pathlib import Path

def fix_pdf_rotation(input_path, output_path):
    """Fix rotation for a single PDF file."""
    try:
        pdf = pikepdf.open(input_path)
        fixed_pages = []
        
        for i, page in enumerate(pdf.pages):
            current_rotation = page.get('/Rotate', 0)
            if current_rotation in [90, 180, 270]:
                print(f'  Page {i+1}: Correcting rotation from {current_rotation}° to 0°')
                page.Rotate = 0
                fixed_pages.append(i+1)
        
        if fixed_pages:
            pdf.save(output_path)
            print(f'  ✓ Fixed {len(fixed_pages)} page(s): {fixed_pages}')
            print(f'  ✓ Saved to: {output_path}')
            return True
        else:
            print(f'  ℹ No rotation corrections needed')
            return False
            
    except Exception as e:
        print(f'  ✗ Error processing file: {e}')
        return False

def main():
    if len(sys.argv) < 2:
        print("Usage: python fix-pdf-rotation.py <directory_or_file> [output_suffix]")
        print("Example: python fix-pdf-rotation.py C:\\Documents\\pdfs")
        print("Example: python fix-pdf-rotation.py input.pdf _fixed")
        sys.exit(1)
    
    input_arg = sys.argv[1]
    output_suffix = sys.argv[2] if len(sys.argv) > 2 else "_fixed"
    
    # Check if input is a directory or file
    path = Path(input_arg)
    
    if path.is_dir():
        # Process all PDFs in directory
        pdf_files = list(path.glob("*.pdf"))
        if not pdf_files:
            print(f"No PDF files found in {path}")
            sys.exit(1)
        
        print(f"Found {len(pdf_files)} PDF file(s) in {path}\n")
        
        processed = 0
        fixed = 0
        for pdf_file in pdf_files:
            print(f"Processing: {pdf_file.name}")
            output_name = pdf_file.stem + output_suffix + pdf_file.suffix
            output_path = pdf_file.parent / output_name
            
            if fix_pdf_rotation(str(pdf_file), str(output_path)):
                fixed += 1
            processed += 1
            print()
        
        print(f"Summary: Processed {processed} files, fixed {fixed} files")
        
    elif path.is_file():
        # Process single file
        if not path.suffix.lower() == '.pdf':
            print(f"Error: {path} is not a PDF file")
            sys.exit(1)
        
        output_name = path.stem + output_suffix + path.suffix
        output_path = path.parent / output_name
        
        print(f"Processing: {path.name}")
        fix_pdf_rotation(str(path), str(output_path))
        
    else:
        print(f"Error: {input_arg} is not a valid file or directory")
        sys.exit(1)

if __name__ == "__main__":
    main()
