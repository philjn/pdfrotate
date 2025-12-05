#!/usr/bin/env python3
"""
Fix PDF page rotations - processes all PDFs in a directory.
Uses OCR (Tesseract) to detect optimal page orientation and corrects rotation.
"""

import pikepdf
import sys
import os
from pathlib import Path
from pdf2image import convert_from_path
import pytesseract
from PIL import Image

def detect_orientation(image):
    """
    Detect the rotation needed to make text upright using OCR.
    Returns rotation angle (0, 90, 180, 270) and confidence score.
    """
    try:
        # Use Tesseract's OSD (Orientation and Script Detection)
        osd = pytesseract.image_to_osd(image)
        
        # Parse OSD output
        rotation = 0
        confidence = 0
        for line in osd.split('\n'):
            if line.startswith('Rotate:'):
                rotation = int(line.split(':')[1].strip())
            elif line.startswith('Orientation confidence:'):
                confidence = float(line.split(':')[1].strip())
        
        return rotation, confidence
    except Exception as e:
        # If OSD fails, try manual detection with different rotations
        best_angle = 0
        best_confidence = 0
        
        for angle in [0, 90, 180, 270]:
            try:
                rotated = image.rotate(-angle, expand=True)  # Negative for counter-clockwise
                result = pytesseract.image_to_data(rotated, output_type=pytesseract.Output.DICT)
                
                # Calculate average confidence of detected text
                confidences = [int(conf) for conf in result['conf'] if conf != '-1']
                if confidences:
                    avg_conf = sum(confidences) / len(confidences)
                    if avg_conf > best_confidence:
                        best_confidence = avg_conf
                        best_angle = angle
            except:
                continue
        
        return best_angle, best_confidence

def fix_pdf_rotation(input_path, output_path, dpi=200):
    """Fix rotation for a single PDF file using OCR detection."""
    try:
        # Convert PDF pages to images for OCR analysis
        print(f'  Converting PDF to images (DPI={dpi})...')
        images = convert_from_path(input_path, dpi=dpi)
        
        # Open PDF for modification
        pdf = pikepdf.open(input_path)
        fixed_pages = []
        
        print(f'  Analyzing {len(images)} page(s) with OCR...')
        
        for i, (image, page) in enumerate(zip(images, pdf.pages)):
            rotation_needed, confidence = detect_orientation(image)
            
            if rotation_needed != 0:
                print(f'  Page {i+1}: Detected {rotation_needed}° rotation (confidence: {confidence:.1f}%) - correcting')
                # Apply counter-rotation to make upright
                page.Rotate = rotation_needed
                fixed_pages.append((i+1, rotation_needed, confidence))
            else:
                print(f'  Page {i+1}: Already upright (confidence: {confidence:.1f}%)')
        
        if fixed_pages:
            pdf.save(output_path)
            print(f'  ✓ Fixed {len(fixed_pages)} page(s)')
            for page_num, rotation, conf in fixed_pages:
                print(f'    - Page {page_num}: rotated {rotation}° (confidence: {conf:.1f}%)')
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
