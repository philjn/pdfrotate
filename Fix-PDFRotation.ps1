param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Path,
    
    [Parameter(Position=1)]
    [string]$OutputSuffix = "_fixed"
)

<#
.SYNOPSIS
    Fixes PDF page rotations using OCR to detect optimal orientation.

.DESCRIPTION
    Processes PDF files in a directory or a single PDF file, using Tesseract OCR
    to detect which pages are rotated incorrectly and automatically corrects them.
    Original files are preserved; corrected files are saved with a suffix.

.PARAMETER Path
    Path to a PDF file or directory containing PDF files to process.

.PARAMETER OutputSuffix
    Suffix to append to output filenames (default: "_fixed")

.EXAMPLE
    .\Fix-PDFRotation.ps1 C:\Documents\pdfs
    Process all PDFs in the directory, creating *_fixed.pdf files

.EXAMPLE
    .\Fix-PDFRotation.ps1 C:\Documents\report.pdf -OutputSuffix "_corrected"
    Process single file, creating report_corrected.pdf

.NOTES
    PREREQUISITES:
    
    1. Python 3 (https://www.python.org/)
       - Download and install Python 3.8 or later
       - Make sure to check "Add Python to PATH" during installation
    
    2. Tesseract OCR (Required for text detection)
       - Download: https://github.com/UB-Mannheim/tesseract/wiki
       - Windows installer: tesseract-ocr-w64-setup-*.exe
       - Install to default location: C:\Program Files\Tesseract-OCR
       - The installer should add Tesseract to PATH automatically
       - Alternative: Manual setup if not in PATH:
         $env:PATH += ";C:\Program Files\Tesseract-OCR"
         Or in Python script: pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    
    3. Poppler (Required for PDF to image conversion)
       - Download: https://github.com/oschwartz10612/poppler-windows/releases
       - Extract poppler-*.zip to a folder (e.g., C:\Program Files\poppler)
       - Add the bin folder to PATH:
         $env:PATH += ";C:\Program Files\poppler\Library\bin"
       - Or permanently via System Environment Variables:
         Search "Environment Variables" → Edit System Environment Variables → 
         Environment Variables → Path → New → Add: C:\Program Files\poppler\Library\bin
    
    4. Python packages (auto-installed by this script):
       - pikepdf: PDF manipulation
       - pdf2image: Convert PDF pages to images
       - pytesseract: Python wrapper for Tesseract
       - Pillow: Image processing
       
       Manual installation if needed:
       python -m pip install pikepdf pdf2image pytesseract Pillow
    
    TROUBLESHOOTING:
    - "Tesseract not found": Install Tesseract OCR and ensure it's in PATH
    - "Unable to get page count": Install poppler and add bin folder to PATH
    - "No module named 'PIL'": Run: python -m pip install Pillow
    - "Permission denied": Run PowerShell as Administrator
    
    QUICK SETUP (copy and paste into PowerShell as Administrator):
    # Install Python packages
    python -m pip install pikepdf pdf2image pytesseract Pillow
    
    # Add Tesseract and Poppler to PATH (update paths as needed)
    [Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Program Files\Tesseract-OCR", "Machine")
    [Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Program Files\poppler\Library\bin", "Machine")
#>

# Check if Python is available
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Using: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "Error: Python is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install Python from https://www.python.org/" -ForegroundColor Yellow
    Exit 1
}

# Check if required packages are installed
Write-Host "Checking for required Python packages..." -ForegroundColor Cyan

$requiredPackages = @("pikepdf", "pdf2image", "pytesseract", "PIL")
$missingPackages = @()

foreach ($package in $requiredPackages) {
    $checkCmd = if ($package -eq "PIL") { "import PIL; print('OK')" } else { "import $package; print('OK')" }
    $result = python -c $checkCmd 2>&1
    
    if ($result -ne "OK") {
        $installName = if ($package -eq "PIL") { "Pillow" } else { $package }
        $missingPackages += $installName
    }
}

if ($missingPackages.Count -gt 0) {
    Write-Host "Missing packages: $($missingPackages -join ', ')" -ForegroundColor Yellow
    Write-Host "Installing..." -ForegroundColor Yellow
    python -m pip install $missingPackages --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Failed to install packages" -ForegroundColor Red
        Exit 1
    }
    Write-Host "✓ Packages installed successfully" -ForegroundColor Green
} else {
    Write-Host "✓ All required packages are installed" -ForegroundColor Green
}

# Check for Tesseract OCR
Write-Host "Checking for Tesseract OCR..." -ForegroundColor Cyan
$tesseractCheck = python -c "import pytesseract; pytesseract.get_tesseract_version(); print('OK')" 2>&1

if ($tesseractCheck -notlike "*OK*") {
    Write-Host "Warning: Tesseract OCR not found or not configured" -ForegroundColor Yellow
    Write-Host "Please install Tesseract from: https://github.com/tesseract-ocr/tesseract" -ForegroundColor Yellow
    Write-Host "After installation, you may need to set the path in your script" -ForegroundColor Yellow
    $continue = Read-Host "Continue anyway? (y/n)"
    if ($continue -ne "y") {
        Exit 1
    }
} else {
    Write-Host "✓ Tesseract OCR is installed" -ForegroundColor Green
}

# Get the Python script path (same directory as this PowerShell script)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonScript = Join-Path $scriptDir "fix-pdf-rotation.py"

# Check if Python script exists
if (-not (Test-Path $pythonScript)) {
    Write-Host "Error: Python script not found at $pythonScript" -ForegroundColor Red
    Exit 1
}

# Validate input path
if (-not (Test-Path $Path)) {
    Write-Host "Error: Path not found: $Path" -ForegroundColor Red
    Exit 1
}

Write-Host "`nStarting PDF rotation fix..." -ForegroundColor Cyan
Write-Host "Input: $Path" -ForegroundColor Cyan
Write-Host "Output suffix: $OutputSuffix`n" -ForegroundColor Cyan

# Call Python script
python $pythonScript $Path $OutputSuffix

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n✓ Processing completed successfully" -ForegroundColor Green
} else {
    Write-Host "`n✗ Processing failed with exit code $LASTEXITCODE" -ForegroundColor Red
    Exit $LASTEXITCODE
}
