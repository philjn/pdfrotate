param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Path,
    
    [Parameter(Position=1)]
    [string]$OutputSuffix = "_fixed"
)

<#
.SYNOPSIS
    Fixes PDF page rotations by detecting and correcting rotated pages.

.DESCRIPTION
    Processes PDF files in a directory or a single PDF file, detecting pages with 
    rotation metadata (90°, 180°, 270°) and correcting them to 0° (upright).
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
    Requires Python 3 with pikepdf package installed.
    Run: python -m pip install pikepdf
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

# Check if pikepdf is installed
Write-Host "Checking for pikepdf package..." -ForegroundColor Cyan
$pikepdfCheck = python -c "import pikepdf; print('OK')" 2>&1

if ($pikepdfCheck -ne "OK") {
    Write-Host "pikepdf not found. Installing..." -ForegroundColor Yellow
    python -m pip install pikepdf --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Failed to install pikepdf" -ForegroundColor Red
        Exit 1
    }
    Write-Host "✓ pikepdf installed successfully" -ForegroundColor Green
} else {
    Write-Host "✓ pikepdf is installed" -ForegroundColor Green
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
