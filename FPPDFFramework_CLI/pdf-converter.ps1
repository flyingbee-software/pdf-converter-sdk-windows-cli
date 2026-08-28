# -*- coding: utf-8-with-bom -*-
# ====================================================
#   PowerShell PDF Conversion Script
#   Author: Flyingbee
#   Supports: Custom output folder, multi-threading, logging
# ====================================================

# ============ CONFIGURATION ============
$PDF_DIR        = "."                             # Input directory ("." = current)
$TOOL_PATH      = "FPPDFConverter.exe"            # Path to converter
$OUTPUT_FORMAT  = "docx"                          # Output format: docx, pptx, xlsx, etc.
$OUTPUT_DIR     = "converted"                     # Output subdirectory (default: "converted")
$LOG_FILE       = "conversion.log"                # Log file name
$SUCCESS_MARKER = "Successfully converted!"       # Success keyword
$NUM_THREADS    = 4                               # Thread count per conversion
# =======================================

# Change to PDF directory
try {
    Set-Location $PDF_DIR -ErrorAction Stop
} catch {
    Write-Host "❌ Failed to access directory: $PDF_DIR" -ForegroundColor Red
    "❌ Failed to access directory: $PDF_DIR" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
    exit 1
}

# Create output directory if not exists
$OutputPath = Join-Path -Path $PWD -ChildPath $OUTPUT_DIR
if (-not (Test-Path -Path $OutputPath -PathType Container)) {
    New-Item -Path $OutputPath -ItemType Directory | Out-Null
    Write-Host "📁 Created output directory: $OutputPath"
}
"📁 Output directory: $OutputPath" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

# Initialize log
$START_TIME_TOTAL = Get-Date
"📄 PDF to $OUTPUT_FORMAT Conversion Log" | Out-File -FilePath $LOG_FILE -Encoding UTF8
"📅 Started at: $START_TIME_TOTAL" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
"========================================" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
"⚙️  Threads: $NUM_THREADS" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
"📤 Output folder: .\$OUTPUT_DIR/" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
"" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

# Clean existing outputs in the output directory
$OutputGlob = Join-Path -Path $OutputPath -ChildPath "*.${OUTPUT_FORMAT}"
Remove-Item -Path $OutputGlob -ErrorAction SilentlyContinue
Write-Host "✅ Cleaned existing .${OUTPUT_FORMAT} files in '$OUTPUT_DIR'."

# Get PDF files (force array)
$PDF_FILES = @(Get-ChildItem -Path "*.pdf" -ErrorAction Stop)
$TOTAL = $PDF_FILES.Length

if ($TOTAL -eq 0) {
    Write-Host "⚠️ No PDF files found."
    "⚠️ No PDF files found." | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
    exit 0
}

Write-Host "📦 Found $TOTAL PDF file(s). Starting conversion..."

# Counters
$SUCCESS_COUNT = 0
$FAIL_COUNT = 0

# ============ Conversion Loop ============
foreach ($i in 0..($TOTAL - 1)) {
    $PDF_FILE = $PDF_FILES[$i]
    $INDEX = $i + 1
    $PDF_NAME = $PDF_FILE.Name
    $MESSAGE = "Converting ($INDEX/$TOTAL): $PDF_NAME"
    Write-Host $MESSAGE
    $MESSAGE | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

    $SEPARATOR = "--------------------------------------------------------------------------------"
    $SEPARATOR | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

    $START_FILE = Get-Date

    # Build output path
    $OUTPUT_FILE_NAME = [System.IO.Path]::ChangeExtension($PDF_NAME, $OUTPUT_FORMAT)
    $OUTPUT_FILE_PATH = Join-Path -Path $OutputPath -ChildPath $OUTPUT_FILE_NAME

    # Build command arguments
    $ARGUMENTS = "-a PDF2Files -i `"$($PDF_FILE.FullName)`" -o `"$OUTPUT_FILE_PATH`" -f $OUTPUT_FORMAT -p all -t $NUM_THREADS"

    # Log command
    "🔧 Running: `"$TOOL_PATH`" $ARGUMENTS" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

    # Execute conversion via cmd /c for better path handling
    $OUTPUT = cmd /c "`"$TOOL_PATH`" $ARGUMENTS" 2>&1 | Out-String

    # Write output to log
    $OUTPUT | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

    # Check success
    if ($OUTPUT -match [regex]::Escape($SUCCESS_MARKER)) {
        $END_FILE = Get-Date
        $DURATION = [math]::Round(($END_FILE - $START_FILE).TotalSeconds)
        Write-Host "✅ Done ($DURATION s)"
        "⏱️  Succeeded in $DURATION seconds." | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
        $SUCCESS_COUNT++
    } else {
        $END_FILE = Get-Date
        $DURATION = [math]::Round(($END_FILE - $START_FILE).TotalSeconds)
        Write-Host "❌ Failed ($DURATION s)" -ForegroundColor Red
        "❌ Failed: '$SUCCESS_MARKER' not found." | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
        # Output preview for debugging
        $PREVIEW = ($OUTPUT -split "`n") | Select-Object -First 3
        "➡️  Output preview:" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
        $PREVIEW | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
        $FAIL_COUNT++
    }

    "" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8  # Empty line
}

# ============ Final Summary ============
$END_TIME_TOTAL = Get-Date
$DURATION_TOTAL = [math]::Round(($END_TIME_TOTAL - $START_TIME_TOTAL).TotalSeconds)
$MINUTES = [Math]::Floor($DURATION_TOTAL / 60)
$SECONDS = $DURATION_TOTAL % 60

Write-Host ""
Write-Host "📊 Conversion Summary:" -ForegroundColor Cyan
Write-Host "------------------------"
Write-Host "✅ $SUCCESS_COUNT succeeded"
Write-Host "❌ $FAIL_COUNT failed"
Write-Host "📤 $TOTAL total"

if ($DURATION_TOTAL -ge 60) {
    Write-Host "⏱️  Total time: ${MINUTES}m ${SECONDS}s"
} else {
    Write-Host "⏱️  Total time: ${DURATION_TOTAL}s"
}

if ($FAIL_COUNT -eq 0 -and $TOTAL -gt 0) {
    Write-Host "🎉 All conversions completed successfully!" -ForegroundColor Green
} elseif ($SUCCESS_COUNT -gt 0) {
    Write-Host "⚠️  Some conversions failed. Check $LOG_FILE for details." -ForegroundColor Yellow
} else {
    Write-Host "💥 All conversions failed!" -ForegroundColor Red
}

# Append summary to log
@"
 
📊 FINAL SUMMARY
---------------
✅ $SUCCESS_COUNT succeeded
❌ $FAIL_COUNT failed
📤 $TOTAL total
⏱️  Total time: $($MINUTES)m $($SECONDS)s
🧵  Threads: $NUM_THREADS
📁  Output: .\$OUTPUT_DIR/
📅 Finished at: $END_TIME_TOTAL
"@ | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8