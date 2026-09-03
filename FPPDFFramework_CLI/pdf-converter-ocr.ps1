# -*- coding: utf-8-with-bom -*-
# ====================================================
#   PowerShell PDF OCR Conversion Script
#   Author: Flyingbee
#   Features: Same as pdf-converter.ps1, PLUS OCR enabled by default.
#             Each PDF is converted with: -r 1 -g <OCR_LANG>
#             (-r 1 starts OCR; scanned/image-based PDFs become searchable text).
#   Requires: The matching Tesseract language data (*.traineddata) shipped in
#             the "Resources.bundle" folder next to the EXE.
# ====================================================

# ============ CONFIGURATION ============
$PDF_DIR        = "."                             # Input directory ("." = current)
$TOOL_PATH      = "FPPDFConverter.exe"            # Path to converter
$OUTPUT_FORMAT  = "docx"                          # Output format: docx, pptx, xlsx, html, csv, txt, png, jpeg, etc.
$OUTPUT_DIR     = "converted_ocr"                 # Output subdirectory (kept separate from the non-OCR script)
$LOG_FILE       = "conversion-ocr.log"            # Log file name (kept separate from the non-OCR script)
$SUCCESS_MARKER = "Successfully converted!"       # Success keyword
$NUM_THREADS    = 4                               # Thread count per conversion
$OCR_ENABLE     = 1                               # 1 = start OCR (-r 1), 0 = do not use OCR
$OCR_LANG       = "eng"                           # OCR language for Tesseract: e.g. eng, chi_sim, jpn,
                                                  #                       osd, deu, fra, spa, rus, ...
# =======================================

# ============ Helper: validate a real PDF ============
# Reads the first 4 bytes and checks for the "%PDF" magic header.
# Files that merely carry a ".pdf" name (empty files, Git LFS pointers,
# partially-downloaded files, etc.) are therefore skipped.
function Test-ValidPdf {
    param([string]$Path)
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        try {
            $bytes = New-Object byte[] 4
            if ($fs.Read($bytes, 0, 4) -lt 4) { return $false }
            return ([System.Text.Encoding]::ASCII.GetString($bytes) -eq "%PDF")
        } finally { $fs.Dispose() }
    } catch { return $false }
}
# ====================================================

# Change to the PDF directory
try {
    Set-Location $PDF_DIR -ErrorAction Stop
} catch {
    Write-Host "❌ Failed to access directory: $PDF_DIR" -ForegroundColor Red
    "❌ Failed to access directory: $PDF_DIR" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
    exit 1
}

# Initialize log
$START_TIME_TOTAL = Get-Date
"📄 PDF to $OUTPUT_FORMAT Conversion Log (OCR: $OCR_LANG)" | Out-File -FilePath $LOG_FILE -Encoding UTF8
"📅 Started at: $START_TIME_TOTAL" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
"========================================" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
"⚙️  Threads: $NUM_THREADS" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
"🔤 OCR language: $OCR_LANG" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
"📤 Output folder: .\$OUTPUT_DIR/" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
"" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

# Create output directory if it does not exist, then clean old output files
$OutputPath = Join-Path -Path $PWD -ChildPath $OUTPUT_DIR
if (-not (Test-Path -Path $OutputPath -PathType Container)) {
    New-Item -Path $OutputPath -ItemType Directory | Out-Null
    Write-Host "📁 Created output directory: $OutputPath"
    "📁 Created output directory: $OutputPath" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
}
$OutputGlob = Join-Path -Path $OutputPath -ChildPath "*.${OUTPUT_FORMAT}"
Remove-Item -Path $OutputGlob -ErrorAction SilentlyContinue
Write-Host "✅ Cleaned existing .${OUTPUT_FORMAT} files in '$OUTPUT_DIR'."
"✅ Cleaned existing .${OUTPUT_FORMAT} files in '$OUTPUT_DIR'." | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

# Enumerate *.pdf files (files only, sorted by name)
$ALL_FILES = @(Get-ChildItem -Path "*.pdf" -File -ErrorAction SilentlyContinue | Sort-Object Name)

if ($ALL_FILES.Count -eq 0) {
    Write-Host "⚠️ No PDF files found."
    "⚠️ No PDF files found." | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
    exit 0
}

# Keep only valid PDFs: verify the "%PDF" magic header so files that merely
# carry a .pdf name (empty files, Git LFS pointers, etc.) are skipped.
$PDF_FILES     = @()
$SKIPPED_COUNT = 0
foreach ($File in $ALL_FILES) {
    if (Test-ValidPdf -Path $File.FullName) {
        $PDF_FILES += $File
    } else {
        Write-Host "⚠️ Skipping invalid PDF file: $($File.Name)"
        "⚠️ Skipping invalid PDF file: $($File.Name)" | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
        $SKIPPED_COUNT++
    }
}

if ($PDF_FILES.Count -eq 0) {
    Write-Host "⚠️ No valid PDF files found (all *.pdf files failed the %PDF check)."
    "⚠️ No valid PDF files found (all *.pdf files failed the %PDF check)." | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
    exit 0
}

$TOTAL = $PDF_FILES.Count
Write-Host "📦 Found $TOTAL valid PDF file(s). Starting conversion (OCR: $OCR_LANG)..."
"📦 Found $TOTAL valid PDF file(s)." | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
if ($SKIPPED_COUNT -gt 0) {
    Write-Host "⚠️ Skipped $SKIPPED_COUNT invalid file(s)."
    "⚠️ Skipped $SKIPPED_COUNT invalid file(s)." | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
}

# ============ Counters ============
$SUCCESS_COUNT = 0
$FAIL_COUNT = 0

# ============ Conversion Loop ============
foreach ($i in 0..($TOTAL - 1)) {
    $PDF_FILE   = $PDF_FILES[$i]
    $INDEX      = $i + 1
    $PDF_NAME   = $PDF_FILE.Name

    $MESSAGE = "------------------"
    Write-Host $MESSAGE
    $MESSAGE | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

    $MESSAGE = "【$PDF_NAME】"
    Write-Host $MESSAGE
    $MESSAGE | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

    $MESSAGE = "🔄 OCR Converting ($INDEX/$TOTAL)..."
    Write-Host $MESSAGE
    $MESSAGE | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

    $SEPARATOR = "--------------------------------------------------------------------------------"
    $SEPARATOR | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

    $START_FILE = Get-Date

    # Build output path: -o expects a full output file path (the SDK treats it
    # as a file name), so pass "<outputdir>/<name>.<format>" rather than only
    # the directory.
    $OUTPUT_FILE_NAME = [System.IO.Path]::ChangeExtension($PDF_NAME, $OUTPUT_FORMAT)
    $OUTPUT_FILE_PATH = Join-Path -Path $OutputPath -ChildPath $OUTPUT_FILE_NAME

    # Build command arguments. "-r 1" starts OCR and "-g" selects the OCR language.
    $OCR_ARG = ""
    if ($OCR_ENABLE -eq 1) { $OCR_ARG = "-r 1 -g $OCR_LANG" }
    $ARGUMENTS = "-a PDF2Files -i `"$($PDF_FILE.FullName)`" -o `"$OUTPUT_FILE_PATH`" -f $OUTPUT_FORMAT -p all -t $NUM_THREADS $OCR_ARG"

    # Print the exact command line to the console and the log before running it
    $RUN_LINE = "▶️ Running: `"$TOOL_PATH`" $ARGUMENTS"
    Write-Host $RUN_LINE
    $RUN_LINE | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

    # Execute conversion via cmd /c for better path handling.
    # NOTE: keep "2>&1" INSIDE the cmd command string. If it is applied at the
    # PowerShell level instead, PowerShell 5.1 wraps the tool's stderr output
    # (e.g. license warnings) into ErrorRecord objects and the log/console gets
    # polluted with "cmd : Error ... + CategoryInfo ..." noise.
    $OUTPUT = cmd /c "`"$TOOL_PATH`" $ARGUMENTS 2>&1" | Out-String

    # Write conversion output to log
    $OUTPUT | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8

    # Check success marker in output
    if ($OUTPUT -match [regex]::Escape($SUCCESS_MARKER)) {
        $END_FILE = Get-Date
        $DURATION = [math]::Round(($END_FILE - $START_FILE).TotalSeconds)
        Write-Host "✅ Done ($DURATION s)"
        "✅ Succeeded in $DURATION seconds." | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
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
$MINUTES = [math]::Floor($DURATION_TOTAL / 60)
$SECONDS = $DURATION_TOTAL % 60

Write-Host ""
Write-Host "------------------------------------------"
Write-Host "FINAL SUMMARY"
Write-Host "------------------------------------------"
Write-Host "✅ $SUCCESS_COUNT succeeded"
Write-Host "❌ $FAIL_COUNT failed"
Write-Host "📤 $TOTAL total"

if ($DURATION_TOTAL -ge 60) {
    Write-Host "⏱️  Total time: ${MINUTES}m ${SECONDS}s"
} else {
    Write-Host "⏱️  Total time: ${DURATION_TOTAL}s"
}

# 🎉 Final completion message
if ($FAIL_COUNT -eq 0 -and $TOTAL -gt 0) {
    Write-Host "🎉 All conversions completed successfully!"
} elseif ($SUCCESS_COUNT -gt 0) {
    Write-Host "⚠️  Some conversions failed. Check $LOG_FILE for details."
} else {
    Write-Host "💥 All conversions failed!"
}

# Append summary to log
@"

------------------------------------------
FINAL SUMMARY
------------------------------------------
✅ $SUCCESS_COUNT succeeded
❌ $FAIL_COUNT failed
📤 $TOTAL total
⏱️  Total time: ${MINUTES}m ${SECONDS}s
📁  Output: .\$OUTPUT_DIR/
📅 Finished at: $END_TIME_TOTAL
------------------------------------------
"@ | Out-File -Append -FilePath $LOG_FILE -Encoding UTF8
