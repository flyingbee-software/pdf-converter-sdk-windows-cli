# PDF to Word, Excel & PowerPoint Conversion SDK for Windows Server (CLI)

Welcome to the official documentation for the **Flyingbee PDF Conversion SDK for Windows Server**. This comprehensive guide provides developers with step-by-step instructions for deploying and using our high-performance, command-line PDF conversion library on Windows systems. Convert PDFs to Word, Excel, PowerPoint, images, and more — directly from the terminal or integrated into your application.

---

## 🕹️Try It Online

Want to try before deploying? Use the **Flyingbee PDF Converter Online** to convert PDFs to MS Office formats directly in your browser — no installation required.

[🚀 Launch Web Demo](https://www.flyingbee.com/pdf-converter/?utm_source=github_readme_conversion_sdk_windows_cli&utm_medium=referral&utm_campaign=github_readme_conversion_sdk_windows_cli)

---

## What is Flyingbee PDF Conversion SDK?

Flyingbee PDF Conversion SDK is a flexible, high-performance **PDF to MS Office conversion library** for Windows, Linux, and Web. It helps you convert PDFs into editable and well-formatted documents while accurately preserving:

- Original text, images, and layouts
- Hyperlinks and tables
- Bezier graphics and vector shapes
- Complex formatting and styling

Whether you need **PDF to DOCX**, **PDF to XLSX**, **PDF to PPTX**, or image conversion, this SDK delivers reliable results with minimal integration effort.

---

## System Requirements

Before deployment, ensure your Windows server meets the following requirements:

- **Operating System**: Windows 7 / Server 2008 R2 or later (64-bit recommended)
- **Runtime**: Microsoft Visual C++ Redistributable (included in the package)
- **Files**: EXE, DLL, and resource files (including the `Resources.bundle` folder used for OCR) must reside in the same directory
- **Permissions**: The EXE file must have executable permissions

---

## Quick Start Guide

### 1. Navigate to the Executable Directory

Open PowerShell or Command Prompt and navigate to the folder containing the EXE file:

```
cd "C:/Users/Administrator/Desktop/Exe-Folder/"
```

> **Note:** Use forward slashes (`/`) in file paths even though Windows typically uses backslashes (`\`). Enclose paths in double quotes to avoid issues with spaces or special characters.

### 2. Execute a Conversion Command

```
./FPPDFConverter -a PDF2Files -i "Test.pdf" -f docx -p all
```

### 3. Display Help Information

Show the general help screen:

```
./FPPDFConverter -h
```

Show command-specific help (recommended — lists every option that command accepts):

```
./FPPDFConverter -a PDF2Files -h
./FPPDFConverter -a Images2PDF -h
./FPPDFConverter -a Text2Word -h
```

Sample general help output:

```
FPPDFConverter - High-performance PDF to Word/Excel/PPT/Image Converter

USAGE:
  FPPDFConverter -a <command> [options]

COMMANDS:
  PDF2Files       Convert PDF documents to Word, Excel, PPT, Images, etc.
  Images2PDF      Merge multiple images into a single PDF document.
  Text2Word       Convert plain text files to formatted Word documents.

GENERAL OPTIONS:
  -h              Show this help message and exit.
  -c              Open the converted file automatically upon completion.

EXAMPLES:
  FPPDFConverter -a PDF2Files -i input.pdf -o output.docx
  FPPDFConverter -a Images2PDF -i ./images -o merged.pdf -s A4

Run 'FPPDFConverter -a <command> -h' for command-specific options.
```

All supported options are summarized in the [Command-Line Reference](#command-line-reference) section below.

---

## Supported Conversion Apps

The CLI ships with three conversion tasks, selected with the `-a` switch:

- **`PDF2Files`** — Convert PDF documents to Word, Excel, PPT, HTML, TXT, CSV, or images.
- **`Images2PDF`** — Merge multiple images into a single PDF document.
- **`Text2Word`** — Convert plain text files to formatted Word documents.

### PDF2Files — Convert PDF to Word, Excel, PPT, Images, and More

Convert a PDF document into Word, Excel, PowerPoint, HTML, TXT, CSV, or raster image formats:

```
./FPPDFConverter -a PDF2Files -i "Test.pdf" -o "Test.docx" -f docx
./FPPDFConverter -a PDF2Files -i "Test.pdf" -o "Page" -f png -p "1-3,5"
./FPPDFConverter -a PDF2Files -i "encrypted.pdf" -o "decrypted.docx" -w "your-password"
./FPPDFConverter -a PDF2Files -i "scanned.pdf" -o "ocr.docx" -f docx -r -g eng
```

**Supported output formats:** `docx` (default), `pptx`, `xlsx`, `html`, `csv`, `txt`, `jpeg`, `jpg`, `png`, `bmp`, `tif`, `tiff`, `gif`.

See [PDF2Files Options](#pdf2files-options) for the complete option reference.

### Images2PDF — Convert Images to PDF

Merge multiple images into a single PDF document.

**Supported image formats:** `.jpg`, `.jpeg`, `.png`, `.gif`, `.tif`, `.tiff`, `.bmp`, `.ico`

**Sample commands:**

```
./FPPDFConverter -a Images2PDF -h
./FPPDFConverter -a Images2PDF -i "C:/Users/Administrator/Desktop/images/" -o "C:/Users/Administrator/Desktop/Images2PDF.pdf"
./FPPDFConverter -a Images2PDF -i "C:/Users/Administrator/Desktop/images/" -o "C:/Users/Administrator/Desktop/Images2PDF.pdf" -s auto
./FPPDFConverter -a Images2PDF -i "C:/Users/Administrator/Desktop/images/" -o "C:/Users/Administrator/Desktop/Images2PDF.pdf" -s A4 -m 0.5 -t "My Album" -A "Author"
```

> **Tip:** Control the paper size with `-s` (e.g., `-s A4` or `-s auto`) and the page margins with `-m`. Full option details are in [Images2PDF Options](#images2pdf-options).

### Text2Word — Convert Text File to Word

Convert plain text files to Word documents with formatting:

```
./FPPDFConverter -a Text2Word -h
./FPPDFConverter -a Text2Word -i "C:/Users/Administrator/Desktop/Txt Tests/text-cn-gb.txt" -o "C:/Users/Administrator/Desktop/Text2Word.docx"
./FPPDFConverter -a Text2Word -i "C:/Users/Administrator/Desktop/Txt Tests/text-en-ansi.txt" -o "C:/Users/Administrator/Desktop/Text2Word.docx"
./FPPDFConverter -a Text2Word -i "C:/Users/Administrator/Desktop/Txt Tests/text-cn-gb.txt" -o "C:/Users/Administrator/Desktop/Text2Word.docx" -s A4 -b "SimSun" -d 12
```

See [Text2Word Options](#text2word-options) for the complete option reference.

---

## Command-Line Reference

Every task is invoked through the single `FPPDFConverter` executable. Select the task with the `-a` switch and pass its options as shown below:

```
FPPDFConverter -a <command> [options]
```

The option tables mirror the help text printed by `FPPDFConverter -a <command> -h`.

> **Note:** Option switches are interpreted in the context of each command. For example, `-w` is the PDF *password* for `PDF2Files`, but the custom paper *width* for `Images2PDF` and `Text2Word`. Always run `FPPDFConverter -a <command> -h` if you are unsure.

### Common Options (all commands)

| Option | Description |
| :--- | :--- |
| `-h` | Show the help message for the current command and exit. |
| `-c` | Open the converted file automatically upon completion. |

### PDF2Files Options

```
FPPDFConverter -a PDF2Files -i <input.pdf> -o <output> [options]
```

**Required Options**

| Option | Description |
| :--- | :--- |
| `-i <path>` | Input PDF file path. |
| `-o <path>` | Output file or directory path. |

**Conversion Settings**

| Option | Description |
| :--- | :--- |
| `-f <format>` | Target format: `docx` (default), `pptx`, `xlsx`, `html`, `csv`, `txt`, `jpeg`, `jpg`, `png`, `bmp`, `tif`, `tiff`, `gif`. |
| `-p <ranges>` | Page ranges to convert (e.g., `1-3,5,8-10`). |
| `-t <threads>` | Number of processing threads (1–10, default: 1). |
| `-x` | (XLSX only) Merge multiple sheets into a single sheet. |
| `-w <password>` | Password for opening encrypted PDF files. |

**HTML-Specific Options**

| Option | Description |
| :--- | :--- |
| `-l <mode>` | Layout mode: `0` = Exact page layout (default), `1` = Text flow. |
| `-e <mode>` | Paragraph style: `0` = Line break (default), `1` = Indent. |
| `-b <mode>` | Navigation bar: `0` = Disabled, `1` = Enabled (default). |
| `-m <level>` | Resource merge: `0` = None (default), `1` = CSS/JS, `2` = CSS/JS/Small Images, `3` = CSS/JS/All Images. |
| `-z` | Package the HTML output and its resources into a ZIP file. |

**Image-Specific Options**

| Option | Description |
| :--- | :--- |
| `-d <dpi>` | Image resolution (72–600, default: 144). |
| `-n <mode>` | Anti-aliasing: `0` = None, `1` = Font smoothing (default). |

**OCR Options**

| Option | Description |
| :--- | :--- |
| `-r` | Enable OCR (Optical Character Recognition). |
| `-g <lang>` | OCR language for Tesseract (e.g., `eng`, `chi_sim`, `jpn`). |

> OCR relies on the Tesseract language packs shipped inside the `Resources.bundle` folder — keep that folder alongside the EXE (see [System Requirements](#system-requirements)).

### Images2PDF Options

```
FPPDFConverter -a Images2PDF -i <folder> -o <output.pdf> [options]
```

**Required Options**

| Option | Description |
| :--- | :--- |
| `-i <folder>` | Source folder containing images. |
| `-o <path>` | Output PDF file path. |

**Page & Layout Options**

| Option | Description |
| :--- | :--- |
| `-s <size>` | Paper size: `A0`–`A10`, `Letter`, `Legal`, `Tabloid`, `4x6`, `5x7`, or `auto` (default: none). |
| `-n <orient>` | Orientation: `portrait`, `landscape`, `auto` (default: `auto`). |
| `-w <inches>` | Custom paper width in inches (used when the size is custom). |
| `-l <inches>` | Custom paper height in inches (used when the size is custom). |
| `-m <margin>` | Page margins in inches (append `mm` for millimeters). |
| `-z <scale>` | Scale mode: `fit` (default), `fw`, `fh`, `reduce`, `rw`, `rh`, `none`. |
| `-r <crop>` | Crop/Expand: `none` (default), `height`, `width`, `both`. |

**Metadata Options**

| Option | Description |
| :--- | :--- |
| `-t <title>` | Document title. |
| `-A <author>` | Document author. |
| `-k <keywords>` | Document keywords. |
| `-S <subject>` | Document subject. |
| `-C <creator>` | Document creator. |

### Text2Word Options

```
FPPDFConverter -a Text2Word -i <input.txt> -o <output.docx> [options]
```

**Required Options**

| Option | Description |
| :--- | :--- |
| `-i <path>` | Input plain text file path. |
| `-o <path>` | Output Word document path. |

**Page & Layout Options**

| Option | Description |
| :--- | :--- |
| `-s <size>` | Paper size: `A0`–`A10`, `Letter`, `Legal`, `Tabloid`, `4x6`, `5x7`, or `auto` (default: none). |
| `-n <orient>` | Orientation: `portrait`, `landscape`, `auto` (default: `auto`). |
| `-w <inches>` | Custom paper width in inches (used when the size is custom). |
| `-l <inches>` | Custom paper height in inches (used when the size is custom). |
| `-m <margin>` | Page margins in inches (append `mm` for millimeters). |
| `-e <columns>` | Number of columns (1–10, default: 1). |

**Font Options**

| Option | Description |
| :--- | :--- |
| `-b <font>` | Font name: `Arial` (default), `Calibri`, `Courier`, `Times New Roman`, `Helvetica`, `Verdana`, `Consolas`, `SimSun`, `SimHei`, `FangSong`, `KaiTi`, `Microsoft YaHei`. |
| `-d <size>` | Font size in points (8–72, default: 12). |

### Command Examples at a Glance

| Task | Command |
| :--- | :--- |
| PDF → Word | `FPPDFConverter -a PDF2Files -i "input.pdf" -o "output.docx" -f docx` |
| PDF → PowerPoint | `FPPDFConverter -a PDF2Files -i "input.pdf" -o "output.pptx" -f pptx` |
| PDF → Excel | `FPPDFConverter -a PDF2Files -i "input.pdf" -o "output.xlsx" -f xlsx` |
| PDF → HTML (packed as ZIP) | `FPPDFConverter -a PDF2Files -i "input.pdf" -o "output" -f html -z` |
| Convert a page range to PNG | `FPPDFConverter -a PDF2Files -i "input.pdf" -o "Page" -f png -p "1-3,5"` |
| OCR a scanned PDF | `FPPDFConverter -a PDF2Files -i "scanned.pdf" -o "ocr.docx" -f docx -r -g eng` |
| Open encrypted PDF | `FPPDFConverter -a PDF2Files -i "private.pdf" -o "out.docx" -w "your-password"` |
| Images → PDF (auto page size) | `FPPDFConverter -a Images2PDF -i "./images" -o "merged.pdf" -s auto` |
| Images → PDF with title/author | `FPPDFConverter -a Images2PDF -i "./images" -o "merged.pdf" -s A4 -t "My Album" -A "Author"` |
| TXT → Word (Chinese font) | `FPPDFConverter -a Text2Word -i "notes.txt" -o "notes.docx" -b "SimSun" -d 12` |

---

## Integrating with Your Application

The SDK provides a console-based executable that can be invoked programmatically from your application code. Below are integration guides for the most popular development languages:

- [**How to invoke a console program in C++ on Windows?**](https://www.flyingbee.com/pdf-sdk/documentation/guides/server-windows/archives/65.html) — Learn how to launch and communicate with FPPDFConverter from C++ applications using Win32 API or _popen.
- [**How to invoke a console program in .NET programming?**](https://www.flyingbee.com/pdf-sdk/documentation/guides/server-windows/archives/64.html) — Integrate PDF conversion into your C# or VB.NET applications using Process.Start and standard I/O streams.
- [**How to invoke a console program in JAVA?**](https://www.flyingbee.com/pdf-sdk/documentation/guides/server-windows/archives/66.html) — Call the conversion executable from Java using ProcessBuilder and handle conversion feedback in real time.
- [**How you can invoke a console program in PHP?**](https://www.flyingbee.com/pdf-sdk/documentation/guides/server-windows/archives/67.html) — Trigger PDF conversions from your PHP web applications using shell_exec or proc_open for server-side document processing.

These guides cover process creation, argument passing, standard input/output capture, and error handling — everything you need to embed PDF conversion into your own software.

---

## Conversion Feedback

During the conversion process, FPPDFConverter provides real-time feedback on conversion progress. Upon successful completion, you will see a confirmation message along with the elapsed time:

```
Successfully converted! (Elapsed time: 2.34 seconds)
```

---

## PowerShell Execution Policy

If you encounter script execution restrictions in PowerShell, allow script execution for the current user:

```
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## Platform Command Differences

PowerShell and Command Prompt differ slightly in how you launch the executable and quote arguments:

| Platform | Command |
| :--- | :--- |
| **PowerShell** | `./FPPDFConverter -a PDF2Files -i "Test.pdf" -o "Test.docx"` |
| **Command Prompt (CMD)** | `FPPDFConverter.exe -a PDF2Files -i "Test.pdf" -o "Test.docx"` |

> In PowerShell, `./` prefixes the executable to run it from the current directory; in CMD you can call `FPPDFConverter.exe` directly when it is in the current directory or on `PATH`.

---

## Support

Should you have any questions, please contact Flyingbee Support:

- **Email:** [support@flyingbee.com](mailto:support@flyingbee.com)
- **Contact Page:** [https://www.flyingbee.com/contact-us/](https://www.flyingbee.com/contact-us/?utm_source=github_readme_conversion_sdk_windows_cli&utm_medium=referral&utm_campaign=github_readme_conversion_sdk_windows_cli)

---

## Release History

| Date | Version |
| :--- | :--- |
| 2026-08-26 | v10.3.6.0 |
| 2025-10-15 | v2.0.2.0 |
| 2024-08-24 | v1.6.8.0 |

---

*Powered by Flyingbee PDF Conversion SDK — Your trusted partner for PDF document conversion on Windows servers.*
