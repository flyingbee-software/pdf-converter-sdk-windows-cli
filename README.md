# PDF to Word, Excel & PowerPoint Conversion SDK for Windows Server (CLI)

Welcome to the official documentation for the **Flyingbee PDF Conversion SDK for Windows Server**. This comprehensive guide provides developers with step-by-step instructions for deploying and using our high-performance, command-line PDF conversion library on Windows systems. Convert PDFs to Word, Excel, PowerPoint, images, and more — directly from the terminal or integrated into your application.

---

## 🕹️Try It Online

Want to try before deploying? Use the **Flyingbee PDF Converter Online** to convert PDFs to MS Office formats directly in your browser — no installation required.

[🚀 Launch Web Demo](https://www.flyingbee.com/pdf-sdk/documentation/guides/server-windows/)

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
- **Files**: EXE, DLL, and resource files must reside in the same directory
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

```
./FPPDFConverter -h
```

Sample output:

```
=== Welcome to use Flyingbee PDF Converter! ===
App exe folder: C:UsersAdministratorDesktopApps PDF Console
FPPDFConverter v1.6.1.0
FPPDFConverter - Convert PDF to Word/Excel/PPT, etc.
Usage: FPPDFConverter -a CMD_Name [input]filepath.pdf [options]docx [output]filepath.docx
Where options are:
  -a [app]     The apps: "PDF2Files", "Images2PDF", "Text2Word"
  -h [help]    print this help message
  -c [open]    open converted file [todo]
=== Thank you for using our product! ===
```

---

## Supported Conversion Apps

### PDF2Files — Convert PDF to Office Documents, TXT, Images

Convert all PDF files in a folder to Word, Excel, PowerPoint, or other formats:

```
./FPPDFConverter -a PDF2Files -i "Test.pdf" -f docx -p all
```

### Images2PDF — Convert Images to PDF

Merge multiple images into a single PDF document.

**Supported image formats:** `.jpg`, `.jpeg`, `.png`, `.gif`, `.tif`, `.tiff`, `.bmp`, `.ico`

**Sample commands:**

```
./FPPDFConverter -a Images2PDF -h
./FPPDFConverter -a Images2PDF -i "C:/Users/Administrator/Desktop/images/" -o "C:/Users/Administrator/Desktop/Images2PDF.pdf"
./FPPDFConverter -a Images2PDF -i "C:/Users/Administrator/Desktop/images/" -o "C:/Users/Administrator/Desktop/Images2PDF.pdf" -p auto
```

### Text2Word — Convert Text File to Word

Convert plain text files to Word documents with formatting:

```
./FPPDFConverter -a Text2Word -h
./FPPDFConverter -a Text2Word -i "C:/Users/Administrator/Desktop/Txt Tests/text-cn-gb.txt" -o "C:/Users/Administrator/Desktop/Text2Word.docx"
./FPPDFConverter -a Text2Word -i "C:/Users/Administrator/Desktop/Txt Tests/text-en-ansi.txt" -o "C:/Users/Administrator/Desktop/Text2Word.docx"
```

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

| Platform | Command |
| :--- | :--- |
| **PowerShell** | `./FPPDFConverter -i xxxx.pdf` |
| **Command Prompt (CMD)** | `FPPDFConverter -i xxxx.pdf` |

---

## Support

Should you have any questions, please contact Flyingbee Support:

- **Email:** [support@flyingbee.com](mailto:support@flyingbee.com)
- **Contact Page:** [https://www.flyingbee.com/contact-us/](https://www.flyingbee.com/contact-us/)

---

## Release History

| Date | Version |
| :--- | :--- |
| 2026-08-26 | v10.3.6.0 |
| 2025-10-15 | v2.0.2.0 |
| 2024-08-24 | v1.6.8.0 |

---

*Powered by Flyingbee PDF Conversion SDK — Your trusted partner for PDF document conversion on Windows servers.*
