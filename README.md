# PDF Merge Context Menu Tool

This project provides a Python-based tool for merging multiple PDF files with Windows File Explorer context menu integration.

## Quick Install (One-Line Command)

Open PowerShell and run:

```powershell
irm https://raw.githubusercontent.com/YOUR_USERNAME/pdftool/main/install.ps1 | iex
```

Replace `YOUR_USERNAME` with your GitHub username after uploading this repository to GitHub.

## Alternative: Manual Setup

If you've cloned the repository, run the setup script from an elevated PowerShell prompt:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

Use `-SkipRegister` if you just want to install dependencies, or `-VerbName "Combine PDFs"` to customize the context-menu label.

## Features

- **Merge multiple PDFs** into a single document
- **Alphabetical ordering** of input files for predictable results
- **Windows 11-style notifications** - Modern toast alerts
- **Silent operation** - No console windows
- **Simple save dialog** - Choose output location easily
- **Multi-select support** - Select and merge many files at once

## Prerequisites

If you prefer to set things up manually, make sure you have:

- Windows 10 or later
- Python 3.8+ installed with the standard `py`/`pyw` launcher registered under `C:\Windows`
- Required packages:
  ```powershell
  py -m pip install pypdf winotify filelock
  ```

## Files

- `merge_pdfs.py` - Core merge logic with save dialog
- `merge_pdf_handler.py` - Context menu handler for file collection
- `setup.ps1` - Automated installation (Python, packages, context menu)
- `register_context_menu.ps1` - Registers the Merge PDF context menu
- `unregister_context_menu.ps1` - Removes the context menu entry
- `install.ps1` - One-line web installer script

## Register the context menu entry

1. Launch PowerShell in this folder.
2. Allow the script to run (if needed) and register the verb:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\register_context_menu.ps1
   ```
   The script locates `pyw.exe` automatically; pass `-PythonLauncher` if you need a custom Python interpreter.

### Using the Context Menu

1. Select two or more PDF files in File Explorer
2. Right-click on one of the selected files
3. Click **Merge PDF**
4. Choose where to save the merged file
5. Toast notification shows when complete

## Removing the entry

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\unregister_context_menu.ps1
```

## Command-line usage

You can also run the merger directly from the command line:

```powershell
py .\merge_pdfs.py file1.pdf file2.pdf -o merged.pdf
```

Omit the `-o` flag to trigger the graphical save dialog.

## Troubleshooting

**Error logs:**

- Merge: `%TEMP%\merge_pdfs.log`
- Handler errors: `%TEMP%\pdf_merge_handler_error.log`

**Context menu not appearing?**

- Re-run `register_context_menu.ps1`
- Check that Python and packages are installed

**Merge failing?**

- Ensure all selected files are valid PDFs
- Check the log file for error details
# pdftool
