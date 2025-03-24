# MD to DOCX Converter

A tool to convert Markdown (.md) files to Word (.docx) format using a customizable Word template.

## Overview

The MD to DOCX Converter is designed to simplify the process of converting Markdown files to properly formatted Word documents. It maintains formatting integrity and applies appropriate styles from your chosen Word template.

Key features:
* Convert single Markdown files or entire folders of .md files
* Customize using Word templates for consistent branding and styling
* Intelligently apply heading styles, lists, and code formatting
* Simple, user-friendly interface

## Installation

### System Requirements
* macOS operating system
* Python 3.6+
* Internet connection (for downloading dependencies)

### Installation Steps

1. Download the application files
2. Open Terminal and navigate to the folder containing the files
3. Run the installer:
   ```
   python -m venv venv
   source venv/bin/activate
   python install.py
   ```
4. The installer will:
   * Check for and install required Python packages
   * Create an application directory
   * Set up a desktop shortcut
   * Configure necessary files

If you encounter permission issues, you may need to run:
```
chmod +x install.py
python install.py
```

## Using the Application

1. Launch the application by clicking the "MD to DOCX Converter" icon on your desktop

2. Select Input:
   * Choose "Single File" to convert one Markdown file
   * Choose "Folder" to convert all .md files in a directory
   * Click "Browse..." to select your file or folder

3. Select Template:
   * Click "Browse..." to select a Word template (.docx)
   * The application comes with a default template

4. Select Output Location (Optional):
   * By default, converted files are saved in the same location as the input
   * Click "Browse..." to choose a different output location

5. Click "Convert" to begin the conversion process

6. Once complete, a success message will appear, and you can find your converted files in the output location

7. Click "Close" to exit the application when finished

## Troubleshooting

* **Error: Missing dependencies** - Run the installer again to check for and install missing packages