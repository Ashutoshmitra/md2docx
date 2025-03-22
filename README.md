# CBS Markdown to Word Converter

A professional utility for converting Markdown (.md) files to Microsoft Word (.docx) format while applying styles from a specified Word template.

## Overview

This application provides a simple and efficient way to convert Markdown content to properly formatted Word documents. It is especially designed for handling AI-generated Markdown content and applying consistent styling from Word templates.

## Features

- Convert single Markdown files or entire directories
- Apply styles from Word templates
- Intelligent handling of headings, lists, and formatting
- Removal of fixed numbering when styles auto-number
- Multi-level bullet handling
- Proper handling of soft carriage returns
- User-friendly graphical interface
- Command-line interface for batch processing
- Desktop shortcut for easy access

## System Requirements

- macOS (tested on macOS 10.15+)
- Python 3.6 or higher
- 50MB of free disk space

## Installation

### Option 1: Easy Install (Recommended)

1. Download the disk image (.dmg) file
2. Open the disk image
3. Drag the "CBS Markdown Converter" app to your Applications folder
4. Open the app from Applications or Launchpad

### Option 2: Installation Script

1. Download and extract the source code
2. Open Terminal
3. Navigate to the extracted directory
4. Run the installation script:
   ```
   ./install.sh
   ```
5. Follow the on-screen instructions
6. After installation, a shortcut will be created on your desktop

## Usage

### Using the Graphical Interface

1. Launch the application by double-clicking the desktop shortcut or opening it from Applications
2. Select a single Markdown file or a folder containing multiple .md files
3. Optionally, select a Word template (.dotx) file to apply its styles
4. Choose an output location (or use the same as input by checking the corresponding option)
5. Click "Convert" to start the conversion process
6. Once complete, the converted file(s) will be available at the specified location

### Using the Command Line

For advanced users or batch processing, a command-line interface is available:

```
python main.py --cli --input <input_file_or_dir> [--output <output_file_or_dir>] [--template <template_file>]
```

#### Parameters

- `--cli`: Run in command-line mode (no GUI)
- `--input`: Path to input Markdown file or directory
- `--output`: (Optional) Path to output Word file or directory
- `--template`: (Optional) Path to Word template file (.dotx)

## Creating Compatible Markdown

To ensure the best conversion results, follow these guidelines when creating Markdown content:

1. Use standard Markdown syntax for headings (`#`, `##`, etc.)
2. Use consistent list markers (preferably `-` for unordered lists)
3. Use proper indentation for nested lists (2 spaces per level)
4. Leave a blank line before and after lists, code blocks, and headings
5. For more detailed guidelines, refer to the included LLM prompt template

## LLM Prompt Template

For generating AI content that works seamlessly with this converter, use the included prompt template. This ensures that the Markdown follows the conventions expected by the converter.

The prompt template is available in the application directory as `llm_prompt_template.md`.

## Troubleshooting

### Common Issues

1. **Application doesn't start**
   - Ensure you have Python 3.6+ installed
   - Try running the application from Terminal using `./launch_converter.sh`

2. **Conversion fails**
   - Check that your Markdown file is valid
   - Ensure the template file exists and is a valid .dotx file
   - Check file permissions for the output directory

3. **Styling issues in output document**
   - Ensure your template contains the styles referenced in the Markdown
   - Check that the Markdown follows the conventions in the LLM prompt template

### Getting Help

For additional help or to report issues:
- Contact CBS IT support
- Check the documentation folder for detailed guides

## For Developers

### Project Structure

- `md_to_docx_converter.py`: Core conversion functionality
- `md_to_docx_gui.py`: Graphical user interface
- `main.py`: Main entry point
- `launcher.py`: Application launcher with service management
- `install.sh`: Installation script
- `package.sh`: macOS packaging script
- `llm_prompt_template.md`: Template for LLM prompt

### Building from Source

To build the application from source:

1. Clone the repository
2. Install dependencies:
   ```
   pip install python-docx markdown beautifulsoup4 Pillow
   ```
3. Run the packaging script:
   ```
   ./package.sh
   ```
4. The packaged application will be available in the `dist` directory

## License

This software is proprietary and for internal use by CBS only.

---

© 2025 CBS. All rights reserved.