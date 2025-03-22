#!/usr/bin/env python3
"""
CBS Markdown to Word Converter
Main entry point for the application
"""

import os
import sys
import argparse
from md_to_docx_converter import MarkdownToWordConverter
from md_to_docx_gui import ConverterApp
import tkinter as tk

def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description='Convert Markdown files to Word documents with template support'
    )
    
    parser.add_argument(
        '--cli', 
        action='store_true',
        help='Run in command-line mode (no GUI)'
    )
    
    parser.add_argument(
        '--input', 
        type=str,
        help='Input Markdown file or directory'
    )
    
    parser.add_argument(
        '--output', 
        type=str,
        help='Output Word file or directory'
    )
    
    parser.add_argument(
        '--template', 
        type=str,
        help='Word template file (.dotx)'
    )
    
    return parser.parse_args()

def run_cli_mode(args):
    """Run the converter in command-line mode."""
    if not args.input:
        print("Error: --input is required in CLI mode")
        sys.exit(1)
    
    try:
        # Initialize converter
        converter = MarkdownToWordConverter(args.template)
        
        # Check if input is a file or directory
        if os.path.isfile(args.input):
            # Convert single file
            output = converter.convert_file(args.input, args.output)
            print(f"Conversion complete: {output}")
        elif os.path.isdir(args.input):
            # Convert directory
            outputs = converter.convert_directory(args.input, args.output)
            print(f"Converted {len(outputs)} files.")
            for output in outputs:
                print(f"  - {output}")
        else:
            print(f"Error: Input '{args.input}' not found")
            sys.exit(1)
            
    except Exception as e:
        print(f"Error during conversion: {e}")
        sys.exit(1)

def run_gui_mode():
    """Run the converter in GUI mode."""
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()

def main():
    """Main entry point."""
    # Parse command-line arguments
    args = parse_arguments()
    
    # Check mode
    if args.cli:
        run_cli_mode(args)
    else:
        run_gui_mode()

if __name__ == "__main__":
    main()