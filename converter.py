#!/usr/bin/env python3
"""
MD to DOCX Converter GUI Application
Provides a GUI for converting Markdown files to Word documents using a specified template
"""

import os
import sys
import shutil
import subprocess
import threading
import time
from pathlib import Path
import base64
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QFileDialog, QRadioButton, QButtonGroup,
    QLineEdit, QProgressBar, QMessageBox, QGroupBox, QScrollArea,
    QTextEdit, QSplashScreen
)
from PyQt5.QtGui import QIcon, QPixmap, QFont, QFontMetrics
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# Import the converter module
from md2docx import convert_file, convert_folder

# Base64 encoded CBS logo (placeholder)
# This will be replaced with the actual logo later
CBS_LOGO_BASE64 = """
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
"""

class ConversionThread(QThread):
    """Thread for running file conversions without blocking the UI"""
    update_progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, input_path, template_path, output_path, is_folder):
        super().__init__()
        self.input_path = input_path
        self.template_path = template_path
        self.output_path = output_path
        self.is_folder = is_folder
    
    def run(self):
        try:
            if self.is_folder:
                # For folder conversion, make sure output_path is a directory
                if self.output_path and not os.path.isdir(self.output_path):
                    os.makedirs(self.output_path, exist_ok=True)
                
                results = convert_folder(self.input_path, self.template_path, self.output_path)
                success_count = sum(1 for success, _ in results if success)
                total_count = len(results)
                
                for i, (success, path) in enumerate(results):
                    progress = int((i + 1) / total_count * 100)
                    status = f"Converted {i+1} of {total_count}: {os.path.basename(path)}"
                    self.update_progress.emit(progress, status)
                
                all_success = all(success for success, _ in results)
                if all_success:
                    message = f"Successfully converted {success_count} files to {self.output_path}"
                else:
                    message = f"Converted {success_count} of {total_count} files with some errors"
                
                self.finished.emit(all_success, message)
            else:
                # For single file conversion, ensure we have a file path, not just a directory
                output_path = self.output_path
                if output_path and os.path.isdir(output_path):
                    # If output_path is a directory, create a file path based on input filename
                    file_name = os.path.splitext(os.path.basename(self.input_path))[0]
                    output_path = os.path.join(output_path, f"{file_name}.docx")
                
                self.update_progress.emit(50, f"Converting {os.path.basename(self.input_path)}...")
                success, actual_output_path = convert_file(self.input_path, self.template_path, output_path)
                
                if success:
                    message = f"Successfully converted to {actual_output_path}"
                else:
                    message = f"Error converting {os.path.basename(self.input_path)}"
                
                # Just update the progress without a message, or with a simple status
                self.update_progress.emit(100, "Conversion complete")
                
                # Let finished signal handle the success/error message
                self.finished.emit(success, message)
        
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.update_progress.emit(0, f"Error: {str(e)}")
            self.log_area.append(f"Detailed error: {error_details}")
            self.finished.emit(False, f"Error: {str(e)}")


class MDConverterApp(QMainWindow):
    """Main application window for MD to DOCX Converter"""
    def __init__(self):
        super().__init__()
        
        # Window setup
        self.setWindowTitle("MD to DOCX Converter")
        self.setMinimumSize(600, 500)
        
        # Application state
        self.input_path = ""
        self.template_path = ""
        self.output_path = ""
        self.is_folder = False
        
        # Create the UI
        self.setup_ui()
        
        # Set default template path - look for templates folder
        possible_template_locations = [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates"),
            os.path.expanduser("~/Documents/MD2DOCXConverter/templates")
        ]
        
        for location in possible_template_locations:
            if os.path.exists(location):
                template_files = [f for f in os.listdir(location) if f.endswith('.docx')]
                if template_files:
                    default_template = os.path.join(location, template_files[0])
                    self.template_input.setText(default_template)
                    self.template_path = default_template
                    break
    
    def setup_ui(self):
        """Set up the user interface"""
        # Main widget and layout
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # Logo section
        logo_section = QHBoxLayout()
        logo_section.setAlignment(Qt.AlignCenter)

        # Decode Base64 logo
        try:
            logo_data = base64.b64decode(CBS_LOGO_BASE64)
            logo_pixmap = QPixmap()
            if logo_pixmap.loadFromData(logo_data):
                logo_label = QLabel()
                logo_label.setPixmap(logo_pixmap.scaled(200, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                logo_section.addWidget(logo_label)
            else:
                raise ValueError("Failed to load pixmap from Base64 data")
        except Exception as e:
            print(f"Error loading logo: {str(e)}")
            # Fallback to text-based logo
            logo_label = QLabel("MD to DOCX Converter")
            logo_font = QFont()
            logo_font.setPointSize(18)
            logo_font.setBold(True)
            logo_label.setFont(logo_font)
            logo_label.setStyleSheet("color: #2c3e50;")
            logo_section.addWidget(logo_label)

        main_layout.addLayout(logo_section)
        
        # Title and description
        title_label = QLabel("Markdown to Word Document Converter")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        
        desc_label = QLabel("Convert your Markdown files to formatted Word documents using templates")
        desc_label.setAlignment(Qt.AlignCenter)
        desc_font = QFont()
        desc_font.setPointSize(10)
        desc_label.setFont(desc_font)
        
        main_layout.addWidget(title_label)
        main_layout.addWidget(desc_label)
        
        # Input Type Selection
        input_type_group = QGroupBox("Select Input Type")
        input_type_layout = QHBoxLayout()
        
        self.file_radio = QRadioButton("Single File")
        self.file_radio.setChecked(True)
        self.folder_radio = QRadioButton("Folder (All .md files)")
        
        self.file_radio.toggled.connect(self.update_input_type)
        
        input_type_layout.addWidget(self.file_radio)
        input_type_layout.addWidget(self.folder_radio)
        input_type_group.setLayout(input_type_layout)
        
        main_layout.addWidget(input_type_group)
        
        # Input file/folder selection
        input_layout = QHBoxLayout()
        input_label = QLabel("Input:")
        input_label.setMinimumWidth(60)
        self.input_input = QLineEdit()
        self.input_input.setReadOnly(True)
        browse_input_btn = QPushButton("Browse...")
        browse_input_btn.clicked.connect(self.browse_input)
        
        input_layout.addWidget(input_label)
        input_layout.addWidget(self.input_input)
        input_layout.addWidget(browse_input_btn)
        
        main_layout.addLayout(input_layout)
        
        # Template selection
        template_layout = QHBoxLayout()
        template_label = QLabel("Template:")
        template_label.setMinimumWidth(60)
        self.template_input = QLineEdit()
        self.template_input.setReadOnly(True)
        browse_template_btn = QPushButton("Browse...")
        browse_template_btn.clicked.connect(self.browse_template)
        
        template_layout.addWidget(template_label)
        template_layout.addWidget(self.template_input)
        template_layout.addWidget(browse_template_btn)
        
        main_layout.addLayout(template_layout)
        
        # Output location selection
        output_layout = QHBoxLayout()
        output_label = QLabel("Output:")
        output_label.setMinimumWidth(60)
        self.output_input = QLineEdit()
        self.output_input.setReadOnly(True)
        self.output_input.setPlaceholderText("Same location as input (default)")
        browse_output_btn = QPushButton("Browse...")
        browse_output_btn.clicked.connect(self.browse_output)
        
        output_layout.addWidget(output_label)
        output_layout.addWidget(self.output_input)
        output_layout.addWidget(browse_output_btn)
        
        main_layout.addLayout(output_layout)
        
        # Progress area
        progress_group = QGroupBox("Conversion Progress")
        progress_layout = QVBoxLayout()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        
        self.status_text = QLabel("Ready")
        
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(100)
        
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_text)
        progress_layout.addWidget(self.log_area)
        
        progress_group.setLayout(progress_layout)
        main_layout.addWidget(progress_group)
        
        # Action buttons
        button_layout = QHBoxLayout()
        
        self.convert_btn = QPushButton("Convert")
        self.convert_btn.setMinimumHeight(40)
        self.convert_btn.clicked.connect(self.start_conversion)
        
        close_btn = QPushButton("Close")
        close_btn.setMinimumHeight(40)
        close_btn.clicked.connect(self.close_application)
        
        button_layout.addWidget(self.convert_btn)
        button_layout.addWidget(close_btn)
        
        main_layout.addLayout(button_layout)
        
        # Set the main widget
        self.setCentralWidget(main_widget)
    
    def update_input_type(self):
        """Update the input type based on radio button selection"""
        self.is_folder = self.folder_radio.isChecked()
        # Clear the input when changing type
        self.input_path = ""
        self.input_input.setText("")
        self.output_path = ""
        self.output_input.setText("")
    
    def browse_input(self):
        """Open file dialog to select input file or folder"""
        if self.is_folder:
            folder_path = QFileDialog.getExistingDirectory(
                self, "Select Folder Containing Markdown Files"
            )
            if folder_path:
                self.input_path = folder_path
                self.input_input.setText(folder_path)
                # Set default output to same location
                self.output_path = folder_path
        else:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Select Markdown File", "", "Markdown Files (*.md);;All Files (*)"
            )
            if file_path:
                self.input_path = file_path
                self.input_input.setText(file_path)
                # Set default output to same folder
                self.output_path = os.path.dirname(file_path)
    
    def browse_template(self):
        """Open file dialog to select template file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Word Template", "", "Word Documents (*.docx);;All Files (*)"
        )
        if file_path:
            self.template_path = file_path
            self.template_input.setText(file_path)

    def browse_output(self):
        """Open file dialog to select output location"""
        initial_dir = ""
        if self.input_path:
            if os.path.isdir(self.input_path):
                initial_dir = self.input_path
            else:
                initial_dir = os.path.dirname(self.input_path)
        
        # For both single file and folder mode, let the user select a directory
        folder_path = QFileDialog.getExistingDirectory(
            self, "Select Output Folder", initial_dir
        )
        if folder_path:
            self.output_path = folder_path
            self.output_input.setText(folder_path)

    def start_conversion(self):
        """Start the conversion process"""
        # Validate inputs
        if not self.input_path:
            QMessageBox.warning(self, "Input Required", "Please select an input file or folder.")
            return
        
        if not self.template_path:
            QMessageBox.warning(self, "Template Required", "Please select a Word template.")
            return
        
        if not os.path.exists(self.template_path):
            QMessageBox.warning(self, "Template Not Found", f"The template file '{self.template_path}' does not exist.")
            return
            
        # If input is a folder, make sure it contains at least one MD file
        if self.is_folder:
            md_files = [f for f in os.listdir(self.input_path) if f.lower().endswith('.md')]
            if not md_files:
                QMessageBox.warning(self, "No Markdown Files", "The selected folder doesn't contain any .md files.")
                return
        else:
            # For single file, check if it exists
            if not os.path.exists(self.input_path):
                QMessageBox.warning(self, "File Not Found", f"The input file '{self.input_path}' does not exist.")
                return
        
        # Disable UI during conversion
        self.set_ui_enabled(False)
        
        # Reset progress
        self.progress_bar.setValue(0)
        self.status_text.setText("Starting conversion...")
        self.log_area.clear()
        
        # If output path is empty, use input folder or generate appropriate output path
        output_path = self.output_path
        if not output_path:
            if self.is_folder:
                output_path = self.input_path
            else:
                # For single file, create output path in same directory with .docx extension
                file_name = os.path.splitext(os.path.basename(self.input_path))[0]
                output_path = os.path.join(os.path.dirname(self.input_path), f"{file_name}.docx")
        
        # Start conversion thread
        self.conversion_thread = ConversionThread(
            self.input_path,
            self.template_path,
            output_path,
            self.is_folder
        )
        
        self.conversion_thread.update_progress.connect(self.update_progress)
        self.conversion_thread.finished.connect(self.conversion_finished)
        self.conversion_thread.start()

    
    def update_progress(self, value, message):
        """Update progress bar and status message"""
        self.progress_bar.setValue(value)
        self.status_text.setText(message)
        self.log_area.append(message)
    
    def conversion_finished(self, success, message):
        """Handle conversion completion"""
        # Re-enable UI
        self.set_ui_enabled(True)
        
        # Update status
        self.status_text.setText(message)
        self.log_area.append(message)
        
        # Show completion message
        if success:
            QMessageBox.information(self, "Conversion Complete", message)
        else:
            QMessageBox.warning(self, "Conversion Error", message)
    
    def set_ui_enabled(self, enabled):
        """Enable or disable UI elements during conversion"""
        self.file_radio.setEnabled(enabled)
        self.folder_radio.setEnabled(enabled)
        self.convert_btn.setEnabled(enabled)
        
        # Update cursor if needed
        if enabled:
            QApplication.restoreOverrideCursor()
        else:
            QApplication.setOverrideCursor(Qt.WaitCursor)
    
    def close_application(self):
        """Close the application and ensure cleanup"""
        # Make sure any running conversions are stopped
        if hasattr(self, 'conversion_thread') and self.conversion_thread.isRunning():
            self.conversion_thread.terminate()
            self.conversion_thread.wait()
        
        # Close the application
        self.close()
        QApplication.quit()


def check_required_services():
    """
    Check if required services are running and restart them if needed
    This is just a placeholder - actual implementation would depend on 
    specific services required
    """
    # For example, check if certain Python processes need to be available
    return True

def main():
    """Application entry point"""
    # Check and restart services if needed
    check_required_services()
    
    # Create and run application
    app = QApplication(sys.argv)
    
    # Create splash screen
    splash_pixmap = QPixmap(1, 1)  # Create minimal pixmap for the splash screen
    splash = QSplashScreen(splash_pixmap)
    splash.show()
    
    # Initialize application
    window = MDConverterApp()
    
    # Close splash and show main window
    splash.finish(window)
    window.show()
    
    # Run application
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()