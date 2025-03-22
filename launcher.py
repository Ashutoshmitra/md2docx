#!/usr/bin/env python3
"""
Launcher for the CBS Markdown to Word Converter.
This script ensures proper environment setup and service management.
"""

import os
import sys
import subprocess
import signal
import time

def get_script_dir():
    """Get the directory of the current script."""
    return os.path.dirname(os.path.abspath(__file__))

def check_python_environment():
    """Check if Python environment is properly set up."""
    script_dir = get_script_dir()
    venv_dir = os.path.join(script_dir, "venv")
    
    if not os.path.exists(venv_dir):
        print("Error: Virtual environment not found.")
        print("Please run the installation script first.")
        sys.exit(1)
    
    # Check if required modules are installed
    try:
        # Try to import required modules
        import tkinter
        from PIL import Image
        import markdown
        from bs4 import BeautifulSoup
        import docx
    except ImportError as e:
        print(f"Error: Required module not found: {e}")
        print("Please run the installation script to install all dependencies.")
        sys.exit(1)
    
    return True

def kill_existing_instances():
    """Kill any existing instances of the converter application."""
    # This is a simplified approach for macOS
    try:
        # Look for python processes running our GUI
        result = subprocess.run(
            ["pgrep", "-f", "python.*md_to_docx_gui.py"],
            capture_output=True,
            text=True
        )
        
        if result.stdout:
            # Get PIDs
            pids = result.stdout.strip().split('\n')
            
            # Kill each process
            for pid in pids:
                if pid.isdigit():
                    try:
                        os.kill(int(pid), signal.SIGTERM)
                        print(f"Terminated existing process with PID {pid}")
                    except OSError:
                        pass
            
            # Give processes time to terminate
            time.sleep(1)
    except Exception as e:
        print(f"Warning: Could not check for existing processes: {e}")

def restart_services():
    """Restart any services required by the application."""
    # For this application, we don't need to restart any system services
    # This function is included for future expansion
    pass

def launch_application():
    """Launch the converter application."""
    script_dir = get_script_dir()
    app_dir = os.path.join(script_dir, "md_to_docx_app")
    
    if not os.path.exists(app_dir):
        print("Error: Application files not found.")
        print("Please run the installation script first.")
        sys.exit(1)
    
    # Change to the application directory
    os.chdir(app_dir)
    
    # Launch the application
    try:
        subprocess.run([sys.executable, "md_to_docx_gui.py"])
    except Exception as e:
        print(f"Error launching application: {e}")
        sys.exit(1)

def main():
    """Main function."""
    print("Starting CBS Markdown to Word Converter...")
    
    # Check Python environment
    check_python_environment()
    
    # Kill existing instances
    kill_existing_instances()
    
    # Restart required services
    restart_services()
    
    # Launch the application
    launch_application()

if __name__ == "__main__":
    main()