#!/usr/bin/env python3
import os
import sys
import subprocess
import shutil
from pathlib import Path

def check_install_pip():
    """Check if pip is installed and install if not."""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "--version"], 
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("✅ pip is already installed")
        return True
    except:
        print("❌ pip is not installed. Attempting to install...")
        try:
            subprocess.check_call([sys.executable, "-m", "ensurepip", "--upgrade"],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print("✅ pip has been installed")
            return True
        except:
            print("❌ Failed to install pip. Please install pip manually.")
            return False

def install_requirements():
    """Install required packages."""
    requirements = [
        'python-docx',
        'markdown',
        'beautifulsoup4',
        'PyQt5'
    ]
    
    print("\n📦 Checking and installing required packages...")
    
    for package in requirements:
        try:
            __import__(package.replace('-', '_').split('>=')[0])
            print(f"✅ {package} is already installed")
        except ImportError:
            print(f"📥 Installing {package}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package], 
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print(f"✅ {package} has been installed")
            except:
                print(f"❌ Failed to install {package}")
                return False
    
    return True

def create_converter_script(app_dir):
    """Create the converter.py script in the application directory."""
    # Source and destination paths
    current_dir = os.path.dirname(os.path.abspath(__file__))
    source_converter_path = os.path.join(current_dir, "converter.py")
    dest_converter_path = os.path.join(app_dir, "converter.py")
    
    # Check if source file exists
    if not os.path.exists(source_converter_path):
        print(f"❌ Error: converter.py not found at {source_converter_path}")
        return None
    
    # Copy the file
    shutil.copy2(source_converter_path, dest_converter_path)
    
    # Make executable
    os.chmod(dest_converter_path, 0o755)
    print(f"✅ Copied converter script to {dest_converter_path}")
    
    return dest_converter_path

def create_templates_folder(app_dir):
    """Create templates folder and add a default template."""
    templates_dir = os.path.join(app_dir, "templates")
    os.makedirs(templates_dir, exist_ok=True)
    print(f"✅ Created templates directory at {templates_dir}")
    
    # Check if there's a default template to copy
    current_dir = os.path.dirname(os.path.abspath(__file__))
    default_template = os.path.join(current_dir, "templates", "default.docx")
    
    if os.path.exists(default_template):
        template_dest = os.path.join(templates_dir, "default.docx")
        shutil.copy2(default_template, template_dest)
        print(f"✅ Copied default template to {template_dest}")
    else:
        # Create a simple default template using python-docx
        try:
            from docx import Document
            doc = Document()
            doc.add_heading('Default Template', 0)
            
            # Save the template
            template_dest = os.path.join(templates_dir, "default.docx")
            doc.save(template_dest)
            print(f"✅ Created default template at {template_dest}")
        except Exception as e:
            print(f"⚠️ Could not create default template: {str(e)}")
            print("ℹ️ You'll need to provide a Word template file (.docx) when using the converter.")

def create_app_directory():
    """Create application directory structure."""
    # Create app directory in the current folder
    current_dir = os.path.dirname(os.path.abspath(__file__))
    app_dir = os.path.join(current_dir, "MD2DOCXConverter")
    
    if not os.path.exists(app_dir):
        os.makedirs(app_dir)
        print(f"✅ Created application directory at {app_dir}")
    else:
        print(f"✅ Application directory already exists at {app_dir}")
    
    # Ensure md2docx.py is copied to the app directory
    module_path = os.path.join(current_dir, "md2docx.py")
    if os.path.exists(module_path):
        module_dest = os.path.join(app_dir, "md2docx.py")
        shutil.copy2(module_path, module_dest)
        print(f"✅ Copied converter module to {module_dest}")
    else:
        print(f"❌ Error: md2docx.py not found at {module_path}")
        return None
    
    # Create converter.py script
    create_converter_script(app_dir)
    
    # Create templates folder
    create_templates_folder(app_dir)
    
    return app_dir

def create_desktop_shortcut(app_dir):
    """Create a desktop shortcut for macOS that doesn't show terminal window."""
    desktop_path = os.path.expanduser("~/Desktop")
    app_launcher = os.path.join(app_dir, "converter.py")
    
    # Create an AppleScript application instead of a .command file
    app_name = "MD to DOCX Converter"
    app_folder = os.path.join(desktop_path, f"{app_name}.app")
    contents_folder = os.path.join(app_folder, "Contents")
    macos_folder = os.path.join(contents_folder, "MacOS")
    resources_folder = os.path.join(contents_folder, "Resources")
    
    # Create the directory structure
    os.makedirs(macos_folder, exist_ok=True)
    os.makedirs(resources_folder, exist_ok=True)
    
    # Create Info.plist file
    info_plist = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>launcher</string>
    <key>CFBundleIdentifier</key>
    <string>com.md2docx.converter</string>
    <key>CFBundleName</key>
    <string>{app_name}</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>CFBundleInfoDictionaryVersion</key>
    <string>6.0</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.13</string>
</dict>
</plist>'''
    
    with open(os.path.join(contents_folder, "Info.plist"), 'w') as f:
        f.write(info_plist)
    
    # Create the launcher script
    launcher_script = f'''#!/bin/bash
cd "{app_dir}"
"{sys.executable}" "{app_launcher}" > /dev/null 2>&1 &
'''
    
    launcher_path = os.path.join(macos_folder, "launcher")
    with open(launcher_path, 'w') as f:
        f.write(launcher_script)
    
    # Make the launcher executable
    os.chmod(launcher_path, 0o755)
    
    print(f"✅ Created desktop application at {app_folder}")
    return app_folder

def main():
    print("=" * 50)
    print("🚀 MD to DOCX Converter Installation")
    print("=" * 50)
    
    # Check if running on macOS
    if sys.platform != 'darwin':
        print("❌ This installer is designed for macOS only.")
        return False
    
    # Check and install pip if needed
    if not check_install_pip():
        return False
    
    # Install required packages
    if not install_requirements():
        print("\n❌ Installation failed due to package installation errors.")
        return False
    
    # Create application directory and copy files
    app_dir = create_app_directory()
    if app_dir is None:
        print("\n❌ Installation failed because required files are missing.")
        return False
    
    # Create desktop shortcut
    app_path = create_desktop_shortcut(app_dir)
    
    print("\n✅ Installation completed successfully!")
    print(f"📝 You can now use the 'MD to DOCX Converter' application on your desktop.")
    print(f"📁 The application files are located at: {app_dir}")
    
    return True

if __name__ == "__main__":
    success = main()
    if not success:
        print("\n⚠️ Installation encountered some issues. Please resolve them and try again.")
        sys.exit(1)
    sys.exit(0)