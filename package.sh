#!/bin/bash
# Packaging script for creating a standalone macOS application

# Display banner
echo "============================================="
echo "   CBS Markdown Converter - macOS Packager   "
echo "============================================="
echo

# Define colors for messages
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Check for required tools
echo -n "Checking for required tools... "

# Check for py2app
if ! pip list | grep -q "py2app"; then
    echo -e "${RED}py2app not found${NC}"
    echo "Installing py2app..."
    pip install py2app
    if [ $? -ne 0 ]; then
        echo -e "${RED}Failed to install py2app. Aborting.${NC}"
        exit 1
    fi
else
    echo -e "${GREEN}py2app found${NC}"
fi

# Get script directory
SCRIPT_DIR=$(cd "$(dirname "${BASH_SOURCE[0]}")" &> /dev/null && pwd)
cd "$SCRIPT_DIR"

# Create setup.py for py2app
echo "Creating setup.py for py2app..."
cat > setup.py << EOL
"""
Setup script for creating a standalone macOS application
"""
from setuptools import setup

APP = ['launcher.py']
DATA_FILES = [
    'md_to_docx_converter.py',
    'md_to_docx_gui.py',
    'README.txt'
]

OPTIONS = {
    'argv_emulation': True,
    'packages': ['tkinter', 'PIL', 'markdown', 'bs4', 'docx'],
    'iconfile': 'icon.icns',
    'plist': {
        'CFBundleName': 'CBS Markdown Converter',
        'CFBundleDisplayName': 'CBS Markdown Converter',
        'CFBundleIdentifier': 'au.com.cbs.markdownconverter',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHumanReadableCopyright': 'Copyright © 2025 CBS. All rights reserved.',
        'NSHighResolutionCapable': True,
    },
}

setup(
    name='CBS Markdown Converter',
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
EOL

# Create a basic icon using ImageMagick if available
echo -n "Creating application icon... "
if command -v convert >/dev/null 2>&1; then
    # Create a simple colored square icon with text
    convert -size 1024x1024 xc:navy -fill white -gravity center \
        -pointsize 120 -annotate 0 "CBS\nMD→DOCX" icon.png
    
    # Convert to icns format
    mkdir -p icon.iconset
    convert icon.png -resize 16x16 icon.iconset/icon_16x16.png
    convert icon.png -resize 32x32 icon.iconset/icon_16x16@2x.png
    convert icon.png -resize 32x32 icon.iconset/icon_32x32.png
    convert icon.png -resize 64x64 icon.iconset/icon_32x32@2x.png
    convert icon.png -resize 128x128 icon.iconset/icon_128x128.png
    convert icon.png -resize 256x256 icon.iconset/icon_128x128@2x.png
    convert icon.png -resize 256x256 icon.iconset/icon_256x256.png
    convert icon.png -resize 512x512 icon.iconset/icon_256x256@2x.png
    convert icon.png -resize 512x512 icon.iconset/icon_512x512.png
    convert icon.png -resize 1024x1024 icon.iconset/icon_512x512@2x.png
    
    # Use iconutil to create .icns file
    iconutil -c icns icon.iconset
    
    # Clean up
    rm -rf icon.iconset icon.png
    
    echo -e "${GREEN}Done${NC}"
else
    echo -e "${YELLOW}ImageMagick not found, skipping custom icon${NC}"
    
    # Create a dummy icon file
    touch icon.icns
fi

# Run py2app to build the application
echo "Building application with py2app..."
python setup.py py2app
if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to build application. Aborting.${NC}"
    exit 1
fi

# Create disk image (DMG) if hdiutil is available
echo -n "Creating disk image... "
if command -v hdiutil >/dev/null 2>&1; then
    # Create a temporary directory for DMG contents
    DMG_DIR=$(mktemp -d)
    APP_PATH="$SCRIPT_DIR/dist/CBS Markdown Converter.app"
    
    # Copy the app to the temporary directory
    cp -R "$APP_PATH" "$DMG_DIR/"
    
    # Add a symlink to Applications folder
    ln -s /Applications "$DMG_DIR/Applications"
    
    # Create the DMG
    hdiutil create -volname "CBS Markdown Converter" \
                  -srcfolder "$DMG_DIR" \
                  -ov -format UDZO \
                  "dist/CBS Markdown Converter.dmg"
    
    # Clean up
    rm -rf "$DMG_DIR"
    
    echo -e "${GREEN}Done${NC}"
    echo "Disk image created: dist/CBS Markdown Converter.dmg"
else
    echo -e "${YELLOW}hdiutil not found, skipping DMG creation${NC}"
fi

# Copy application to desktop
echo -n "Creating desktop shortcut... "
APP_PATH="$SCRIPT_DIR/dist/CBS Markdown Converter.app"
DESKTOP_PATH="$HOME/Desktop/CBS Markdown Converter.app"

if [ -d "$APP_PATH" ]; then
    # Create a symlink on the desktop
    ln -sf "$APP_PATH" "$DESKTOP_PATH"
    echo -e "${GREEN}Done${NC}"
else
    echo -e "${RED}Application not found. Shortcut not created.${NC}"
fi

# Final instructions
echo
echo -e "${GREEN}Package build completed!${NC}"
echo
echo "The application has been packaged and is available at:"
echo "  $APP_PATH"
echo
echo "A shortcut has been placed on your desktop."
echo
echo "To distribute this application:"
if command -v hdiutil >/dev/null 2>&1; then
    echo "1. Use the disk image (DMG) file: dist/CBS Markdown Converter.dmg"
else
    echo "1. Zip the application: zip -r \"dist/CBS Markdown Converter.zip\" \"$APP_PATH\""
fi
echo "2. Send the package to users"
echo "3. Users can drag the app to their Applications folder"
echo

# Exit with success
exit 0