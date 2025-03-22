#!/bin/bash
# Installation script for the Markdown to Word Converter
# This script installs the required dependencies and creates a desktop shortcut

# Display welcome message
echo "========================================================"
echo "   Markdown to Word Converter Installation"
echo "   (CBS Application)"
echo "========================================================"
echo

# Define colors for messages
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Check if Python is installed
echo -n "Checking for Python 3.6+ installation... "
if command -v python3 >/dev/null 2>&1; then
    PYTHON_VERSION=$(python3 -c 'import sys; print("{}.{}".format(sys.version_info.major, sys.version_info.minor))')
    PYTHON_MAJOR=$(echo $PYTHON_VERSION | cut -d. -f1)
    PYTHON_MINOR=$(echo $PYTHON_VERSION | cut -d. -f2)
    
    if [ "$PYTHON_MAJOR" -ge 3 ] && [ "$PYTHON_MINOR" -ge 6 ]; then
        echo -e "${GREEN}Found Python $PYTHON_VERSION${NC}"
        PYTHON_CMD="python3"
    else
        echo -e "${RED}Python 3.6+ is required, but found $PYTHON_VERSION${NC}"
        exit 1
    fi
else
    echo -e "${RED}Python 3 not found. Please install Python 3.6 or later.${NC}"
    exit 1
fi

# Create a virtual environment
echo -n "Creating virtual environment... "
$PYTHON_CMD -m venv venv
if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to create virtual environment.${NC}"
    exit 1
fi
echo -e "${GREEN}Done${NC}"

# Activate the virtual environment
echo -n "Activating virtual environment... "
source venv/bin/activate
if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to activate virtual environment.${NC}"
    exit 1
fi
echo -e "${GREEN}Done${NC}"

# Install required packages
echo "Installing required packages..."
pip install --upgrade pip
pip install python-docx markdown beautifulsoup4 Pillow
if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to install required packages.${NC}"
    exit 1
fi
echo -e "${GREEN}All packages installed successfully${NC}"

# Get the script directory
SCRIPT_DIR=$(cd "$(dirname "${BASH_SOURCE[0]}")" &> /dev/null && pwd)

# Copy application files
echo "Setting up application files..."
APP_DIR="$SCRIPT_DIR/md_to_docx_app"
mkdir -p "$APP_DIR"

# Copy Python files to app directory
cp "$SCRIPT_DIR/md_to_docx_converter.py" "$APP_DIR/"
cp "$SCRIPT_DIR/md_to_docx_gui.py" "$APP_DIR/"

# Create run script
RUN_SCRIPT="$APP_DIR/run.sh"
cat > "$RUN_SCRIPT" << EOL
#!/bin/bash
cd "\$(dirname "\${BASH_SOURCE[0]}")"
source ../venv/bin/activate
python md_to_docx_gui.py
EOL
chmod +x "$RUN_SCRIPT"

# Create launcher script (to be used by the desktop shortcut)
LAUNCHER_SCRIPT="$SCRIPT_DIR/launch_converter.sh"
cat > "$LAUNCHER_SCRIPT" << EOL
#!/bin/bash
cd "$SCRIPT_DIR"
"$APP_DIR/run.sh"
EOL
chmod +x "$LAUNCHER_SCRIPT"

# Create desktop shortcut
echo -n "Creating desktop shortcut... "
DESKTOP_DIR="$HOME/Desktop"
SHORTCUT_PATH="$DESKTOP_DIR/CBS Markdown Converter.command"

cat > "$SHORTCUT_PATH" << EOL
#!/bin/bash
"$LAUNCHER_SCRIPT"
EOL
chmod +x "$SHORTCUT_PATH"
echo -e "${GREEN}Done${NC}"

# Create simple README file
README_PATH="$SCRIPT_DIR/README.txt"
cat > "$README_PATH" << EOL
=================================================
CBS Markdown to Word Converter
=================================================

This application converts Markdown (.md) files to Word (.docx) format 
using a specified Word template.

To use the converter:
1. Double-click the "CBS Markdown Converter" shortcut on your desktop
2. Select a Markdown file or folder containing Markdown files
3. Optionally, select a Word template (.dotx) file
4. Choose an output location (or use the same as input)
5. Click "Convert" to start the conversion process

For support, please contact CBS IT support.
EOL

# Final message
echo
echo -e "${GREEN}Installation completed successfully!${NC}"
echo "A shortcut has been created on your desktop."
echo "To start the application, double-click the 'CBS Markdown Converter' icon."
echo
echo -e "${YELLOW}Note: If the shortcut doesn't work, try running the application directly:${NC}"
echo "  1. Open Terminal"
echo "  2. Run: $LAUNCHER_SCRIPT"
echo

# Exit with success
exit 0