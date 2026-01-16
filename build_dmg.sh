#!/bin/bash

# =============================================================================
# UCSC Result Fetcher - macOS DMG Builder Script
# =============================================================================
# This script builds a macOS .app bundle and packages it into a .dmg file
# =============================================================================

set -e  # Exit on error

# Configuration
APP_NAME="UCSC Result Fetcher"
DMG_NAME="UCSC_Result_Fetcher"
VERSION="2.0.0"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "=============================================="
echo "  Building ${APP_NAME} for macOS"
echo "=============================================="
echo ""

# Step 1: Check for Python and pip
echo "[1/6] Checking Python installation..."
if ! command -v python3 &> /dev/null; then
    echo "❌ Error: Python 3 is not installed. Please install Python 3.8 or higher."
    exit 1
fi
PYTHON_VERSION=$(python3 --version)
echo "✓ Found: $PYTHON_VERSION"

# Step 2: Install/upgrade PyInstaller
echo ""
echo "[2/6] Installing/upgrading PyInstaller..."
pip3 install --upgrade pyinstaller
echo "✓ PyInstaller is ready"

# Step 3: Install application dependencies
echo ""
echo "[3/6] Installing application dependencies..."
pip3 install pandas requests beautifulsoup4 openpyxl
echo "✓ Dependencies installed"

# Step 4: Clean previous builds
echo ""
echo "[4/6] Cleaning previous builds..."
cd "$SCRIPT_DIR"
rm -rf build/ dist/ *.dmg
echo "✓ Cleaned previous builds"

# Step 5: Build the .app bundle with PyInstaller
echo ""
echo "[5/6] Building the .app bundle with PyInstaller..."
python3 -m PyInstaller UCSC_Result_Fetcher.spec --noconfirm

if [ ! -d "dist/${APP_NAME}.app" ]; then
    echo "❌ Error: Failed to create the .app bundle"
    exit 1
fi
echo "✓ Created ${APP_NAME}.app"

# Step 6: Create the DMG file
echo ""
echo "[6/6] Creating DMG file..."

DMG_TEMP="temp_dmg"
DMG_FINAL="${DMG_NAME}_${VERSION}.dmg"

# Create a temporary directory for DMG contents
rm -rf "$DMG_TEMP"
mkdir -p "$DMG_TEMP"

# Copy the app to temporary directory
cp -R "dist/${APP_NAME}.app" "$DMG_TEMP/"

# Copy credits.csv alongside the app (for reference)
if [ -f "credits.csv" ]; then
    cp "credits.csv" "$DMG_TEMP/credits.csv"
fi

# Create a symbolic link to Applications folder for easy installation
ln -s /Applications "$DMG_TEMP/Applications"

# Create a README file for the DMG
cat > "$DMG_TEMP/README.txt" << 'EOF'
===========================================
UCSC Result Fetcher - Installation Guide
===========================================

To install:
1. Drag "UCSC Result Fetcher.app" to the Applications folder
2. Double-click to run from Applications

Note: On first run, you may need to right-click the app 
and select "Open" to bypass Gatekeeper security.

If you see "App is damaged" error:
Open Terminal and run:
  xattr -cr /Applications/UCSC\ Result\ Fetcher.app

For batch processing, place your CSV files in the same 
folder or use the Browse button to select them.

===========================================
EOF

# Create the DMG using hdiutil
echo "Creating DMG image..."
hdiutil create -volname "${APP_NAME}" \
    -srcfolder "$DMG_TEMP" \
    -ov -format UDZO \
    "$DMG_FINAL"

# Clean up
rm -rf "$DMG_TEMP"

# Verify the DMG was created
if [ -f "$DMG_FINAL" ]; then
    DMG_SIZE=$(du -h "$DMG_FINAL" | cut -f1)
    echo ""
    echo "=============================================="
    echo "  ✓ BUILD SUCCESSFUL!"
    echo "=============================================="
    echo "  DMG File: ${DMG_FINAL}"
    echo "  Size: ${DMG_SIZE}"
    echo "  Location: ${SCRIPT_DIR}/${DMG_FINAL}"
    echo "=============================================="
    echo ""
    echo "To distribute this app, share the DMG file."
    echo "Users can open it and drag the app to Applications."
else
    echo "❌ Error: Failed to create DMG file"
    exit 1
fi
