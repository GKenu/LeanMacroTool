#!/bin/bash
# LeanMacroTools Installer for macOS
# Double-click this file to install the add-in
# Version 1.0.7

echo ""
echo "═══════════════════════════════════════════════════"
echo "  LeanMacroTools v1.0.7 Installer"
echo "═══════════════════════════════════════════════════"
echo ""

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Try multiple path variations (for localized macOS systems)
ADDINS_PATHS=(
    "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized"
    "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins"
    "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins"
)

# Find which Add-ins path exists
ADDINS_PATH=""
for path in "${ADDINS_PATHS[@]}"; do
    if [ -d "$path" ]; then
        ADDINS_PATH="$path"
        echo "✓ Found Excel Add-ins folder"
        break
    fi
done

# Check if Add-ins folder was found
if [ -z "$ADDINS_PATH" ]; then
    echo "❌ Error: Excel Add-ins folder not found!"
    echo ""
    echo "Tried these locations:"
    for path in "${ADDINS_PATHS[@]}"; do
        echo "   - $path"
    done
    echo ""
    echo "Please ensure Microsoft Excel is installed."
    echo ""
    read -p "Press Enter to close..."
    exit 1
fi

# Check if the xlam file exists in the script directory
XLAM_SOURCE="$SCRIPT_DIR/LeanMacroTools_v1.0.7.xlam"
if [ ! -f "$XLAM_SOURCE" ]; then
    echo "❌ Error: LeanMacroTools_v1.0.7.xlam not found!"
    echo ""
    echo "Expected location: $XLAM_SOURCE"
    echo ""
    echo "Please ensure the .xlam file is in the same folder as this installer."
    echo ""
    read -p "Press Enter to close..."
    exit 1
fi

# Copy the xlam file to Add-ins folder
echo "Installing LeanMacroTools..."
cp "$XLAM_SOURCE" "$ADDINS_PATH/"

if [ $? -eq 0 ]; then
    echo "✓ Installation successful!"
    echo ""
    echo "═══════════════════════════════════════════════════"
    echo "  NEXT STEPS:"
    echo "═══════════════════════════════════════════════════"
    echo ""
    echo "1. Open Microsoft Excel"
    echo "2. Go to Tools > Excel Add-ins..."
    echo "3. Check ☑ LeanMacroTools_v1.0.7"
    echo "4. Click OK"
    echo ""
    echo "You should see a 'Lean Macros' tab in the ribbon!"
    echo ""
    echo "Keyboard Shortcuts:"
    echo "  • Ctrl+Shift+N - Cycle number formats"
    echo "  • Ctrl+Shift+V - Cycle font colors"
    echo "  • Ctrl+Shift+B - Cycle fill patterns"
    echo "  • Ctrl+Shift+T - Trace precedents"
    echo "  • Ctrl+Shift+Y - Trace dependents"
    echo ""
else
    echo "❌ Installation failed!"
    echo ""
    echo "Please check file permissions and try again."
    echo ""
fi

read -p "Press Enter to close..."
