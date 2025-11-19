#!/bin/bash
# Install ribbon into LeanMacroTools add-in
# This script properly handles paths with spaces and localized folder names on macOS

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
        echo "✓ Found Add-ins folder: $path"
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
    echo "Please check your Excel installation."
    exit 1
fi

# Try to find the xlam file (check multiple versions)
XLAM_FILE=""
for version in "v1.0.3" "v1.0.2" "v1.0.1" ""; do
    if [ -z "$version" ]; then
        TEST_FILE="$ADDINS_PATH/LeanMacroTools.xlam"
    else
        TEST_FILE="$ADDINS_PATH/LeanMacroTools_$version.xlam"
    fi

    if [ -f "$TEST_FILE" ]; then
        XLAM_FILE="$TEST_FILE"
        echo "✓ Found add-in file: $(basename "$TEST_FILE")"
        break
    fi
done

# Check if xlam was found
if [ -z "$XLAM_FILE" ]; then
    echo ""
    echo "❌ Error: LeanMacroTools .xlam file not found in:"
    echo "   $ADDINS_PATH"
    echo ""
    echo "Looking for one of:"
    echo "   - LeanMacroTools_v1.0.3.xlam"
    echo "   - LeanMacroTools_v1.0.2.xlam"
    echo "   - LeanMacroTools.xlam"
    echo ""
    echo "Files currently in Add-ins folder:"
    ls -1 "$ADDINS_PATH" | grep -i "leanmacro" || echo "   (none found)"
    echo ""
    echo "Please create the add-in first (see README.md Part 1: Installation)"
    exit 1
fi

# Run inject script with properly quoted paths
python3 inject_ribbon.py \
  "$XLAM_FILE" \
  customUI14.xml \
  _rels_dot_rels_for_customUI.xml

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ Success! Please restart Excel to see the Lean Macros tab."
else
    echo ""
    echo "❌ Installation failed. Check error messages above."
    exit 1
fi
