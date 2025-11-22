#!/bin/bash
# Build LeanMacroTools release package
# This script creates a distribution-ready package for users
# Version 2.1.0

set -e  # Exit on error

VERSION="v2.1.0"
DIST_DIR="dist"
PACKAGE_NAME="LeanMacroTools_${VERSION}"
XLAM_FILE="LeanMacroTools_${VERSION}.xlam"

echo ""
echo "═══════════════════════════════════════════════════"
echo "  Building LeanMacroTools ${VERSION}"
echo "═══════════════════════════════════════════════════"
echo ""

# Get script directory and project root
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"

# Find Add-ins folder
ADDINS_PATHS=(
    "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized"
    "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins"
    "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins"
)

ADDINS_PATH=""
for path in "${ADDINS_PATHS[@]}"; do
    if [ -d "$path" ]; then
        ADDINS_PATH="$path"
        break
    fi
done

if [ -z "$ADDINS_PATH" ]; then
    echo "❌ Error: Excel Add-ins folder not found!"
    exit 1
fi

XLAM_SOURCE="$ADDINS_PATH/$XLAM_FILE"
TEMPLATE_PATH="$PROJECT_ROOT/templates/LeanMacroTools_template.xlam"

# Step 1: Check if template exists
if [ ! -f "$TEMPLATE_PATH" ]; then
    echo "❌ Error: Template file not found!"
    echo ""
    echo "Expected location: $TEMPLATE_PATH"
    echo ""
    echo "Please create the template first:"
    echo "See templates/README.md for instructions"
    exit 1
fi

echo "✓ Found template: $TEMPLATE_PATH"

# Step 2: Copy template to Add-ins folder
echo ""
echo "Step 1: Copying template to Add-ins folder..."
cp "$TEMPLATE_PATH" "$XLAM_SOURCE"

if [ $? -ne 0 ]; then
    echo "❌ Failed to copy template!"
    exit 1
fi

echo "✓ Template copied to: $XLAM_SOURCE"
echo ""

# Step 2: Inject ribbon into xlam file
echo "Step 2: Injecting ribbon UI..."
python3 "$SCRIPT_DIR/inject_ribbon.py" \
    "$XLAM_SOURCE" \
    "$PROJECT_ROOT/ribbon/customUI14.xml" \
    "$PROJECT_ROOT/ribbon/_rels_dot_rels_for_customUI.xml"

if [ $? -ne 0 ]; then
    echo "❌ Ribbon injection failed!"
    exit 1
fi

echo "✓ Ribbon injected successfully"
echo ""

# Step 3: Create distribution package
echo "Step 3: Creating distribution package..."
mkdir -p "$PROJECT_ROOT/$DIST_DIR/$PACKAGE_NAME"

# Copy files to dist package
cp "$XLAM_SOURCE" "$PROJECT_ROOT/$DIST_DIR/$PACKAGE_NAME/"
cp "$PROJECT_ROOT/install.command" "$PROJECT_ROOT/$DIST_DIR/$PACKAGE_NAME/"
cp "$PROJECT_ROOT/README.md" "$PROJECT_ROOT/$DIST_DIR/$PACKAGE_NAME/"
cp "$PROJECT_ROOT/LICENSE" "$PROJECT_ROOT/$DIST_DIR/$PACKAGE_NAME/"
cp "$PROJECT_ROOT/CHANGELOG.md" "$PROJECT_ROOT/$DIST_DIR/$PACKAGE_NAME/"

# Make install.command executable in the package
chmod +x "$PROJECT_ROOT/$DIST_DIR/$PACKAGE_NAME/install.command"

echo "✓ Package created in $DIST_DIR/$PACKAGE_NAME/"

# Create zip archive
cd "$PROJECT_ROOT/$DIST_DIR"
zip -r "${PACKAGE_NAME}.zip" "$PACKAGE_NAME"
cd "$PROJECT_ROOT"

echo "✓ Archive created: $DIST_DIR/${PACKAGE_NAME}.zip"
echo ""
echo "═══════════════════════════════════════════════════"
echo "  Build Complete!"
echo "═══════════════════════════════════════════════════"
echo ""
echo "Distribution package: $DIST_DIR/${PACKAGE_NAME}.zip"
echo ""
echo "Package contents:"
echo "  • $XLAM_FILE (with ribbon UI embedded)"
echo "  • install.command (double-click installer)"
echo "  • README.md"
echo "  • LICENSE"
echo "  • CHANGELOG.md"
echo ""
echo "Users can now:"
echo "  1. Download and unzip ${PACKAGE_NAME}.zip"
echo "  2. Double-click install.command"
echo "  3. Enable add-in in Excel"
echo ""
