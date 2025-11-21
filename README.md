# LEAN MACRO TOOLS FOR EXCEL MAC

**Version 2.0.0** - 5 Powerful Features via Keyboard Shortcuts & Ribbon Tab

I missed TTS Macros for personal use, so I built my own. Not perfect yet, but feel free to use and contribute!

**Attribution:** The precedent/dependent tracer UserForm interface is adapted from [TTS Turbo Macros](https://sites.google.com/site/ttsturbo/). We are grateful for their excellent UX design and apologize for using their UserForm. The tracing logic and integration are our own implementation.

## Features

1. **Cycle Number Formats** (Ctrl+Shift+N)
   - Cycles through: Original → Thousands → Percentage → Multiples → USD → BRL → (wraps back to Original)
   - Returns to cell's original format!
   - Customizable format list via ribbon button

2. **Cycle Font Colors** (Ctrl+Shift+V)
   - Cycles through preset colors: Blue → Green → Red → Grey → Black → Original
   - Changes text color (not background)
   - Remembers and restores original font color
   - Perfect for highlighting important cells!

3. **Cycle Fill Patterns** (Ctrl+Shift+B) **NEW in v1.0.6**
   - Cycles through: Color+Border → Pattern → Original
   - First press: Applies beige background with outline border
   - Second press: Applies dotted pattern fill
   - Third press: Restores original formatting
   - Great for highlighting sections and creating visual separation!

4. **Trace Precedents** (Ctrl+Shift+T) **MAJOR UPGRADE in v2.0.0**
   - **Interactive panel with list-based navigation** - much better UX!
   - Shows cells that feed into formulas
   - **Click or use arrow keys** to navigate through cells instantly
   - **Auto-navigation** - selecting items automatically jumps to cells
   - Real-time display of cell address, value, and formula
   - Panel stays open while you work in Excel
   - Works cross-sheet perfectly

5. **Trace Dependents** (Ctrl+Shift+Y) **MAJOR UPGRADE in v2.0.0**
   - **Interactive panel with list-based navigation** - much better UX!
   - Shows cells that use the current cell
   - **Click or use arrow keys** to navigate through cells instantly
   - **Auto-navigation** - selecting items automatically jumps to cells
   - Real-time display of cell address, value, and formula
   - Panel stays open while you work in Excel
   - Works cross-sheet perfectly

---

## Installation

### Quick Install (2 minutes) ⚡

**Step 1: Download and Extract**
1. Download the latest release: `LeanMacroTools_v2.0.0.zip`
2. Double-click to extract the zip file

**Step 2: Run Installer**
1. Double-click **install.command** in the extracted folder
2. The installer will automatically find your Excel Add-ins folder
3. Follow the on-screen instructions

**Step 3: Enable in Excel**
1. Open Excel
2. Go to **Tools > Excel Add-ins...**
3. Check ☑ **LeanMacroTools_v1.0.7**
4. Click **OK**

**That's it!** You should see a "Lean Macros" tab in the ribbon with all features ready to use.

---

### Advanced: Manual Installation

<details>
<summary>Click to expand manual installation instructions</summary>

#### Part 1: Create the Add-In (5 minutes)

**Step 1: Create New Workbook**
1. Open Excel
2. Create new blank workbook

**Step 2: Import VBA Modules**
1. Press **Option+F11** (VBA Editor)
2. **File > Import File...** (or **Option+Cmd+I**) → Select **modNumberFormats.bas** → Open
3. **File > Import File...** (or **Option+Cmd+I**) → Select **modColorFormats.bas** → Open
4. **File > Import File...** (or **Option+Cmd+I**) → Select **modFillFormats.bas** → Open
5. **File > Import File...** (or **Option+Cmd+I**) → Select **modTraceTools.bas** → Open

You should see all four modules in the left panel.

**Step 3: Add Keyboard Shortcuts**
1. In VBA Editor, double-click **ThisWorkbook** (left panel)
2. Paste this code:

```vba
Private Sub Workbook_Open()
    ' Register keyboard shortcuts for macro functions
    ' Syntax: Application.OnKey "^+[Key]", "[MacroName]"
    ' Where ^ = Ctrl, + = Shift

    ' Number Format Cycling (Ctrl+Shift+N)
    Application.OnKey "^+N", "CycleFormatsKeyboard"

    ' Font Color Cycling (Ctrl+Shift+V)
    Application.OnKey "^+V", "CycleColorsKeyboard"

    ' Fill Pattern Cycling (Ctrl+Shift+B)
    Application.OnKey "^+B", "CycleFillKeyboard"

    ' Trace Precedents (Ctrl+Shift+T)
    Application.OnKey "^+T", "TracePrecedentsKeyboard"

    ' Trace Dependents (Ctrl+Shift+Y)
    Application.OnKey "^+Y", "TraceDependentsKeyboard"
End Sub
```

**Note:** The add-in uses keyboard wrapper functions to support both ribbon buttons and keyboard shortcuts.

3. **File > Save** (Cmd+S)

**Step 4: Save as Add-In**
1. Close VBA Editor (Cmd+Q)
2. **File > Save As...**
3. **Where:** Navigate to Add-ins folder:
   ```
   /Users/[YourName]/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins/
   ```
   **Tip:** Press **Cmd+Shift+G**, paste path above, replace [YourName]

4. **File Format:** **Excel Macro-Enabled Add-In (.xlam)**
5. **Name:** `LeanMacroTools_v1.0.7`
6. **Save**
7. Close the workbook

#### Part 2: Add Ribbon Tab (3 minutes)

This adds a "Lean Macros" tab to Excel with buttons for all features.

**Step 1: Install Python** (if not installed)
```bash
# Check if Python is installed
python3 --version

# If not, install via Homebrew:
brew install python3
```

**Step 2: Run the Ribbon Injector Script**

**Easy way (recommended):**
```bash
cd /path/to/LeanMacroTool  # Or wherever you saved the files
./install_ribbon.sh
```

The script will automatically:
- Detect your Add-ins folder (even if localized like `Add-Ins.localized`)
- Find your .xlam file (searches for v1.0.7, v1.0.6, v1.0.5, etc.)
- Inject the ribbon XML

**Manual way (if needed):**
```bash
cd /path/to/LeanMacroTool

# For English macOS:
python3 inject_ribbon.py \
  "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins/LeanMacroTools_v1.0.7.xlam" \
  customUI14.xml \
  _rels_dot_rels_for_customUI.xml

# For localized macOS (Portuguese, etc.):
python3 inject_ribbon.py \
  "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/LeanMacroTools_v1.0.7.xlam" \
  customUI14.xml \
  _rels_dot_rels_for_customUI.xml
```

**Note:** The script auto-detects localized folder names. If it can't find the file, check your exact path:
```bash
# Find your actual Add-ins folder:
find ~/Library -name "Add-*ns*" -type d 2>/dev/null | grep Office
```

**Step 3: Restart Excel**
- Quit Excel completely
- Reopen Excel
- You should see a **"Lean Macros"** tab in the ribbon!

#### Part 3: Enable the Add-In

1. In Excel: **Tools > Excel Add-ins...**
2. Check ☑ **LeanMacroTools**
3. Click **OK**

</details>

---

## Usage

### Via Ribbon Tab

Click the **"Lean Macros"** tab, then click any button:
- **Cycle Formats** - Change number format
- **Cycle Colors** - Change font color
- **Cycle Fill** - Change fill pattern and borders
- **Trace Precedents** - See formula inputs
- **Trace Dependents** - See what uses this cell

### Via Keyboard (Faster!)

- **Ctrl+Shift+N** - Cycle formats
- **Ctrl+Shift+V** - Cycle font colors
- **Ctrl+Shift+B** - Cycle fill patterns
- **Ctrl+Shift+T** - Trace precedents (opens navigator dialog)
- **Ctrl+Shift+Y** - Trace dependents (opens navigator dialog)

(Note: Use Control key, not Command)

### Trace Navigator Controls

When you open the Trace Precedents/Dependents dialog:

**Navigate automatically through list:**
- Type **+** or **n** (next) - Jump to next cell
- Type **-** or **p** (previous) - Jump to previous cell
- Type **0** - Go to current/origin cell
- Type **1**, **2**, **3**, etc. - Jump directly to that cell
- Press **ESC** or **Cancel** - Close dialog

The dialog stays open so you can explore multiple cells without reopening it!

---

## For Developers

### Building and Releasing

The project uses a template-based build system for easy development and distribution.

**Quick Start:**
```bash
./scripts/build_release.sh
```

This automatically:
1. Copies `templates/LeanMacroTools_template.xlam`
2. Injects ribbon UI
3. Creates distribution package in `dist/`
4. Generates `.zip` file ready for GitHub Releases

**When You Update Code:**

1. **Edit VBA modules** in `src/*.bas` files
2. **Update the template:**
   - Open `templates/LeanMacroTools_template.xlam` in Excel
   - Press Option+F11 (VBA Editor)
   - Re-import the changed module(s)
   - Save and close
3. **Build release:** `./scripts/build_release.sh`

**Development Scripts:**
- `scripts/build_release.sh` - Build complete distribution package
- `scripts/install_ribbon.sh` - Inject ribbon into existing .xlam (for manual testing)
- `install.command` - End-user installer (in distribution package)

**Repository Structure:**
```
LeanMacroTool/
├── src/                    # VBA source modules (.bas files)
├── ribbon/                 # Ribbon UI definition files
├── scripts/                # Build and installation scripts
├── templates/              # Template .xlam with VBA modules
└── install.command         # User-facing installer
```

---

## Customizing Number Formats

To add, remove, or modify number formats, edit the `LoadFormats` function in `modNumberFormats.bas`:

**Method 1: Edit source file**
1. Open `modNumberFormats.bas` in a text editor
2. Find the `allFormats = Array(...)` section (around line 137)
3. Add, remove, or modify format strings in the array
4. Re-import the module into your `.xlam` file

**Method 2: Edit within Excel VBA**
1. Open Excel and press **Option+F11** (VBA Editor)
2. Find `modNumberFormats` module in your add-in
3. Locate the `LoadFormats` function
4. Edit the `allFormats = Array(...)` section
5. Save (Cmd+S) and restart Excel

The array automatically calculates the format count, so just add or remove lines as needed!

---

## Default Formats

1. `#,##0.00_);(#,##0.00);"-"_);@_)` - Thousands with 2 decimals (1,234.56)
2. `0.0%_);(0.0%);"-"_);@_)` - Percentage (12.3%)
3. `#,##0.0x_);(#,##0.0)x;"-"_);@_)` - Multiple (2.5x)
4. `[$R$-416]#,##0.0_);([$R$-416]#,##0.0);"-"_);@_)` - Brazilian Reals (R$1,234.5)
5. `[$$-409]#,##0.0_);([$$-409]#,##0.0);"-"_);@_)` - US Dollars ($1,234.5)
6. `dd-mmm-yy_)` - Date format (15-Jan-25)
7. `mmm-yy_)` - Month-year format (Jan-25)
8. `General_)` - General number format

Pressing Ctrl+Shift+N cycles through all formats and wraps back to the original cell format.

---

## Troubleshooting

### "Can't find Add-ins folder"
Press **Cmd+Shift+G** in Finder, paste:
```
/Users/[YourName]/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins/
```
Replace [YourName] with your Mac username.

### "Ribbon tab not showing"
- Make sure you ran `inject_ribbon.py` successfully
- Restart Excel completely (Cmd+Q then reopen)
- Check that the .xlam file exists in Add-ins folder

### "Keyboard shortcuts not working"
- Make sure add-in is enabled (**Tools > Excel Add-ins**)
- Use **Control** key (not Command)
- Restart Excel

### "Macros don't appear in macro list"
- That's normal for add-ins
- They work via ribbon buttons and keyboard shortcuts
- Or change "Macros in:" to "All Open Workbooks"

### "Python script failed"
Make sure all 3 files are in the same folder:
- `inject_ribbon.py`
- `customUI14.xml`
- `_rels_dot_rels_for_customUI.xml`

---

## Files Included

### For Users (in distribution package):
1. **LeanMacroTools_v1.0.7.xlam** - Pre-built add-in with ribbon UI embedded
2. **install.command** - Double-click installer (auto-detects Add-ins folder)
3. **README.md** - This file
4. **CHANGELOG.md** - Version history
5. **LICENSE** - MIT License

### For Developers (in source repository):
1. **modNumberFormats.bas** - Number formatting code
2. **modColorFormats.bas** - Font color cycling code
3. **modFillFormats.bas** - Fill pattern and border cycling code
4. **modTraceTools.bas** - Tracing code with keyboard navigation
5. **customUI14.xml** - Ribbon tab definition
6. **_rels_dot_rels_for_customUI.xml** - Ribbon relationships
7. **inject_ribbon.py** - Script to add ribbon to .xlam
8. **install_ribbon.sh** - Manual ribbon installer (for development)
9. **build_release.sh** - Creates distribution packages
10. **ThisWorkbook_KeyboardShortcuts.txt** - Keyboard shortcut registration guide

---

## System Requirements

- macOS 12+ (Monterey or newer)
- Excel for Mac 16.x
- Macros enabled in Excel

**Note:** Python is NOT required for users! The pre-built .xlam already includes the ribbon UI. Python is only needed for developers building new releases.

---

## Tips for Additional Useful Shortcuts

While this add-in provides custom macros, Excel for Mac also lets you customize keyboard shortcuts for built-in commands. Here are some recommended shortcuts that complement the add-in features:

### How to Set Up Custom Shortcuts

1. Go to **Tools > Customize Keyboard...**
2. In the **Categories** list, select **Home Tab**
3. Search for and add shortcuts for these useful commands:

### Recommended Shortcuts

| Command | Suggested Shortcut | Description |
|---------|-------------------|-------------|
| **Increase Indent** | Ctrl+Shift+] | Move content farther from cell border |
| **Decrease Indent** | Ctrl+Shift+[ | Move content closer to cell border |
| **Increase Decimal** | Ctrl+Shift++ | Show more decimal places |
| **Decrease Decimal** | Ctrl+Shift+. | Show fewer decimal places |

### Setup Instructions

**For Indent Controls:**
1. In Customize Keyboard, search for "inde"
2. Select **Increase Indent** → Press **Ctrl+Shift+]** → Click **Add**
3. Select **Decrease Indent** → Press **Ctrl+Shift+[** → Click **Add**

**For Decimal Controls:**
1. Search for "dec"
2. Select **Increase Decimal** → Press **Ctrl+Shift++** → Click **Add**
3. Select **Decrease Decimal** → Press **Ctrl+Shift+.** → Click **Add**

**Note:** These shortcuts are Excel native features and work independently of the LeanMacroTools add-in.

---

## Version History

### v1.0.7 (Current)
- ✨ **NEW:** Simplified installation process! Now just 2 steps
- ✨ **NEW:** install.command - Double-click installer script
- ✨ **NEW:** build_release.sh - Creates distribution packages
- ✨ Auto-detects Add-ins folder (handles localized paths)
- ✨ Pre-built .xlam with ribbon UI already embedded
- 📝 No Python required for users anymore!
- 📝 Manual installation moved to "Advanced" section
- 📝 Updated documentation for simpler workflow

### v1.0.6
- ✨ **NEW:** Fill pattern cycling feature (Ctrl+Shift+B)
- ✨ Cycles through: Color+Border → Pattern → Original
- ✨ First format: Beige background with outline border
- ✨ Second format: Fine dots pattern fill
- ✨ Remembers and restores original formatting
- 📝 Added Cycle Fill ribbon button
- 📝 Supports both keyboard shortcut and ribbon button

### v1.0.5
- ✨ **NEW:** Font color cycling feature (Ctrl+Shift+V)
- ✨ Cycles through: Blue → Green → Red → Grey → Black → Original
- ✨ Remembers and restores original font color for each cell
- 📝 Added Cycle Colors ribbon button
- 📝 Supports both keyboard shortcut and ribbon button

### v1.0.4
- ✨ **NEW:** Original format tracking - cycling returns to cell's original format!
- 🐛 **FIXED:** Ribbon buttons now work (fixed callback signatures)
- 🐛 **FIXED:** Configure button in ribbon functional
- 🐛 **FIXED:** Format cycling returns to original instead of getting stuck
- 📝 Cycle order: Original → Thousands → Percentage → Multiples → USD → BRL → Original
- 📝 Updated Workbook_Open to use keyboard wrapper functions

### v1.0.3
- ✨ **NEW:** Interactive keyboard navigation for trace dialogs
- ✨ **NEW:** Dialog stays open for exploring multiple cells
- ✨ **NEW:** Current cell included in trace list (index 0)
- ✨ **NEW:** Navigate with +/- or n/p keys through cells automatically
- 🐛 **FIXED:** Cross-sheet navigation now works correctly
- 🐛 **FIXED:** macOS path handling in install script
- 🐛 **FIXED:** Dialog formatting (removed emoji characters)

### v1.0.2
- Initial release with basic trace functionality
- Cross-sheet reference support

### v1.0.1
- Added ribbon UI integration
- Keyboard shortcuts

### v1.0.0
- Number format cycling
- Basic trace precedents/dependents

---

## Questions?

Check the troubleshooting section above. All features are documented in this README.

Enjoy your faster Excel workflow! 🚀
