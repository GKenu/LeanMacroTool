# LEAN MACRO TOOLS FOR EXCEL MAC

**3 Powerful Features via Keyboard Shortcuts & Ribbon Tab**

## Features

1. **Cycle Number Formats** (Ctrl+Shift+N)
   - Cycles through: Thousands → Percentage → Multiples → USD → BRL → (wraps back to Thousands)
   - Customizable format list
   
2. **Trace Precedents** (Ctrl+Shift+T)
   - Shows cells that feed into formulas
   - **Mac-native list dialog** - click to select, no typing needed!
   - Works cross-sheet perfectly
   
3. **Trace Dependents** (Ctrl+Shift+Y)
   - Shows cells that use the current cell
   - **Mac-native list dialog** - click to select!
   - Works cross-sheet perfectly

---

## Installation

### Part 1: Create the Add-In (5 minutes)

**Step 1: Create New Workbook**
1. Open Excel
2. Create new blank workbook

**Step 2: Import VBA Modules**
1. Press **Option+F11** (VBA Editor)
2. **File > Import File...** → Select **modNumberFormats.bas** → Open
3. **File > Import File...** → Select **modTraceTools.bas** → Open

You should see both modules in the left panel.

**Step 3: Add Ribbon Callback**
1. In VBA Editor, double-click **ThisWorkbook** (left panel)
2. Paste this code:

```vba
Private Sub Workbook_Open()
    Application.OnKey "^+N", "CycleCustomNumberFormats"
    Application.OnKey "^+T", "TracePrecedentsDialog"
    Application.OnKey "^+Y", "TraceDependentsDialog"
End Sub
```

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
5. **Name:** `LeanMacroTools`
6. **Save**
7. Close the workbook

---

### Part 2: Add Ribbon Tab (3 minutes)

This adds a "Lean Macros" tab to Excel with buttons for all features.

**Step 1: Install Python** (if not installed)
```bash
# Check if Python is installed
python3 --version

# If not, install via Homebrew:
brew install python3
```

**Step 2: Run the Ribbon Injector Script**
```bash
cd ~/Downloads  # Or wherever you saved the files

python3 inject_ribbon.py \
  ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content/Add-ins/LeanMacroTools.xlam \
  customUI14.xml \
  _rels_dot_rels_for_customUI.xml
```

**Note:** Replace `LeanMacroTools.xlam` with your actual filename if different (e.g., `LeanMacroTools_v1.0.1.xlam`).

If the script can't find the file, check the exact filename:
```bash
ls ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content/Add-ins/
```

**Step 3: Restart Excel**
- Quit Excel completely
- Reopen Excel
- You should see a **"Lean Macros"** tab in the ribbon!

---

### Part 3: Enable the Add-In

1. In Excel: **Tools > Excel Add-ins...**
2. Check ☑ **LeanMacroTools**
3. Click **OK**

---

## Usage

### Via Ribbon Tab

Click the **"Lean Macros"** tab, then click any button:
- **Cycle Formats** - Change number format
- **Configure** - Customize formats
- **Trace Precedents** - See formula inputs
- **Trace Dependents** - See what uses this cell

### Via Keyboard (Faster!)

- **Ctrl+Shift+N** - Cycle formats
- **Ctrl+Shift+T** - Trace precedents
- **Ctrl+Shift+Y** - Trace dependents

(Note: Use Control key, not Command)

---

## Customizing Number Formats

**Method 1: Via Ribbon**
1. Click **Lean Macros** tab
2. Click **Configure** button
3. Edit the sheet that appears
4. Click OK when done

**Method 2: Via Macro**
1. **Tools > Macro > Macros**
2. Run `ConfigureNumberFormats`
3. Edit the sheet
4. Click OK

The sheet shows:
- Column A: Number format codes
- Column B: TRUE (enabled) or FALSE (disabled)

---

## Default Formats

1. `#,##0.00_);(#,##0.00);"-"_);@_)` - Thousands with 2 decimals (1,234.56)
2. `0.0%_);(0.0%);"-"_);@_)` - Percentage (12.3%)
3. `#,##0.0x_);(#,##0.0)x;"-"_);@_)` - Multiple (2.5x)
4. `$#,##0.0_);$(#,##0.0);"-"_);@_)` - US Dollars ($1,234.5)
5. `R$#,##0.0_);R$(#,##0.0);"-"_);@_)` - Brazilian Reals (R$1,234.5)

After the 5th format, pressing Ctrl+Shift+N wraps back to the first format.

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

1. **modNumberFormats.bas** - Number formatting code
2. **modTraceTools.bas** - Tracing code
3. **customUI14.xml** - Ribbon tab definition
4. **_rels_dot_rels_for_customUI.xml** - Ribbon relationships
5. **inject_ribbon.py** - Script to add ribbon to .xlam
6. **README.md** - This file

---

## System Requirements

- macOS 12+ (Monterey or newer)
- Excel for Mac 16.x
- Python 3 (for ribbon injection)
- Macros enabled in Excel

---

## Questions?

Check the troubleshooting section above. All features are documented in this README.

Enjoy your faster Excel workflow! 🚀
