# LEAN MACRO TOOLS - INSTALLATION

## What You Get
3 keyboard shortcuts for Excel Mac:
- **Ctrl+Shift+N** - Cycle number formats
- **Ctrl+Shift+T** - Trace precedents  
- **Ctrl+Shift+Y** - Trace dependents

---

## Installation (5 minutes)

### Step 1: Create New Workbook
1. Open Excel
2. Create a **new blank workbook**

### Step 2: Import the Modules
1. Press **Option+F11** (opens VBA Editor)
2. Go to **File > Import File...**
3. Select **modNumberFormats.bas** → Open
4. Go to **File > Import File...** again
5. Select **modTraceTools.bas** → Open

You should now see in the left panel:
```
VBAProject (Book1)
├─ Microsoft Excel Objects
│  └─ ThisWorkbook
└─ Modules
   ├─ modNumberFormats
   └─ modTraceTools
```

### Step 3: Save as Add-In
1. Press **Cmd+Q** to close VBA Editor
2. Go to **File > Save As...**
3. **Where:** Navigate to:
   ```
   /Users/[YourName]/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins/
   ```
   **Tip:** Press **Cmd+Shift+G** in the save dialog, paste the path above, change [YourName] to your username
   
4. **File Format:** Select **Excel Macro-Enabled Add-In (.xlam)**
5. **Name:** `LeanMacroTools`
6. Click **Save**
7. **Close** the workbook (don't need it anymore)

### Step 4: Enable the Add-In
1. In Excel, go to **Tools > Excel Add-ins...**
2. Find and check ☑ **LeanMacroTools**
3. Click **OK**

### Step 5: Assign Keyboard Shortcuts
1. Go to **Tools > Macro > Macros** (or press **Option+F8**)
2. You'll see your macros listed:

**For Number Format Cycler:**
- Select `LeanMacroTools.xlam!CycleCustomNumberFormats`
- Click **Options...**
- In "Shortcut key" box, type capital **N**
- Click **OK**

**For Trace Precedents:**
- Select `LeanMacroTools.xlam!TracePrecedentsDialog`  
- Click **Options...**
- Type capital **T**
- Click **OK**

**For Trace Dependents:**
- Select `LeanMacroTools.xlam!TraceDependentsDialog`
- Click **Options...**
- Type capital **Y**
- Click **OK**

### Step 6: Test It!
1. Open any workbook
2. Click on a cell
3. Press **Ctrl+Shift+N** (Control, not Command)
4. The format should change!

---

## Usage

### Cycle Number Formats (Ctrl+Shift+N)
1. Select cell(s)
2. Press **Ctrl+Shift+N** repeatedly
3. Format cycles through: thousands → percentage → multiples → dollars → reals

**To customize formats:**
- Run macro: `ConfigureNumberFormats`
- A hidden sheet appears where you can edit formats
- Click OK when done

### Trace Precedents (Ctrl+Shift+T)
1. Click on cell **with a formula**
2. Press **Ctrl+Shift+T**
3. See numbered list of cells that feed into the formula
4. Type a number and press Enter to jump to that cell

### Trace Dependents (Ctrl+Shift+Y)
1. Click on any cell
2. Press **Ctrl+Shift+Y**
3. See numbered list of cells that use this cell
4. Type a number to jump to that cell

---

## Troubleshooting

**"Can't find the Add-ins folder"**
- Press **Cmd+Shift+G** in Finder
- Paste: `/Users/[YourName]/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins/`
- Replace [YourName] with your actual Mac username

**"Macros are disabled"**
- Go to **Excel > Preferences > Security & Privacy**
- Under "Macro Security", select **"Enable all macros"**
- Restart Excel

**"Add-in not showing up"**
- Make sure you saved it in the Add-ins folder
- Try **Tools > Excel Add-ins** → uncheck and recheck it
- Restart Excel

**"Shortcuts not working"**
- Make sure you're using **Control** key (not Command)
- Re-assign shortcuts: **Tools > Macro > Macros > Options**
- Make sure you typed CAPITAL letters (N, T, Y)

**"Only modNumberFormats imported"**
- You must import BOTH .bas files separately
- Each import adds one module
- Check VBA Editor to confirm both are there

---

## Files Included

1. **modNumberFormats.bas** - Number format cycling code
2. **modTraceTools.bas** - Precedent/dependent tracing code
3. **INSTALLATION.md** - This file

---

## Why Two Files?

Excel VBA .bas files can only contain ONE module each. That's why you have to import both files to get all features.

---

## Default Number Formats

1. `#,##0.00_);(#,##0.00);"-"_);@_)` - Thousands with 2 decimals
2. `0.0%_);(0.0%);"-"_);@_)` - Percentage
3. `#,##0.0x_);(#,##0.0)x;"-"_);@_)` - Multiples (2.5x)
4. `$#,##0.0_);$(#,##0.0)"x";"-"_);@_)` - US Dollars
5. `R$#,##0.0_);R$(#,##0.0)"x";"-"_);@_)` - Brazilian Reals

Run `ConfigureNumberFormats` to add more or modify these.

---

## That's It!

You now have 3 powerful keyboard shortcuts. Enjoy your faster Excel workflow!
