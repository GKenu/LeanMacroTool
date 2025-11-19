# SIMPLIFIED LEAN MACRO TOOLS FOR EXCEL MAC
## NO UserForms Required - Easy Installation!

This version uses native Excel dialogs instead of custom UserForms, making it:
- ✅ Much easier to install
- ✅ More Mac-compatible
- ✅ Simpler to maintain
- ✅ Just as functional!

---

## WHAT YOU GET

**3 Keyboard Shortcuts:**

1. **Ctrl+Shift+N** - Cycle Number Formats
   - Cycles through your custom formats
   - Works on single cells or selections
   
2. **Ctrl+Shift+T** - Trace Precedents  
   - Shows all cells that feed into the formula
   - Lists them numbered
   - Enter a number to jump to that cell
   
3. **Ctrl+Shift+Y** - Trace Dependents
   - Shows all cells that depend on the current cell
   - Enter a number to jump to that cell

**Bonus:** `ConfigureNumberFormats` macro to customize your format cycle

---

## SUPER SIMPLE INSTALLATION

### Step 1: Open Your File

1. Open the file you saved: `kenu-tts.xlsm`
2. Press `Option + F11` to open VBA Editor

### Step 2: Delete What's There (if anything)

- If you have any modules already, right-click them and select `Remove`
- We're starting fresh!

### Step 3: Import the Code File

1. In VBA Editor, go to **File > Import File...**
2. Navigate to where you saved `LeanMacroTools_Complete_Code.bas`
3. Select it and click **Open**

That's it! The code is now imported as TWO modules:
- `modNumberFormats`
- `modTraceTools`

### Step 4: Save as Add-In

1. Close VBA Editor (`Cmd + Q`)
2. In Excel: **File > Save As**
3. Choose location: 
   ```
   /Users/[YourUsername]/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins/
   ```
   **TIP:** Press `Cmd + Shift + G` in the save dialog and paste that path
4. File Format: **Excel Add-In (*.xlam)**
5. Filename: `LeanMacroTools.xlam`
6. Click **Save**

### Step 5: Enable the Add-In

1. In Excel: **Tools > Excel Add-ins...**
2. Check the box next to `LeanMacroTools`
3. Click **OK**

### Step 6: Assign Keyboard Shortcuts

1. **Tools > Macro > Macros** (or `Option + F8`)

2. **For Number Format Cycler:**
   - Select `LeanMacroTools.xlam!CycleCustomNumberFormats`
   - Click **Options...**
   - Shortcut key: Type capital **N**
   - Click **OK**

3. **For Trace Precedents:**
   - Select `LeanMacroTools.xlam!TracePrecedentsDialog`
   - Click **Options...**
   - Shortcut key: Type capital **T**
   - Click **OK**

4. **For Trace Dependents:**
   - Select `LeanMacroTools.xlam!TraceDependentsDialog`
   - Click **Options...**
   - Shortcut key: Type capital **Y**
   - Click **OK**

---

## HOW TO USE

### 🔢 Number Format Cycling (Ctrl+Shift+N)

1. Select any cell(s)
2. Press **Ctrl+Shift+N**
3. Format changes to next in cycle
4. Press again to cycle to next format
5. After last format, wraps back to first

**Default Formats:**
1. `#,##0.00_);(#,##0.00);"-"_);@_)` - Numbers with 2 decimals
2. `0.0%_);(0.0%);"-"_);@_)` - Percentages
3. `#,##0.0x_);(#,##0.0)x;"-"_);@_)` - Multiples (e.g., 2.5x)
4. `$#,##0.0_);$(#,##0.0)"x";"-"_);@_)` - Dollars
5. `R$#,##0.0_);R$(#,##0.0)"x";"-"_);@_)` - Reals (R$)

### ⚙️ Configure Formats

To customize which formats are in the cycle:

1. **Tools > Macro > Macros**
2. Run: `ConfigureNumberFormats`
3. A hidden sheet becomes visible with:
   - Column A: Format codes
   - Column B: TRUE (enabled) or FALSE (disabled)
4. Edit as needed:
   - Change format strings
   - Toggle TRUE/FALSE to enable/disable
   - Add new rows for more formats
5. Click **OK** when done
6. Sheet is hidden again, changes are saved

### 🔍 Trace Precedents (Ctrl+Shift+T)

1. Click on a cell **with a formula**
2. Press **Ctrl+Shift+T**
3. A dialog shows:
   ```
   TRACE PRECEDENTS
   ==================================================
   
   Origin: Sheet1!C5
   Value: 150
   Formula: =SUM(A1:A10)
   
   Precedent Cells:
     1. Sheet1!A1
     2. Sheet1!A2
     3. Sheet1!A3
     ...
   
   Enter cell number to jump to (or Cancel):
   ```
4. Type a number (e.g., `2`) and press Enter to jump to that cell
5. Or click **Cancel** to close

### 🔍 Trace Dependents (Ctrl+Shift+Y)

Same as precedents, but shows cells that **use** the current cell in their formulas.

1. Click on any cell
2. Press **Ctrl+Shift+Y**
3. See list of dependent cells
4. Enter number to jump to that cell

---

## TROUBLESHOOTING

### "Can't find the macro"
- Make sure add-in is enabled: **Tools > Excel Add-ins**
- Check the box next to `LeanMacroTools`

### "Keyboard shortcut not working"
- Excel for Mac uses **Control** key (not Command)
- Make sure you typed capital letters when assigning shortcuts
- Try re-assigning: **Tools > Macro > Macros > Options**

### "No precedents/dependents found"
- For precedents: Cell must have a formula
- For dependents: Other cells must reference this cell
- Excel only shows direct links (one level)

### "Configuration sheet not saving"
- After editing, make sure you clicked **OK** in the dialog
- The sheet is automatically hidden after saving

---

## WHY THIS IS BETTER THAN USERFORMS

❌ **UserForms require:**
- Manual creation of each control
- Precise positioning
- More code complexity
- Can have Mac compatibility issues

✅ **This approach uses:**
- Native Excel InputBox and MsgBox
- Automatic formatting
- Less code to maintain
- 100% Mac compatible
- Faster to install!

The functionality is identical - you just see the information in a different format.

---

## TESTING CHECKLIST

After installation, test each feature:

- [ ] Press Ctrl+Shift+N on a cell - format changes
- [ ] Press Ctrl+Shift+N again - cycles to next format
- [ ] Press multiple times - wraps back to first format
- [ ] Select multiple cells - all get formatted
- [ ] Run ConfigureNumberFormats - sheet appears
- [ ] Edit a format - change is saved
- [ ] Click on cell with formula - press Ctrl+Shift+T
- [ ] See list of precedents - enter number - jumps to cell
- [ ] Click on any cell - press Ctrl+Shift+Y  
- [ ] See dependents (if any) - can jump to them

---

## WHAT'S DIFFERENT FROM THE ORIGINAL PLAN?

**Original:** Complex UserForms with ListBoxes, buttons, labels, etc.

**Simplified:** Native Excel dialogs (InputBox/MsgBox)

**What you gain:**
- 10x easier to install (just import one file!)
- No manual UI creation
- Better Mac compatibility
- Same functionality

**What you lose:**
- Slightly less pretty UI (text-based instead of graphical)
- No mouse clicking in lists (type numbers instead)

**Bottom line:** For power users who prefer keyboards anyway, this is actually BETTER!

---

## CUSTOMIZATION TIPS

### Add More Number Formats

1. Run `ConfigureNumberFormats`
2. Add new rows in the sheet:
   - Column A: Your format code
   - Column B: TRUE
3. Click OK

### Change Keyboard Shortcuts

1. **Tools > Macro > Macros**
2. Select macro
3. **Options...**
4. Change to any capital letter

**Good alternatives:**
- Ctrl+Shift+P for Precedents
- Ctrl+Shift+D for Dependents  
- Ctrl+Shift+F for Format cycling

---

## FILE LOCATIONS

**Add-in file:**
```
/Users/[Username]/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins/LeanMacroTools.xlam
```

**To find it:**
1. Open Finder
2. Press `Cmd + Shift + G`
3. Paste the path above
4. You'll see your .xlam file

---

## GETTING HELP

If something doesn't work:

1. **Check Security Settings:**
   - Excel > Preferences > Security & Privacy
   - Under Macro Security: "Enable all macros"

2. **Verify Add-in is Loaded:**
   - Tools > Excel Add-ins
   - LeanMacroTools should be checked

3. **Check Macros Exist:**
   - Tools > Macro > Macros
   - Should see all three macros listed

4. **Restart Excel:**
   - Sometimes a fresh start helps!

---

## NEXT STEPS

Once this is working, you can:

1. **Customize formats** to match your workflow
2. **Add more formats** (up to as many as you want)
3. **Change shortcuts** to your preference
4. **Share the .xlam** file with colleagues

The code is clean, well-commented, and easy to modify if you want to extend it later!

---

**Enjoy your lean, keyboard-driven Excel workflow! 🚀**
