# LEAN MACRO TOOLS - QUICK REFERENCE

## 🎯 KEYBOARD SHORTCUTS

| Shortcut | Function | What It Does |
|----------|----------|--------------|
| **Ctrl+Shift+N** | Number Format Cycle | Cycles through your custom number formats |
| **Ctrl+Shift+T** | Trace Precedents | Shows cells that feed into current cell's formula |
| **Ctrl+Shift+Y** | Trace Dependents | Shows cells that use current cell in their formulas |

---

## 📋 QUICK START

### Installation (5 minutes)
1. Open `kenu-tts.xlsm` in Excel
2. Press `Option+F11` (VBA Editor)
3. **File > Import File...** → Select `LeanMacroTools_Complete_Code.bas`
4. Save as `.xlam` in Add-ins folder
5. Enable in **Tools > Excel Add-ins**
6. Assign keyboard shortcuts in **Tools > Macro > Macros > Options**

### First Use
1. Select a cell
2. Press **Ctrl+Shift+N** to format
3. Press again to cycle to next format
4. Click cell with formula, press **Ctrl+Shift+T** to see precedents
5. Click any cell, press **Ctrl+Shift+Y** to see dependents

---

## 🔧 CONFIGURATION

### Change Number Formats
1. **Tools > Macro > Macros**
2. Run: `ConfigureNumberFormats`
3. Edit the visible sheet (Column A = format, Column B = TRUE/FALSE)
4. Click OK when done

### Default Formats
1. `#,##0.00_);(#,##0.00);"-"_);@_)` - Thousands with 2 decimals
2. `0.0%_);(0.0%);"-"_);@_)` - Percentage
3. `#,##0.0x_);(#,##0.0)x;"-"_);@_)` - Multiple (2.5x)
4. `$#,##0.0_);$(#,##0.0)"x";"-"_);@_)` - USD
5. `R$#,##0.0_);R$(#,##0.0)"x";"-"_);@_)` - BRL (Reals)

---

## ⚡ USAGE TIPS

### Number Formatting
- Works on single cells or multi-cell selections
- Detects current format and moves to next
- Wraps around after last format

### Tracing
- **Precedents:** Cell MUST have a formula
- **Dependents:** Shows direct dependencies only (one level)
- Enter number to jump to that cell, or Cancel to close
- Works across sheets (shows as `Sheet1!A1`)
- Works across workbooks (if open)

---

## 🐛 COMMON ISSUES

| Problem | Solution |
|---------|----------|
| Shortcut not working | Use Control key (not Cmd), reassign if needed |
| Can't find macro | Check add-in is enabled in Tools > Excel Add-ins |
| No precedents found | Cell must contain a formula |
| Configuration not saving | Click OK in the dialog after editing |
| Macro security error | Excel > Preferences > Security > Enable all macros |

---

## 📂 FILE LOCATIONS

**Add-in:**
```
~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins/LeanMacroTools.xlam
```

**Access:** Press `Cmd+Shift+G` in Finder, paste path above

---

## 🎨 CUSTOMIZE

### Add New Format
1. Run `ConfigureNumberFormats`
2. Add row: Format code in column A, `TRUE` in column B
3. Click OK

### Change Shortcuts
1. Tools > Macro > Macros
2. Select macro > Options
3. Type new letter (capital)

### Disable Format
1. Run `ConfigureNumberFormats`
2. Change `TRUE` to `FALSE` for that format
3. Click OK

---

## 💡 WORKFLOW EXAMPLES

### Financial Model Building
1. Enter formulas
2. Select range of numbers
3. **Ctrl+Shift+N** → format as thousands
4. Select range of percentages  
5. **Ctrl+Shift+N** (multiple times) → format as %
6. Need to trace error? **Ctrl+Shift+T** to see inputs

### Formula Auditing
1. Complex formula not working?
2. **Ctrl+Shift+T** → see all inputs
3. Jump to suspicious input
4. Check its formula with **Ctrl+Shift+T**
5. Fix error
6. **Ctrl+Shift+Y** on fixed cell → verify all dependents updated

### Quick Formatting
1. Copy format cycle from old model
2. **Ctrl+Shift+N** repeatedly to find matching format
3. Much faster than manual Format Cells dialog!

---

## 🚀 ADVANCED TIPS

1. **Combine with Excel shortcuts:**
   - Format cells, then `Cmd+C`, select range, `Cmd+V` (pastes format)
   - Or use Ctrl+Shift+N for instant format

2. **Chain tracing:**
   - Trace precedents, jump to one
   - Trace ITS precedents, jump again
   - Walk up/down the formula chain

3. **Multi-format selections:**
   - Select non-contiguous ranges (`Cmd+Click`)
   - Ctrl+Shift+N formats all at once

4. **Speed formatting:**
   - Create 10-20 custom formats
   - Disable all but current 3-5 you're using today
   - Tomorrow's model? Re-enable different set

---

## 📝 NOTES

- **Mac-specific:** Uses Control key (bottom-left), not Command
- **Lightweight:** ~14KB of code, no bloat
- **Pure VBA:** No external dependencies
- **Open source:** Edit the code if you want to customize further

---

**Questions? Check the full INSTALLATION_GUIDE_SIMPLIFIED.md**

Print this page and keep it next to your keyboard! 📄
