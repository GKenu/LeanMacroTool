# VISUAL WALKTHROUGH - What You'll See

## 📸 STEP-BY-STEP WITH EXAMPLES

---

## 1️⃣ NUMBER FORMAT CYCLING (Ctrl+Shift+N)

### Before:
```
Cell A1 contains: 1234.56
Current format: General
Displays as: 1234.56
```

### Press Ctrl+Shift+N once:
```
Cell A1 contains: 1234.56
New format: #,##0.00_);(#,##0.00);"-"_);@_)
Displays as: 1,234.56
```

### Press Ctrl+Shift+N again:
```
Cell A1 contains: 1234.56
New format: 0.0%_);(0.0%);"-"_);@_)
Displays as: 123456.0%
```

### Press Ctrl+Shift+N again:
```
Cell A1 contains: 1234.56
New format: #,##0.0x_);(#,##0.0)x;"-"_);@_)
Displays as: 1,234.6x
```

### Keep pressing → cycles through all 5 formats → wraps back to first

---

## 2️⃣ TRACE PRECEDENTS (Ctrl+Shift+T)

### Your Spreadsheet:
```
     A          B          C
1   100        200        =SUM(A1:B1)
2   150        250        =SUM(A2:B2)
3   =A1+A2    =B1+B2      =C1+C2
```

### Click on C3, Press Ctrl+Shift+T:

You see this dialog:

```
┌────────────────────────────────────────────────────┐
│                TRACE PRECEDENTS                    │
│                                                    │
│  Origin: Sheet1!C3                                │
│  Value: 1000                                      │
│  Formula: =C1+C2                                  │
│                                                    │
│  Precedent Cells:                                 │
│    1. Sheet1!C1                                   │
│    2. Sheet1!C2                                   │
│                                                    │
│  Enter cell number to jump to (or Cancel):        │
│  [ 1                                         ]    │
│                                                    │
│            [    OK    ]  [   Cancel   ]           │
└────────────────────────────────────────────────────┘
```

### Type "1" and press Enter:
- Excel jumps to cell C1
- Dialog closes
- You can now see what's in C1

### Click C1, Press Ctrl+Shift+T again:
```
┌────────────────────────────────────────────────────┐
│                TRACE PRECEDENTS                    │
│                                                    │
│  Origin: Sheet1!C1                                │
│  Value: 300                                       │
│  Formula: =SUM(A1:B1)                             │
│                                                    │
│  Precedent Cells:                                 │
│    1. Sheet1!A1                                   │
│    2. Sheet1!B1                                   │
│                                                    │
│  Enter cell number to jump to (or Cancel):        │
└────────────────────────────────────────────────────┘
```

### Now you can trace back further!

---

## 3️⃣ TRACE DEPENDENTS (Ctrl+Shift+Y)

### Same Spreadsheet:
```
     A          B          C
1   100        200        =SUM(A1:B1)
2   150        250        =SUM(A2:B2)
3   =A1+A2    =B1+B2      =C1+C2
```

### Click on A1, Press Ctrl+Shift+Y:

You see:

```
┌────────────────────────────────────────────────────┐
│                TRACE DEPENDENTS                    │
│                                                    │
│  Origin: Sheet1!A1                                │
│  Value: 100                                       │
│                                                    │
│  Dependent Cells:                                 │
│    1. Sheet1!C1                                   │
│    2. Sheet1!A3                                   │
│                                                    │
│  Enter cell number to jump to (or Cancel):        │
│  [                                           ]    │
│                                                    │
│            [    OK    ]  [   Cancel   ]           │
└────────────────────────────────────────────────────┘
```

### This shows:
- C1 uses A1 in its formula: =SUM(A1:B1)
- A3 uses A1 in its formula: =A1+A2

### Type "2" and press Enter:
- Jumps to A3
- Now you can see how A1 flows through the model

---

## 4️⃣ CONFIGURE NUMBER FORMATS

### Run: Tools > Macro > Macros > ConfigureNumberFormats

### First, you see this message box:
```
┌────────────────────────────────────────────────────┐
│           Configure Number Formats                 │
├────────────────────────────────────────────────────┤
│                                                    │
│  The NumberFormatConfig sheet is now visible.     │
│                                                    │
│  Column A: Number format codes                    │
│  Column B: TRUE to enable, FALSE to disable       │
│                                                    │
│  Edit the formats as needed, then click OK to     │
│  save and hide the sheet.                         │
│                                                    │
│            [    OK    ]  [   Cancel   ]           │
└────────────────────────────────────────────────────┘
```

### Click OK, and a sheet appears:

```
Sheet: NumberFormatConfig
┌─────────────────────────────────────────────┬──────────┐
│                   Format                    │ Enabled  │
├─────────────────────────────────────────────┼──────────┤
│ #,##0.00_);(#,##0.00);"-"_);@_)            │  TRUE    │
│ 0.0%_);(0.0%);"-"_);@_)                    │  TRUE    │
│ #,##0.0x_);(#,##0.0)x;"-"_);@_)            │  TRUE    │
│ $#,##0.0_);$(#,##0.0)"x";"-"_);@_)         │  TRUE    │
│ R$#,##0.0_);R$(#,##0.0)"x";"-"_);@_)       │  TRUE    │
└─────────────────────────────────────────────┴──────────┘
```

### You can:
1. **Change a format string** - Edit column A
2. **Disable a format** - Change TRUE to FALSE
3. **Add new format** - Add new row with format in A, TRUE in B
4. **Delete format** - Delete the row

### When done, the message box appears again - Click OK:
- Sheet is hidden
- Changes are saved
- Next time you press Ctrl+Shift+N, it uses your new config!

---

## 🎯 CROSS-SHEET TRACING

### Your Workbook has 2 sheets:

**Sheet: Revenue**
```
     A          B          
1   Q1         Q2         
2   1000       1500       
```

**Sheet: Summary**  
```
     A          
1   Total         
2   =SUM(Revenue!A2:B2)
```

### Click on Summary!A2, Press Ctrl+Shift+T:

```
┌────────────────────────────────────────────────────┐
│                TRACE PRECEDENTS                    │
│                                                    │
│  Origin: Summary!A2                               │
│  Value: 2500                                      │
│  Formula: =SUM(Revenue!A2:B2)                     │
│                                                    │
│  Precedent Cells:                                 │
│    1. Revenue!A2                                  │
│    2. Revenue!B2                                  │
│                                                    │
│  Enter cell number to jump to (or Cancel):        │
└────────────────────────────────────────────────────┘
```

### Type "1" and press Enter:
- Excel switches to Revenue sheet
- Selects cell A2
- You can see the source data!

---

## ⚠️ ERROR MESSAGES YOU MIGHT See

### No Formula (when tracing precedents):
```
┌────────────────────────────────────────────┐
│         Trace Precedents             [X]   │
├────────────────────────────────────────────┤
│                                            │
│  ⓘ  The selected cell does not contain   │
│      a formula.                            │
│                                            │
│              [    OK    ]                  │
└────────────────────────────────────────────┘
```

### No Precedents Found:
```
┌────────────────────────────────────────────┐
│         Trace Precedents             [X]   │
├────────────────────────────────────────────┤
│                                            │
│  ⓘ  No precedent cells found for          │
│      Sheet1!A1                             │
│                                            │
│              [    OK    ]                  │
└────────────────────────────────────────────┘
```

### No Dependents Found:
```
┌────────────────────────────────────────────┐
│         Trace Dependents             [X]   │
├────────────────────────────────────────────┤
│                                            │
│  ⓘ  No dependent cells found for          │
│      Sheet1!Z99                            │
│                                            │
│              [    OK    ]                  │
└────────────────────────────────────────────┘
```

### No Formats Enabled:
```
┌────────────────────────────────────────────┐
│         No Formats                   [X]   │
├────────────────────────────────────────────┤
│                                            │
│  ⚠  No number formats are enabled.        │
│      Please configure formats first.       │
│                                            │
│              [    OK    ]                  │
└────────────────────────────────────────────┘
```

---

## 🎬 COMPLETE WORKFLOW EXAMPLE

### Scenario: Building a Financial Model

**Step 1: Set up revenue assumptions**
```
     A              B
1   Revenue        1234567
2   Growth Rate    0.15
3   Year 2         =A1*(1+A2)
```

**Step 2: Format the numbers**
- Select A1
- Press **Ctrl+Shift+N** until you see: `1,234,567.00`
- Select A2  
- Press **Ctrl+Shift+N** until you see: `15.0%`
- Select A3
- Press **Ctrl+Shift+N** → matches A1 format: `1,420,000.00`

**Step 3: Audit the calculation**
- Click on A3
- Press **Ctrl+Shift+T**
- See: "Precedents: 1. Sheet1!A1, 2. Sheet1!A2"
- Type "2" → Jump to A2
- Verify growth rate is correct

**Step 4: Check what uses this cell**
- Click on A1 (original revenue)
- Press **Ctrl+Shift+Y**
- See: "Dependents: 1. Sheet1!A3"
- Confirms A3 depends on A1

**Step 5: Make a change**
- Change A1 to 2000000
- Press **Ctrl+Shift+Y** to verify A3 updated
- Jump to A3 and check value

**Done in seconds with just keyboard shortcuts!**

---

## 💪 POWER USER MOVES

### Move 1: Rapid Formatting
```
1. Select A1:A100 (all revenue numbers)
2. Ctrl+Shift+N → all formatted as thousands
3. Select B1:B100 (all percentages)
4. Ctrl+Shift+N (multiple times) → all formatted as %
5. 2 seconds vs 30 seconds with Format Cells dialog!
```

### Move 2: Formula Chain Navigation
```
1. Start at final output cell
2. Ctrl+Shift+T → see inputs
3. Jump to suspicious input (type its number)
4. Ctrl+Shift+T → see ITS inputs
5. Keep going until you find the error
6. Fix it
7. Ctrl+Shift+Y → see what updated
8. Walk back up the chain verifying
```

### Move 3: Format Library
```
1. Create 20 custom formats in config
2. Enable only 5 for current project
3. Ctrl+Shift+N cycles through just those 5
4. Next project: Disable those 5, enable different 5
5. Never lose your format library!
```

---

## 🎓 COMPARISON TO MANUAL METHODS

### Format Cells Dialog (Manual Way):
1. Select cell
2. Right-click
3. Choose "Format Cells..."
4. Click "Number" tab
5. Scroll through categories
6. Select category
7. Type custom format code
8. Click OK
**Time: ~15 seconds per cell**

### Ctrl+Shift+N (This Add-in):
1. Select cell
2. Press Ctrl+Shift+N
**Time: ~1 second per cell**
**15x faster!**

---

### Excel's Built-in Trace (Manual Way):
1. Click cell
2. Go to Formulas tab in ribbon
3. Click "Trace Precedents" button
4. See blue arrows on sheet
5. Follow arrows to find cells
6. Click on arrow to navigate
7. Click "Remove Arrows" to clean up
**Time: ~10 seconds per trace**

### Ctrl+Shift+T (This Add-in):
1. Click cell
2. Press Ctrl+Shift+T
3. See numbered list
4. Type number to jump
**Time: ~2 seconds per trace**
**5x faster + cleaner!**

---

## ✨ FINAL NOTES

- All dialogs are **native Excel InputBox/MsgBox** = Mac-compatible
- No custom UI = No positioning/sizing issues
- Text-based = Easy to read
- Keyboard-driven = Perfect for power users
- Lightweight = Fast and reliable

**The simplified approach is actually BETTER for pros! 🚀**

---

See INSTALLATION_GUIDE_SIMPLIFIED.md for setup instructions.
See QUICK_REFERENCE.md for command reference.
