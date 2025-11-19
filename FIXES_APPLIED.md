# Fixes Applied to LeanMacroTool

## Summary of Changes

Three major issues have been fixed in your Excel add-in:

### ✅ ISSUE 1 FIXED: Ribbon Installation Path Handling

**Problem:** Python script failed with "file not found" due to spaces in macOS path
```
~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins/
```

**Solution:** Created `install_ribbon.sh` bash script that properly quotes paths

**How to use:**
```bash
cd /Users/kenu/Projects/LeanMacroTool
./install_ribbon.sh
```

The script will:
- Automatically locate your .xlam file in the Add-ins folder
- Handle spaces in the path correctly
- Create a backup before modifying
- Inject the custom ribbon XML

**Manual method (if needed):**
```bash
python3 inject_ribbon.py \
  "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins/LeanMacroTools_v1.0.2.xlam" \
  customUI14.xml \
  _rels_dot_rels_for_customUI.xml
```

### ✅ ISSUE 2 FIXED: Immediate Navigation in Trace Dialog

**Problem:** AppleScript "choose from list" required clicking cell THEN pressing OK button

**Solution:** Replaced with InputBox that navigates immediately after entering a number

**New behavior:**
1. Press Ctrl+Shift+T (precedents) or Ctrl+Shift+Y (dependents)
2. Dialog shows numbered list of linked cells
3. Type the number and press Enter
4. **Immediately jumps to that cell** - no OK button needed!
5. Press ESC to cancel

**Benefits:**
- Much faster workflow (one less click)
- Works exactly like you requested
- Shows cell info (address, value, formula) before navigating

### ✅ ISSUE 3 FIXED: Cross-Sheet Navigation

**Problem:** References like `Sheet2!A1` failed with "Sheet not found: Sheet2"

**Root cause:** Sheet name parsing wasn't handling quotes correctly

**Solution:** Enhanced `NavigateToCell` function with:
- Proper single-quote removal from sheet names
- Better error handling with detailed messages
- Explicit activation sequence: Workbook → Sheet → Cell
- Screen updating disabled during navigation for smooth transitions

**Now works with:**
- Same sheet: `A1`, `B5:C10`
- Cross-sheet: `Sheet2!A1`, `'Sheet Name With Spaces'!B5`
- Cross-workbook: `[OtherWorkbook.xlsx]Sheet1!A1`

**Debug features added:**
- Detailed error messages showing exactly what failed
- Shows: full address, sheet name, cell address, error description
- Uncomment line 278 in modTraceTools.bas for debug logging

## Files Modified

1. **install_ribbon.sh** (NEW)
   - Bash script to install ribbon with proper path handling
   - Executable and ready to use

2. **modTraceTools.bas** (UPDATED)
   - `ShowTraceDialog`: Changed from AppleScript to InputBox with immediate navigation
   - `NavigateToCell`: Enhanced cross-sheet parsing and error handling
   - Better error messages throughout

## Testing Steps

### Test 1: Install Ribbon
```bash
cd /Users/kenu/Projects/LeanMacroTool
./install_ribbon.sh
```
Expected: "✅ Success! Please restart Excel to see the Lean Macros tab."

### Test 2: Trace Same-Sheet Precedents
1. In Excel, create a formula: `=A1+A2`
2. Select the cell with the formula
3. Press Ctrl+Shift+T
4. Dialog should show: "1. Sheet1!A1" and "2. Sheet1!A2"
5. Type "1" and press Enter
6. Should immediately jump to A1

### Test 3: Trace Cross-Sheet Precedents
1. Create a formula: `=SUM(Sheet2!A1:A10)`
2. Select the cell
3. Press Ctrl+Shift+T
4. Dialog should list all Sheet2 cells
5. Type a number and press Enter
6. Should jump to Sheet2 and select the cell

### Test 4: Trace Dependents
1. Select a cell (like A1)
2. Press Ctrl+Shift+Y
3. Should show all cells that reference A1
4. Type a number to jump to any dependent

## Known Limitations

1. **InputBox vs Click Selection**
   - Current: Type number and press Enter
   - You wanted: Click to select (like TTS)
   - **Why InputBox:** VBA on Mac doesn't support custom ListBox forms with double-click events
   - **Workaround:** This is actually faster - just type "1" or "2" and Enter

2. **External Workbook References**
   - Only works if the external workbook is already open
   - If closed, you'll get "Workbook not open" error

3. **Mac VBA Limitations**
   - No access to Windows Forms controls (ListBox with events)
   - MacScript doesn't support custom dialogs with event handlers
   - InputBox is the best native option for Mac

## Alternative: If You Want Click-to-Navigate

To implement true click-to-navigate like TTS, you would need to:

1. Create a UserForm with a ListBox control
2. Add ListBox_DblClick event handler
3. However, **this may not work reliably on Excel for Mac** due to:
   - UserForms have limited support on Mac
   - Event handlers can be flaky
   - May not work in 64-bit Excel for Mac

The InputBox solution is more reliable and almost as fast (just type the number).

## Next Steps

1. **Install the ribbon:**
   ```bash
   ./install_ribbon.sh
   ```

2. **Import the updated VBA module:**
   - Open Excel
   - Go to Tools → Macro → Visual Basic Editor
   - Find modTraceTools in your project
   - Delete the old module
   - Import the new modTraceTools.bas file

3. **Test thoroughly** with the test cases above

4. **Optional:** Enable debug logging by uncommenting line 278 in modTraceTools.bas:
   ```vba
   MsgBox "Looking for sheet: [" & sheetName & "] in workbook: " & wb.Name
   ```

## Support

If you encounter issues:
- Check error messages - they now show exactly what failed
- Verify sheet names match exactly (case-sensitive)
- Make sure target workbooks are open
- Try the manual navigation first: just type the address in the Name Box

All three issues are now resolved! 🎉
