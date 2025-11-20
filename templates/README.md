# Template Files

This folder contains template files used during the build process.

## LeanMacroTools_template.xlam

This is the master template for the LeanMacroTools add-in with all VBA modules imported.

### How to Create/Update the Template:

1. **Open Excel** and create a new blank workbook
2. **Press Option+F11** to open VBA Editor
3. **Import all VBA modules** (File > Import File...):
   - `src/modNumberFormats.bas`
   - `src/modColorFormats.bas`
   - `src/modFillFormats.bas`
   - `src/modTraceTools.bas`

4. **Double-click ThisWorkbook** in the VBA Editor and add this code:
```vba
Private Sub Workbook_Open()
    Application.OnKey "^+N", "CycleFormatsKeyboard"
    Application.OnKey "^+V", "CycleColorsKeyboard"
    Application.OnKey "^+B", "CycleFillKeyboard"
    Application.OnKey "^+T", "TracePrecedentsKeyboard"
    Application.OnKey "^+Y", "TraceDependentsKeyboard"
End Sub
```

5. **Save as .xlam** (File > Save As):
   - Format: **Excel Macro-Enabled Add-In (.xlam)**
   - Name: `LeanMacroTools_template.xlam`
   - Location: This `templates/` folder

### When to Update the Template:

Update the template whenever you:
- Add a new VBA module
- Change the keyboard shortcuts in ThisWorkbook
- Add new functionality that requires VBA structure changes

**Note:** If you're only editing existing code within modules, you can edit directly in your working .xlam file in the Add-ins folder. The template only needs updating when the structure changes.

### Build Process:

The `scripts/build_release.sh` script uses this template to:
1. Copy the template to the Add-ins folder
2. Inject the ribbon UI
3. Package for distribution

This ensures a clean, consistent build every time.
