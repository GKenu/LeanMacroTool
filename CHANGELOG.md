# Changelog

All notable changes to LeanMacroTool will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] - 2025-01-21

### Added
- **UserForm-based Precedent/Dependent Tracers** - Major UX overhaul replacing InputBox dialogs
- Interactive panel with professional list-based interface
- Click or use arrow keys to navigate through cells instantly
- Real-time cell preview showing address, value, and formula
- Auto-navigation on selection change (no need to press Enter)
- Modeless panels allow working in Excel while tracer is open
- `modLeanTracer.bas` - New adapter module bridging GetPrecedents/GetDependents with UserForms
- `FPrecedentAnalyzer.frm/.frx` - Precedent tracer UserForm
- `zFPrecedentAnalyzer.frm/.frx` - Dependent tracer UserForm
- Support for keyboard shortcuts within forms (Ctrl+Shift+Home to switch to Excel)
- Form state tracking with module-level instances (persistent across calls)

### Changed
- **BREAKING:** Replaced InputBox-based trace dialogs with professional UserForm interface
- **Major UX Improvement:** No more number-based navigation - click cells directly in list
- **Major UX Improvement:** Forms stay positioned and visible while navigating
- Precedent/Dependent tracers now show all cells in scrollable list format
- Original cell marked with "original cell" label in list
- Keyboard shortcuts (Ctrl+Shift+T/Y) now open UserForms instead of InputBox dialogs
- Ribbon buttons now call `ShowLeanPrecedents`/`ShowLeanDependents` directly (no wrapper functions)
- Consolidated tracing logic into single `modLeanTracer.bas` module
- `GetPrecedents()` and `GetDependents()` moved from modTraceTools to modLeanTracer

### Removed
- **BREAKING:** Old InputBox-based trace dialogs removed
- `TracePrecedentsDialog`/`TraceDependentsDialog` wrapper functions removed
- `TracePrecedentsImpl`/`TraceDependentsImpl` implementation functions removed
- `ShowTraceDialog` function removed
- `modTraceTools.bas` module deprecated (functionality moved to modLeanTracer)

### Fixed
- Type mismatch bug when converting string addresses to Range objects
- Address formatting for cells with special characters in sheet names
- Detection logic for dependent vs precedent mode (using flag instead of form visibility)
- GetFullAddress now removes workbook brackets correctly so Range() can parse addresses

### Technical Details
- **Hybrid approach:** TTS UserForm UI + our simpler GetPrecedents/GetDependents logic
- UserForms require Windows Excel to create/modify, but work on Mac once imported
- Module-level flag (`mbDependentMode`) determines whether to show precedents or dependents
- Forms call `NewPrecedents()` which adapts Collection output to 2D array format
- Address format: Column 0 = full address (for navigation), Column 1 = short address (for display)
- `bOriginalCellAtEnd = False` configuration places original cell at index 0
- Forms use WithEvents for Worksheet/Workbook to update display when switching contexts
- Binary .frx files control visual layout and cannot be regenerated on Mac

### Attribution
- UserForm interface adapted from TTS Turbo Macros with permission
- Original TTS forms: FPrecedentAnalyzer by TTS Turbo team
- Modified caption and integrated with LeanMacroTools tracing engine

## [1.0.7] - 2025-01-20

### Added
- **Simplified Installation System** - Reduced installation from 6 complex steps to 2 simple steps
- `install.command` - Double-click installer script with automatic Add-ins folder detection
- `build_release.sh` - Automated release package builder for developers
- **Template-based build system** - Uses `templates/LeanMacroTools_template.xlam` for consistent builds
- `templates/` folder with template .xlam and documentation
- Auto-detection of Add-ins folder (handles both English and localized folder names)
- Pre-built distribution packages with ribbon UI already embedded
- Distribution folder structure with all necessary files for users
- "For Developers" section in README with build workflow documentation

### Changed
- **Major UX Improvement:** Users no longer need Python, VBA Editor, or manual ribbon injection
- **Major Dev Improvement:** Template-based builds - just run `./scripts/build_release.sh`
- Installation process now: Download → Double-click install.command → Enable in Excel
- README.md restructured with "Quick Install" as primary method
- Manual installation moved to collapsible "Advanced" section
- System Requirements updated to clarify Python not needed for users
- Files list now separated into "For Users" and "For Developers" sections
- Updated .gitignore to exclude dist/ folder and .zip files
- Repository reorganized: src/, ribbon/, scripts/, templates/ folders

### Technical Details
- **Template system:** Pre-built .xlam with VBA modules → copy → inject ribbon → package
- install.command uses same proven path detection as install_ribbon.sh
- Tries multiple Add-ins path variations for localized macOS systems
- build_release.sh automates: copy template → inject ribbon → create package → zip
- Distribution package includes: .xlam, install.command, README, LICENSE, CHANGELOG
- Ribbon UI is embedded during release build, not during user installation
- Maintains backward compatibility with manual installation method
- Repository structure reorganized for clarity (src/, ribbon/, scripts/, templates/)

## [1.0.6] - 2025-01-20

### Added
- **Fill Pattern Cycling Feature** - New fill/border cycling functionality with Ctrl+Shift+B keyboard shortcut
- `modFillFormats.bas` module for fill pattern and border management
- Cycle through formats: Color+Border → Pattern → Original (3 states)
- First format: Beige background (RGB 255, 242, 204) with outline border
- Second format: Fine dots pattern (xlGray8 - 6.25% pattern)
- Original format tracking and restoration (per-cell memory)
- Cycle Fill button in ribbon UI (Number Formatting group)
- Support for both ribbon button and keyboard shortcut (Ctrl+Shift+B)
- `CycleFillKeyboard()` wrapper function for keyboard shortcut compatibility

### Changed
- Updated ribbon UI to include Cycle Fill button with CellFillColorPicker icon
- Updated README.md with fill cycling feature and usage instructions
- Updated ThisWorkbook_KeyboardShortcuts.txt with Ctrl+Shift+B shortcut
- Installation instructions now include modFillFormats.bas import step
- Version references updated from v1.0.5 to v1.0.6

### Technical Details
- Fill cycling uses same architecture pattern as number/color format cycling
- Module-level variables track original cell address, fill color, pattern, and border state
- Uses `Selection.Interior` property for fill and `BorderAround` for borders
- Stores original fill properties (Color, Pattern, PatternColor) for restoration
- Works on entire selection (multi-cell support)
- Resets cycle when moving to different cell
- `HasBorders()` helper function checks if cell has any borders

## [1.0.5] - 2025-01-20

### Added
- **Font Color Cycling Feature** - New color cycling functionality with Ctrl+Shift+V keyboard shortcut
- `modColorFormats.bas` module for font color management
- Cycle through preset colors: Blue → Green → Red → Grey → Black → Original (6 colors)
- Original font color tracking and restoration (per-cell memory)
- Cycle Colors button in ribbon UI (Number Formatting group)
- Support for both ribbon button and keyboard shortcut (Ctrl+Shift+V)
- `CycleColorsKeyboard()` wrapper function for keyboard shortcut compatibility
- `ThisWorkbook_KeyboardShortcuts.txt` documentation file

### Changed
- Updated ribbon UI to include Cycle Colors button with FontColor icon
- Enhanced keyboard shortcut registration documentation
- Updated README.md with color cycling feature and usage instructions
- Installation instructions now include modColorFormats.bas import step

### Technical Details
- Color cycling uses same architecture pattern as number format cycling
- Module-level variables track original cell address and font color
- Uses `Selection.Font.Color` property with RGB values
- Index-based cycling ensures reliable color transitions
- Works on entire selection (multi-cell support)
- Resets cycle when moving to different cell

## [1.0.4] - 2025-01-18

### Added
- Original format tracking: cycling now returns to cell's original format
- Keyboard wrapper functions for compatibility with both ribbon and shortcuts
- Support for `IRibbonControl` parameter in all ribbon callbacks

### Changed
- Format cycling now includes original format at index 0
- Cycle order: **Original** → Thousands → Percentage → Multiples → USD → BRL → **Original**
- Ribbon callbacks now use proper Office 2007+ signature: `Sub Name(control As IRibbonControl)`
- Workbook_Open now calls wrapper functions: `CycleFormatsKeyboard`, `TracePrecedentsKeyboard`, `TraceDependentsKeyboard`

### Fixed
- **CRITICAL:** Ribbon buttons now work! Fixed "Wrong number of arguments" error
- **CRITICAL:** Configure button in ribbon now functional
- **CRITICAL:** Format cycling now returns to original format instead of getting stuck
- All ribbon callbacks properly accept IRibbonControl parameter

### Technical Details
- Split public functions into ribbon callbacks (with IRibbonControl) and implementation functions
- Module-level variables track original cell address and format
- Backward compatible: keyboard shortcuts still work through wrapper functions
- Architecture supports both ribbon button clicks and keyboard shortcuts

## [1.0.3] - 2025-01-18

### Added
- Interactive keyboard navigation for trace precedents/dependents dialogs
- Current cell is now included in trace list (index 0)
- Modeless navigation: dialog stays open for exploring multiple cells
- Navigate automatically through cell list using +/- or n/p keys
- Visual indicator (<--) shows current position in list
- Direct jump to any cell by typing its index number
- `install_ribbon.sh` automated installer script for macOS
- Formula parsing fallback for cross-sheet reference detection
- Support for localized folder names (e.g., `Add-Ins.localized`)

### Changed
- Trace dialog now uses persistent InputBox loop instead of single prompt
- Dialog displays full formula and origin cell information
- Enhanced cross-sheet navigation with better error handling
- Updated navigation commands: +/n (next), -/p (previous), ESC (close)
- Replaced Unicode emoji with ASCII characters for better compatibility
- Dialog formatting now displays correctly without garbled characters
- `GetPrecedents()` now parses formula text when DirectPrecedents fails

### Fixed
- **CRITICAL:** Cross-sheet precedents now detected correctly on Mac Excel
- **CRITICAL:** Dialog formatting fixed - removed emoji characters causing "ΔΔΔΔΔΔ" display
- **CRITICAL:** Install script now detects localized Add-ins paths
- macOS path handling in installation script (spaces in path now handled)
- Single quote parsing in sheet names
- Improved error messages show exactly which sheet/cell failed
- Screen updating disabled during navigation for smoother transitions
- Formula references like `=SUM(Sheet2!A1:A10)` now properly detected

## [1.0.2] - 2025-01-17

### Added
- Initial public release with ribbon UI
- Cross-sheet reference support for trace tools
- MIT License

### Changed
- Reorganized repository structure (moved files to root)
- Removed archive directory

## [1.0.1] - 2025-01-16

### Added
- Custom Ribbon UI integration (Lean Macros tab)
- Ribbon buttons for all features
- Keyboard shortcuts: Ctrl+Shift+N, Ctrl+Shift+T, Ctrl+Shift+Y
- `inject_ribbon.py` script for ribbon installation

### Changed
- Improved installation process with ribbon support

## [1.0.0] - 2025-01-15

### Added
- Number format cycling feature (Ctrl+Shift+N)
- Configurable number format list
- Support for Thousands, Percentage, Multiples, USD, BRL formats
- Trace precedents feature (Ctrl+Shift+T)
- Trace dependents feature (Ctrl+Shift+Y)
- Basic dialog for selecting precedent/dependent cells
- VBA modules: modNumberFormats.bas, modTraceTools.bas
- README.md with installation instructions

### Features
- Excel for Mac compatibility (16.x)
- Cross-workbook reference support
- Error handling and user-friendly messages

---

## Release Notes

### How to Use This Changelog

- **Added** - New features
- **Changed** - Changes to existing functionality
- **Deprecated** - Features that will be removed in future versions
- **Removed** - Features that have been removed
- **Fixed** - Bug fixes
- **Security** - Security vulnerability fixes

### Version Numbering

We use [Semantic Versioning](https://semver.org/):
- **MAJOR** version: Incompatible API changes
- **MINOR** version: New functionality (backwards-compatible)
- **PATCH** version: Bug fixes (backwards-compatible)

Current version: **1.0.7**
