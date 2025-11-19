# Changelog

All notable changes to LeanMacroTool will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

Current version: **1.0.3**
