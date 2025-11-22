# Development Backlog

This file tracks known issues and planned improvements for LeanMacroTools.

---

## Known Issues

### High Priority

1. **Tracer Panel - Improvements**
   - **Issue:** Show in order of formula instead of order of cells with smaller number to larger + Show range when it is a range instead of individual cells
   - **Impact:** Much better user experience
   - **Status:** Needs investigation on how to implement this feature
   - **Workaround:** User as it is. It works.

2. **Dependent Tracer - Incomplete Cell Detection**
   - **Issue:** The dependent tracer doesn't always detect and display all dependent cells
   - **Impact:** Users may miss some cells that reference the selected cell
   - **Status:** Investigating root cause in `GetDependents()` function
   - **Workaround:** Cross-verify with Excel's built-in "Trace Dependents" feature

3. **Tracer Panel - Selection Conflict**
   - **Issue:** Selecting cells that are shown in the tracer panel list causes unexpected behavior
   - **Impact:** Panel may jump or refresh incorrectly when navigating to listed cells
   - **Status:** Needs investigation - likely related to WithEvents triggering on selection change
   - **Workaround:** Avoid manually selecting cells that appear in the panel list; use the panel's click/arrow navigation instead

---

## Planned Improvements

### Tracer Enhancement

- **Allow editing cells while tracer panel is open**
  - **Goal:** Enable users to modify cells that appear in the tracer panel without causing conflicts
  - **Technical Challenge:** Need to prevent selection event loops between Excel and UserForm
  - **Approach:**
    - Add flag to temporarily disable auto-navigation when user manually selects a listed cell
    - Detect when user is editing vs. navigating
    - Possibly refresh panel list after cell value changes

### UX Improvements

- **Change tracer panel caption from "TTS Turbo" to "Lean Macro Tools"**
  - **Requirement:** Need Windows Excel to edit UserForm .frx binary files
  - **Current:** Caption reads "TTS Turbo Precedent Tracer" / "Lean Macro Dependant Tracer"
  - **Desired:** Both should say "Lean Macro Tools - Precedent Tracer" / "Dependent Tracer"

---

## Future Features

*(Items to consider for future releases)*

- Cross-workbook dependent tracing (currently only works within same workbook)
- Option to export precedent/dependent lists to worksheet
- Recursive tracing (trace precedents of precedents)
- Visual formula tree diagram
- Support for named ranges in tracer

---

## Contributing

If you'd like to help tackle any of these issues, please:

1. Check if there's already an open GitHub issue for the item
2. If not, create a new issue referencing this backlog item
3. Fork the repo and create a feature branch
4. Submit a PR with your fix/improvement

---

**Last Updated:** 2025-01-21 (v2.0.0)
