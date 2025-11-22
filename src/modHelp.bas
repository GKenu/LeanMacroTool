Attribute VB_Name = "modHelp"
Option Explicit

' Lean Macro Tools - Help Module
' Displays quick guide for all features

' ============================================================================
' PUBLIC INTERFACE - Call from ribbon button
' ============================================================================

Public Sub ShowHelp(Optional control As Object = Nothing)
    ' Display help dialog with feature overview
    ShowHelpImpl
End Sub

' ============================================================================
' IMPLEMENTATION
' ============================================================================

Private Sub ShowHelpImpl()
    Dim helpText As String

    helpText = "Lean Macro Tools - Quick Guide" & vbCrLf & vbCrLf & _
               "FORMATTING:" & vbCrLf & _
               "• Cycle Formats (Ctrl+Shift+N)" & vbCrLf & _
               "  Change number formats (thousands, %, multiples, currency)" & vbCrLf & vbCrLf & _
               "• Cycle Colors (Ctrl+Shift+V)" & vbCrLf & _
               "  Change font colors (blue, green, red, grey, black)" & vbCrLf & vbCrLf & _
               "• Cycle Fill (Ctrl+Shift+B)" & vbCrLf & _
               "  Change fill patterns and borders" & vbCrLf & vbCrLf & _
               "TRACING:" & vbCrLf & _
               "• Trace Precedents (Ctrl+Shift+T)" & vbCrLf & _
               "  Show formula inputs (cells feeding into formulas)" & vbCrLf & vbCrLf & _
               "• Trace Dependents (Ctrl+Shift+Y)" & vbCrLf & _
               "  Show what uses this cell (cells referencing current cell)" & vbCrLf & vbCrLf & _
               "EXPORT:" & vbCrLf & _
               "• Export to PDF" & vbCrLf & _
               "  One-click PDF export with auto-range detection" & vbCrLf & _
               "  Perfect for sharing or uploading to AI tools!" & vbCrLf & vbCrLf & _
               "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & _
               "Find all features in the ""Lean Macros"" ribbon tab!" & vbCrLf & _
               "Tip: Use keyboard shortcuts for faster workflow!"

    MsgBox helpText, vbInformation, "Lean Macro Tools v2.1.0"
End Sub
