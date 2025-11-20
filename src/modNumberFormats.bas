Attribute VB_Name = "modNumberFormats"
Option Explicit

' ================================================================
' MODULE: modNumberFormats
' Purpose: Cycle through custom number formats
' ================================================================

' Module-level variables to track original format and position in cycle
Private originalCellAddress As String
Private originalCellFormat As String
Private lastAppliedIndex As Integer
Private lastAppliedAddress As String

' Ribbon callback - called when ribbon loads
Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    ' Ribbon loaded successfully
End Sub

' Ribbon callback wrapper - called by ribbon buttons
Public Sub CycleCustomNumberFormats(Optional control As IRibbonControl = Nothing)
    CycleNumberFormatsImpl
End Sub

' Keyboard shortcut wrapper - called by Application.OnKey
Public Sub CycleFormatsKeyboard()
    CycleNumberFormatsImpl
End Sub

' Implementation: Cycle through number formats with original format tracking
Private Sub CycleNumberFormatsImpl()
    On Error GoTo ErrorHandler

    Dim formats() As String
    Dim formatEnabled() As Boolean
    Dim enabledFormats() As String
    Dim enabledCount As Integer
    Dim i As Integer, j As Integer
    Dim currentFormat As String
    Dim nextIndex As Integer
    Dim targetRange As Range
    Dim cellAddr As String

    ' Get current selection
    Set targetRange = Selection
    If targetRange Is Nothing Then Exit Sub

    ' Track the cell address to detect if user changed cells
    cellAddr = targetRange.Cells(1, 1).Address(External:=True)

    ' Get current format of active cell
    currentFormat = targetRange.Cells(1, 1).NumberFormat

    ' If this is a different cell, store its current format as the new original and reset cycle
    If cellAddr <> originalCellAddress Then
        originalCellAddress = cellAddr
        originalCellFormat = currentFormat
        lastAppliedIndex = -1  ' Reset cycle tracking
        lastAppliedAddress = ""
    End If

    ' IMPORTANT: originalCellFormat is now locked for this cell until we move to a different cell

    ' Load formats from configuration
    LoadFormats formats, formatEnabled

    ' Build array of enabled formats, INCLUDING original format
    enabledCount = 0
    For i = LBound(formatEnabled) To UBound(formatEnabled)
        If formatEnabled(i) Then enabledCount = enabledCount + 1
    Next i

    If enabledCount = 0 Then
        MsgBox "No number formats are enabled. Please check format configuration.", vbExclamation, "No Formats"
        Exit Sub
    End If

    ' Build array with 1-based indexing (1 to enabledCount+1)
    ' Index 1 = original format, Indices 2 onwards = configured formats
    ReDim enabledFormats(1 To enabledCount + 1)

    ' Add original format as index 1
    enabledFormats(1) = originalCellFormat

    ' Add configured formats starting at index 2
    j = 2
    For i = LBound(formats) To UBound(formats)
        If formatEnabled(i) Then
            enabledFormats(j) = formats(i)
            j = j + 1
        End If
    Next i

    ' Use index tracking for reliable cycling (avoids format string comparison issues)
    If cellAddr = lastAppliedAddress And lastAppliedIndex >= 1 Then
        ' We're on the same cell where we last applied a format - use index tracking
        nextIndex = lastAppliedIndex + 1
        If nextIndex > UBound(enabledFormats) Then
            ' Wrap back to original format (index 1)
            nextIndex = 1
        End If
    Else
        ' First cycle on this cell, or format was changed externally
        ' Try to find current format in the list
        nextIndex = -1
        For i = LBound(enabledFormats) To UBound(enabledFormats)
            If currentFormat = enabledFormats(i) Then
                nextIndex = i + 1
                If nextIndex > UBound(enabledFormats) Then
                    nextIndex = 1
                End If
                Exit For
            End If
        Next i

        ' If not found, start from first configured format (index 2)
        If nextIndex = -1 Then nextIndex = 2
    End If

    ' Apply the next format to entire selection
    targetRange.NumberFormat = enabledFormats(nextIndex)

    ' Track this application for next cycle
    lastAppliedIndex = nextIndex
    lastAppliedAddress = cellAddr

    Exit Sub

ErrorHandler:
    MsgBox "Error cycling number formats: " & Err.Description, vbCritical, "Error"
End Sub

' Load hardcoded number formats
Public Sub LoadFormats(ByRef formats() As String, ByRef enabled() As Boolean)
    ' Define all formats in an array (automatically calculates count)
    Dim allFormats As Variant
    allFormats = Array( _
        "#,##0.00_);(#,##0.00);""-""_);@_)", _
        "0.0%_);(0.0%);""-""_);@_)", _
        "#,##0.0x_);(#,##0.0)x;""-""_);@_)", _
        "[$R$-416]#,##0.0_);([$R$-416]#,##0.0);""-""_);@_)", _
        "[$$-409]#,##0.0_);([$$-409]#,##0.0);""-""_);@_)", _
        "dd-mmm-yy_)", _
        "mmm-yy_)", _
        "General_)" _
    )

    ' Calculate format count dynamically
    Dim formatCount As Integer
    formatCount = UBound(allFormats) - LBound(allFormats) + 1

    ' Size output arrays based on actual format count
    ReDim formats(1 To formatCount)
    ReDim enabled(1 To formatCount)

    ' Copy formats and enable all
    Dim i As Integer
    For i = 1 To formatCount
        formats(i) = allFormats(i - 1)  ' Variant array starts at 0
        enabled(i) = True
    Next i
End Sub

' ================================================================
