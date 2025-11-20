Attribute VB_Name = "modColorFormats"
Option Explicit

' ================================================================
' MODULE: modColorFormats
' Purpose: Cycle through font colors
' ================================================================

' Module-level variables to track original color and position in cycle
Private originalCellAddress As String
Private originalCellColor As Long
Private lastAppliedIndex As Integer
Private lastAppliedAddress As String

' Ribbon callback wrapper - called by ribbon buttons
Public Sub CycleCustomColors(Optional control As IRibbonControl = Nothing)
    CycleColorsImpl
End Sub

' Keyboard shortcut wrapper - called by Application.OnKey
Public Sub CycleColorsKeyboard()
    CycleColorsImpl
End Sub

' Implementation: Cycle through font colors with original color tracking
Private Sub CycleColorsImpl()
    On Error GoTo ErrorHandler

    Dim colors() As Long
    Dim i As Integer
    Dim currentColor As Long
    Dim nextIndex As Integer
    Dim targetRange As Range
    Dim cellAddr As String

    ' Get current selection
    Set targetRange = Selection
    If targetRange Is Nothing Then Exit Sub

    ' Track the cell address to detect if user changed cells
    cellAddr = targetRange.Cells(1, 1).Address(External:=True)

    ' Get current font color of active cell
    currentColor = targetRange.Cells(1, 1).Font.Color

    ' If this is a different cell, store its current color as the new original and reset cycle
    If cellAddr <> originalCellAddress Then
        originalCellAddress = cellAddr
        originalCellColor = currentColor
        lastAppliedIndex = -1  ' Reset cycle tracking
        lastAppliedAddress = ""
    End If

    ' IMPORTANT: originalCellColor is now locked for this cell until we move to a different cell

    ' Load colors: Blue → Green → Red → Grey → Black → Original
    LoadColors colors, originalCellColor

    ' Use index tracking for reliable cycling
    If cellAddr = lastAppliedAddress And lastAppliedIndex >= 1 Then
        ' We're on the same cell where we last applied a color - use index tracking
        nextIndex = lastAppliedIndex + 1
        If nextIndex > UBound(colors) Then
            ' Wrap back to first color (Blue)
            nextIndex = 1
        End If
    Else
        ' First cycle on this cell, or color was changed externally
        ' Try to find current color in the list
        nextIndex = -1
        For i = LBound(colors) To UBound(colors)
            ' For black color, also check if current color is 0 (RGB black)
            If currentColor = colors(i) Or (i = 5 And currentColor = 0) Then
                nextIndex = i + 1
                If nextIndex > UBound(colors) Then
                    nextIndex = 1
                End If
                Exit For
            End If
        Next i

        ' If not found, start from first color (Blue)
        If nextIndex = -1 Then nextIndex = 1
    End If

    ' Skip "Original" if it's the same as the current color (avoid invisible cycle)
    If nextIndex = UBound(colors) And colors(nextIndex) = currentColor Then
        nextIndex = 1  ' Jump to Blue instead
    End If

    ' Apply the next color to entire selection
    targetRange.Font.Color = colors(nextIndex)

    ' Track this application for next cycle
    lastAppliedIndex = nextIndex
    lastAppliedAddress = cellAddr

    Exit Sub

ErrorHandler:
    MsgBox "Error cycling font colors: " & Err.Description, vbCritical, "Error"
End Sub

' Load color array: Blue → Green → Red → Grey → Black → Original
Private Sub LoadColors(ByRef colors() As Long, originalColor As Long)
    ' Build array with 1-based indexing (1 to 6)
    ' Cycle order: Blue → Green → Red → Grey → Black → Original
    ReDim colors(1 To 6)

    colors(1) = RGB(0, 0, 255)      ' Blue (#0000FF)
    colors(2) = RGB(0, 128, 0)      ' Green (#008000)
    colors(3) = RGB(255, 0, 0)      ' Red (#FF0000)
    colors(4) = RGB(128, 128, 128)  ' Grey (#808080)
    colors(5) = RGB(0, 0, 0)        ' Black (#000000)
    colors(6) = originalColor       ' Original color
End Sub

' ================================================================
