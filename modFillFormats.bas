Attribute VB_Name = "modFillFormats"
Option Explicit

' ================================================================
' MODULE: modFillFormats
' Purpose: Cycle through cell fill patterns and borders
' ================================================================

' Module-level variables to track original format and position in cycle
Private originalCellAddress As String
Private originalFillColor As Long
Private originalPattern As Integer
Private originalPatternColor As Long
Private originalBorders As Boolean
Private lastAppliedIndex As Integer
Private lastAppliedAddress As String

' Ribbon callback wrapper - called by ribbon buttons
Public Sub CycleFillFormats(Optional control As IRibbonControl = Nothing)
    CycleFillImpl
End Sub

' Keyboard shortcut wrapper - called by Application.OnKey
Public Sub CycleFillKeyboard()
    CycleFillImpl
End Sub

' Implementation: Cycle through fill formats with original format tracking
Private Sub CycleFillImpl()
    On Error GoTo ErrorHandler

    Dim targetRange As Range
    Dim cellAddr As String
    Dim nextIndex As Integer

    ' Get current selection
    Set targetRange = Selection
    If targetRange Is Nothing Then Exit Sub

    ' Track the cell address to detect if user changed cells
    cellAddr = targetRange.Cells(1, 1).Address(External:=True)

    ' If this is a different cell, store its current format as the new original and reset cycle
    If cellAddr <> originalCellAddress Then
        originalCellAddress = cellAddr
        ' Store original fill properties
        With targetRange.Cells(1, 1).Interior
            originalFillColor = .Color
            originalPattern = .Pattern
            originalPatternColor = .PatternColor
        End With
        ' Store original border state (check if any borders exist)
        originalBorders = HasBorders(targetRange.Cells(1, 1))
        lastAppliedIndex = -1  ' Reset cycle tracking
        lastAppliedAddress = ""
    End If

    ' IMPORTANT: original format is now locked for this cell until we move to a different cell

    ' Use index tracking for reliable cycling
    If cellAddr = lastAppliedAddress And lastAppliedIndex >= 1 Then
        ' We're on the same cell where we last applied a format - use index tracking
        nextIndex = lastAppliedIndex + 1
        If nextIndex > 3 Then
            ' Wrap back to first format (Color + Border)
            nextIndex = 1
        End If
    Else
        ' First cycle on this cell, or format was changed externally
        ' Determine current state and set next
        nextIndex = 1  ' Default to first format
    End If

    ' Apply the next format
    Select Case nextIndex
        Case 1
            ' Format 1: Beige fill + Outline border
            ApplyColorAndBorder targetRange
        Case 2
            ' Format 2: Dotted pattern fill (no color)
            ApplyPattern targetRange
        Case 3
            ' Format 3: Original format
            RestoreOriginal targetRange
    End Select

    ' Track this application for next cycle
    lastAppliedIndex = nextIndex
    lastAppliedAddress = cellAddr

    Exit Sub

ErrorHandler:
    MsgBox "Error cycling fill formats: " & Err.Description, vbCritical, "Error"
End Sub

' Apply beige fill color with outline border
Private Sub ApplyColorAndBorder(targetRange As Range)
    ' Apply beige/cream background color (RGB 255, 242, 204)
    With targetRange.Interior
        .Pattern = xlSolid
        .Color = RGB(255, 242, 204)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    ' Apply borders to all cells at once (much faster)
    With targetRange.Borders
        .LineStyle = xlDot
        .Weight = xlHairline
        .Color = RGB(183, 183, 183)  ' Grey color
    End With
End Sub

' Apply dotted pattern fill (no background color)
Private Sub ApplyPattern(targetRange As Range)
    ' Apply 16% gray dotted pattern with no background color
    ' Order matters: ColorIndex first, then Pattern, then PatternColor
    With targetRange.Interior
        .ColorIndex = RGB(0, 0, 0)  ' Set background first
        .Pattern = xlGray16  ' 16% gray dotted pattern
        .PatternColor = RGB(183, 183, 183)  ' Grey pattern color
    End With

    ' Clear borders
    targetRange.Borders.LineStyle = xlNone
End Sub

' Restore original format
Private Sub RestoreOriginal(targetRange As Range)
    ' Restore original fill
    With targetRange.Interior
        .Pattern = originalPattern
        If originalPattern <> xlNone Then
            .Color = originalFillColor
            .PatternColor = originalPatternColor
        End If
    End With

    ' Restore original borders (or clear if none)
    If Not originalBorders Then
        targetRange.Borders.LineStyle = xlNone
    End If
    ' Note: We don't restore exact border configuration, just clear them
    ' Full border restoration would require saving all 4 border properties
End Sub

' Helper function to check if cell has any borders
Private Function HasBorders(cell As Range) As Boolean
    Dim hasBorder As Boolean
    hasBorder = False

    On Error Resume Next
    If cell.Borders(xlEdgeLeft).LineStyle <> xlNone Then hasBorder = True
    If cell.Borders(xlEdgeRight).LineStyle <> xlNone Then hasBorder = True
    If cell.Borders(xlEdgeTop).LineStyle <> xlNone Then hasBorder = True
    If cell.Borders(xlEdgeBottom).LineStyle <> xlNone Then hasBorder = True
    On Error GoTo 0

    HasBorders = hasBorder
End Function

' ================================================================
