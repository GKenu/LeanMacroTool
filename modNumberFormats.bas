Attribute VB_Name = "modNumberFormats"
Option Explicit

' ================================================================
' MODULE: modNumberFormats
' Purpose: Cycle through custom number formats
' ================================================================

Private Const CONFIG_SHEET_NAME As String = "NumberFormatConfig"

' Module-level variables to track original format
Private originalCellAddress As String
Private originalCellFormat As String

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

    ' If this is a different cell or first time, store original format
    If cellAddr <> originalCellAddress Then
        originalCellAddress = cellAddr
        originalCellFormat = currentFormat
    End If

    ' Load formats from configuration
    LoadFormats formats, formatEnabled
    
    ' Build array of enabled formats, INCLUDING original format
    enabledCount = 0
    For i = LBound(formatEnabled) To UBound(formatEnabled)
        If formatEnabled(i) Then enabledCount = enabledCount + 1
    Next i

    If enabledCount = 0 Then
        MsgBox "No number formats are enabled. Run ConfigureNumberFormats to set up.", vbExclamation, "No Formats"
        Exit Sub
    End If

    ' Add space for original format at beginning of array
    ReDim enabledFormats(0 To enabledCount)

    ' Add original format as index 0
    enabledFormats(0) = originalCellFormat

    ' Add configured formats starting at index 1
    j = 1
    For i = LBound(formats) To UBound(formats)
        If formatEnabled(i) Then
            enabledFormats(j) = formats(i)
            j = j + 1
        End If
    Next i

    ' Find current format in enabled list and move to next
    nextIndex = 1 ' Default to first configured format if not found
    For i = LBound(enabledFormats) To UBound(enabledFormats)
        If currentFormat = enabledFormats(i) Then
            nextIndex = i + 1
            If nextIndex > UBound(enabledFormats) Then
                ' Wrap back to original format (index 0)
                nextIndex = 0
            End If
            Exit For
        End If
    Next i

    ' Apply the next format to entire selection
    targetRange.NumberFormat = enabledFormats(nextIndex)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error cycling number formats: " & Err.Description, vbCritical, "Error"
End Sub

' Ribbon callback wrapper for Configure button
Public Sub ConfigureNumberFormats(Optional control As IRibbonControl = Nothing)
    ConfigureNumberFormatsImpl
End Sub

' Keyboard shortcut wrapper for Configure
Public Sub ConfigureFormatsKeyboard()
    ConfigureNumberFormatsImpl
End Sub

' Implementation: Configure number formats (simplified version without UserForm)
Private Sub ConfigureNumberFormatsImpl()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim response As VbMsgBoxResult
    Dim msg As String

    ' Get or create config sheet
    Set ws = GetOrCreateConfigSheet()

    ' Make it visible temporarily for editing
    ws.Visible = xlSheetVisible
    ws.Activate

    ' Show instructions
    msg = "The NumberFormatConfig sheet is now visible." & vbCrLf & vbCrLf & _
          "Column A: Number format codes" & vbCrLf & _
          "Column B: TRUE to enable, FALSE to disable" & vbCrLf & vbCrLf & _
          "Edit the formats as needed, then click OK to save and hide the sheet."

    response = MsgBox(msg, vbOKCancel + vbInformation, "Configure Number Formats")

    If response = vbOK Then
        ' Hide the sheet again
        ws.Visible = xlSheetVeryHidden
        MsgBox "Configuration saved!", vbInformation, "Success"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error configuring formats: " & Err.Description, vbCritical, "Error"
End Sub

' Load formats from hidden config sheet
Public Sub LoadFormats(ByRef formats() As String, ByRef enabled() As Boolean)
    On Error GoTo UseDefaults
    
    Dim ws As Worksheet
    Dim i As Integer
    Dim lastRow As Long
    
    Set ws = GetOrCreateConfigSheet()
    
    ' Read formats from sheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then GoTo UseDefaults
    
    ReDim formats(1 To lastRow - 1)
    ReDim enabled(1 To lastRow - 1)
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            formats(i - 1) = ws.Cells(i, 1).Value
            enabled(i - 1) = (UCase(Trim(CStr(ws.Cells(i, 2).Value))) = "TRUE")
        End If
    Next i
    
    Exit Sub
    
UseDefaults:
    ' Return default formats if error or no config exists
    ReDim formats(1 To 5)
    ReDim enabled(1 To 5)
    
    ' Format 1: Thousands with 2 decimals
    formats(1) = "#,##0.00_);(#,##0.00);""-""_);@_)"
    enabled(1) = True
    
    ' Format 2: Percentage
    formats(2) = "0.0%_);(0.0%);""-""_);@_)"
    enabled(2) = True
    
    ' Format 3: Multiple (2.5x)
    formats(3) = "#,##0.0x_);(#,##0.0)x;""-""_);@_)"
    enabled(3) = True
    
    ' Format 4: USD
    formats(4) = "$#,##0.0_);$(#,##0.0);""-""_);@_)"
    enabled(4) = True
    
    ' Format 5: Brazilian Real
    formats(5) = "R$#,##0.0_);R$(#,##0.0);""-""_);@_)"
    enabled(5) = True
End Sub

' Get or create the configuration sheet
Private Function GetOrCreateConfigSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateConfigSheet = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    
    If GetOrCreateConfigSheet Is Nothing Then
        On Error GoTo 0
        ' Create new config sheet
        Set GetOrCreateConfigSheet = ThisWorkbook.Sheets.Add
        GetOrCreateConfigSheet.Name = CONFIG_SHEET_NAME
        
        ' Set up headers and default data
        With GetOrCreateConfigSheet
            .Cells(1, 1).Value = "Format"
            .Cells(1, 2).Value = "Enabled"
            .Cells(1, 1).Font.Bold = True
            .Cells(1, 2).Font.Bold = True
            
            ' Default formats
            .Cells(2, 1).Value = "#,##0.00_);(#,##0.00);""-""_);@_)"
            .Cells(2, 2).Value = "TRUE"
            
            .Cells(3, 1).Value = "0.0%_);(0.0%);""-""_);@_)"
            .Cells(3, 2).Value = "TRUE"
            
            .Cells(4, 1).Value = "#,##0.0x_);(#,##0.0)x;""-""_);@_)"
            .Cells(4, 2).Value = "TRUE"
            
            .Cells(5, 1).Value = "$#,##0.0_);$(#,##0.0);""-""_);@_)"
            .Cells(5, 2).Value = "TRUE"
            
            .Cells(6, 1).Value = "R$#,##0.0_);R$(#,##0.0);""-""_);@_)"
            .Cells(6, 2).Value = "TRUE"
            
            ' Format columns
            .Columns(1).ColumnWidth = 50
            .Columns(2).ColumnWidth = 12
        End With
        
        ' Hide the sheet
        GetOrCreateConfigSheet.Visible = xlSheetVeryHidden
    End If
End Function


' ================================================================
