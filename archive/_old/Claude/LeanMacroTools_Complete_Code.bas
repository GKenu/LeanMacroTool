Attribute VB_Name = "modNumberFormats"
Option Explicit

' ================================================================
' MODULE: modNumberFormats
' Purpose: Cycle through custom number formats
' ================================================================

Private Const CONFIG_SHEET_NAME As String = "NumberFormatConfig"

' Main macro: Cycle through number formats
Public Sub CycleCustomNumberFormats()
    On Error GoTo ErrorHandler
    
    Dim formats() As String
    Dim formatEnabled() As Boolean
    Dim enabledFormats() As String
    Dim enabledCount As Integer
    Dim i As Integer, j As Integer
    Dim currentFormat As String
    Dim nextIndex As Integer
    Dim targetRange As Range
    
    ' Load formats from configuration
    LoadFormats formats, formatEnabled
    
    ' Build array of only enabled formats
    enabledCount = 0
    For i = LBound(formatEnabled) To UBound(formatEnabled)
        If formatEnabled(i) Then enabledCount = enabledCount + 1
    Next i
    
    If enabledCount = 0 Then
        MsgBox "No number formats are enabled. Run ConfigureNumberFormats to set up.", vbExclamation, "No Formats"
        Exit Sub
    End If
    
    ReDim enabledFormats(1 To enabledCount)
    j = 1
    For i = LBound(formats) To UBound(formats)
        If formatEnabled(i) Then
            enabledFormats(j) = formats(i)
            j = j + 1
        End If
    Next i
    
    ' Get current selection
    Set targetRange = Selection
    If targetRange Is Nothing Then Exit Sub
    
    ' Get current format of active cell
    currentFormat = targetRange.Cells(1, 1).NumberFormat
    
    ' Find current format in enabled list and move to next
    nextIndex = 1 ' Default to first format
    For i = LBound(enabledFormats) To UBound(enabledFormats)
        If currentFormat = enabledFormats(i) Then
            nextIndex = i + 1
            If nextIndex > UBound(enabledFormats) Then nextIndex = LBound(enabledFormats)
            Exit For
        End If
    Next i
    
    ' Apply the next format to entire selection
    targetRange.NumberFormat = enabledFormats(nextIndex)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error cycling number formats: " & Err.Description, vbCritical, "Error"
End Sub

' Configure number formats (simplified version without UserForm)
Public Sub ConfigureNumberFormats()
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
    ' Return default formats if error
    ReDim formats(1 To 5)
    ReDim enabled(1 To 5)
    formats(1) = "#,##0.00_);(#,##0.00);""-""_);@_)"
    formats(2) = "0.0%_);(0.0%);""-""_);@_)"
    formats(3) = "#,##0.0x_);(#,##0.0)x;""-""_);@_)"
    formats(4) = "$#,##0.0_);$(#,##0.0)""x"";""-""_);@_)"
    formats(5) = "R$#,##0.0_);R$(#,##0.0)""x"";""-""_);@_)"
    For i = 1 To 5
        enabled(i) = True
    Next i
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
            
            .Cells(5, 1).Value = "$#,##0.0_);$(#,##0.0)""x"";""-""_);@_)"
            .Cells(5, 2).Value = "TRUE"
            
            .Cells(6, 1).Value = "R$#,##0.0_);R$(#,##0.0)""x"";""-""_);@_)"
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
' MODULE: modTraceTools
' Purpose: Enhanced precedent and dependent tracing
' ================================================================

' Show precedents for active cell (simplified - no UserForm)
Public Sub TracePrecedentsDialog()
    On Error GoTo ErrorHandler
    
    Dim activeCell As Range
    Dim precedents As Collection
    Dim msg As String
    Dim item As Variant
    Dim response As String
    Dim i As Integer
    
    Set activeCell = Application.ActiveCell
    
    If activeCell Is Nothing Then
        MsgBox "No cell selected.", vbExclamation, "Trace Precedents"
        Exit Sub
    End If
    
    ' Check if cell has a formula
    If Not activeCell.HasFormula Then
        MsgBox "The selected cell does not contain a formula.", vbInformation, "Trace Precedents"
        Exit Sub
    End If
    
    ' Get precedents
    Set precedents = GetPrecedents(activeCell)
    
    If precedents.Count = 0 Then
        MsgBox "No precedent cells found for " & GetFullAddress(activeCell), vbInformation, "No Precedents"
        Exit Sub
    End If
    
    ' Build message showing precedents
    msg = "TRACE PRECEDENTS" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    msg = msg & "Origin: " & GetFullAddress(activeCell) & vbCrLf
    msg = msg & "Value: " & GetCellDisplayValue(activeCell) & vbCrLf
    msg = msg & "Formula: " & activeCell.Formula & vbCrLf & vbCrLf
    msg = msg & "Precedent Cells:" & vbCrLf
    
    i = 1
    For Each item In precedents
        msg = msg & "  " & i & ". " & item & vbCrLf
        i = i + 1
    Next item
    
    msg = msg & vbCrLf & "Enter cell number to jump to (or Cancel):"
    
    ' Ask user which cell to jump to
    response = InputBox(msg, "Trace Precedents", "")
    
    If response <> "" And IsNumeric(response) Then
        i = CInt(response)
        If i >= 1 And i <= precedents.Count Then
            NavigateToCell precedents(i)
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error tracing precedents: " & Err.Description, vbCritical, "Error"
End Sub

' Show dependents for active cell (simplified - no UserForm)
Public Sub TraceDependentsDialog()
    On Error GoTo ErrorHandler
    
    Dim activeCell As Range
    Dim dependents As Collection
    Dim msg As String
    Dim item As Variant
    Dim response As String
    Dim i As Integer
    
    Set activeCell = Application.ActiveCell
    
    If activeCell Is Nothing Then
        MsgBox "No cell selected.", vbExclamation, "Trace Dependents"
        Exit Sub
    End If
    
    ' Get dependents
    Set dependents = GetDependents(activeCell)
    
    If dependents.Count = 0 Then
        MsgBox "No dependent cells found for " & GetFullAddress(activeCell), vbInformation, "No Dependents"
        Exit Sub
    End If
    
    ' Build message showing dependents
    msg = "TRACE DEPENDENTS" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    msg = msg & "Origin: " & GetFullAddress(activeCell) & vbCrLf
    msg = msg & "Value: " & GetCellDisplayValue(activeCell) & vbCrLf
    
    If activeCell.HasFormula Then
        msg = msg & "Formula: " & activeCell.Formula & vbCrLf
    End If
    
    msg = msg & vbCrLf & "Dependent Cells:" & vbCrLf
    
    i = 1
    For Each item In dependents
        msg = msg & "  " & i & ". " & item & vbCrLf
        i = i + 1
    Next item
    
    msg = msg & vbCrLf & "Enter cell number to jump to (or Cancel):"
    
    ' Ask user which cell to jump to
    response = InputBox(msg, "Trace Dependents", "")
    
    If response <> "" And IsNumeric(response) Then
        i = CInt(response)
        If i >= 1 And i <= dependents.Count Then
            NavigateToCell dependents(i)
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error tracing dependents: " & Err.Description, vbCritical, "Error"
End Sub

' Get list of precedent cells
Public Function GetPrecedents(sourceCell As Range) As Collection
    On Error GoTo ErrorHandler
    
    Dim precedents As Range
    Dim area As Range
    Dim cell As Range
    Dim result As New Collection
    Dim cellAddress As String
    
    Set GetPrecedents = result
    
    On Error Resume Next
    Set precedents = sourceCell.DirectPrecedents
    On Error GoTo ErrorHandler
    
    If precedents Is Nothing Then Exit Function
    
    For Each area In precedents.Areas
        For Each cell In area.Cells
            cellAddress = GetFullAddress(cell)
            result.Add cellAddress
        Next cell
    Next area
    
    Set GetPrecedents = result
    Exit Function
    
ErrorHandler:
    Set GetPrecedents = New Collection
End Function

' Get list of dependent cells
Public Function GetDependents(sourceCell As Range) As Collection
    On Error GoTo ErrorHandler
    
    Dim dependents As Range
    Dim area As Range
    Dim cell As Range
    Dim result As New Collection
    Dim cellAddress As String
    
    Set GetDependents = result
    
    On Error Resume Next
    Set dependents = sourceCell.DirectDependents
    On Error GoTo ErrorHandler
    
    If dependents Is Nothing Then Exit Function
    
    For Each area In dependents.Areas
        For Each cell In area.Cells
            cellAddress = GetFullAddress(cell)
            result.Add cellAddress
        Next cell
    Next area
    
    Set GetDependents = result
    Exit Function
    
ErrorHandler:
    Set GetDependents = New Collection
End Function

' Get full address including sheet name
Public Function GetFullAddress(cell As Range) As String
    Dim sheetName As String
    Dim workbookName As String
    
    sheetName = cell.Worksheet.Name
    workbookName = cell.Worksheet.Parent.Name
    
    ' Add single quotes around sheet name if it contains spaces
    If InStr(sheetName, " ") > 0 Then
        sheetName = "'" & sheetName & "'"
    End If
    
    ' Include workbook name if not the active workbook
    If workbookName <> ThisWorkbook.Name And workbookName <> ActiveWorkbook.Name Then
        GetFullAddress = "[" & workbookName & "]" & sheetName & "!" & cell.Address
    Else
        GetFullAddress = sheetName & "!" & cell.Address
    End If
End Function

' Navigate to a cell given its full address
Public Sub NavigateToCell(fullAddress As String)
    On Error GoTo ErrorHandler
    
    Dim targetRange As Range
    Dim sheetName As String
    Dim cellAddress As String
    Dim workbookName As String
    Dim pos As Integer
    Dim ws As Worksheet
    Dim wb As Workbook
    
    ' Parse the address: Sheet1!A1 or 'Sheet 1'!A1 or [Book1]Sheet1!A1
    If Left(fullAddress, 1) = "[" Then
        pos = InStr(fullAddress, "]")
        workbookName = Mid(fullAddress, 2, pos - 2)
        fullAddress = Mid(fullAddress, pos + 1)
        Set wb = Workbooks(workbookName)
    Else
        Set wb = ActiveWorkbook
    End If
    
    pos = InStrRev(fullAddress, "!")
    If pos > 0 Then
        sheetName = Left(fullAddress, pos - 1)
        cellAddress = Mid(fullAddress, pos + 1)
        
        ' Remove quotes from sheet name
        sheetName = Replace(sheetName, "'", "")
        
        Set ws = wb.Sheets(sheetName)
        Set targetRange = ws.Range(cellAddress)
        
        Application.Goto targetRange, True
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Could not navigate to: " & fullAddress & vbCrLf & Err.Description, vbExclamation, "Navigation Error"
End Sub

' Get display value of a cell
Public Function GetCellDisplayValue(cell As Range) As String
    On Error GoTo ErrorHandler
    
    If IsEmpty(cell.Value) Then
        GetCellDisplayValue = ""
    ElseIf IsError(cell.Value) Then
        GetCellDisplayValue = CStr(cell.Value)
    Else
        GetCellDisplayValue = cell.Text
    End If
    
    Exit Function
    
ErrorHandler:
    GetCellDisplayValue = "#ERROR#"
End Function
