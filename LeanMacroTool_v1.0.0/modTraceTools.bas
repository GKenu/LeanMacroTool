Attribute VB_Name = "modTraceTools"
Option Explicit

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
