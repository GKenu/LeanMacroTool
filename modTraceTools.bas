Attribute VB_Name = "modTraceTools"
Option Explicit

' Enhanced precedent and dependent tracing with AppleScript dialogs

' Show precedents dialog
Public Sub TracePrecedentsDialog()
    On Error GoTo ErrorHandler
    
    Dim activeCell As Range
    Dim precedents As Collection
    Dim item As Variant
    Dim i As Integer
    Dim listItems As String
    Dim msg As String
    
    Set activeCell = Application.ActiveCell
    
    If activeCell Is Nothing Then
        MsgBox "No cell selected.", vbExclamation, "Trace Precedents"
        Exit Sub
    End If
    
    If Not activeCell.HasFormula Then
        MsgBox "The selected cell does not contain a formula.", vbInformation, "Trace Precedents"
        Exit Sub
    End If
    
    Set precedents = GetPrecedents(activeCell)
    
    If precedents.Count = 0 Then
        MsgBox "No precedent cells found for " & GetFullAddress(activeCell), vbInformation, "No Precedents"
        Exit Sub
    End If
    
    ' Build dialog with AppleScript for better Mac UX
    Call ShowTraceDialog(activeCell, precedents, True)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error tracing precedents: " & Err.Description, vbCritical, "Error"
End Sub

' Show dependents dialog
Public Sub TraceDependentsDialog()
    On Error GoTo ErrorHandler
    
    Dim activeCell As Range
    Dim dependents As Collection
    
    Set activeCell = Application.ActiveCell
    
    If activeCell Is Nothing Then
        MsgBox "No cell selected.", vbExclamation, "Trace Dependents"
        Exit Sub
    End If
    
    Set dependents = GetDependents(activeCell)
    
    If dependents.Count = 0 Then
        MsgBox "No dependent cells found for " & GetFullAddress(activeCell), vbInformation, "No Dependents"
        Exit Sub
    End If
    
    Call ShowTraceDialog(activeCell, dependents, False)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error tracing dependents: " & Err.Description, vbCritical, "Error"
End Sub

' Show trace dialog using AppleScript list picker
Private Sub ShowTraceDialog(originCell As Range, links As Collection, isPrecedents As Boolean)
    On Error GoTo ErrorHandler
    
    Dim asScript As String
    Dim result As String
    Dim item As Variant
    Dim listStr As String
    Dim title As String
    Dim msg As String
    Dim selectedItem As String
    
    ' Build title and message
    If isPrecedents Then
        title = "Trace Precedents"
        msg = "Select a precedent cell to jump to:"
    Else
        title = "Trace Dependents"
        msg = "Select a dependent cell to jump to:"
    End If
    
    msg = msg & vbLf & vbLf & "Origin: " & GetFullAddress(originCell)
    msg = msg & vbLf & "Value: " & GetCellDisplayValue(originCell)
    If originCell.HasFormula Then
        msg = msg & vbLf & "Formula: " & originCell.Formula
    End If
    msg = msg & vbLf & vbLf & "Linked cells:"
    
    ' Build list for AppleScript
    listStr = ""
    For Each item In links
        If listStr <> "" Then listStr = listStr & ", "
        listStr = listStr & """" & Replace(CStr(item), """", """""") & """"
    Next item
    
    ' AppleScript to show list dialog
    asScript = "choose from list {" & listStr & "} " & _
               "with title """ & title & """ " & _
               "with prompt """ & Replace(msg, """", "\""") & """ " & _
               "default items {""" & Replace(CStr(links(1)), """", """""") & """}"
    
    result = MacScript(asScript)
    
    If result <> "false" Then
        ' User selected something
        NavigateToCell result
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Fallback to simpler InputBox method
    Dim response As String
    Dim i As Integer
    
    msg = "TRACE " & IIf(isPrecedents, "PRECEDENTS", "DEPENDENTS") & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    msg = msg & "Origin: " & GetFullAddress(originCell) & vbCrLf
    msg = msg & "Value: " & GetCellDisplayValue(originCell) & vbCrLf
    If originCell.HasFormula Then
        msg = msg & "Formula: " & originCell.Formula & vbCrLf
    End If
    msg = msg & vbCrLf & "Linked Cells:" & vbCrLf
    
    i = 1
    For Each item In links
        msg = msg & "  " & i & ". " & item & vbCrLf
        i = i + 1
    Next item
    
    msg = msg & vbCrLf & "Enter number (1-" & links.Count & "):"
    
    response = InputBox(msg, title, "1")
    
    If response <> "" And IsNumeric(response) Then
        i = CInt(response)
        If i >= 1 And i <= links.Count Then
            NavigateToCell links(i)
        End If
    End If
End Sub

' Get precedents with cross-sheet support
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
            On Error Resume Next
            result.Add cellAddress
            On Error GoTo ErrorHandler
        Next cell
    Next area
    
    Set GetPrecedents = result
    Exit Function
    
ErrorHandler:
    Set GetPrecedents = New Collection
End Function

' Get dependents with cross-sheet support
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
            On Error Resume Next
            result.Add cellAddress
            On Error GoTo ErrorHandler
        Next cell
    Next area
    
    Set GetDependents = result
    Exit Function
    
ErrorHandler:
    Set GetDependents = New Collection
End Function

' Get full address with sheet and workbook
Public Function GetFullAddress(cell As Range) As String
    Dim sheetName As String
    Dim workbookName As String
    Dim wb As Workbook
    
    On Error Resume Next
    
    Set wb = cell.Worksheet.Parent
    sheetName = cell.Worksheet.Name
    workbookName = wb.Name
    
    ' Quote sheet name if it has spaces
    If InStr(sheetName, " ") > 0 Then
        sheetName = "'" & sheetName & "'"
    End If
    
    ' Include workbook if different from active
    If wb.Name <> ActiveWorkbook.Name Then
        GetFullAddress = "[" & workbookName & "]" & sheetName & "!" & cell.Address(False, False)
    Else
        GetFullAddress = sheetName & "!" & cell.Address(False, False)
    End If
End Function

' Navigate to cell with robust cross-sheet handling
Public Sub NavigateToCell(fullAddress As String)
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim workbookName As String
    Dim sheetName As String
    Dim cellAddress As String
    Dim pos As Integer
    
    ' Clean up address
    fullAddress = Trim(fullAddress)
    
    ' Parse workbook if present [WorkbookName]
    If Left(fullAddress, 1) = "[" Then
        pos = InStr(fullAddress, "]")
        If pos > 0 Then
            workbookName = Mid(fullAddress, 2, pos - 2)
            fullAddress = Mid(fullAddress, pos + 1)
            
            ' Find workbook
            On Error Resume Next
            Set wb = Workbooks(workbookName)
            On Error GoTo ErrorHandler
            
            If wb Is Nothing Then
                MsgBox "Workbook not open: " & workbookName, vbExclamation
                Exit Sub
            End If
        End If
    Else
        Set wb = ActiveWorkbook
    End If
    
    ' Parse sheet!address
    pos = InStrRev(fullAddress, "!")
    If pos > 0 Then
        sheetName = Left(fullAddress, pos - 1)
        cellAddress = Mid(fullAddress, pos + 1)
        
        ' Remove quotes
        sheetName = Replace(sheetName, "'", "")
        
        ' Get worksheet
        On Error Resume Next
        Set ws = wb.Worksheets(sheetName)
        
        If ws Is Nothing Then
            ' Try by index or name variations
            Set ws = wb.Sheets(sheetName)
        End If
        On Error GoTo ErrorHandler
        
        If ws Is Nothing Then
            MsgBox "Sheet not found: " & sheetName & " in " & wb.Name, vbExclamation
            Exit Sub
        End If
        
        ' Get range
        On Error Resume Next
        Set targetRange = ws.Range(cellAddress)
        On Error GoTo ErrorHandler
        
        If targetRange Is Nothing Then
            MsgBox "Invalid cell address: " & cellAddress, vbExclamation
            Exit Sub
        End If
        
        ' Navigate
        wb.Activate
        ws.Activate
        targetRange.Select
        Application.Goto targetRange, True
    Else
        ' No sheet specified
        Set targetRange = Range(fullAddress)
        Application.Goto targetRange, True
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Could not navigate to: " & fullAddress & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & " (" & Err.Number & ")", vbExclamation, "Navigation Error"
End Sub

' Get cell display value
Public Function GetCellDisplayValue(cell As Range) As String
    On Error Resume Next
    
    If IsEmpty(cell.Value) Then
        GetCellDisplayValue = "(empty)"
    ElseIf IsError(cell.Value) Then
        GetCellDisplayValue = CStr(cell.Value)
    Else
        GetCellDisplayValue = cell.Text
        If GetCellDisplayValue = "" Then GetCellDisplayValue = CStr(cell.Value)
    End If
    
    If Err.Number <> 0 Then GetCellDisplayValue = "#ERROR#"
End Function
