Attribute VB_Name = "modTraceTools"
Option Explicit

' Enhanced precedent and dependent tracing - Mac optimized
' Uses immediate navigation without OK button requirement

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

    ' Show dialog with immediate navigation
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

' Show trace dialog with keyboard navigation (stays open, auto-navigates)
Private Sub ShowTraceDialog(originCell As Range, links As Collection, isPrecedents As Boolean)
    On Error GoTo ErrorHandler

    Dim response As String
    Dim currentIndex As Integer
    Dim msg As String
    Dim title As String
    Dim item As Variant
    Dim cellNum As Integer
    Dim instruction As String

    ' Start at first item
    currentIndex = 0

    ' Build title
    If isPrecedents Then
        title = "Trace Precedents Navigator"
    Else
        title = "Trace Dependents Navigator"
    End If

    ' Navigate through list with persistent dialog
    Do
        ' Build message showing current position
        msg = "Origin: " & GetFullAddress(originCell) & vbCrLf
        If originCell.HasFormula Then
            msg = msg & "Formula: " & originCell.Formula & vbCrLf
        End If
        msg = msg & vbCrLf & "📍 LINKED CELLS:" & vbCrLf & vbCrLf

        ' Add current cell to list first (index 0)
        msg = msg & "  0. [CURRENT] " & GetFullAddress(originCell)
        If currentIndex = 0 Then msg = msg & " ◀"
        msg = msg & vbCrLf

        ' Show all linked cells
        Dim i As Integer
        i = 1
        For Each item In links
            msg = msg & "  " & i & ". " & item
            If i = currentIndex Then msg = msg & " ◀"
            msg = msg & vbCrLf
            i = i + 1
        Next item

        msg = msg & vbCrLf & "━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
        msg = msg & "⌨️  NAVIGATE:" & vbCrLf
        msg = msg & "  • Type number (0-" & links.Count & ") + Enter to jump" & vbCrLf
        msg = msg & "  • Type + or n = Next cell" & vbCrLf
        msg = msg & "  • Type - or p = Previous cell" & vbCrLf
        msg = msg & "  • Press ESC or Cancel = Close" & vbCrLf

        ' Show current cell being viewed
        If currentIndex = 0 Then
            msg = msg & vbCrLf & "👁️  Viewing: [CURRENT] " & GetFullAddress(originCell)
        Else
            msg = msg & vbCrLf & "👁️  Viewing: " & links(currentIndex)
        End If

        response = InputBox(msg, title, "")

        ' User cancelled - exit
        If response = "" Then Exit Sub

        ' Parse response
        response = Trim(LCase(response))

        ' Next cell
        If response = "+" Or response = "n" Or response = "next" Then
            currentIndex = currentIndex + 1
            If currentIndex > links.Count Then currentIndex = 0

            ' Navigate to cell
            If currentIndex = 0 Then
                NavigateToCell GetFullAddress(originCell)
            Else
                NavigateToCell links(currentIndex)
            End If

        ' Previous cell
        ElseIf response = "-" Or response = "p" Or response = "prev" Or response = "previous" Then
            currentIndex = currentIndex - 1
            If currentIndex < 0 Then currentIndex = links.Count

            ' Navigate to cell
            If currentIndex = 0 Then
                NavigateToCell GetFullAddress(originCell)
            Else
                NavigateToCell links(currentIndex)
            End If

        ' Direct number entry
        ElseIf IsNumeric(response) Then
            cellNum = CInt(response)
            If cellNum >= 0 And cellNum <= links.Count Then
                currentIndex = cellNum

                ' Navigate to cell
                If currentIndex = 0 Then
                    NavigateToCell GetFullAddress(originCell)
                Else
                    NavigateToCell links(currentIndex)
                End If
            Else
                MsgBox "Please enter a number between 0 and " & links.Count, vbExclamation, "Invalid Number"
            End If
        Else
            MsgBox "Invalid input. Use number, +, -, n, p, or ESC", vbExclamation, "Invalid Input"
        End If
    Loop

    Exit Sub

ErrorHandler:
    MsgBox "Error showing trace dialog: " & Err.Description, vbCritical, "Error"
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
    Dim tempAddress As String

    ' Clean up address
    fullAddress = Trim(fullAddress)
    tempAddress = fullAddress

    ' Parse workbook if present [WorkbookName]
    If Left(tempAddress, 1) = "[" Then
        pos = InStr(tempAddress, "]")
        If pos > 0 Then
            workbookName = Mid(tempAddress, 2, pos - 2)
            tempAddress = Mid(tempAddress, pos + 1)

            ' Find workbook
            On Error Resume Next
            Set wb = Workbooks(workbookName)
            On Error GoTo ErrorHandler

            If wb Is Nothing Then
                MsgBox "Workbook not open: " & workbookName, vbExclamation, "Navigation Error"
                Exit Sub
            End If
        End If
    Else
        Set wb = ActiveWorkbook
    End If

    ' Parse sheet!address
    pos = InStrRev(tempAddress, "!")
    If pos > 0 Then
        sheetName = Left(tempAddress, pos - 1)
        cellAddress = Mid(tempAddress, pos + 1)

        ' Remove single quotes around sheet name
        If Left(sheetName, 1) = "'" And Right(sheetName, 1) = "'" Then
            sheetName = Mid(sheetName, 2, Len(sheetName) - 2)
        End If

        ' Debug: Log what we're looking for
        ' MsgBox "Looking for sheet: [" & sheetName & "] in workbook: " & wb.Name

        ' Get worksheet - be more careful with error handling
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(sheetName)
        On Error GoTo 0

        If ws Is Nothing Then
            ' Try Sheets collection (includes charts, etc.)
            On Error Resume Next
            Set ws = wb.Sheets(sheetName)
            On Error GoTo 0
        End If

        If ws Is Nothing Then
            MsgBox "Sheet not found: [" & sheetName & "]" & vbCrLf & _
                   "In workbook: " & wb.Name & vbCrLf & vbCrLf & _
                   "Full address: " & fullAddress, vbExclamation, "Navigation Error"
            Exit Sub
        End If

        ' Get range
        Set targetRange = Nothing
        On Error Resume Next
        Set targetRange = ws.Range(cellAddress)
        On Error GoTo 0

        If targetRange Is Nothing Then
            On Error GoTo ErrorHandler
            MsgBox "Invalid cell address: " & cellAddress & vbCrLf & _
                   "On sheet: " & sheetName, vbExclamation, "Navigation Error"
            Exit Sub
        End If

        ' Navigate - ensure we activate workbook, then sheet, then select
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False
        wb.Activate
        ws.Activate
        targetRange.Select
        Application.Goto targetRange, True
        Application.ScreenUpdating = True
    Else
        ' No sheet specified - use current sheet
        Set targetRange = Nothing
        On Error Resume Next
        Set targetRange = ActiveSheet.Range(fullAddress)
        On Error GoTo ErrorHandler

        If Not targetRange Is Nothing Then
            Application.Goto targetRange, True
        Else
            MsgBox "Invalid cell address: " & fullAddress, vbExclamation, "Navigation Error"
        End If
    End If

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Could not navigate to: " & fullAddress & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & " (" & Err.Number & ")" & vbCrLf & vbCrLf & _
           "Sheet: [" & sheetName & "]" & vbCrLf & _
           "Cell: " & cellAddress, vbExclamation, "Navigation Error"
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
