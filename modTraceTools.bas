Attribute VB_Name = "modTraceTools"
Option Explicit

' Enhanced precedent and dependent tracing - Mac optimized
' Uses immediate navigation without OK button requirement

' Ribbon callback wrapper for Trace Precedents
Public Sub TracePrecedentsDialog(Optional control As IRibbonControl = Nothing)
    TracePrecedentsImpl
End Sub

' Keyboard shortcut wrapper for Trace Precedents
Public Sub TracePrecedentsKeyboard()
    TracePrecedentsImpl
End Sub

' Implementation: Show precedents dialog
Private Sub TracePrecedentsImpl()
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

' Ribbon callback wrapper for Trace Dependents
Public Sub TraceDependentsDialog(Optional control As IRibbonControl = Nothing)
    TraceDependentsImpl
End Sub

' Keyboard shortcut wrapper for Trace Dependents
Public Sub TraceDependentsKeyboard()
    TraceDependentsImpl
End Sub

' Implementation: Show dependents dialog
Private Sub TraceDependentsImpl()
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
        msg = msg & vbCrLf & ">> LINKED CELLS:" & vbCrLf & vbCrLf

        ' Add current cell to list first (index 0)
        msg = msg & "  0. [CURRENT] " & GetFullAddress(originCell)
        If currentIndex = 0 Then msg = msg & " <--"
        msg = msg & vbCrLf

        ' Show all linked cells
        Dim i As Integer
        i = 1
        For Each item In links
            msg = msg & "  " & i & ". " & item
            If i = currentIndex Then msg = msg & " <--"
            msg = msg & vbCrLf
            i = i + 1
        Next item

        msg = msg & vbCrLf & "----------------------------------------" & vbCrLf
        msg = msg & ">> HOW TO NAVIGATE:" & vbCrLf
        msg = msg & "  - Type number (0-" & links.Count & ") + Enter to jump" & vbCrLf
        msg = msg & "  - Type + or n = Next cell" & vbCrLf
        msg = msg & "  - Type - or p = Previous cell" & vbCrLf
        msg = msg & "  - Press ESC or Cancel = Close" & vbCrLf

        ' Show current cell being viewed
        If currentIndex = 0 Then
            msg = msg & vbCrLf & ">> Currently viewing: [CURRENT] " & GetFullAddress(originCell)
        Else
            msg = msg & vbCrLf & ">> Currently viewing: " & links(currentIndex)
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

' Get precedents with cross-sheet support (improved for Mac compatibility)
Public Function GetPrecedents(sourceCell As Range) As Collection
    On Error GoTo ErrorHandler

    Dim precedents As Range
    Dim area As Range
    Dim cell As Range
    Dim result As New Collection
    Dim cellAddress As String
    Dim parsedRefs As Collection

    Set result = New Collection
    Set GetPrecedents = result

    ' Try DirectPrecedents first (works for same-sheet references)
    On Error Resume Next
    Set precedents = sourceCell.DirectPrecedents
    On Error GoTo ErrorHandler

    If Not precedents Is Nothing Then
        ' Add cells from DirectPrecedents
        For Each area In precedents.Areas
            For Each cell In area.Cells
                cellAddress = GetFullAddress(cell)
                On Error Resume Next
                result.Add cellAddress
                On Error GoTo ErrorHandler
            Next cell
        Next area
    End If

    ' Only parse formula if DirectPrecedents returned nothing
    ' This catches cross-sheet references that DirectPrecedents might miss on Mac
    ' and prevents duplication when DirectPrecedents works correctly
    If result.Count = 0 And sourceCell.HasFormula Then
        Set parsedRefs = ParseFormulaReferences(sourceCell)

        ' Add parsed references
        Dim ref As Variant
        For Each ref In parsedRefs
            On Error Resume Next
            result.Add CStr(ref)
            On Error GoTo ErrorHandler
        Next ref
    End If

    Set GetPrecedents = result
    Exit Function

ErrorHandler:
    Set GetPrecedents = result  ' Return what we have so far
End Function

' Parse formula string to extract cell references (Mac-compatible fallback)
Private Function ParseFormulaReferences(sourceCell As Range) As Collection
    On Error GoTo ErrorHandler

    Dim result As New Collection
    Dim formula As String
    Dim i As Integer
    Dim char As String
    Dim currentRef As String
    Dim inSheet As Boolean
    Dim inQuote As Boolean
    Dim sheetName As String
    Dim cellRef As String
    Dim fullRef As String

    Set ParseFormulaReferences = result

    If Not sourceCell.HasFormula Then Exit Function

    formula = sourceCell.Formula
    currentRef = ""
    inSheet = False
    inQuote = False
    sheetName = ""

    ' Simple parser: look for patterns like Sheet2!A1 or 'Sheet Name'!B5
    For i = 1 To Len(formula)
        char = Mid(formula, i, 1)

        ' Handle quoted sheet names
        If char = "'" Then
            inQuote = Not inQuote
            currentRef = currentRef & char
        ' Handle sheet separator
        ElseIf char = "!" And Not inQuote Then
            inSheet = True
            sheetName = currentRef
            currentRef = ""
        ' Handle cell reference characters
        ElseIf (char >= "A" And char <= "Z") Or (char >= "a" And char <= "z") Or _
               (char >= "0" And char <= "9") Or char = "$" Or char = ":" Or _
               (inQuote And char <> "'") Then
            currentRef = currentRef & char
        ' End of reference
        Else
            If currentRef <> "" And inSheet Then
                ' We have a cross-sheet reference
                cellRef = currentRef
                fullRef = sheetName & "!" & cellRef

                ' Clean up sheet name quotes
                If Left(sheetName, 1) = "'" Then
                    sheetName = Mid(sheetName, 2)
                End If
                If Right(sheetName, 1) = "'" Then
                    sheetName = Left(sheetName, Len(sheetName) - 1)
                End If

                ' Expand ranges (e.g., A1:A10 becomes A1, A2, ..., A10)
                Dim expandedRefs As Collection
                Set expandedRefs = ExpandCellRange(sheetName, cellRef, sourceCell.Worksheet.Parent)

                Dim ref As Variant
                For Each ref In expandedRefs
                    On Error Resume Next
                    result.Add CStr(ref)
                    On Error GoTo ErrorHandler
                Next ref

                inSheet = False
                sheetName = ""
            ElseIf currentRef <> "" And Not inSheet Then
                ' Same-sheet reference - add sheet name
                cellRef = currentRef

                ' Expand ranges
                Set expandedRefs = ExpandCellRange(sourceCell.Worksheet.Name, cellRef, sourceCell.Worksheet.Parent)

                For Each ref In expandedRefs
                    On Error Resume Next
                    result.Add CStr(ref)
                    On Error GoTo ErrorHandler
                Next ref
            End If

            currentRef = ""
        End If
    Next i

    ' Handle last reference if formula ends with a reference
    If currentRef <> "" And inSheet Then
        cellRef = currentRef
        Set expandedRefs = ExpandCellRange(sheetName, cellRef, sourceCell.Worksheet.Parent)
        For Each ref In expandedRefs
            On Error Resume Next
            result.Add CStr(ref)
            On Error GoTo ErrorHandler
        Next ref
    ElseIf currentRef <> "" Then
        Set expandedRefs = ExpandCellRange(sourceCell.Worksheet.Name, currentRef, sourceCell.Worksheet.Parent)
        For Each ref In expandedRefs
            On Error Resume Next
            result.Add CStr(ref)
            On Error GoTo ErrorHandler
        Next ref
    End If

    Set ParseFormulaReferences = result
    Exit Function

ErrorHandler:
    Set ParseFormulaReferences = result
End Function

' Expand cell range (e.g., A1:A10) into individual cells
Private Function ExpandCellRange(sheetName As String, cellRef As String, wb As Workbook) As Collection
    On Error GoTo ErrorHandler

    Dim result As New Collection
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim fullAddr As String

    Set ExpandCellRange = result

    ' Get worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    If ws Is Nothing Then Set ws = wb.Sheets(sheetName)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then Exit Function

    ' Get range
    On Error Resume Next
    Set rng = ws.Range(cellRef)
    On Error GoTo ErrorHandler

    If rng Is Nothing Then Exit Function

    ' Add each cell in range
    For Each cell In rng.Cells
        fullAddr = GetFullAddress(cell)
        On Error Resume Next
        result.Add fullAddr
        On Error GoTo ErrorHandler
    Next cell

    Set ExpandCellRange = result
    Exit Function

ErrorHandler:
    Set ExpandCellRange = result
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
