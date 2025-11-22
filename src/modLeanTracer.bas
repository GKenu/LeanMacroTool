Attribute VB_Name = "modLeanTracer"
Option Explicit

' Lean Macro Tools - Tracer Adapter Module
' Bridges our GetPrecedents/GetDependents functions with the UserForm interface

' Module-level form instances (persistent)
Dim mfrmPrecedentTracer As FPrecedentAnalyzer
Dim mfrmDependentTracer As zFPrecedentAnalyzer

' Mode tracking - determines which function GetPrecedents/GetDependents to use
Private mbDependentMode As Boolean

' Configuration
Public Const bOriginalCellAtEnd As Boolean = False  ' Original cell at index 0
Public Const TTS_TITLE As String = "Lean Macro Tools"  ' Form title for message boxes

' ============================================================================
' PUBLIC INTERFACE - Call these from keyboard shortcuts
' ============================================================================

Public Sub ShowLeanPrecedents()
    ' Show precedent tracer for active cell
    If ActiveCell Is Nothing Then
        MsgBox "Please select a cell first.", vbExclamation, "Lean Precedent Tracer"
        Exit Sub
    End If

    ' Create form instance if needed
    If mfrmPrecedentTracer Is Nothing Then
        Set mfrmPrecedentTracer = New FPrecedentAnalyzer
    End If

    ' Set mode flag BEFORE setting cell (so NewPrecedents knows which function to use)
    mbDependentMode = False

    ' Set the cell and show form (form handles the rest via Property Set)
    With mfrmPrecedentTracer
        Set .Cell = ActiveCell
        .Show vbModeless
    End With
End Sub

Public Sub ShowLeanDependents()
    ' Show dependent tracer for active cell
    If ActiveCell Is Nothing Then
        MsgBox "Please select a cell first.", vbExclamation, "Lean Dependent Tracer"
        Exit Sub
    End If

    ' Create form instance if needed
    If mfrmDependentTracer Is Nothing Then
        Set mfrmDependentTracer = New zFPrecedentAnalyzer
    End If

    ' Set mode flag BEFORE setting cell (so NewPrecedents knows to use GetDependents)
    mbDependentMode = True

    ' Set the cell and show form (form handles the rest via Property Set)
    With mfrmDependentTracer
        Set .Cell = ActiveCell
        .Show vbModeless
    End With
End Sub

' ============================================================================
' HELPER FUNCTIONS - Address formatting
' ============================================================================

Private Function GetFullAddress(r As Range) As String
    ' Get full address - removes workbook brackets if same workbook
    ' so Range() can parse it
    ' Example: Sheet1!$A$1 or [OtherBook.xlsx]Sheet1!$A$1
    Dim sRange As String
    Dim sAddress As String
    Dim sWorkbook As String

    sRange = r.Address(External:=True)

    ' Extract workbook name from [workbook]Sheet!Address format
    If InStr(sRange, "[") > 0 Then
        sWorkbook = Mid$(sRange, InStr(sRange, "[") + 1)
        sWorkbook = Left$(sWorkbook, InStr(sWorkbook, "]") - 1)

        ' If same workbook, remove the [workbook] part so Range() can parse it
        If sWorkbook = ActiveWorkbook.Name Then
            sAddress = Left$(sRange, InStr(sRange, "[") - 1) & Mid$(sRange, InStr(sRange, "]") + 1)
        Else
            ' Different workbook - keep full format
            sAddress = sRange
        End If
    Else
        sAddress = sRange
    End If

    GetFullAddress = sAddress
End Function

Private Function GetShortAddress(r As Range) As String
    ' Get short address (without workbook if same workbook)
    ' Example: Sheet1!A1 (or just A1 if same sheet)
    Dim fullAddr As String
    Dim result As String

    fullAddr = r.Address(External:=True)

    ' Remove workbook name if same workbook
    If InStr(fullAddr, "[" & ActiveWorkbook.Name & "]") > 0 Then
        result = Replace(fullAddr, "[" & ActiveWorkbook.Name & "]", "")

        ' Remove single quotes if present and no spaces in sheet name
        If Left(result, 1) = "'" Then
            result = Mid(result, 2)
            result = Replace(result, "'!", "!")
        End If

        ' If same sheet, show just cell reference
        If r.Worksheet.Name = ActiveSheet.Name Then
            If InStr(result, "!") > 0 Then
                result = Mid(result, InStr(result, "!") + 1)
            End If
        End If
    Else
        result = fullAddr
    End If

    ' Remove absolute references for cleaner display
    result = Replace(result, "$", "")

    GetShortAddress = result
End Function

Public Function IsRange(sTest As String) As Boolean
    ' Helper function to test if a string is a valid range
    ' Required by FPrecedentAnalyzer UserForm
    On Error Resume Next
    IsRange = (TypeName(Range(sTest)) = "Range")
    On Error GoTo 0
End Function

' ============================================================================
' FUNCTIONS REQUIRED BY FPrecedentAnalyzer USERFORM
' ============================================================================

Public Function GetAddress(rRange As Range) As String
    ' Required by FPrecedentAnalyzer.Cell property
    ' Returns: "full_address|short_address"
    Dim fullAddr As String
    Dim shortAddr As String

    fullAddr = GetFullAddress(rRange)
    shortAddr = GetShortAddress(rRange)

    GetAddress = fullAddr & "|" & shortAddr
End Function

Public Function NewPrecedents(rCell As Range) As Variant
    ' Required by both UserForms (FPrecedentAnalyzer and zFPrecedentAnalyzer)
    ' Returns 2D array in format expected by UserForm
    ' Detects which form is active to determine precedents vs dependents

    Dim precedents As Collection
    Dim precedentArray As Variant
    Dim i As Long
    Dim prec As Range
    Dim cellAddress As String
    Dim isDependentForm As Boolean

    ' Use the mode flag to determine which function to call
    ' (Flag is set in ShowLeanPrecedents/ShowLeanDependents before Set .Cell is called)
    isDependentForm = mbDependentMode

    ' Get precedents OR dependents based on mode flag
    If isDependentForm Then
        Set precedents = GetDependents(rCell)
    Else
        Set precedents = GetPrecedents(rCell)
    End If

    If precedents Is Nothing Or precedents.Count = 0 Then
        ' No items found - return empty
        NewPrecedents = ""
        Exit Function
    End If

    ' Build 2D array: (rows, 3 columns)
    ' Add 1 extra row for original cell
    ReDim precedentArray(1 To precedents.Count + 1, 1 To 3)

    If bOriginalCellAtEnd Then
        ' Precedents first (1 to Count), then original cell (Count+1)
        For i = 1 To precedents.Count
            ' v2.1.0: precedents(i) is already a string address with formula order preserved
            ' Convert to Range to get proper formatting, but maintain order
            On Error Resume Next
            Set prec = Range(precedents(i))
            On Error GoTo 0
            If Not prec Is Nothing Then
                ' Get formatted addresses (full and short versions)
                cellAddress = GetAddress(prec)
                precedentArray(i, 1) = Left$(cellAddress, InStr(cellAddress, "|") - 1)  ' Full
                precedentArray(i, 2) = Mid$(cellAddress, InStr(cellAddress, "|") + 1)   ' Short
                precedentArray(i, 3) = ""  ' Display (filled by form)
            Else
                ' If Range() failed, use the string address directly
                ' This handles cases where the range might not be valid
                precedentArray(i, 1) = precedents(i)  ' Full address
                precedentArray(i, 2) = precedents(i)  ' Short address (same)
                precedentArray(i, 3) = ""
            End If
        Next i

        ' Original cell at end
        cellAddress = GetAddress(rCell)
        precedentArray(precedents.Count + 1, 1) = Left$(cellAddress, InStr(cellAddress, "|") - 1)
        precedentArray(precedents.Count + 1, 2) = Mid$(cellAddress, InStr(cellAddress, "|") + 1)
        precedentArray(precedents.Count + 1, 3) = ""
    Else
        ' Original cell first (1), then precedents (2 to Count+1)
        cellAddress = GetAddress(rCell)
        precedentArray(1, 1) = Left$(cellAddress, InStr(cellAddress, "|") - 1)
        precedentArray(1, 2) = Mid$(cellAddress, InStr(cellAddress, "|") + 1)
        precedentArray(1, 3) = ""

        For i = 1 To precedents.Count
            ' v2.1.0: precedents(i) is already a string address with formula order preserved
            ' Convert to Range to get proper formatting, but maintain order
            On Error Resume Next
            Set prec = Range(precedents(i))
            On Error GoTo 0
            If Not prec Is Nothing Then
                ' Get formatted addresses (full and short versions)
                cellAddress = GetAddress(prec)
                precedentArray(i + 1, 1) = Left$(cellAddress, InStr(cellAddress, "|") - 1)
                precedentArray(i + 1, 2) = Mid$(cellAddress, InStr(cellAddress, "|") + 1)
                precedentArray(i + 1, 3) = ""
            Else
                ' If Range() failed, use the string address directly
                ' This handles cases where the range might not be valid
                precedentArray(i + 1, 1) = precedents(i)  ' Full address
                precedentArray(i + 1, 2) = precedents(i)  ' Short address (same)
                precedentArray(i + 1, 3) = ""
            End If
        Next i
    End If

    NewPrecedents = precedentArray
End Function

' ============================================================================
' CORE TRACING FUNCTIONS - Get precedents and dependents
' ============================================================================

Public Function GetPrecedents(sourceCell As Range) As Collection
    On Error GoTo ErrorHandler

    Dim result As New Collection
    Dim parsedRefs As Collection

    Set GetPrecedents = result

    ' v2.1.0: Always use ParseFormulaReferences to preserve formula order
    ' DirectPrecedents returns cells in position order (A1, B1, C1...) not formula order
    ' ParseFormulaReferences preserves the order references appear in the formula
    If sourceCell.HasFormula Then
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

    formula = sourceCell.formula
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

        ' Clean up sheet name quotes
        If Left(sheetName, 1) = "'" Then
            sheetName = Mid(sheetName, 2)
        End If
        If Right(sheetName, 1) = "'" Then
            sheetName = Left(sheetName, Len(sheetName) - 1)
        End If

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

Private Function ExpandCellRange(sheetName As String, cellRef As String, wb As Workbook) As Collection
    On Error GoTo ErrorHandler

    Dim result As New Collection
    Dim ws As Worksheet
    Dim rng As Range
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

    ' v2.1.0: Return entire range as single item instead of expanding
    ' This preserves ranges like "A1:A10" for cleaner display
    ' If it's a single cell, return it; if it's a range, return the range notation
    fullAddr = GetFullAddress(rng)
    On Error Resume Next
    result.Add fullAddr
    On Error GoTo ErrorHandler

    Set ExpandCellRange = result
    Exit Function

ErrorHandler:
    Set ExpandCellRange = result
End Function

' ============================================================================
' FORM LIFECYCLE - Clean up
' ============================================================================

Public Sub CleanupTracerForms()
    ' Clean up form instances (call on workbook close)
    If Not mfrmPrecedentTracer Is Nothing Then
        Unload mfrmPrecedentTracer
        Set mfrmPrecedentTracer = Nothing
    End If
    If Not mfrmDependentTracer Is Nothing Then
        Unload mfrmDependentTracer
        Set mfrmDependentTracer = Nothing
    End If
End Sub

' ============================================================================
' FORM HELPER FUNCTIONS - Called by UserForm keyboard shortcuts
' ============================================================================

Public Sub ActivatePrecedentAnalyzer()
    ' Activate the precedent analyzer form (called by Ctrl+Shift+Home from within form)
    If Not mfrmPrecedentTracer Is Nothing Then
        On Error Resume Next
        mfrmPrecedentTracer.SetFocus
        On Error GoTo 0
    End If
    If Not mfrmDependentTracer Is Nothing Then
        On Error Resume Next
        mfrmDependentTracer.SetFocus
        On Error GoTo 0
    End If
End Sub

Public Sub ResetTracerDialogSettings()
    ' Reset tracer dialog settings after keyboard shortcut activation
    Application.ScreenUpdating = True
    If Not mfrmPrecedentTracer Is Nothing Then
        On Error Resume Next
        mfrmPrecedentTracer.ResetListIndex
        On Error GoTo 0
    End If
    If Not mfrmDependentTracer Is Nothing Then
        On Error Resume Next
        mfrmDependentTracer.ResetListIndex
        On Error GoTo 0
    End If
End Sub
