' ============================================================================
' LEAN MACRO TOOLS FOR EXCEL MAC
' ============================================================================
' Version: 1.0
' Date: November 2025
' Compatible with: Excel for Mac 16.102.2 on macOS Sequoia 15.6.1
' 
' FEATURES:
' 1. Number Format Cycling (Ctrl+Shift+N)
' 2. Trace Precedents (Ctrl+Shift+T)
' 3. Trace Dependents (Ctrl+Shift+Y)
'
' INSTALLATION:
' 1. Import this file in VBA Editor: File > Import File...
' 2. Save workbook as .xlam in Add-ins folder
' 3. Enable in Tools > Excel Add-ins
' 4. Assign keyboard shortcuts in Tools > Macro > Macros > Options
'
' ============================================================================


' ============================================================================
' MODULE 1: modNumberFormats
' ============================================================================
Attribute VB_Name = "modNumberFormats"
Option Explicit

Private Const CONFIG_SHEET_NAME As String = "NumberFormatConfig"

' ============================================================================
' PUBLIC MACROS
' ============================================================================

' ----------------------------------------------------------------------------
' Macro: CycleCustomNumberFormats
' Shortcut: Ctrl+Shift+N
' Description: Cycles through enabled number formats for selected cell(s)
' ----------------------------------------------------------------------------
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
        MsgBox "No number formats are enabled." & vbCrLf & vbCrLf & _
               "Run ConfigureNumberFormats to set up your formats.", _
               vbExclamation, "No Formats"
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
            If nextIndex > UBound(enabledFormats) Then
                nextIndex = LBound(enabledFormats) ' Wrap around
            End If
            Exit For
        End If
    Next i
    
    ' Apply the next format to entire selection
    targetRange.NumberFormat = enabledFormats(nextIndex)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error cycling number formats: " & Err.Description, _
           vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Macro: ConfigureNumberFormats
' Description: Opens the configuration sheet for editing number formats
' ----------------------------------------------------------------------------
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

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: LoadFormats
' Description: Loads number formats and their enabled status from config sheet
' Parameters:
'   - formats: Output array of format strings
'   - enabled: Output array of boolean values (TRUE = enabled)
' ----------------------------------------------------------------------------
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
    
    ' Default format #1: Thousands with 2 decimals
    formats(1) = "#,##0.00_);(#,##0.00);""-""_);@_)"
    enabled(1) = True
    
    ' Default format #2: Percentage with 1 decimal
    formats(2) = "0.0%_);(0.0%);""-""_);@_)"
    enabled(2) = True
    
    ' Default format #3: Multiple (e.g., 2.5x)
    formats(3) = "#,##0.0x_);(#,##0.0)x;""-""_);@_)"
    enabled(3) = True
    
    ' Default format #4: USD
    formats(4) = "$#,##0.0_);$(#,##0.0)""x"";""-""_);@_)"
    enabled(4) = True
    
    ' Default format #5: Brazilian Real
    formats(5) = "R$#,##0.0_);R$(#,##0.0)""x"";""-""_);@_)"
    enabled(5) = True
End Sub

' ----------------------------------------------------------------------------
' Function: GetOrCreateConfigSheet
' Description: Returns the configuration sheet, creating it if it doesn't exist
' Returns: Worksheet object
' ----------------------------------------------------------------------------
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
            ' Headers
            .Cells(1, 1).Value = "Format"
            .Cells(1, 2).Value = "Enabled"
            .Cells(1, 1).Font.Bold = True
            .Cells(1, 2).Font.Bold = True
            
            ' Default format #1: Thousands with 2 decimals
            .Cells(2, 1).Value = "#,##0.00_);(#,##0.00);""-""_);@_)"
            .Cells(2, 2).Value = "TRUE"
            
            ' Default format #2: Percentage
            .Cells(3, 1).Value = "0.0%_);(0.0%);""-""_);@_)"
            .Cells(3, 2).Value = "TRUE"
            
            ' Default format #3: Multiple (2.5x)
            .Cells(4, 1).Value = "#,##0.0x_);(#,##0.0)x;""-""_);@_)"
            .Cells(4, 2).Value = "TRUE"
            
            ' Default format #4: USD
            .Cells(5, 1).Value = "$#,##0.0_);$(#,##0.0)""x"";""-""_);@_)"
            .Cells(5, 2).Value = "TRUE"
            
            ' Default format #5: Brazilian Real
            .Cells(6, 1).Value = "R$#,##0.0_);R$(#,##0.0)""x"";""-""_);@_)"
            .Cells(6, 2).Value = "TRUE"
            
            ' Format columns for readability
            .Columns(1).ColumnWidth = 50
            .Columns(2).ColumnWidth = 12
        End With
        
        ' Hide the sheet (users access it via ConfigureNumberFormats macro)
        GetOrCreateConfigSheet.Visible = xlSheetVeryHidden
    End If
End Function


' ============================================================================
' MODULE 2: modTraceTools
' ============================================================================
Attribute VB_Name = "modTraceTools"
Option Explicit

' ============================================================================
' PUBLIC MACROS
' ============================================================================

' ----------------------------------------------------------------------------
' Macro: TracePrecedentsDialog
' Shortcut: Ctrl+Shift+T
' Description: Shows precedent cells (cells that feed into current formula)
' ----------------------------------------------------------------------------
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
        MsgBox "The selected cell does not contain a formula.", _
               vbInformation, "Trace Precedents"
        Exit Sub
    End If
    
    ' Get precedents
    Set precedents = GetPrecedents(activeCell)
    
    If precedents.Count = 0 Then
        MsgBox "No precedent cells found for " & GetFullAddress(activeCell), _
               vbInformation, "No Precedents"
        Exit Sub
    End If
    
    ' Build message showing precedents
    msg = "TRACE PRECEDENTS" & vbCrLf & _
          String(50, "=") & vbCrLf & vbCrLf & _
          "Origin: " & GetFullAddress(activeCell) & vbCrLf & _
          "Value: " & GetCellDisplayValue(activeCell) & vbCrLf & _
          "Formula: " & activeCell.Formula & vbCrLf & vbCrLf & _
          "Precedent Cells:" & vbCrLf
    
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

' ----------------------------------------------------------------------------
' Macro: TraceDependentsDialog
' Shortcut: Ctrl+Shift+Y
' Description: Shows dependent cells (cells that use current cell in formulas)
' ----------------------------------------------------------------------------
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
        MsgBox "No dependent cells found for " & GetFullAddress(activeCell), _
               vbInformation, "No Dependents"
        Exit Sub
    End If
    
    ' Build message showing dependents
    msg = "TRACE DEPENDENTS" & vbCrLf & _
          String(50, "=") & vbCrLf & vbCrLf & _
          "Origin: " & GetFullAddress(activeCell) & vbCrLf & _
          "Value: " & GetCellDisplayValue(activeCell) & vbCrLf
    
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

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetPrecedents
' Description: Returns collection of precedent cell addresses
' Parameters:
'   - sourceCell: The cell to analyze
' Returns: Collection of address strings (e.g., "Sheet1!A1")
' ----------------------------------------------------------------------------
Public Function GetPrecedents(sourceCell As Range) As Collection
    On Error GoTo ErrorHandler
    
    Dim precedents As Range
    Dim area As Range
    Dim cell As Range
    Dim result As New Collection
    Dim cellAddress As String
    
    Set GetPrecedents = result
    
    ' Get direct precedents using Excel's built-in method
    On Error Resume Next
    Set precedents = sourceCell.DirectPrecedents
    On Error GoTo ErrorHandler
    
    If precedents Is Nothing Then Exit Function
    
    ' Build collection of precedent addresses
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

' ----------------------------------------------------------------------------
' Function: GetDependents
' Description: Returns collection of dependent cell addresses
' Parameters:
'   - sourceCell: The cell to analyze
' Returns: Collection of address strings (e.g., "Sheet1!A1")
' ----------------------------------------------------------------------------
Public Function GetDependents(sourceCell As Range) As Collection
    On Error GoTo ErrorHandler
    
    Dim dependents As Range
    Dim area As Range
    Dim cell As Range
    Dim result As New Collection
    Dim cellAddress As String
    
    Set GetDependents = result
    
    ' Get direct dependents using Excel's built-in method
    On Error Resume Next
    Set dependents = sourceCell.DirectDependents
    On Error GoTo ErrorHandler
    
    If dependents Is Nothing Then Exit Function
    
    ' Build collection of dependent addresses
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

' ----------------------------------------------------------------------------
' Function: GetFullAddress
' Description: Returns full address of cell including sheet and workbook name
' Parameters:
'   - cell: The cell to get address for
' Returns: String like "Sheet1!A1" or "[Book1]Sheet1!A1"
' ----------------------------------------------------------------------------
Public Function GetFullAddress(cell As Range) As String
    Dim sheetName As String
    Dim workbookName As String
    
    sheetName = cell.Worksheet.Name
    workbookName = cell.Worksheet.Parent.Name
    
    ' Add single quotes around sheet name if it contains spaces
    If InStr(sheetName, " ") > 0 Then
        sheetName = "'" & sheetName & "'"
    End If
    
    ' Include workbook name if different from current workbook
    If workbookName <> ThisWorkbook.Name And _
       workbookName <> ActiveWorkbook.Name Then
        GetFullAddress = "[" & workbookName & "]" & sheetName & "!" & cell.Address
    Else
        GetFullAddress = sheetName & "!" & cell.Address
    End If
End Function

' ----------------------------------------------------------------------------
' Function: NavigateToCell
' Description: Navigates to a cell given its full address string
' Parameters:
'   - fullAddress: Address string like "Sheet1!A1" or "[Book1]Sheet1!A1"
' ----------------------------------------------------------------------------
Public Sub NavigateToCell(fullAddress As String)
    On Error GoTo ErrorHandler
    
    Dim targetRange As Range
    Dim sheetName As String
    Dim cellAddress As String
    Dim workbookName As String
    Dim pos As Integer
    Dim ws As Worksheet
    Dim wb As Workbook
    
    ' Parse the address
    ' Formats: Sheet1!A1 or 'Sheet 1'!A1 or [Book1]Sheet1!A1
    
    ' Check for workbook reference
    If Left(fullAddress, 1) = "[" Then
        pos = InStr(fullAddress, "]")
        workbookName = Mid(fullAddress, 2, pos - 2)
        fullAddress = Mid(fullAddress, pos + 1)
        Set wb = Workbooks(workbookName)
    Else
        Set wb = ActiveWorkbook
    End If
    
    ' Split on "!" to separate sheet and cell address
    pos = InStrRev(fullAddress, "!")
    If pos > 0 Then
        sheetName = Left(fullAddress, pos - 1)
        cellAddress = Mid(fullAddress, pos + 1)
        
        ' Remove quotes from sheet name if present
        sheetName = Replace(sheetName, "'", "")
        
        ' Navigate to the cell
        Set ws = wb.Sheets(sheetName)
        Set targetRange = ws.Range(cellAddress)
        Application.Goto targetRange, True
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Could not navigate to: " & fullAddress & vbCrLf & _
           Err.Description, vbExclamation, "Navigation Error"
End Sub

' ----------------------------------------------------------------------------
' Function: GetCellDisplayValue
' Description: Returns the formatted display value of a cell
' Parameters:
'   - cell: The cell to get value from
' Returns: String representation of the cell's displayed value
' ----------------------------------------------------------------------------
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


' ============================================================================
' END OF LEAN MACRO TOOLS
' ============================================================================
'
' TO USE:
' 1. Import this file in Excel VBA Editor
' 2. Save as .xlam add-in
' 3. Enable in Tools > Excel Add-ins
' 4. Assign shortcuts:
'    - CycleCustomNumberFormats → Ctrl+Shift+N
'    - TracePrecedentsDialog → Ctrl+Shift+T
'    - TraceDependentsDialog → Ctrl+Shift+Y
'
' For detailed instructions, see INSTALLATION_GUIDE_SIMPLIFIED.md
' ============================================================================
