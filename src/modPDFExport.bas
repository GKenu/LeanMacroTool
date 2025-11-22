Attribute VB_Name = "modPDFExport"
Option Explicit

' Lean Macro Tools - PDF Export Module
' Exports entire workbook to PDF with auto-range detection
' One page per sheet, high quality output

' Type to store original PageSetup settings
Private Type PageSettings
    PrintArea As String
    FitToPagesWide As Variant
    FitToPagesTall As Variant
    Zoom As Variant
End Type

' ============================================================================
' PUBLIC INTERFACE - Call from ribbon button
' ============================================================================

Public Sub ExportWorkbookToPDF(Optional control As Object = Nothing)
    ' Ribbon callback wrapper
    ' Note: control parameter uses Object instead of IRibbonControl for Mac compatibility
    ExportWorkbookToPDFImpl
End Sub

' ============================================================================
' IMPLEMENTATION
' ============================================================================

Private Sub ExportWorkbookToPDFImpl()
    On Error GoTo ErrorHandler

    Dim savePath As Variant
    Dim ws As Worksheet
    Dim originalSettings() As PageSettings
    Dim settingsCount As Long
    Dim i As Long
    Dim visibleSheets As Long

    ' Check if workbook has any sheets
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Export to PDF"
        Exit Sub
    End If

    If ActiveWorkbook.Worksheets.Count = 0 Then
        MsgBox "The workbook has no worksheets to export.", vbExclamation, "Export to PDF"
        Exit Sub
    End If

    ' Count visible sheets
    visibleSheets = 0
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            visibleSheets = visibleSheets + 1
        End If
    Next ws

    If visibleSheets = 0 Then
        MsgBox "The workbook has no visible worksheets to export.", vbExclamation, "Export to PDF"
        Exit Sub
    End If

    ' Build save path in same directory as workbook
    Dim wbPath As String
    Dim wbName As String

    ' Get workbook path and name
    wbPath = ActiveWorkbook.Path
    wbName = ActiveWorkbook.Name

    ' Remove extension from workbook name
    If InStr(wbName, ".") > 0 Then
        wbName = Left(wbName, InStrRev(wbName, ".") - 1)
    End If

    ' If workbook hasn't been saved yet, use Desktop
    If wbPath = "" Then
        wbPath = Environ("HOME") & "/Desktop"
    End If

    ' Build full PDF path
    savePath = wbPath & "/" & wbName & ".pdf"

    ' Confirm with user
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("Export workbook to PDF?" & vbCrLf & vbCrLf & _
                          savePath, _
                          vbOKCancel + vbQuestion, "Export to PDF")

    ' User cancelled
    If userResponse <> vbOK Then Exit Sub

    ' Store original settings for all visible worksheets
    settingsCount = 0
    ReDim originalSettings(1 To visibleSheets)

    Application.ScreenUpdating = False

    ' Store original settings and configure for PDF export
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            settingsCount = settingsCount + 1

            ' Store original settings
            With originalSettings(settingsCount)
                On Error Resume Next
                .PrintArea = ws.PageSetup.PrintArea
                .FitToPagesWide = ws.PageSetup.FitToPagesWide
                .FitToPagesTall = ws.PageSetup.FitToPagesTall
                .Zoom = ws.PageSetup.Zoom
                On Error GoTo ErrorHandler
            End With

            ' Configure for one page per sheet with used range
            ConfigureSheetForPDF ws
        End If
    Next ws

    ' Export to PDF
    On Error Resume Next
    ActiveWorkbook.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' Check if export succeeded
    If Err.Number <> 0 Then
        MsgBox "Error exporting to PDF: " & Err.Description, vbCritical, "PDF Export Error"
        On Error GoTo ErrorHandler
    Else
        ' Restore original settings
        settingsCount = 0
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Visible = xlSheetVisible Then
                settingsCount = settingsCount + 1
                RestoreSheetSettings ws, originalSettings(settingsCount)
            End If
        Next ws

        Application.ScreenUpdating = True

        ' Show success message
        MsgBox "Workbook exported to PDF successfully!" & vbCrLf & vbCrLf & _
               CStr(savePath), vbInformation, "Export Complete"
    End If

    Exit Sub

ErrorHandler:
    ' Restore original settings even if export failed
    If settingsCount > 0 Then
        Dim wsCount As Long
        wsCount = 0
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Visible = xlSheetVisible Then
                wsCount = wsCount + 1
                If wsCount <= settingsCount Then
                    On Error Resume Next
                    RestoreSheetSettings ws, originalSettings(wsCount)
                    On Error GoTo 0
                End If
            End If
        Next ws
    End If

    Application.ScreenUpdating = True
    MsgBox "Error exporting to PDF: " & Err.Description, vbCritical, "PDF Export Error"
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Sub ConfigureSheetForPDF(ws As Worksheet)
    ' Configure worksheet for PDF export:
    ' - Set print area to used range (auto-detect data boundaries)
    ' - Fit to one page (width and height)

    On Error Resume Next

    With ws.PageSetup
        ' Clear any existing print area
        .PrintArea = ""

        ' Set print area to used range (only exports cells with content)
        ' UsedRange automatically finds the last row and column with data
        If Not ws.UsedRange Is Nothing Then
            ' Only set print area if there's actual content
            If ws.UsedRange.Address <> "$A$1" Or ws.Range("A1").Value <> "" Then
                .PrintArea = ws.UsedRange.Address
            End If
        End If

        ' Fit to one page (width and height)
        .Zoom = False  ' Must set to False to enable FitToPages
        .FitToPagesWide = 1
        .FitToPagesTall = 1

        ' Optional: Set orientation to automatic or landscape for better fit
        ' .Orientation = xlLandscape  ' Uncomment if you want landscape by default
    End With

    On Error GoTo 0
End Sub

Private Sub RestoreSheetSettings(ws As Worksheet, settings As PageSettings)
    ' Restore original PageSetup settings

    On Error Resume Next

    With ws.PageSetup
        .PrintArea = settings.PrintArea

        ' Restore zoom/fit settings
        If IsNumeric(settings.Zoom) Then
            If settings.Zoom > 0 Then
                .Zoom = settings.Zoom
            Else
                .Zoom = False
                If Not IsEmpty(settings.FitToPagesWide) Then .FitToPagesWide = settings.FitToPagesWide
                If Not IsEmpty(settings.FitToPagesTall) Then .FitToPagesTall = settings.FitToPagesTall
            End If
        End If
    End With

    On Error GoTo 0
End Sub
