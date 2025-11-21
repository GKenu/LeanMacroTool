VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zFPrecedentAnalyzer
   Caption         =   "LeanMacro Dependant Tracer"
   ClientHeight    =   5805
   ClientLeft      =   40
   ClientTop       =   440
   ClientWidth     =   5460
   OleObjectBlob   =   "zFPrecedentAnalyzer.frx":0000
End
Attribute VB_Name = "zFPrecedentAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

'' Required Win32 API Declarations
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
'Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
'Private Declare Function GetForegroundWindow Lib "user32" () As Long
'Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
''
'' Constants used with APIs
'Private Const SW_SHOW = 5
'Private Const SW_RESTORE = 9

Private mbJump As Boolean
Private msFullAddress As String
Private msAddress As String
Private miListIndex As Long
Private mrActive As Range

Private WithEvents mWS As Worksheet
Attribute mWS.VB_VarHelpID = -1
Private WithEvents mWB As Workbook
Attribute mWB.VB_VarHelpID = -1

Public Property Set Cell(rCell As Range)
  Dim sGetAddress As String

  Set mWS = rCell.Parent
  Set mWB = mWS.Parent

  sGetAddress = GetAddress(rCell)
  msFullAddress = Left$(sGetAddress, InStr(sGetAddress, "|") - 1)
  msAddress = Mid$(sGetAddress, InStr(sGetAddress, "|") + 1)

  Me.lblCellAddress.Caption = "  " & msAddress
  Me.lblCellValue.Caption = "  " & rCell.Text
  Me.txtCellFormula.Text = rCell.formula

  PopulatePrecedents rCell

End Property

Private Sub btnClose_Click()
  Set mWS = Nothing
  Set mWB = Nothing
  Me.Hide
End Sub

Private Sub lblCellAddress_DblClick(ByVal Cancel As msforms.ReturnBoolean)
  VisitRange Me.lblCellAddress.Caption
End Sub

Private Sub lstPrecedents_Change()
  If mbJump Then
    If Me.lstPrecedents.ListIndex >= 0 Then
      VisitRange Me.lstPrecedents.List(Me.lstPrecedents.ListIndex, 0), False
    End If
  End If
End Sub

Private Sub lstPrecedents_DblClick(ByVal Cancel As msforms.ReturnBoolean)
  VisitRange Me.lstPrecedents.List(Me.lstPrecedents.ListIndex, 0)
End Sub

Private Sub lstPrecedents_KeyDown(ByVal KeyCode As msforms.ReturnInteger, ByVal Shift As Integer)
  If KeyCode = 13 Then
    If Me.lstPrecedents.ListIndex >= 0 Then
      VisitRange Me.lstPrecedents.List(Me.lstPrecedents.ListIndex, 0)
    End If
  ElseIf KeyCode = vbKeyHome And Shift = 6 Then
    Application.OnKey "^*{HOME}", "'" & ThisWorkbook.Name & "'!ActivatePrecedentAnalyzer"
    miListIndex = Me.lstPrecedents.ListIndex
    Set mrActive = ActiveCell
    On Error Resume Next
    AppActivate Application.Caption, False
    'MoveExcelToFront
    Beep
    Application.OnTime Now + TimeValue("0:0:1"), "'" & ThisWorkbook.Name & "'!ResetTracerDialogSettings"
    ''Application.ScreenUpdating = False
  End If
End Sub

Private Sub mWB_Activate()
  Dim iList As Long

  mbJump = False
  Me.lblCellAddress.Caption = "  " & msAddress
  With Me.lstPrecedents
    If .ListCount > 0 Then
      If bOriginalCellAtEnd Then
        For iList = 0 To .ListCount - 2
          .List(iList, 2) = .List(iList, 1)
        Next
        .List(.ListCount - 1, 2) = .List(.ListCount - 1, 1) & "     original cell"
      Else
        For iList = 1 To .ListCount - 1
          .List(iList, 2) = .List(iList, 1)
        Next
        .List(0, 2) = .List(0, 1) & "     original cell"
      End If
    End If
  End With
  mbJump = True
End Sub

Private Sub mWB_Deactivate()
  Dim iList As Long

  mbJump = False
  Me.lblCellAddress.Caption = "  " & msFullAddress
  With Me.lstPrecedents
    If .ListCount > 0 Then
      If bOriginalCellAtEnd Then
        For iList = 0 To .ListCount - 2
          .List(iList, 2) = .List(iList, 0)
        Next
        .List(.ListCount - 1, 2) = .List(.ListCount - 1, 0) & "     original cell"
      Else
        For iList = 1 To .ListCount - 1
          .List(iList, 2) = .List(iList, 0)
        Next
        .List(0, 2) = .List(0, 0) & "     original cell"
      End If
    End If
  End With
  mbJump = True
End Sub

Private Sub mWS_Calculate()
  Dim sAddress As String
  Dim rCell As Range
  Dim sPrecedent As String

  sAddress = Me.lblCellAddress.Caption
  If IsRange(sAddress) Then
    Set rCell = Range(sAddress)
    Me.lblCellValue.Caption = "  " & rCell.Text
    Me.txtCellFormula.Text = rCell.formula
    Me.lblCellValue.Caption = "  " & rCell.Text
  End If

  sPrecedent = Me.lstPrecedents.Value
  PopulatePrecedents rCell
  Me.lstPrecedents.Value = sPrecedent

End Sub

Private Sub PopulatePrecedents(rngCell As Range)
  Dim vPrecedents As Variant
  Dim iList As Long

  mbJump = False

  vPrecedents = NewPrecedents(rngCell)
  If IsArray(vPrecedents) Then
    With Me.lstPrecedents
      .List = vPrecedents
      '      If .ListIndex < 0 Then
      '        .ListIndex = 0
      '      End If
      .SetFocus
      If bOriginalCellAtEnd Then
        'If .ListIndex < 0 Then
        .ListIndex = .ListCount - 1
        'End If
        For iList = 0 To .ListCount - 2
          .List(iList, 2) = .List(iList, 1)
        Next
        .List(.ListCount - 1, 2) = .List(.ListCount - 1, 1) & "     original cell"
        .List(.ListCount - 1, 2) = .List(.ListCount - 1, 1) & "     original cell"
      Else
        'If .ListIndex < 0 Then
        .ListIndex = 0
        'End If
        For iList = 1 To .ListCount - 1
          .List(iList, 2) = .List(iList, 1)
        Next
        .List(0, 2) = .List(0, 1) & "     original cell"
        .List(0, 2) = .List(0, 1) & "     original cell"
      End If
    End With
    'VisitRange Me.lstPrecedents.List(Me.lstPrecedents.ListIndex, 0)
  Else
    Me.lstPrecedents.Clear
  End If

  Me.lstPrecedents.SetFocus
  mbJump = True

End Sub

Private Sub UserForm_Activate()
    lstPrecedents.Width = 277
End Sub

Private Sub UserForm_Deactivate()
  Application.OnKey "^*{HOME}", "'" & ThisWorkbook.Name & "'!ActivatePrecedentAnalyzer"
End Sub

Private Sub UserForm_Initialize()
  Me.Left = Application.Left + Application.Width - Me.Width - 12
  Me.Top = Application.Top + Application.Height - Me.Height - 12
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As msforms.ReturnInteger, ByVal Shift As Integer)
  If KeyCode = vbKeyHome And Shift = 6 Then
    'Me.lstPrecedents.SetFocus
    Application.OnKey "^*{HOME}", "'" & ThisWorkbook.Name & "'!ActivatePrecedentAnalyzer"
    miListIndex = Me.lstPrecedents.ListIndex
    Set mrActive = ActiveCell
    On Error Resume Next
    AppActivate Application.Caption
    'MoveExcelToFront
    Beep
    Application.OnTime Now + TimeValue("0:0:1"), "'" & ThisWorkbook.Name & "'!ResetTracerDialogSettings"
    ''Application.ScreenUpdating = False
  End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    btnClose_Click
  End If
End Sub

Private Sub VisitRange(sRangeAddress As String, Optional bAppActivate As Boolean = True)

  On Error GoTo ErrorTrap

  Me.lstPrecedents.SetFocus

  Application.Goto Range(sRangeAddress)

  If ActiveWorkbook.Name = mWB.Name Then
    mWB_Activate
  Else
    mWB_Deactivate
  End If

  If err.Number = 0 Then
    If bAppActivate Then
      AppActivate Application.Caption
    End If
  Else
    MsgBox "Error accessing '" & sRangeAddress & "'."
  End If

ExitProcedure:
  Exit Sub

ErrorTrap:
  If err.Description = "Method 'Goto' of object '_Application' failed" Then
    MsgBox "Cannot activate a range on a hidden worksheet.   ", vbOKOnly + vbExclamation, TTS_TITLE
    GoTo ExitProcedure
  End If
  Resume
End Sub

Public Sub ResetListIndex()
  mbJump = False
  Me.lstPrecedents.ListIndex = miListIndex
  Application.Goto mrActive
  mbJump = True
End Sub

'Private Sub MoveExcelToFront()
'  Dim xlHwnd As Long
'
'  Debug.Print "w " & GetForegroundWindow
'  xlHwnd = FindWindow(vbNullString, Application.Caption)
'  Debug.Print "x " & xlHwnd
'  'ForceForegroundWindow xlHwnd
'  Debug.Print "y " & GetForegroundWindow
'  SetForegroundWindow xlHwnd
'  Debug.Print "z " & GetForegroundWindow
'
'End Sub

'Private Function ForceForegroundWindow(ByVal hWnd As Long) As Boolean
'   Dim ThreadID1 As Long
'   Dim ThreadID2 As Long
'   Dim nRet As Long
'   '
'   ' Nothing to do if already in foreground.
'   '
'
'   Debug.Print "A " & hWnd
'   Debug.Print "A " & GetForegroundWindow
'
'   If hWnd = GetForegroundWindow() Then
'      ForceForegroundWindow = True
'   Else
'      '
'      ' First need to get the thread responsible for this window,
'      ' and the thread for the foreground window.
'      '
'      ThreadID1 = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
'      ThreadID2 = GetWindowThreadProcessId(hWnd, ByVal 0&)
'      '
'      ' By sharing input state, threads share their concept of
'      ' the active window.
'      '
'      DoEvents
'      If ThreadID1 <> ThreadID2 Then
'         Call AttachThreadInput(ThreadID1, ThreadID2, True)
'         nRet = SetForegroundWindow(hWnd)
'         Call AttachThreadInput(ThreadID1, ThreadID2, False)
'      Else
'         nRet = SetForegroundWindow(hWnd)
'      End If
'
'      Debug.Print "B " & GetForegroundWindow
'      '
'      ' Restore and repaint
'      '
'      DoEvents
'      If IsIconic(hWnd) Then
'         Call ShowWindow(hWnd, SW_RESTORE)
'      Else
'         Call ShowWindow(hWnd, SW_SHOW)
'      End If
'      '
'      ' SetForegroundWindow return accurately reflects success.
'      '
'      Debug.Print "C " & GetForegroundWindow
'      ForceForegroundWindow = CBool(nRet)
'   End If
'End Function
