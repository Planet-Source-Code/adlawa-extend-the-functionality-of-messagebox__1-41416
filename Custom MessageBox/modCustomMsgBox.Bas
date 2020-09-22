Attribute VB_Name = "modCustomMsgBox"
'=========================================================================================
'  modCustomMsgBox
'  routines and functions to allow customization to a msgbox
'=========================================================================================
'  Created By:Amer
'  Published Date: 02/23/2001
'  Legal Copyright: Amer © 02/23/2001
'=========================================================================================
Option Explicit

Dim MsgBoxTitle As String
Dim MsgBoxText As String
Dim MsgBoxX As Long
Dim MsgBoxY As Long
Dim MsgBoxTimeOut As Integer
Dim MsgBoxHWND As Long
Dim MsgBoxButtons As Long
Dim MsgBoxCheckBoxText As String
Dim MsgBoxCheckBoxHWND As Long
Dim MsgBoxCheckBoxValue As Boolean
Dim MsgBoxButtonCount As Integer
Dim MsgBoxButtonText() As String
Dim MsgBoxOriginalButtonText() As String

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const MB_CLOSEMSGBOX = &H5000&
Private Const MB_MOVEMSGBOX = &H5001&
Private Const MB_CHECK_CHECKBOXVALUE = &H5003&
Private Const MB_CHANGEBUTTONTEXT = &H5002&
Private Const MB_ADDCHECKBOX = &H5004&
Private Const SWP_NOSIZE = &H1
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_TABSTOP = &H10000
Private Const WM_SETFONT = &H30
Private Const BS_AUTOCHECKBOX = &H3&
Private Const HWND_TOPMOST = -1
Private Const BM_GETSTATE = &HF2
Private Const WM_GETFONT = &H31

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal bRepaint&) As Long
Private Declare Function GetFocus Lib "user32" () As Long
'=========================================================================================
Public Function CreateSpecialMsgbox(Text As String, Buttons As Long, Title As String, Optional ByVal TimeOut As Integer = -1, Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1, Optional CheckBoxText As String = "", Optional CheckBoxValue As Boolean, Optional ByVal ButtonCount As Integer, Optional ButtonText As String, Optional OriginalButtonText As String) As VbMsgBoxResult
' routine to create the special msgbox based on passed parameters
  On Error Resume Next
  MsgBoxHWND = Screen.ActiveForm.hwnd
  MsgBoxText = Text
  MsgBoxTitle = Title
  MsgBoxButtons = Buttons
  MsgBoxTimeOut = TimeOut * 1000
  MsgBoxX = X
  MsgBoxY = Y
  MsgBoxCheckBoxText = CheckBoxText
  MsgBoxButtonCount = ButtonCount
  MsgBoxButtonText = Split(ButtonText, ",")
  MsgBoxOriginalButtonText = Split(OriginalButtonText, ",")

  'set the timer to fire for various effects
  If MsgBoxX <> -1 And MsgBoxY <> -1 Then
    SetTimer MsgBoxHWND, MB_MOVEMSGBOX, 0&, AddressOf NewTimerProc
  End If
  If MsgBoxTimeOut > 0 Then
    SetTimer MsgBoxHWND, MB_CLOSEMSGBOX, MsgBoxTimeOut, AddressOf NewTimerProc
  End If
  If ButtonText <> "" And OriginalButtonText <> "" Then
    SetTimer MsgBoxHWND, MB_CHANGEBUTTONTEXT, 0&, AddressOf NewTimerProc
  End If
  If MsgBoxCheckBoxText <> "" Then
    SetTimer MsgBoxHWND, MB_CHECK_CHECKBOXVALUE, 0&, AddressOf NewTimerProc
    SetTimer MsgBoxHWND, MB_ADDCHECKBOX, 0&, AddressOf NewTimerProc
  End If

  ' creat the special Msgbox
  CreateSpecialMsgbox = MessageBox(MsgBoxHWND, MsgBoxText, MsgBoxTitle, MsgBoxButtons)

  If MsgBoxCheckBoxText <> "" Then
    'cancel the checkbox fire timer
    MsgBoxCheckBoxHWND = 0
    KillTimer MsgBoxHWND, MB_CHECK_CHECKBOXVALUE
    'return its state
    CheckBoxValue = MsgBoxCheckBoxValue
  End If
End Function 'CreateSpecialMsgbox(Text As String, Buttons As Long, Title As String, Optional ByVal TimeOut As Integer = -1, Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1, Optional ByVal CheckBoxText As String = "", Optional ByRef CheckBoxValue As Boolean, Optional ByVal ButtonCount As Integer, Optional ButtonText As String, Optional OriginalButtonText As String) As VbMsgBoxResult
'=========================================================================================
Public Function NewTimerProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' routine to fire for timer events which determine what should be done for the custom msgbox
  Dim TempForm As Form
  On Error Resume Next
  If wParam = MB_CHECK_CHECKBOXVALUE Then
    If MsgBoxCheckBoxHWND > 0 Then
      'Returns the value of the checkbox on extended MsgBox
      MsgBoxCheckBoxValue = (SendMessage(MsgBoxCheckBoxHWND, BM_GETSTATE, 0, 0&) <> 0)
    Else
      KillTimer hwnd, wParam
    End If
  Else
    'Cancel timer
    KillTimer hwnd, wParam
    MsgBoxHWND = FindWindow("#32770", MsgBoxTitle)
    If MsgBoxHWND <> 0 Then
      Select Case wParam
        Case MB_CLOSEMSGBOX
          SetForegroundWindow MsgBoxHWND
          If MsgBoxCheckBoxText <> "" Then
            If MsgBoxCheckBoxHWND = GetFocus&() Then
              If MsgBoxButtonCount = 1 Then SendKeys "{TAB}"
              If MsgBoxButtonCount > 1 Then SetForegroundWindow MsgBoxHWND
            End If
          End If
          SendKeys "{enter}"
        Case MB_MOVEMSGBOX
          Dim W As Single
          Dim H As Single
          Dim rBox As RECT
          W = Screen.Width / Screen.TwipsPerPixelX
          H = Screen.Height / Screen.TwipsPerPixelY
          GetWindowRect MsgBoxHWND, rBox
          With rBox
            If MsgBoxX > (W - (.Right - .Left) - 1) Then MsgBoxX = (W - (.Right - .Left) - 1)
            If MsgBoxY > (H - (.Bottom - .Top) - 1) Then MsgBoxY = (H - (.Bottom - .Top) - 1)
          End With
          If MsgBoxX < 1 Then MsgBoxX = 1: If MsgBoxY < 1 Then MsgBoxY = 1
          SetWindowPos MsgBoxHWND, HWND_TOPMOST, MsgBoxX, MsgBoxY, 0, 0, SWP_NOSIZE
        Case MB_ADDCHECKBOX
          Dim hFont As Long
          Dim lCaptionHwnd As Long
          Dim R As RECT
          Dim nHeight As Integer
          'Find the window
          lCaptionHwnd = FindWindowEx(MsgBoxHWND, 0, "Static", MsgBoxText)
          GetWindowRect MsgBoxHWND, R
          Set TempForm = Screen.ActiveForm
          nHeight = TempForm.TextHeight(MsgBoxCheckBoxText) / Screen.TwipsPerPixelY
          'Create the checkbox control
          MsgBoxCheckBoxHWND = CreateWindowEx(0, "Button", MsgBoxCheckBoxText, WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_AUTOCHECKBOX, 3, (R.Bottom - R.Top) - nHeight - 15, (TempForm.TextWidth(MsgBoxCheckBoxText) / Screen.TwipsPerPixelX) + 22, nHeight, MsgBoxHWND, 0, App.hInstance, ByVal 0&)
          ' set the font of the checkbox to the same as the messagebox
          hFont = SendMessage(lCaptionHwnd, WM_GETFONT, 0, 0&)
          SendMessage MsgBoxCheckBoxHWND, WM_SETFONT, hFont, 0&
          'move the new checkbox to the correct position
          MoveWindow MsgBoxHWND, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top + nHeight, 1&
        Case MB_CHANGEBUTTONTEXT
          Dim hButton As Long
          Dim Counter As Integer
          'replace the captions where required
          For Counter = LBound(MsgBoxOriginalButtonText) To UBound(MsgBoxOriginalButtonText)
            If Len(MsgBoxButtonText(Counter)) > 0 Then
              hButton = FindWindowEx(MsgBoxHWND, 0&, "Button", MsgBoxOriginalButtonText(Counter))
              If hButton <> 0 Then
                SetWindowText hButton, MsgBoxButtonText(Counter)
              End If
            End If
          Next Counter
      End Select
    End If
  End If
  'clear objects
  Set TempForm = Nothing
End Function 'NewTimerProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'=========================================================================================

