Attribute VB_Name = "mBrowseFolder"
Option Explicit

' ===========================================================
' Name:     mBrowseFolder.bas
' Author:   Ravi Kumar  (bugfree@sify.com)
' Date:     23 April 2005
'
' Requires: - None
' Copyright © 2005-2006 Ravi Kumar (bugfree@sify.com)
' ============================================================

Public Const MAX_PATH As Long = 260
Private Const WM_USER As Long = &H400

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_VALIDATEFAILEDA As Long = 3
Private Const BFFM_VALIDATEFAILEDW As Long = 4
Public Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)

Public Const GWL_HINSTANCE = (-6)
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const HCBT_ACTIVATE = 5
Public Const WH_CBT = 5

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Declare Function UnhookWindowsHookEx Lib "user32" ( _
   ByVal hHook As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias _
   "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) _
   As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias _
   "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, _
   ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function SetWindowPos Lib "user32" ( _
   ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
   ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
   ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd _
   As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" ( _
   ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Public sInitDir As String, lHook As Long, lHwndOwner As Long
Public lpLeft As Long, lpTop As Long

Public Function CenterForm( _
  ByVal lMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long) As Long

Dim rectForm As RECT, rectMsg As RECT
Dim x As Long, y As Long

   'On HCBT_ACTIVATE, show the DialogBox centered over Form1
   If lMsg = HCBT_ACTIVATE Then
      'Get the coordinates of the form and the DialogBox so that
      'you can determine where the center of the form is located
      GetWindowRect lHwndOwner, rectForm
      GetWindowRect wParam, rectMsg
      x = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - _
          ((rectMsg.Right - rectMsg.Left) / 2)
      y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - _
          ((rectMsg.Bottom - rectMsg.Top) / 2)
      'Position the DialogBox
      SetWindowPos wParam, 0, x, y, 0, 0, _
                   SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
      'Release the CBT hook
      UnhookWindowsHookEx lHook
   End If
   CenterForm = False
End Function

Public Function CenterScreen( _
  ByVal lMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long) As Long

Dim rectForm As RECT, rectMsg As RECT
Dim x As Long, y As Long

   'On HCBT_ACTIVATE, show the DialogBox centered over Form1
   If lMsg = HCBT_ACTIVATE Then
      'Get the coordinates of the form and the DialogBox so that
      'you can determine where the center of the form is located
      rectForm.Left = 0: rectForm.Top = 0: rectForm.Right = Screen.Width \ Screen.TwipsPerPixelX: rectForm.Bottom = Screen.Height \ Screen.TwipsPerPixelY
      GetWindowRect wParam, rectMsg
      x = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - _
          ((rectMsg.Right - rectMsg.Left) / 2)
      y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - _
          ((rectMsg.Bottom - rectMsg.Top) / 2)
      'Position the DialogBox
      SetWindowPos wParam, 0, x, y, 0, 0, _
                   SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
      'Release the CBT hook
      UnhookWindowsHookEx lHook
   End If
   CenterScreen = False
End Function

Public Function CenterCustom( _
  ByVal lMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long) As Long

Dim rectForm As RECT, rectMsg As RECT
Dim x As Long, y As Long

   'On HCBT_ACTIVATE, show the DialogBox centered over Form1
   If lMsg = HCBT_ACTIVATE Then
      'Position the DialogBox
      SetWindowPos wParam, 0, lpLeft, lpTop, 0, 0, _
                   SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
      'Release the CBT hook
      UnhookWindowsHookEx lHook
   End If
   CenterCustom = False
End Function

Public Function BrowseCallbackProc( _
      ByVal hWnd As Long, _
      ByVal uMsg As Long, _
      ByVal lParam As Long, _
      ByVal lpData As Long) As Long

  Dim lPidl As Long, lRet As Long
  Dim sBuffer  As String
  Dim cBF As cBrowseFolder
  
  Debug.Print hWnd, uMsg, lParam, lpData
    sBuffer = Space$(MAX_PATH)
    Select Case uMsg
        'Set the initial selection
        Case BFFM_INITIALIZED
          Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, sInitDir)
        Case BFFM_SELCHANGED
        'Set the selection status of path
          If (SHGetPathFromIDList(lParam, sBuffer)) Then _
            Call SendMessage(hWnd, BFFM_SETSTATUSTEXTA, 0&, sBuffer)
        Case BFFM_VALIDATEFAILEDA
         'To Do
    End Select
    BrowseCallbackProc = 0
End Function

Public Function GetAddressOfFunction(Address As Long) As Long
    GetAddressOfFunction = Address
End Function
