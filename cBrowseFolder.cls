VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrowseFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ===========================================================
' Name:     cBrowseFolder.cls
' Author:   Ravi Kumar  (bugfree@sify.com)
' Date:     23 April 2005
'
' Requires: - mBrowseFolder.bas
' Copyright © 2005-2006 Ravi Kumar (bugfree@sify.com)
' ============================================================
'
' Add [+] : Remove [-] : Fixed [*]
' 23 April: First Release
' 26 April: [*] : Initial Directory into BrowseCallbackProc
'               : BFFM_SELCHANGED into BrowseCallbackProc
' 27 April: [+] : Position,Left,Top Properties
'           [*] : Some minor bugs
'               : Last Revision [Final Release]
'-------------------------------------------------------------
' TO DO:
' [+] : Validation field
' [+] : Set any path to root [don't know how to convert path
'     : to pidl
' [*] : Some minor problems [like variable declaration in bas
'     : module]
'=============================================================
Option Explicit

'-- API's------------
Private Const NO_ERROR As Long = 0

Private Const BIF_EDITBOX As Long = &H10
Private Const BIF_VALIDATE As Long = &H20
Private Const BIF_USENEWUI As Long = &H40
Private Const BIF_STATUSTEXT As Long = &H4
Private Const BIF_SHAREABLE As Long = &H8000
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_BROWSEINCLUDEURLS As Long = &H80
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfnCallBack As Long
  lParam As Long
  iImage As Long
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" ( _
   ByVal lpString1 As String, _
   ByVal lpString2 As String) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
   ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   ByRef lParam As Any) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" ( _
   ByRef lpbi As BROWSEINFO) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" ( _
   ByVal hwndOwner As Long, _
   ByVal nFolder As Long, _
   pidl As ITEMIDLIST) As Long
Private Declare Sub SHGetDesktopFolder Lib "shell32.dll" ( _
   ByRef ppshf As Long)
Private Declare Sub CoTaskMemFree Lib "ole32.dll" ( _
   ByRef pv As Any)
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Enum SFConstants
   SF_DESKTOP = &H0                   '(desktop)
   SF_INTERNET = &H1                  'Internet Explorer (icon on desktop)
   SF_PROGRAMS = &H2                  'Start Menu\Programs
   SF_CONTROLS = &H3                  'My Computer\Control Panel
   SF_PRINTERS = &H4                  'My Computer\Printers
   SF_PERSONAL = &H5                  'My Documents
   SF_FAVORITES = &H6                 '(user name)\Favorites
   SF_STARTUP = &H7                   'Start Menu\Programs\Startup
   SF_RECENT = &H8                    '(user name)\Recent
   SF_SENDTO = &H9                    '(user name)\SendTo
   SF_BITBUCKET = &HA                 '(desktop)\Recycle Bin
   SF_STARTMENU = &HB                 '(user name)\Start Menu
   SF_DESKTOPDIRECTORY = &H10         '(user name)\Desktop
   SF_DRIVES = &H11                   'My Computer
   SF_NETWORK = &H12                  'Network Neighborhood
   SF_NETHOOD = &H13                  '(user name)\nethood
   SF_FONTS = &H14                    'windows\fonts
   SF_TEMPLATES = &H15
   SF_COMMON_STARTMENU = &H16         'All Users\Start Menu
   SF_COMMON_PROGRAMS = &H17          'All Users\Programs
   SF_COMMON_STARTUP = &H18           'All Users\Startup
   SF_COMMON_DESKTOPDIRECTORY = &H19  'All Users\Desktop
   SF_APPDATA = &H1A                  '(user name)\Application Data
   SF_PRINTHOOD = &H1B                '(user name)\PrintHood
   SF_LOCAL_APPDATA = &H1C            '(user name)\Local Settings\Applicaiton Data (non roaming)
   SF_ALTSTARTUP = &H1D               'non localized startup
   SF_COMMON_ALTSTARTUP = &H1E        'non localized common startup
   SF_COMMON_FAVORITES = &H1F
   SF_INTERNET_CACHE = &H20
   SF_COOKIES = &H21
   SF_HISTORY = &H22
   SF_COMMON_APPDATA = &H23           'All Users\Application Data
   SF_WINDOWS = &H24                  'GetWindowsDirectory()
   SF_SYSTEM = &H25                   'GetSystemDirectory()
   SF_PROGRAM_FILES = &H26            'C:\Program Files
   SF_MYPICTURES = &H27               'C:\Program Files\My Pictures
   SF_PROFILE = &H28                  'USERPROFILE
   SF_PROGRAM_FILES_COMMON = &H2B     'C:\Program Files\Common
   SF_COMMON_TEMPLATES = &H2D         'All Users\Templates
   SF_COMMON_DOCUMENTS = &H2E         'All Users\Documents
   SF_COMMON_ADMINTOOLS = &H2F        'All Users\Start Menu\Programs\Administrative Tools
   SF_ADMINTOOLS = &H30               '(user name)\Start Menu\Programs\Administrative Tools
End Enum

Public Enum BFPositionConstants
  [bfUseDefault] = 0
  [bfUseCenterOwner] = 1
  [bfUseCenterScreen] = 2
  [bfUseCustom] = 3
End Enum

'-- Private Variables
Private m_hWndOwner As Long
Private m_sTitle As String
Private m_sInitialDir As String
Private m_eRootDir As SFConstants
Private m_sDisplayName As String
Private m_bEditBox As Boolean
Private m_bStatusText As Boolean
Private m_bFileSystemOnly As Boolean
Private m_bUseNewUI As Boolean
Private m_ePosition As BFPositionConstants

'-- Functions
Public Function BrowseFolder() As String
  Dim lPidlList As Long
  Dim sBuffer As String
  Dim tBI As BROWSEINFO
  Dim sDisplay As String
  Dim sTitle As String
  Dim lRet As Long, lPidl As ITEMIDLIST
  Dim hInst As Long, lThread As Long
  
  BrowseFolder = vbNullString
  
  If (hwndOwner And Not m_ePosition = bfUseDefault) Then
    hInst = GetWindowLong(IIf(hwndOwner, hwndOwner, 0&), GWL_HINSTANCE)
    lThread = GetCurrentThreadId()
    If (m_ePosition = bfUseCenterOwner) Then lHook = SetWindowsHookEx(WH_CBT, AddressOf CenterForm, hInst, lThread)
    If (m_ePosition = bfUseCenterScreen) Then lHook = SetWindowsHookEx(WH_CBT, AddressOf CenterScreen, hInst, lThread)
    If (m_ePosition = bfUseCustom) And (lpLeft) And (lpTop) Then lHook = SetWindowsHookEx(WH_CBT, AddressOf CenterCustom, hInst, lThread)
  End If
   
  With tBI
    .hOwner = IIf(hwndOwner, hwndOwner, 0&)
    sDisplay = Space$(MAX_PATH)
    .pszDisplayName = sDisplay
    .lpszTitle = IIf(Len(m_sTitle), m_sTitle, "Choose Folder")
    .ulFlags = plGetOptions()
    .iImage = 0
    lRet = SHGetSpecialFolderLocation(IIf(hwndOwner, hwndOwner, 100), RootDir, lPidl)
    .pidlRoot = lPidl.mkid.cb
    '.lParam
    .lpfnCallBack = GetAddressOfFunction(AddressOf BrowseCallbackProc)
  End With
  
  lPidlList = SHBrowseForFolder(tBI)
  
  If (lPidlList) Then
    sBuffer = Space$(MAX_PATH)
    SHGetPathFromIDList lPidlList, sBuffer
    sBuffer = VBA.Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    m_sDisplayName = VBA.Left$(tBI.pszDisplayName, InStr(tBI.pszDisplayName, vbNullChar) - 1)
    BrowseFolder = sBuffer
  End If
End Function

Private Function plGetOptions() As Long
Dim lOpt As Long
   If m_bEditBox Then lOpt = BIF_EDITBOX Or BIF_VALIDATE
   If m_bStatusText Then lOpt = lOpt Or BIF_STATUSTEXT
   If m_bFileSystemOnly Then lOpt = lOpt Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN
   If m_bUseNewUI Then lOpt = lOpt Or BIF_USENEWUI
   plGetOptions = lOpt
End Function

Private Sub Class_Initialize()
  'FileSystemOnly = True
End Sub

'-- Properties
Public Property Get EditBox() As Boolean
  EditBox = m_bEditBox
End Property
Public Property Let EditBox(ByVal bState As Boolean)
  m_bEditBox = bState
End Property

Public Property Get StatusText() As Boolean
  StatusText = m_bStatusText
End Property
Public Property Let StatusText(ByVal bState As Boolean)
  m_bStatusText = bState
End Property

Public Property Get FileSystemOnly() As Boolean
  FileSystemOnly = m_bFileSystemOnly
End Property
Public Property Let FileSystemOnly(ByVal bState As Boolean)
  m_bFileSystemOnly = bState
End Property

Public Property Get UseNewUI() As Boolean
  UseNewUI = m_bUseNewUI
End Property
Public Property Let UseNewUI(ByVal bState As Boolean)
  m_bUseNewUI = bState
End Property

Public Property Get Title() As String
   Title = m_sTitle
End Property
Public Property Let Title(ByVal sTitle As String)
   m_sTitle = sTitle
End Property

Public Property Get InitialDir() As String
   InitialDir = m_sInitialDir
End Property
Public Property Let InitialDir(ByVal sDir As String)
   m_sInitialDir = sDir
   sInitDir = sDir
End Property

Public Property Get RootDir() As SFConstants  ' String
   RootDir = m_eRootDir
End Property
Public Property Let RootDir(ByVal sDir As SFConstants)  'String)
   m_eRootDir = sDir
End Property

Public Property Get hwndOwner() As Long
   hwndOwner = m_hWndOwner
End Property
Public Property Let hwndOwner(ByVal lHwnd As Long)
   m_hWndOwner = lHwnd
   lHwndOwner = lHwnd
End Property

Public Property Get DisplayName() As String
  DisplayName = m_sDisplayName
End Property

Public Property Let Left(ByVal vNewLeft As Long)
  lpLeft = vNewLeft
End Property
Public Property Let Top(ByVal vNewTop As Long)
  lpTop = vNewTop
End Property

Public Property Let Position(ByVal vNewPosition As BFPositionConstants)
  m_ePosition = vNewPosition
End Property

Public Property Get SpecialFolder(ByVal eFolder As SFConstants) As String
  Dim lPidl As ITEMIDLIST
  Dim sPath As String, lRet As Long
  On Error Resume Next
  SpecialFolder = vbNullString
  lRet = SHGetSpecialFolderLocation(100, eFolder, lPidl)
  If lRet = NO_ERROR Then
    sPath = Space$(MAX_PATH)
    lRet = SHGetPathFromIDList(ByVal lPidl.mkid.cb, ByVal sPath)
    SpecialFolder = VBA.Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    CoTaskMemFree lPidl
  End If
End Property
