VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse for Folder Demonstration"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSFolder 
      Height          =   315
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   1440
      Width           =   3345
   End
   Begin VB.TextBox txtDisplayName 
      Height          =   315
      Left            =   1830
      TabIndex        =   26
      Top             =   600
      Width           =   5865
   End
   Begin VB.CommandButton cmdSpecialFolder 
      Caption         =   "Show Special Folder"
      Height          =   735
      Left            =   7890
      TabIndex        =   19
      Top             =   1050
      Width           =   1395
   End
   Begin VB.CommandButton cmdBrowseFolder 
      Caption         =   "Browse Folder"
      Height          =   615
      Left            =   7800
      TabIndex        =   18
      Top             =   240
      Width           =   1485
   End
   Begin VB.TextBox txtSpecialFolder 
      Height          =   315
      Left            =   1830
      TabIndex        =   17
      Top             =   1020
      Width           =   5865
   End
   Begin VB.TextBox txtFolderPath 
      Height          =   315
      Left            =   1830
      TabIndex        =   16
      Top             =   150
      Width           =   5865
   End
   Begin VB.Frame Frame1 
      Caption         =   "BrowseFolder Properties"
      Height          =   3255
      Left            =   90
      TabIndex        =   3
      Top             =   1890
      Width           =   9315
      Begin VB.ComboBox cboPosition 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1320
         Width           =   3585
      End
      Begin VB.ComboBox cboRootDir 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   840
         Width           =   3555
      End
      Begin VB.CommandButton cmdInitDir 
         Caption         =   "Select"
         Height          =   345
         Left            =   7710
         TabIndex        =   21
         Top             =   330
         Width           =   1065
      End
      Begin VB.TextBox txtInitDir 
         Height          =   345
         Left            =   1680
         TabIndex        =   20
         Top             =   330
         Width           =   5865
      End
      Begin VB.TextBox txtTop 
         Enabled         =   0   'False
         Height          =   345
         Left            =   7230
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtLeft 
         Enabled         =   0   'False
         Height          =   345
         Left            =   5880
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtTitle 
         Height          =   345
         Left            =   1680
         TabIndex        =   12
         Top             =   1800
         Width           =   2205
      End
      Begin VB.CheckBox chkProp 
         Caption         =   "Use New User Interface"
         Height          =   525
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   2670
         Width           =   2115
      End
      Begin VB.CheckBox chkProp 
         Caption         =   "Show Status Text"
         Height          =   525
         Index           =   3
         Left            =   4890
         TabIndex        =   7
         Top             =   2190
         Width           =   2115
      End
      Begin VB.CheckBox chkProp 
         Caption         =   "Use Owner HWND"
         Height          =   525
         Index           =   2
         Left            =   2580
         TabIndex        =   6
         Top             =   2640
         Width           =   2115
      End
      Begin VB.CheckBox chkProp 
         Caption         =   "Show File System Only"
         Height          =   525
         Index           =   1
         Left            =   2565
         TabIndex        =   5
         Top             =   2190
         Width           =   2115
      End
      Begin VB.CheckBox chkProp 
         Caption         =   "Show Edit Box"
         Height          =   525
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   2190
         Width           =   2115
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Top :"
         Height          =   195
         Left            =   6750
         TabIndex        =   25
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Left :"
         Height          =   195
         Left            =   5460
         TabIndex        =   24
         Top             =   1410
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Title :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label Label6 
         Caption         =   "Root Directory :"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Position :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Initial Directory :"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   1125
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note : Blank Special Folder means (VIRTUAL FOLDER)"
      Height          =   195
      Left            =   180
      TabIndex        =   28
      Top             =   1530
      Width           =   3960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Folder :"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Display Name :"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   690
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Folder Path :"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================
' BrowseFolder
' Ravi Kumar (bugfree) - 2005.12.24
'================================================
'
' UI for cBrowseFolder.cls
' Gives 1.  New UI look
'       2.  Position set:
'             Default
'             Center Owner
'             Center Screen
'             Custom (Left,Top)
'       3.  Special Folder Location
'       4.  Set Initial Directory
'
' Last revision: 2005.12.24
'================================================

Private m_BF As New cBrowseFolder
Private m_IDir As New cBrowseFolder

Private Sub cboPosition_Click()
  txtLeft.Enabled = False
  txtTop.Enabled = False
  
  If cboPosition.Text = cboPosition.List(3) Then
    txtLeft.Enabled = True:    txtTop.Enabled = True
  End If
End Sub

Private Sub cmdBrowseFolder_Click()
  On Error Resume Next
  With m_BF
    .InitialDir = txtInitDir
    .RootDir = cboRootDir.ItemData(cboRootDir.ListIndex)
    .Position = cboPosition.ListIndex
    If cboPosition.ListIndex = 3 Then .Left = Val(txtLeft): .Top = Val(txtTop)
    .Title = txtTitle
    
    .EditBox = CBool(chkProp(0).Value)
    .FileSystemOnly = CBool(chkProp(1).Value)
    .StatusText = CBool(chkProp(3).Value)
    .UseNewUI = CBool(chkProp(4).Value)
    If CBool(chkProp(2).Value) Then .hwndOwner = hWnd
    
  txtFolderPath = m_BF.BrowseFolder
  txtDisplayName = m_BF.DisplayName
  End With
  Set m_BF = New cBrowseFolder
End Sub

Private Sub cmdInitDir_Click()
  m_IDir.FileSystemOnly = True
  m_IDir.hwndOwner = hWnd
  m_IDir.Position = bfUseCenterOwner
  txtInitDir = m_IDir.BrowseFolder
End Sub

Private Sub cmdSpecialFolder_Click()
  txtSpecialFolder = m_BF.SpecialFolder(cboSFolder.ItemData(cboSFolder.ListIndex))
End Sub

Private Sub Form_Load()
  Call pvSetSpecialFolder
  Call pvSetRootDir
  Call pvSetPosition
End Sub

Private Sub pvSetPosition()
  cboPosition.AddItem "Default"
  cboPosition.AddItem "Center Owner"
  cboPosition.AddItem "Center Screen"
  cboPosition.AddItem "Custom"
  cboPosition.Text = cboPosition.List(0)
End Sub

Private Sub pvSetRootDir()
  Dim i As Integer
  For i = 0 To 41
    cboRootDir.AddItem cboSFolder.List(i)
    cboRootDir.ItemData(i) = cboSFolder.ItemData(i)
  Next
End Sub

Private Sub pvSetSpecialFolder()
  cboSFolder.AddItem "SF_ADMINTOOLS": cboSFolder.ItemData(0) = &H30
  cboSFolder.AddItem "SF_ALTSTARTUP": cboSFolder.ItemData(1) = &H1D
  cboSFolder.AddItem "SF_APPDATA": cboSFolder.ItemData(2) = &H1A
  cboSFolder.AddItem "SF_BITBUCKET": cboSFolder.ItemData(3) = &HA
  cboSFolder.AddItem "SF_COMMON_ADMINTOOLS": cboSFolder.ItemData(4) = &H2F
  cboSFolder.AddItem "SF_COMMON_ALTSTARTUP": cboSFolder.ItemData(5) = &H1E
  cboSFolder.AddItem "SF_COMMON_APPDATA": cboSFolder.ItemData(6) = &H23
  cboSFolder.AddItem "SF_COMMON_DESKTOPDIRECTORY": cboSFolder.ItemData(7) = &H19
  cboSFolder.AddItem "SF_COMMON_DOCUMENTS": cboSFolder.ItemData(8) = &H2E
  cboSFolder.AddItem "SF_COMMON_FAVORITES": cboSFolder.ItemData(9) = &H1F
  cboSFolder.AddItem "SF_COMMON_PROGRAMS": cboSFolder.ItemData(10) = &H17
  cboSFolder.AddItem "SF_COMMON_STARTMENU": cboSFolder.ItemData(11) = &H16
  cboSFolder.AddItem "SF_COMMON_STARTUP": cboSFolder.ItemData(12) = &H18
  cboSFolder.AddItem "SF_COMMON_TEMPLATES": cboSFolder.ItemData(13) = &H2D
  cboSFolder.AddItem "SF_CONTROLS": cboSFolder.ItemData(14) = &H3
  cboSFolder.AddItem "SF_COOKIES": cboSFolder.ItemData(15) = &H21
  cboSFolder.AddItem "SF_DESKTOP": cboSFolder.ItemData(16) = &H0
  cboSFolder.AddItem "SF_DESKTOPDIRECTORY": cboSFolder.ItemData(17) = &H10
  cboSFolder.AddItem "SF_DRIVES": cboSFolder.ItemData(18) = &H11
  cboSFolder.AddItem "SF_FAVORITES": cboSFolder.ItemData(19) = &H6
  cboSFolder.AddItem "SF_FONTS": cboSFolder.ItemData(20) = &H14
  cboSFolder.AddItem "SF_HISTORY": cboSFolder.ItemData(21) = &H22
  cboSFolder.AddItem "SF_INTERNET": cboSFolder.ItemData(22) = &H1
  cboSFolder.AddItem "SF_INTERNET_CACHE": cboSFolder.ItemData(23) = &H20
  cboSFolder.AddItem "SF_LOCAL_APPDATA": cboSFolder.ItemData(24) = &H1C
  cboSFolder.AddItem "SF_MYPICTURES": cboSFolder.ItemData(25) = &H27
  cboSFolder.AddItem "SF_NETHOOD": cboSFolder.ItemData(26) = &H13
  cboSFolder.AddItem "SF_NETWORK": cboSFolder.ItemData(27) = &H12
  cboSFolder.AddItem "SF_PERSONAL": cboSFolder.ItemData(28) = &H5
  cboSFolder.AddItem "SF_PRINTERS": cboSFolder.ItemData(29) = &H4
  cboSFolder.AddItem "SF_PRINTHOOD": cboSFolder.ItemData(30) = &H1B
  cboSFolder.AddItem "SF_PROFILE": cboSFolder.ItemData(31) = &H28
  cboSFolder.AddItem "SF_PROGRAM_FILES": cboSFolder.ItemData(32) = &H26
  cboSFolder.AddItem "SF_PROGRAM_FILES_COMMON": cboSFolder.ItemData(33) = &H2B
  cboSFolder.AddItem "SF_PROGRAMS": cboSFolder.ItemData(34) = &H2
  cboSFolder.AddItem "SF_RECENT": cboSFolder.ItemData(35) = &H8
  cboSFolder.AddItem "SF_SENDTO": cboSFolder.ItemData(36) = &H9
  cboSFolder.AddItem "SF_STARTMENU": cboSFolder.ItemData(37) = &HB
  cboSFolder.AddItem "SF_STARTUP": cboSFolder.ItemData(38) = &H7
  cboSFolder.AddItem "SF_SYSTEM": cboSFolder.ItemData(39) = &H25
  cboSFolder.AddItem "SF_TEMPLATES": cboSFolder.ItemData(40) = &H15
  cboSFolder.AddItem "SF_WINDOWS": cboSFolder.ItemData(41) = &H24
  cboSFolder.Text = cboSFolder.List(41)
End Sub

