VERSION 5.00
Object = "{AA7302EC-EFCC-475A-B6B4-FF7979954FD3}#14.0#0"; "LogicFade.ocx"
Object = "{AC75A81E-CB5D-460E-8DCC-0B58DF856B4F}#35.0#0"; "OpenDefaultbrowser.ocx"
Object = "{0F5B07D7-006B-11D3-A40D-A54B783FE719}#23.0#0"; "Line3D.ocx"
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#5.0#0"; "Toolbar2.ocx"
Object = "{8467CD72-6937-11D4-9C7F-00E02917505E}#1.0#0"; "FlatButton.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LogicSoft : WinRun Manager [Beta 1] for Win98"
   ClientHeight    =   7035
   ClientLeft      =   780
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Config.sys"
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Top             =   480
      Width           =   4515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Autoexec.bat"
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   480
      Width           =   4335
   End
   Begin FlatButtonControl.FlatButton mndExit 
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   6600
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      Caption         =   "Exit"
      HasFocusRect    =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton optStartMenu 
      Caption         =   "StartUp Folder"
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Program Shortcuts that are launched from the 'Start Menu\Programs\Start Up\' Folder"
      Top             =   120
      Width           =   1635
   End
   Begin VB.OptionButton optWinINI 
      Caption         =   "Win.ini"
      Height          =   375
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Program that are launched from the windows initlization file 'Win.ini'"
      Top             =   120
      Width           =   1635
   End
   Begin FadeText.LogicFade lblVertion 
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   6660
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      Color1          =   0
      Color2          =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Logicsoft FadeText"
   End
   Begin AIFCmp1.asxToolbar asxToolbar4 
      Height          =   735
      Left            =   60
      Top             =   5640
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   1296
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonCount     =   0
      Begin FlatButtonControl.FlatButton cmdDel 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Caption         =   "Delete Selection"
         HasFocusRect    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FlatButtonControl.FlatButton FlatButton4 
         Height          =   255
         Left            =   3480
         TabIndex        =   26
         Top             =   380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Caption         =   "Credits ..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FlatButtonControl.FlatButton FlatButton1 
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Caption         =   "Language"
         Enabled         =   0   'False
         HasFocusRect    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FlatButtonControl.FlatButton FlatButton3 
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Caption         =   "Scan this task"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FlatButtonControl.FlatButton FlatButton2 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Caption         =   "Execute this task"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FadeText.LogicFade LogicFade4 
         Height          =   255
         Left            =   6240
         TabIndex        =   21
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Color1          =   0
         Color2          =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Task to Run in this Section :"
      End
      Begin FlatButtonControl.FlatButton cmdBackup 
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Caption         =   "Regedit Back-Up"
         HasFocusRect    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblnumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "###"
         Height          =   195
         Left            =   8760
         TabIndex        =   22
         Top             =   285
         Width           =   315
      End
   End
   Begin AIFCmp1.asxToolbar asxToolbar2 
      Height          =   735
      Left            =   60
      Top             =   4815
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   1296
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonCount     =   0
      Begin FlatButtonControl.FlatButton cmdAdd 
         Height          =   615
         Left            =   7920
         TabIndex        =   20
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         HasBorder       =   -1  'True
         Caption         =   "Add/Set"
         HasFocusRect    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FlatButtonControl.FlatButton cmdBrowse 
         Height          =   285
         Left            =   6720
         TabIndex        =   16
         Top             =   390
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         HasBorder       =   -1  'True
         Caption         =   "Browse..."
         HasFocusRect    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1260
         TabIndex        =   12
         Top             =   60
         Width           =   6645
      End
      Begin VB.TextBox txtCmdLine 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1260
         TabIndex        =   11
         Top             =   390
         Width           =   5445
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Program Name:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   105
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Command Line:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   435
         Width           =   1095
      End
   End
   Begin VB.OptionButton optRun2 
      Caption         =   "HKEY_CURRENT_USER Run"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Programs that are found in the 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunServices' Registry Key"
      Top             =   120
      Width           =   2835
   End
   Begin VB.OptionButton optRunServices 
      Caption         =   "Run Services"
      Height          =   375
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Service Programs that are found in the 'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices' Registry Key"
      Top             =   120
      Width           =   1635
   End
   Begin VB.OptionButton optRun 
      Caption         =   " Run"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Programs that are found in the 'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run' Registry Key"
      Top             =   120
      Value           =   -1  'True
      Width           =   1155
   End
   Begin LogicSoftLine3D.Line3D Line3D1 
      Height          =   45
      Left            =   0
      TabIndex        =   2
      Top             =   6480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   79
      Enabled         =   -1  'True
   End
   Begin VB.ListBox lstCmdLine 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   3180
      Left            =   3240
      TabIndex        =   1
      Top             =   1290
      Width           =   6015
   End
   Begin VB.ListBox lstName 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   3180
      Left            =   60
      TabIndex        =   0
      Top             =   1290
      Width           =   3165
   End
   Begin Default_internet_Browser.URLLink URLLink1 
      Height          =   195
      Left            =   420
      Top             =   6690
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   344
      Text            =   "LogicSoft Online"
      URL             =   "http://pages.infinit.net/ramstein/"
   End
   Begin AIFCmp1.asxToolbar asxToolbar19 
      Height          =   255
      Left            =   60
      Top             =   1035
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   450
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonCount     =   0
      Begin FadeText.LogicFade LogicFade1 
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   15
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   397
         Color1          =   0
         Color2          =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Task Name:"
      End
   End
   Begin AIFCmp1.asxToolbar asxToolbar1 
      Height          =   255
      Left            =   3240
      Top             =   1035
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonCount     =   0
      Begin FadeText.LogicFade LogicFade2 
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   15
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   397
         Color1          =   0
         Color2          =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Command Line:"
      End
   End
   Begin AIFCmp1.asxToolbar asxToolbar3 
      Height          =   255
      Left            =   60
      Top             =   4560
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   450
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonCount     =   0
      Begin FadeText.LogicFade LogicFade3 
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   15
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   397
         Color1          =   0
         Color2          =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Properties :"
      End
   End
   Begin FlatButtonControl.FlatButton FlatButton5 
      Height          =   255
      Left            =   6015
      TabIndex        =   27
      Top             =   1035
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   450
      HasBorder       =   -1  'True
      Caption         =   "Active/Desactive this Task (Not Delete)"
      HasFocusRect    =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frmMain.frx":030A
      Top             =   6660
      Width           =   240
   End
   Begin VB.Menu mnuBk 
      Caption         =   "Backup"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "Make HKEY_LOCAL_MACHINE Run Reg File"
      End
      Begin VB.Menu mnu2 
         Caption         =   "Make HKEY_LOCAL_MACHINE RunServices Reg File"
      End
      Begin VB.Menu mnu3 
         Caption         =   "Make HKEY_CURRENT_USER Run Reg File"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim hKey As Long, hKey2 As Long, lCount As Long, i As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const LB_SETHORIZONTALEXTENT = (1045)
Dim CurKey As String
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


'Enumerate from HKEY_LOCAL_MACHINE , Run
Private Sub RMEnumRegRun()
    On Error Resume Next
    ClrLists
    hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        lstName.AddItem EnumValue(hKey, i)
        lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
    Next i
    lstName.ListIndex = 0
    lstCmdLine.ListIndex = 0
End Sub


'Enumerate from HKEY_LOCAL_MACHINE , RunServices
Private Sub RMEnumRegRunServices()
    On Error Resume Next
    ClrLists
    hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        lstName.AddItem EnumValue(hKey, i)
        lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
    Next i
    lstName.ListIndex = 0
    lstCmdLine.ListIndex = 0
End Sub

'Enumerate from HKEY_CURRENT_USER , Run
Private Sub RMEnumRegRun2()
    On Error Resume Next
    ClrLists
    lstName.ListIndex = lstCmdLine.ListIndex = 1
    hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        lstName.AddItem EnumValue(hKey, i)
        lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
    Next i
    lstName.ListIndex = 0
    lstCmdLine.ListIndex = 0
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    Dim prvidx As Integer
    prvidx = lstName.ListIndex
    If Trim(txtName.Text) = "" Then
        MsgBox "Enter Name for this RunTask", vbInformation, "Info: No Name"
        txtName.SetFocus
        Exit Sub
    End If
    
    If optWinINI.Value = False Then
      If Trim(txtCmdLine.Text) = "" Then
          MsgBox "Enter Command Line for this RunTask", vbInformation, "Info: No Command Line"
          txtCmdLine.SetFocus
          Exit Sub
      End If
    End If
    If optRun.Value = True Then
        If CurKey <> txtName.Text Then
            SetValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", txtName.Text, txtCmdLine.Text
        Else
            SetValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", CurKey, txtCmdLine.Text
        End If
        RMEnumRegRun
    End If
    If optRunServices.Value = True Then
        If CurKey <> txtName.Text Then
            SetValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", txtName.Text, txtCmdLine.Text
        Else
            SetValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", CurKey, txtCmdLine.Text
        End If
        RMEnumRegRunServices
    End If
    If optRun2.Value = True Then
        If CurKey <> txtName.Text Then
            SetValue RegistryKeys.HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", txtName.Text, txtCmdLine.Text
        Else
            SetValue RegistryKeys.HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", CurKey, txtCmdLine.Text
        End If
        RMEnumRegRun2
    End If
    
    If optWinINI.Value = True Then
     If UCase$(txtName) <> "LOAD" Or UCase$(txtName) <> "RUN" Then
'        WritePrivateProfileByKeyName "windows", ".", ".", TxtappIniFilename.Text
'        WritePrivateProfileToDeleteKey EncryptText$("PersoNames" & LstClasse.Text, "LogicSoftQC99"), ".", 0&, TxtappIniFilename.Text
        WritePrivateProfileByKeyName& "windows", txtName, txtCmdLine, WinDir & "\win.ini"
     
        optWinINI_Click
     End If
    End If
    
    lstCmdLine.ListIndex = prvidx
    lstName.ListIndex = prvidx
    
    lblnumber.Caption = lstName.ListCount
    
End Sub

Private Sub cmdBackup_Click()
    Me.PopupMenu mnuBk, 1, 1850, 6020, mnu1
End Sub

Private Sub cmdBrowse_Click()
    Dim r As String
    r = OpenFile
    If r <> "" Then txtCmdLine.Text = r
End Sub

Private Sub cmdDel_Click()
    On Error Resume Next
    Dim prvidx As Integer, msgResult As VbMsgBoxResult
    prvidx = lstName.ListIndex
    If optRun.Value = False Then
        If optRunServices.Value = False Then
            If optRun2.Value = False Then
                Exit Sub
            End If
        End If
    End If
    msgResult = MsgBox("Are you sure you want to delete the program/service '" & lstName.List(prvidx) & "' from the run sequence ?", vbQuestion Or vbYesNo, "Confirm Delete")
    If msgResult = vbNo Then
        Exit Sub
    Else
        'Do Nothing and continue
    End If
    If optRun.Value = True Then
        hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
        DeleteValue hKey, CurKey
        RMEnumRegRun
    End If
    If optRunServices.Value = True Then
        hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices")
        DeleteValue hKey, CurKey
        RMEnumRegRunServices
    End If
    If optRun2.Value = True Then
        hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
        DeleteValue hKey, CurKey
        RMEnumRegRun2
    End If
    lstCmdLine.ListIndex = prvidx - 1
    lstName.ListIndex = prvidx - 1
    
    lblnumber.Caption = lstName.ListCount
End Sub


Private Sub Command1_Click()
  Shell "notepad.exe c:\autoexec.bat", vbNormalFocus
End Sub

Private Sub Command2_Click()
  Shell "notepad.exe c:\config.sys", vbNormalFocus
End Sub

Private Sub FlatButton2_Click()
 On Error Resume Next
  Shell lstCmdLine
End Sub

Private Sub FlatButton3_Click()
  frmFileInfo.Show 1
End Sub

Private Sub FlatButton4_Click()

  frmcredits.Show 1
End Sub

Private Sub FlatButton5_Click()
 If UCase$(Mid(txtCmdLine, 1, 1)) = "*" Then
    txtCmdLine.Text = Mid(txtCmdLine, 2, Len(txtCmdLine) - 1)
    cmdAdd_Click
  Else
    txtCmdLine.Text = "*" & txtCmdLine
    cmdAdd_Click
 End If
End Sub

Private Sub Form_Load()

DrawGradient Me, 255, 255, 255, 151, 50, 0, 6480, 1, True, 1, 1, 1 'Argent

AppVertion lblVertion
    LstAddScroll lstName
    LstAddScroll lstCmdLine
    RMEnumRegRun
    lblnumber.Caption = lstName.ListCount
    txtCmdLine_Change
End Sub


Private Sub LstAddScroll(Listbox As Listbox)
    SendMessage Listbox.hwnd, LB_SETHORIZONTALEXTENT, 600, 0
End Sub

Private Sub ClrLists()
    lstName.Clear
    lstCmdLine.Clear
End Sub

Private Sub lstCmdLine_Click()
    On Error Resume Next
    lstName.ListIndex = lstCmdLine.ListIndex
    txtCmd.Text = lstCmdLine.List(lstCmdLine.ListIndex)
    CurKey = lstName.List(lstName.ListIndex)
End Sub

Private Sub lstCmdLine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lstCmdLine_Click
End Sub

Private Sub lstCmdLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lstCmdLine_Click
End Sub

Private Sub lstName_Click()
    On Error Resume Next
    lstCmdLine.ListIndex = lstName.ListIndex
    txtName.Text = lstName.List(lstName.ListIndex)
    txtCmdLine.Text = lstCmdLine.List(lstCmdLine.ListIndex)
    CurKey = lstName.List(lstName.ListIndex)
End Sub

Private Sub lstName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lstName_Click
End Sub

Private Sub lstName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lstName_Click
End Sub

Private Sub mndExit_Click()
 A% = MsgBox("Do you want really quit WinRun Manager ?", vbQuestion + vbYesNo, "Question")
  
 If A% = vbYes Then
  End
  Else
 End If
End Sub

Private Sub mnu1_Click()
    optRun_Click
    optRun.SetFocus
    MakeRegFile
End Sub

Private Sub mnu2_Click()
    optRunServices_Click
    optRunServices.SetFocus
    MakeRegFile , 1
End Sub

Private Sub mnu3_Click()
    optRun2_Click
    optRun2.SetFocus
    MakeRegFile 1, 1
End Sub

Private Sub optRun_Click()
    RMEnumRegRun
    cmdDel.Enabled = True
    lblnumber.Caption = lstName.ListCount
    txtName.Enabled = True
End Sub

Private Sub optRunServices_Click()
    RMEnumRegRunServices
    cmdDel.Enabled = True
    lblnumber.Caption = lstName.ListCount
    txtName.Enabled = True
    FlatButton3.Enabled = True
End Sub


Private Sub optRun2_Click()
    RMEnumRegRun2
    cmdDel.Enabled = True
    lblnumber.Caption = lstName.ListCount
    txtName.Enabled = True
    FlatButton3.Enabled = True
End Sub

Private Sub optStartMenu_Click()
    ShellExecute 0, "open", CheckFolderID(StartUp), "", CheckFolderID(StartUp), 1
    
    lblnumber.Caption = "?"
    ClrLists
    
    FlatButton3.Enabled = False
    cmdDel.Enabled = False
End Sub

Private Sub optWinINI_Click()
'    ShellExecute 0, "open", "notepad.exe", WinDir & "\win.ini", "", 1
ClrLists


lstName.AddItem "Load"
lstName.AddItem "Run"

lstCmdLine.AddItem GetPrivateStringValue("Windows", "Load", WinDir & "\win.ini")
lstCmdLine.AddItem GetPrivateStringValue("Windows", "Run", WinDir & "\win.ini")

cmdDel.Enabled = False
txtName.Enabled = False
FlatButton3.Enabled = False
lblnumber.Caption = lstName.ListCount
    
    lstName.ListIndex = 0
    lstCmdLine.ListIndex = 0
End Sub

'Get Windows's Directory
Public Function WinDir() As String
    Dim RetVal As String
    Dim Tmp As String
    Tmp = Space$(255)
    RetVal = GetWindowsDirectory(Tmp, 255)
    WinDir = Trim$(Left$(Tmp, RetVal))
End Function

Private Function MakeRegFile(Optional hKey As Integer, Optional nType As Integer) As String
    On Error Resume Next
    Dim sKey1 As String, sKey2 As String, r As String
    sKey1 = "HKEY_LOCAL_MACHINE"
    sKey2 = "Run"
    If hKey >= 1 Then sKey1 = "HKEY_CURRENT_USER"
    If nType >= 1 Then sKey2 = "RunServices"
    MakeRegFile = "REGEDIT4" & vbCrLf & vbCrLf & "[" & sKey1 & "\Software\Microsoft\Windows\CurrentVersion\" & sKey2 & "]"
    For i = 0 To lstName.ListCount - 1
        MakeRegFile = MakeRegFile & vbCrLf & Chr(34) & lstName.List(i) & Chr(34) & "=" & Chr(34) & lstCmdLine.List(i) & Chr(34)
    Next i
    r = SaveFile
    If r <> "" Then
        Open r For Binary As #1
            Put #1, , MakeRegFile
        Close #1
    End If
End Function

Private Sub txtCmdLine_Change()
  If txtCmdLine.Text = "" Then
    FlatButton5.Enabled = False
   Else
    FlatButton5.Enabled = True
  End If
End Sub
