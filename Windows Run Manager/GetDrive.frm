VERSION 5.00
Object = "{AA7302EC-EFCC-475A-B6B4-FF7979954FD3}#14.0#0"; "LOGICFADE.OCX"
Object = "{0F5B07D7-006B-11D3-A40D-A54B783FE719}#23.0#0"; "LINE3D.OCX"
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#5.0#0"; "TOOLBAR2.OCX"
Object = "{8467CD72-6937-11D4-9C7F-00E02917505E}#1.0#0"; "FLATBUTTON.OCX"
Begin VB.Form frmFileInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Task Info ..."
   ClientHeight    =   3345
   ClientLeft      =   630
   ClientTop       =   915
   ClientWidth     =   8070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "GetDrive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LogicSoftLine3D.Line3D Line3D1 
      Height          =   45
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   79
      Enabled         =   -1  'True
   End
   Begin FlatButtonControl.FlatButton FlatButton1 
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Ok"
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
   Begin AIFCmp1.asxToolbar asxToolbar1 
      Height          =   2535
      Left            =   120
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4471
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
      Begin FadeText.LogicFade LogicFade4 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2130
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "Status :"
      End
      Begin FadeText.LogicFade LogicFade1 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Task Name :"
      End
      Begin FadeText.LogicFade LogicFade7 
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Info on Selected Task"
      End
      Begin FadeText.LogicFade LogicFade6 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "Command Line :"
      End
      Begin FadeText.LogicFade LogicFade3 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1470
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "Type of Task :"
      End
      Begin FadeText.LogicFade LogicFade2 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1155
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Execute on :"
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   2190
         TabIndex        =   12
         Top             =   2130
         Width           =   5505
      End
      Begin VB.Label lblTaskName 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   2190
         TabIndex        =   10
         Top             =   855
         Width           =   5505
      End
      Begin VB.Label lbPath 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   2190
         TabIndex        =   2
         Top             =   1800
         Width           =   5505
      End
      Begin VB.Label lbBit 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   2190
         TabIndex        =   1
         Top             =   1485
         Width           =   5505
      End
      Begin VB.Label lbwindows 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   2190
         TabIndex        =   0
         Top             =   1170
         Width           =   5505
      End
   End
End
Attribute VB_Name = "frmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FlatButton1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

DrawGradient Me, 255, 255, 255, 151, 50, 0, 2760, 1, True, 1, 1, 1 'Argent

'--------- check file -------------
exe$ = frmMain.lstCmdLine
Path$ = frmMain.lstCmdLine


  ComPath$ = Path$
  fichier$ = Path$
  lbPath = Path$
  lblTaskName = frmMain.lstName
   
  If WinHeader(fichier$) Then
  
    If IsWinProgram(fichier$) Then
        lbwindows.Caption = "Windows 95/98/NT 4.5+"
        
      If Is32bits(fichier$) Then
          lbBit.Caption = "32 Bits"
      Else
          lbBit.Caption = "16 Bits"
      End If
    
    Else
      lbwindows.Caption = "Windows 3.x +"
    End If
Else
  lbwindows.Caption = "DOS"
  lbBit.Caption = "-"
End If
'-------------------------------------------

  If Mid(frmMain.txtCmdLine, 1, 1) = "*" Then
    lblStatus = "Not Load on Windows Star-Up (Desactivate)"
   Else
    lblStatus = "Load on Windows Star-Up (Activate)"
  End If
End Sub


Function WinHeader(fichier$) As Integer
  
  numF = FreeFile
  
    Open fichier$ For Binary As #numF
    Seek #numF, &H19
    Get #numF, , Signature&
  
    If Signature& >= &H40 Then
        WinHeader = True
      
      Else
        WinHeader = False
    End If
      Close #numF
End Function

Function IsWinProgram(fichier$) As String

  numF = FreeFile
   
   Open fichier$ For Binary As #numF
   Seek #numF, &H3C + 1
   Get #numF, , Offset&
   Seek #numF, Offset& + 1
   
   Signature$ = Space$(2)
   Get #numF, , Signature$
   
   If Signature$ = "PE" Or Signature$ = "NE" Then
      IsWinProgram = True
     Else
      IsWinProgram = False
   End If
   Close #numF
   
End Function

Function Is32bits(fichier$) As Integer

 numF = FreeFile
   Open fichier$ For Binary As #numF
   Seek #numF, &H3C + 1
   Get #numF, , Offset&
   Seek #numF, Offset& + 1
   Signature$ = Space(2)
   Get #numF, , Signature$
   
   If Signature$ = "PE" Then
        Is32bits = True
    ElseIf Signature$ = "NE" Then
        Is32bits = False
   End If
  Close #numF
  
End Function



