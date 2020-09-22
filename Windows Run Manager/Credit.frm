VERSION 5.00
Object = "{AC75A81E-CB5D-460E-8DCC-0B58DF856B4F}#35.0#0"; "OpenDefaultbrowser.ocx"
Object = "{0F5B07D7-006B-11D3-A40D-A54B783FE719}#23.0#0"; "Line3D.ocx"
Object = "{8467CD72-6937-11D4-9C7F-00E02917505E}#1.0#0"; "FLATBUTTON.OCX"
Begin VB.Form frmcredits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Credit..."
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   ControlBox      =   0   'False
   Icon            =   "Credit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer CreditTimer 
      Interval        =   1
      Left            =   2520
      Top             =   2760
   End
   Begin FlatButtonControl.FlatButton FlatButton1 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   4200
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
   Begin LogicSoftLine3D.Line3D Line3D1 
      Height          =   45
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   79
      Enabled         =   -1  'True
   End
   Begin Default_internet_Browser.URLLink URLLink1 
      Height          =   195
      Left            =   540
      Top             =   4290
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   344
      Text            =   "LogicSoft Online"
      URL             =   "http://pages.infinit.net/ramstein/"
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   240
      Picture         =   "Credit.frx":000C
      Top             =   4267
      Width           =   240
   End
   Begin VB.Label lblCredit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmcredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FlatButton1_Click()
  Unload Me
End Sub

Private Sub Form_Load()

  DrawGradient Me, 255, 255, 255, 151, 50, 0, 4080, 1, True, 1, 1, 1 'Argent
  ShowString% = 0
  
  vertion$ = App.Major & "." & App.Minor & " Build #" & App.Revision

  lblCredit = "~=(= Logicsoft WinRun Manager =)~" & vbCrLf & vertion$ & " [Beta]" & vbCrLf & vbCrLf & _
            "Originaly 'Windows Run Manager'" & vbCrLf & vbCrLf & _
            "--- Programmer ---" & vbCrLf & "Max Raskin (For the first vertion)" & _
            vbCrLf & "Derek Tremblay (For the second vertion)" & vbCrLf & vbCrLf & _
            "--- Gfx Interface ---" & vbCrLf & "Derek tremblay" & vbCrLf & vbCrLf & vbCrLf & _
            "-----------------------------------------------------------------------" & _
            vbCrLf & "Purpose: Modify Window's StartUp programs in the Registery, Win.INI and the StartUp folder. You have the posibility to Make a Registery Backup." & _
            vbCrLf & vbCrLf & "Note : this software are originaly posted by Max Raskin and Modified by Derek Tremblay"
            
End Sub

