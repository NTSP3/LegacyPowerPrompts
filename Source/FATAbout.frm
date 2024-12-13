VERSION 5.00
Begin VB.Form FATAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About LPP Pointer Stars"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "FATAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   1200
      Width           =   495
      Begin VB.Line Line7 
         X1              =   240
         X2              =   0
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line6 
         X1              =   480
         X2              =   240
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   240
         Y1              =   0
         Y2              =   480
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5280
      Top             =   1920
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   6
      Top             =   3075
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   120
      Picture         =   "FATAbout.frx":014A
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   1
      Top             =   120
      Width           =   750
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   """16"" was a reference to ---- originally"
      Height          =   1095
      Left            =   5040
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   """16"" is a reference to 16-bit."
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   1656.523
      Y2              =   1656.523
   End
   Begin VB.Label lblDescription 
      Caption         =   $"FATAbout.frx":064C
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Legacy Power Prompts - Pointer Stars"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5323.484
      Y1              =   1656.523
      Y2              =   1656.523
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0, 1.0"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Application made by @Novabits on YouTube."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   3
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "FATAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrow As Integer

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    arrow = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Enabled = False
  Me.Hide
  Fatmode.Enabled = True
  Fatmode.Show
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  If (arrow = 0) Then
    Picture1.Top = Picture1.Top - 83
    arrow = 1
  ElseIf (arrow = 1) Then
    Picture1.Top = Picture1.Top - 83
    arrow = 2
  ElseIf (arrow = 2) Then
    Picture1.Top = Picture1.Top + 83
    arrow = 3
  ElseIf (arrow = 3) Then
    Picture1.Top = Picture1.Top + 83
    arrow = 0
  End If
  Timer1.Enabled = True
End Sub
