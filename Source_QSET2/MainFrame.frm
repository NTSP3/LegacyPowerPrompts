VERSION 5.00
Begin VB.Form MainFrame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Legacy Power Prompts Configuration Utility"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   Icon            =   "MainFrame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10305
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   840
      Top             =   6120
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   480
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   6120
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "MainFrame.frx":014A
      ScaleHeight     =   855
      ScaleWidth      =   10335
      TabIndex        =   1
      Top             =   0
      Width           =   10335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   0
      Picture         =   "MainFrame.frx":1DD3C
      ScaleHeight     =   5895
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   840
      Width           =   2535
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "   Sections   "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   2520
      Picture         =   "MainFrame.frx":4EB56
      ScaleHeight     =   5895
      ScaleWidth      =   7815
      TabIndex        =   10
      Top             =   840
      Width           =   7815
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5640
         TabIndex        =   11
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      Picture         =   "MainFrame.frx":E4C88
      ScaleHeight     =   735
      ScaleWidth      =   10335
      TabIndex        =   2
      Top             =   6720
      Width           =   10335
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   4440
         Picture         =   "MainFrame.frx":F771A
         ScaleHeight     =   390
         ScaleWidth      =   1155
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   1155
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Discard "
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   7800
         Picture         =   "MainFrame.frx":F8EEC
         ScaleHeight     =   390
         ScaleWidth      =   1155
         TabIndex        =   6
         Top             =   120
         Width           =   1155
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save "
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   9000
         Picture         =   "MainFrame.frx":FA6BE
         ScaleHeight     =   390
         ScaleWidth      =   1155
         TabIndex        =   3
         Top             =   120
         Width           =   1155
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Quit"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Config Utility 1.0 of LPP Based on NTPToolkit 3.1"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "MainFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cpts As Integer
Dim Pos As Integer
Dim T1 As Boolean
Dim t2 As Boolean

Private Sub Form_Load()
  T1 = False
  t2 = False
  Cpts = 0
  Pos = 1
  Me.Left = Pos
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Cpts = Cpts + 15
  If (Cpts < 255) Then
    ShadowCode.MainShadow Me.hwnd, Cpts
    Timer1.Enabled = True
  Else
    ShadowCode.MainShadow Me.hwnd, 255
    T1 = True
  End If
End Sub

Private Sub Timer2_Timer()
  Timer2.Enabled = False
  If (Pos < (Screen.Width / 2)) Then
    Me.Left = Pos - (Me.Width / 2)
    Pos = Pos + 1000
    Timer2.Enabled = True
  Else
    t2 = True
  End If
End Sub

Private Sub Timer3_Timer()
  Timer3.Enabled = False
  If (T1 = True And t2 = True) Then
    Main.Enabled = True
    Main.Left = Me.Left
    Main.Top = Me.Top
    Main.Show
    Me.Enabled = False
    Me.Hide
    Unload Me
    Exit Sub
  Else
    Timer3.Enabled = True
  End If
End Sub
