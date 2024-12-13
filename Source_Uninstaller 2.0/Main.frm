VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Legacy Power Prompts Configuration Utility"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10305
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "Main.frx":014A
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
      Picture         =   "Main.frx":1DD3C
      ScaleHeight     =   5895
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   840
      Width           =   2535
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   120
         Top             =   5280
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Settings "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
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
         TabIndex        =   5
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
      Picture         =   "Main.frx":4EB56
      ScaleHeight     =   5895
      ScaleWidth      =   7815
      TabIndex        =   11
      Top             =   840
      Width           =   7815
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2280
         TabIndex        =   15
         Text            =   "Windows 2000"
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Saved!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6240
         TabIndex        =   16
         Top             =   4320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Design used:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
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
         TabIndex        =   12
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
      Picture         =   "Main.frx":E4C88
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
         Picture         =   "Main.frx":F771A
         ScaleHeight     =   390
         ScaleWidth      =   1155
         TabIndex        =   9
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
            TabIndex        =   10
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
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   7800
         Picture         =   "Main.frx":F8EEC
         ScaleHeight     =   390
         ScaleWidth      =   1155
         TabIndex        =   7
         Top             =   120
         Width           =   1155
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save "
            Height          =   255
            Left            =   0
            TabIndex        =   8
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
         Picture         =   "Main.frx":FA6BE
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
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Combo1.AddItem "Windows 3.1"
  Combo1.AddItem "Windows 95"
  Combo1.AddItem "Windows 98"
  Combo1.AddItem "Windows Me"
  Combo1.AddItem "Windows NT 3.1"
  Combo1.AddItem "Windows NT 3.5x"
  Combo1.AddItem "Windows NT 4.0"
  Combo1.AddItem "Windows 2000"
  Combo1.AddItem "Incomplete"
  vars.TType = GetSetting("LPowerPrompts", "Type", "LogoffWindow")
  If (vars.TType = 1) Then
    Combo1.ListIndex = 0
  ElseIf (vars.TType = 2) Then
    Combo1.ListIndex = 1
  ElseIf (vars.TType = 3) Then
    Combo1.ListIndex = 2
  ElseIf (vars.TType = 4) Then
    Combo1.ListIndex = 3
  ElseIf (vars.TType = 5) Then
    Combo1.ListIndex = 4
  ElseIf (vars.TType = 6) Then
    Combo1.ListIndex = 5
  ElseIf (vars.TType = 7) Then
    Combo1.ListIndex = 6
  ElseIf (vars.TType = 8) Then
    Combo1.ListIndex = 7
  ElseIf (vars.TType = 9) Then
    Combo1.ListIndex = 8
  ElseIf (vars.TType = 16) Then
    Combo1.Text = "Lorem Ipsum"
  Else
    Combo1.ListIndex = 7
  End If
End Sub

Private Sub Label1_Click()
  End
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Screen.MousePointer = vbCrosshair
End Sub

Private Sub Label25_Click()
  If (Combo1.ListIndex = 0) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "1"
  If (Combo1.ListIndex = 1) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "2"
  If (Combo1.ListIndex = 2) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "3"
  If (Combo1.ListIndex = 3) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "4"
  If (Combo1.ListIndex = 4) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "5"
  If (Combo1.ListIndex = 5) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "6"
  If (Combo1.ListIndex = 6) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "7"
  If (Combo1.ListIndex = 7) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "8"
  If (Combo1.ListIndex = 8) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "9"
  If (Combo1.Text = "Lorem Ipsum") Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "16"
  Label4.Visible = True
  Timer1.Enabled = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not (Screen.MousePointer = vbNormal) Then Screen.MousePointer = vbNormal
End Sub

Private Sub Picture5_Click()
  End
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'If (X > 1100) Then End
End Sub

Private Sub Picture6_Click()
  If (Combo1.ListIndex = 0) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "1"
  If (Combo1.ListIndex = 1) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "2"
  If (Combo1.ListIndex = 2) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "3"
  If (Combo1.ListIndex = 3) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "4"
  If (Combo1.ListIndex = 4) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "5"
  If (Combo1.ListIndex = 5) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "6"
  If (Combo1.ListIndex = 6) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "7"
  If (Combo1.ListIndex = 7) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "8"
  If (Combo1.ListIndex = 8) Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "9"
  If (Combo1.Text = "Lorem Ipsum") Then SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "16"
  Label4.Visible = True
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Label4.Visible = False
End Sub
