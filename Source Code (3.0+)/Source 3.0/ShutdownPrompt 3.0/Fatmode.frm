VERSION 5.00
Begin VB.Form Fatmode 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pointer Stars"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   Enabled         =   0   'False
   Icon            =   "Fatmode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   4035
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4095
      ScaleWidth      =   6015
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   5400
         Top             =   3480
      End
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   5040
         Top             =   3480
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   4680
         Top             =   3480
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   4320
         Top             =   3480
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   3960
         Top             =   3480
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   3600
         Top             =   3480
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3240
         Top             =   3480
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   135
         Left            =   2760
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   9
         Top             =   1800
         Width           =   135
      End
      Begin VB.Image Image8 
         Height          =   1275
         Left            =   -1200
         Picture         =   "Fatmode.frx":030A
         Top             =   3960
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Image Image7 
         Height          =   1275
         Left            =   -1080
         Picture         =   "Fatmode.frx":5D9C
         Top             =   3960
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Image Image6 
         Height          =   1275
         Left            =   -1080
         Picture         =   "Fatmode.frx":B82E
         Top             =   3960
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Image Image5 
         Height          =   1275
         Left            =   -1080
         Picture         =   "Fatmode.frx":112C0
         Top             =   3960
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Image Image4 
         Height          =   1275
         Left            =   -1200
         Picture         =   "Fatmode.frx":16D52
         Top             =   4080
         Visible         =   0   'False
         Width           =   1350
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4095
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "WOW"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5520
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   1200
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "dbg: num = * of 2 #0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Find Me"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4920
         TabIndex        =   6
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   1665
         Left            =   0
         Picture         =   "Fatmode.frx":1C7E4
         Stretch         =   -1  'True
         Top             =   720
         Width           =   6000
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Transparent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.Image Image2 
         Height          =   1635
         Left            =   0
         Picture         =   "Fatmode.frx":228E6
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hidden Window application Version 1.0 Demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   3000
         TabIndex        =   4
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2520
         TabIndex        =   3
         Top             =   3240
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   2520
         X2              =   1800
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   2040
         X2              =   1800
         Y1              =   2760
         Y2              =   3000
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   1800
         X2              =   2040
         Y1              =   3000
         Y2              =   3240
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   2
         Top             =   3600
         Width           =   495
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   4440
         X2              =   4440
         Y1              =   3840
         Y2              =   3720
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   5400
         X2              =   4440
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   5400
         X2              =   5280
         Y1              =   3720
         Y2              =   3600
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   5280
         X2              =   5400
         Y1              =   3840
         Y2              =   3720
      End
   End
End
Attribute VB_Name = "Fatmode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TpWindow As Boolean
Dim Ocreg As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL As Long = 1

Private Sub OpenWebsite(ByVal url As String)
    ' Use ShellExecute to open the default web browser
    ShellExecute Me.hwnd, "open", url, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Check1_Click()
  vars.Fatmode = Check1.Value
End Sub

Private Sub Form_Load()
  TpWindow = True
  vars.Creg = 0
  Ocreg = False
End Sub

Private Sub Image2_Click()
  If Not (vars.Creg = 12) Then
    ShadowCode.MainShadow Me.hwnd, 200
  ElseIf (vars.Creg = 12) Then
    ShadowCode.MainShadow Me.hwnd, 255
  End If
  Picture1.Visible = False
  Picture2.Visible = True
  If Not (vars.Fatmode = 1) Then
    Picture3.BorderStyle = 0
    Picture3.Width = 0
    Picture3.Height = 0
  ElseIf (vars.Fatmode = 1) Then
    Picture3.BorderStyle = 1
  End If
End Sub

Private Sub Image3_Click()
  Check1.Visible = True
End Sub

Private Sub Label1_Click()
  If (vars.Creg = 4) Then Ocreg = True
  If (Ocreg = False) Then
    vars.Creg = vars.Creg + 1
  ElseIf (Ocreg = True) Then
    vars.Creg = vars.Creg + 4
  End If
  Label1.Caption = "dbg: num = * of 2 #" & vars.Creg
End Sub

Private Sub Label5_Click()
  If Not ((vars.Creg) = 16 Or (vars.Creg = 8)) Then
    OpenWebsite "https://boulderbugle.com/legacypowerprompts-help-5eiyk5v3"
  ElseIf (vars.Creg = 8) Then
    If (TpWindow = True) Then
    ShadowCode.MainShadow Me.hwnd, 200
    TpWindow = Not TpWindow
  ElseIf (TpWindow = False) Then
    ShadowCode.MainShadow Me.hwnd, 255
    TpWindow = Not TpWindow
  End If
  ElseIf (vars.Creg = 16) Then
    Me.Enabled = False
    FATAbout.Enabled = True
    FATAbout.Show
  End If
End Sub

Private Sub Label6_Click()
  MsgBox "Info: The number shall be a multiple of 2.", vbInformation, "LPP"
End Sub

Private Sub Picture2_Click()
  Timer1.Enabled = True
End Sub

Private Sub Picture3_Click()
  Timer1.Enabled = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Picture3.Top = Y
  Picture3.Left = X
  vars.YMouse = Y / 1.3
  vars.XMouse = X / 1.3
End Sub

'Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  Picture3.Top = Y
'  Picture3.Left = X
'End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Image8.Visible = False
  Image4.Visible = True
  Image4.Top = vars.YMouse
  Image4.Left = vars.XMouse
  Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
  Timer2.Enabled = False
  Image4.Visible = False
  Image5.Visible = True
  Image5.Top = vars.YMouse
  Image5.Left = vars.XMouse
  Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
  Timer3.Enabled = False
  Image5.Visible = False
  Image6.Visible = True
  Image6.Top = vars.YMouse
  Image6.Left = vars.XMouse
  Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
  Timer4.Enabled = False
  Image5.Visible = False
  Image6.Visible = True
  Image6.Top = vars.YMouse
  Image6.Left = vars.XMouse
  Timer5.Enabled = True
End Sub

Private Sub Timer5_Timer()
  Timer5.Enabled = False
  Image6.Visible = False
  Image7.Visible = True
  Image7.Top = vars.YMouse
  Image7.Left = vars.XMouse
  Timer6.Enabled = True
End Sub

Private Sub Timer6_Timer()
  Timer6.Enabled = False
  Image7.Visible = False
  Image8.Visible = True
  Image8.Top = vars.YMouse
  Image8.Left = vars.XMouse
  Timer7.Enabled = True
End Sub

Private Sub Timer7_Timer()
  Timer7.Enabled = False
  Image8.Visible = False
End Sub
