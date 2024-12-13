VERSION 5.00
Begin VB.MDIForm BkgdWindow 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00808000&
   Caption         =   "Legacy Power Prompts 2.0 Installer"
   ClientHeight    =   5790
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10935
   Enabled         =   0   'False
   Icon            =   "BkgdWindow.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "BkgdWindow.frx":030A
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10080
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10440
      Top             =   120
   End
End
Attribute VB_Name = "BkgdWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim result As VbMsgBoxResult
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As rect) As Long
Private Const WM_NCLBUTTONDBLCLK As Long = &HA3

Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Function TaskbarHeight() As Integer
    Dim rect As rect
    Dim hTaskbar As Long
    hTaskbar = FindWindow("Shell_TrayWnd", vbNullString)
    GetWindowRect hTaskbar, rect
    TaskbarHeight = rect.Bottom - rect.Top
End Function

Private Sub MDIForm_Load()
  vars.Estat = False
  vars.QVal = True
  vars.Finish = False
  Me.Top = -20
  Me.Left = -100
  Me.Width = Screen.Width / 25
  'Me.Height = Screen.Height - TaskbarHeight()
  Me.Height = Screen.Height / 25
  Me.BackColor = RGB(123, 19, 19)
  Timer1.Enabled = True
End Sub

Private Sub MDIForm_Resize()
  If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized And vars.Estat = True Then
    Me.WindowState = vbMaximized
  End If
  If Me.WindowState = vbMinimized And vars.Estat = False Then
    Me.WindowState = vbMaximized
  End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = 1
  If (vars.Finish = True) Then End
  If (vars.QVal = True) Then
    result = MsgBox("If you exit now, LegacyPowerPrompts will not get installed. Are you sure that you want to exit LegacyPowerPrompts Installer?", vbQuestion + vbYesNo, "Exit LegacyPowerPrompts?")
    If result = vbYes Then
      End
    Else
    End If
  End If
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  If Not (Me.Width >= Screen.Width) Then
    Me.Width = Me.Width + (Screen.Width / 10)
  End If
  If Not (Me.Height >= Screen.Height) Then
    Me.Height = Me.Height + (Screen.Height / 10)
  End If
  If Not (Me.Width >= Screen.Width) Or (Me.Height >= Screen.Height) Then
    Timer1.Enabled = True
  Else
    Timer2.Enabled = True
  End If
End Sub

Private Sub Timer2_Timer()
  Timer2.Enabled = False
  Me.WindowState = vbMaximized
  vars.Estat = True
  Transition.Enabled = True
  Transition.Show
End Sub
