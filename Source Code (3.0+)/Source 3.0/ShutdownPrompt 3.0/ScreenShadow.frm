VERSION 5.00
Begin VB.Form ScreenShadow 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6720
      Top             =   4680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6360
      Top             =   4680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6000
      Top             =   4680
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   0
      Picture         =   "ScreenShadow.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "ScreenShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  ShadowCode.MainShadow Me.hwnd, 1
  Me.Left = 0
  Me.Top = 0
  Me.Width = 0
  Me.Height = 0
  If (vars.ShadowMode = "1") Then
    Me.Width = Screen.Width
    Me.Height = Screen.Height
  ElseIf (vars.ShadowMode = "2") Then
    ShadowCode.MainShadow Me.hwnd, 128 ' 128 is 50% transparency (0-255 range)
    Me.Width = Screen.Width
    Timer1.Enabled = True
  ElseIf (vars.ShadowMode = "3") Then
    ShadowCode.MainShadow Me.hwnd, 128
    Me.Height = Screen.Height
    Me.Width = 0
    Timer2.Enabled = True
  ElseIf (vars.ShadowMode = "4") Then
    ShadowCode.MainShadow Me.hwnd, 128
    Me.Height = Screen.Height
    Me.Width = Screen.Width
  ElseIf (vars.ShadowMode = "5") Then
    Me.BackColor = &H808080
    ShadowCode.MainShadow Me.hwnd, 128
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    Image1.Height = Screen.Height / 100
    Image1.Width = Screen.Width / 100
    Image1.Visible = True
    Timer3.Enabled = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  If Not (Me.Height = Screen.Height) Then
    Me.Height = Me.Height + (Screen.Height / 25)
    Timer1.Enabled = True
  End If
End Sub

Private Sub Timer2_Timer()
  Timer2.Enabled = False
  If Not (Me.Width = Screen.Width) Then
    Me.Width = Me.Width + (Screen.Width / 8)
    Timer2.Enabled = True
  End If
End Sub

Private Sub Timer3_Timer()
  Timer3.Enabled = False
  vars.TimerReset = "0"
  If Not (Image1.Height >= (Screen.Height * 2)) Then
    Image1.Height = Image1.Height + (Screen.Height / 5)
    vars.TimerReset = "1"
  End If
  If Not (Image1.Width >= (Screen.Width * 2)) Then
    Image1.Width = Image1.Width + (Screen.Width / 5)
    vars.TimerReset = "1"
  End If
  If (vars.TimerReset = "1") Then
    Timer3.Enabled = True
  End If
End Sub
