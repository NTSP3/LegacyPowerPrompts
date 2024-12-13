VERSION 5.00
Begin VB.Form FSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9225
   Icon            =   "FSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   120
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "FSplash.frx":014A
      Top             =   0
      Width           =   9300
   End
End
Attribute VB_Name = "FSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject
Dim Cpts As Integer

Private Sub Form_Load()
  Cpts = 1
  vars.FileNotFound = False
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Me.Enabled = False
  Me.Visible = False
  Me.Hide
  BkgdWindow.Enabled = True
  BkgdWindow.Show
  Unload Me
End Sub

Private Sub Timer2_Timer()
  Timer2.Enabled = False
  Cpts = Cpts + 10
  If (Cpts < 255) Then
    ShadowCode.MainShadow Me.hwnd, Cpts
    Timer2.Enabled = True
  Else
    If fso.FileExists(".\comdlg32.ocx") Then
    Else
      vars.FileNotFound = True
    End If
    If fso.FileExists(".\ltxdel.sup") Then
    Else
      vars.FileNotFound = True
    End If
    If fso.FileExists(".\ltxmain.sup") Then
    Else
      vars.FileNotFound = True
    End If
    If fso.FileExists(".\ltxdom.sup") Then
    Else
      vars.FileNotFound = True
    End If
    If fso.FileExists(".\nac.sup") Then
    Else
      vars.FileNotFound = True
    End If
    If (vars.FileNotFound = True) Then
      Timer1.Enabled = False
      Timer2.Enabled = False
      Timer3.Enabled = False
      Me.Enabled = False
      Me.Hide
      result = MsgBox("One or more files required by LegacyPowerPrompts were not found. LegacyPowerPrompts cannot continue.", vbCritical + vbOKOnly)
      End
    Else
      Timer3.Enabled = True
    End If
  End If
End Sub

Private Sub Timer3_Timer()
  Timer2.Enabled = False
  Cpts = Cpts - 10
  If (Cpts < 255 And Cpts > 0) Then
    ShadowCode.MainShadow Me.hwnd, Cpts
    Timer3.Enabled = True
  Else
    Me.Hide
  End If
End Sub
