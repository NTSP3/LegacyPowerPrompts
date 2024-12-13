VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Legacy Power Prompts Configuration Utility"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   Enabled         =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10305
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      Picture         =   "Main.frx":014A
      ScaleHeight     =   735
      ScaleWidth      =   10335
      TabIndex        =   2
      Top             =   6720
      Width           =   10335
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   6600
         Picture         =   "Main.frx":12BDC
         ScaleHeight     =   390
         ScaleWidth      =   1155
         TabIndex        =   27
         Top             =   120
         Width           =   1155
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Run"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   4440
         Picture         =   "Main.frx":143AE
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
         Picture         =   "Main.frx":15B80
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
         Picture         =   "Main.frx":17352
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
         Caption         =   "Config Utility 2.0 of LPP Based on NTPToolkit 3.1"
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "Main.frx":18B24
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
      Picture         =   "Main.frx":36716
      ScaleHeight     =   5895
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   840
      Width           =   2535
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   480
         Top             =   5280
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   120
         Top             =   5280
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Config Utility"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Shutdown Prompt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1575
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
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   2520
      Picture         =   "Main.frx":67530
      ScaleHeight     =   5895
      ScaleWidth      =   7815
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   7815
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "User Interface:"
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
         Left            =   480
         TabIndex        =   30
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Configuration Utility Settings "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   360
         Width           =   7815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Enable Animations"
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
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
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
         Height          =   615
         Left            =   6360
         TabIndex        =   22
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   2520
      Picture         =   "Main.frx":FD662
      ScaleHeight     =   5895
      ScaleWidth      =   7815
      TabIndex        =   11
      Top             =   840
      Width           =   7815
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   3000
         Width           =   255
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         Text            =   "Windows 2000"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Shutdown Prompt Settings "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   23
         Top             =   360
         Width           =   7815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Enable Movement"
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
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Shadow Style:"
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
         Left            =   480
         TabIndex        =   17
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
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
         Height          =   615
         Left            =   6360
         TabIndex        =   15
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Shutdown Design used:"
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
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject
Dim QuitBox As VbMsgBoxResult
Dim E_EnableAnimation As String
Dim E_ChadShadow As String
Dim E_WinDesign As String
Dim E_MoveWin As String
Dim E_TType As String

Private Function GoWith(InputType As Integer)
  If (InputType = 1) Then
    If (Combo1.Text = "Lorem Ipsum") Then vars.TType = 16
    E_TType = ""
    E_ChadShadow = ""
    E_MoveWin = ""
    E_EnableAnimation = ""
    E_WinDesign = ""
    E_TType = GetSetting("LPowerPrompts", "Type", "LogoffWindow")
    E_ChadShadow = GetSetting("LPowerPrompts", "Type", "ShadowType")
    E_MoveWin = GetSetting("LPowerPrompts", "Type", "IsMovable")
    E_EnableAnimation = GetSetting("LPowerPrompts", "Configs", "EnableAnimation")
    E_WinDesign = GetSetting("LPowerPrompts", "Configs", "QSetStyle")
    If (E_TType = vars.TType And E_ChadShadow = vars.ChadShadow And E_MoveWin = vars.MoveWin And E_EnableAnimation = vars.RegAnimation And E_WinDesign = vars.WinDesign) Then
      If (vars.EnableAnimations = True) Then
        Timer2.Enabled = True
      ElseIf (vars.EnableAnimations = False) Then
        End
      End If
    Else
      QuitBox = MsgBox("There are some changes that you have made since the last update. Exiting without saving them will result in the loss of them. Do you still want to exit?", vbQuestion & vbOKCancel, "Exit LPP Config Utility")
      If (QuitBox = vbOK) Then
        If (vars.EnableAnimations = True) Then
          Timer2.Enabled = True
        ElseIf (vars.EnableAnimations = False) Then
          End
        End If
      End If
    End If
  ElseIf (InputType = 2) Then
    If (Combo1.Text = "Lorem Ipsum") Then vars.TType = 16
    E_TType = ""
    E_ChadShadow = ""
    E_MoveWin = ""
    E_EnableAnimation = ""
    E_WinDesign = ""
    E_TType = GetSetting("LPowerPrompts", "Type", "LogoffWindow")
    E_ChadShadow = GetSetting("LPowerPrompts", "Type", "ShadowType")
    E_MoveWin = GetSetting("LPowerPrompts", "Type", "IsMovable")
    E_EnableAnimation = GetSetting("LPowerPrompts", "Configs", "EnableAnimation")
    E_WinDesign = GetSetting("LPowerPrompts", "Configs", "QSetStyle")
    If (E_TType = vars.TType And E_ChadShadow = vars.ChadShadow And E_MoveWin = vars.MoveWin And E_EnableAnimation = vars.RegAnimation And E_WinDesign = vars.WinDesign) Then
      Label4.Caption = "No changes needed."
      Label11.Caption = Label4.Caption
      Label4.Visible = True
      Label11.Visible = True
      E_TType = ""
      E_ChadShadow = ""
      E_MoveWin = ""
      E_EnableAnimation = ""
      E_WinDesign = ""
      Timer1.Enabled = True
    Else
      E_TType = ""
      E_ChadShadow = ""
      E_MoveWin = ""
      E_EnableAnimation = ""
      E_WinDesign = ""
      SaveSetting "LPowerPrompts", "Type", "LogoffWindow", vars.TType
      SaveSetting "LPowerPrompts", "Type", "ShadowType", vars.ChadShadow
      SaveSetting "LPowerPrompts", "Type", "IsMovable", vars.MoveWin
      SaveSetting "LPowerPrompts", "Configs", "EnableAnimation", vars.RegAnimation
      SaveSetting "LPowerPrompts", "Configs", "QSetStyle", vars.WinDesign
      E_TType = GetSetting("LPowerPrompts", "Type", "LogoffWindow")
      E_ChadShadow = GetSetting("LPowerPrompts", "Type", "ShadowType")
      E_MoveWin = GetSetting("LPowerPrompts", "Type", "IsMovable")
      E_EnableAnimation = GetSetting("LPowerPrompts", "Configs", "EnableAnimation")
      E_WinDesign = GetSetting("LPowerPrompts", "Configs", "QSetStyle")
      If (E_TType = vars.TType And E_ChadShadow = vars.ChadShadow And E_MoveWin = vars.MoveWin And E_EnableAnimation = vars.RegAnimation And E_WinDesign = vars.WinDesign) Then
        Label4.Caption = "Saved successfully!"
      Else
        Label4.Caption = "Errors occured when saving."
      End If
      Label11.Caption = Label4.Caption
      Label4.Visible = True
      Label11.Visible = True
      E_TType = ""
      E_ChadShadow = ""
      E_MoveWin = ""
      E_EnableAnimation = ""
      E_WinDesign = ""
      Timer1.Enabled = True
    End If
  ElseIf (InputType = 3) Then
    If fso.FileExists(vars.InsPath & "\qscrdlg.exe") Then
      Shell vars.InsPath & "\qscrdlg.exe", vbHide
    Else
      MsgBox "Legacy Power Prompts Shutdown dialog was not found. Reinstalling may fix this problem.", vbExclamation, "Error"
    End If
  Else
    Me.Hide
    MsgBox "Function not defined!", vbCritical, "Error"
    End
  End If
End Function

Private Sub Check1_Click()
  If (Check1.Value = 0) Then
    vars.MoveWin = "False"
  ElseIf (Check1.Value = 1) Then
    vars.MoveWin = "True"
  End If
End Sub

Private Sub Check2_Click()
  If (Check2.Value = 0) Then
    vars.EnableAnimations = "False"
  ElseIf (Check2.Value = 1) Then
    vars.EnableAnimations = "True"
  End If
  vars.RegAnimation = vars.EnableAnimations
End Sub

Private Sub Combo1_Click()
  If (Combo1.ListIndex >= 0 And Combo1.ListIndex < 10) Then
    vars.TType = Combo1.ListIndex + 1
  End If
End Sub

Private Sub Combo2_Click()
  If (Combo2.ListIndex = 0) Then
    vars.ChadShadow = 31
  ElseIf (Combo2.ListIndex = 1) Then
    vars.ChadShadow = 41
  ElseIf (Combo2.ListIndex = 2) Then
    vars.ChadShadow = 40
  ElseIf (Combo2.ListIndex = 3) Then
    vars.ChadShadow = 50
  ElseIf (Combo2.ListIndex = 4) Then
    vars.ChadShadow = 51
  End If
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Combo3_Click()
  If (Combo3.ListIndex = 0) Then
    vars.WinDesign = 51
  End If
End Sub

Private Sub Form_Load()
  E_TType = ""
  E_ChadShadow = ""
  E_MoveWin = ""
  E_EnableAnimations = ""
  E_WinDesign = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = 1
  GoWith (1)
End Sub

Private Sub Label1_Click()
  Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With Label11
     SetCursor LoadCursor(0, IDC_HAND)
   End With
End Sub

Private Sub Label13_Click()
  GoWith (3)
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With Label13
     SetCursor LoadCursor(0, IDC_HAND)
   End With
End Sub

Private Sub Label2_Click()
  Picture1.Visible = True
  Picture8.Visible = False
  Label2.Font.Bold = True
  Label2.Font.Size = 9
  Label5.Font.Bold = False
  Label5.Font.Size = 8
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With Label2
     SetCursor LoadCursor(0, IDC_HAND)
     '.Drag vbBeginDrag
   End With
End Sub

Private Sub Label25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With Label25
     SetCursor LoadCursor(0, IDC_HAND)
   End With
End Sub

Private Sub Label5_Click()
  Picture8.Visible = True
  Picture1.Visible = False
  Label2.Font.Bold = False
  Label2.Font.Size = 8
  Label5.Font.Bold = True
  Label5.Font.Size = 9
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With Label5
     SetCursor LoadCursor(0, IDC_HAND)
     '.Drag vbBeginDrag
   End With
End Sub

Private Sub Label25_Click()
  GoWith (2)
End Sub

Private Sub Label7_Click()
  If (Check2.Value = 0) Then
    Check2.Value = 1
  ElseIf (Check2.Value = 1) Then
    Check2.Value = 0
  End If
End Sub

Private Sub Label9_Click()
  If (Check1.Value = 0) Then
    Check1.Value = 1
  ElseIf (Check1.Value = 1) Then
    Check1.Value = 0
  End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not (Screen.MousePointer = vbNormal) Then Screen.MousePointer = vbNormal
End Sub

Private Sub Picture5_Click()
  Unload Me
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'If (X > 1100) Then End
   With Picture5
     SetCursor LoadCursor(0, IDC_HAND)
   End With
End Sub

Private Sub Picture6_Click()
  GoWith (2)
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With Picture6
     SetCursor LoadCursor(0, IDC_HAND)
   End With
End Sub

Private Sub Picture9_Click()
  GoWith (3)
End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With Picture9
     SetCursor LoadCursor(0, IDC_HAND)
   End With
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Label4.Visible = False
  Label11.Visible = False
  Label4.Caption = "Info"
  Label11.Caption = Label4.Caption
End Sub

Private Sub Timer2_Timer()
  Timer2.Enabled = False
  If Not (Me.Height <= 20) Then
    If ((Me.Height - 750) > 1) Then
      Me.Height = Me.Height - 750
      Timer2.Enabled = True
    Else
      End
    End If
  Else
    End
  End If
End Sub
