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
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
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
Dim Cpts As Integer
Dim LblPtr As Integer
Dim Animation As String
Dim Design As String

Function IsStringInteger(inputString As String) As Boolean
    ' Check if the string is numeric
    If IsNumeric(inputString) Then
        ' Check if the integer value is the same as the original value
        If CStr(Int(inputString)) = inputString Then
            IsStringInteger = True
        Else
            IsStringInteger = False
        End If
    Else
        IsStringInteger = False
    End If
End Function

Private Sub Form_Load()
  'Check InsPath
  vars.InsPath = GetSetting("LPowerPrompts", "Info", "InstalledPath")
  If (vars.InsPath = "" Or (Len(vars.InsPath) < 2)) Then
    Me.Enabled = False
    Me.Hide
    MsgBox "LPP was not able to determine the app's installed location because the path is too short. Reinstalling may fix this problem.", vbCritical, "Error"
    End
  ElseIf Not (Mid(vars.InsPath, 2, 1) = ":") Then
    Me.Enabled = False
    Me.Hide
    MsgBox "LPP was not able to determine the app's installed location because the path is invalid. Reinstalling may fix this problem.", vbCritical, "Error"
    End
  End If
  'Initialization begins
  Cpts = 1
  LblPtr = 0
  vars.TrueOrFalse = True
  vars.EnableAnimations = True
  Design = GetSetting("LPowerPrompts", "Configs", "QSetStyle")
  If IsStringInteger(Design) Then
    vars.WinDesign = Design
  Else
    Me.Enabled = False
    Me.Hide
    MsgBox "A setting required for LegacyPowerPrompts is not numeric. Reinstalling may fix this problem.", vbCritical, "Error"
    End
  End If
  If Not (vars.WinDesign = 50 Or vars.WinDesign = 51 Or vars.WinDesign = 60) Then '5.0 is Tabbed Window (like CTRLPanel Applet), 5.1 is Toolkit interface, and 6.0 is vertical tabs (has modern tabs located in the sidebar)
    Me.Enabled = False
    Me.Hide
    MsgBox "A setting required for LegacyPowerPrompts contains an invalid value. Reinstalling may fix this problem.", vbCritical, "Error"
    End
  End If
  Animation = GetSetting("LPowerPrompts", "Configs", "EnableAnimation")
  If (Animation = "False") Then
    vars.EnableAnimations = False
    Cpts = 256
    Label1.Caption = "Getting Registry Values..."
    Timer1.Interval = 400
    Timer2.Interval = 200
    Timer1.Enabled = True
    Timer2.Enabled = True
  Else
    vars.EnableAnimations = True
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer4.Enabled = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'End
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Me.Enabled = False
  Me.Visible = False
  Me.Hide
  If (vars.EnableAnimations = True) Then
    Main.Check2.Value = 1
    MainFrame.Enabled = True
    MainFrame.Left = (Screen.Width - MainFrame.Width) / 2
    MainFrame.Top = (Screen.Height - MainFrame.Height) / 2
    MainFrame.Show
    Unload Me
  ElseIf (vars.EnableAnimations = False) Then
    Main.Check2.Value = 0
    Main.Enabled = True
    Main.Left = (Screen.Width - Main.Width) / 2
    Main.Top = (Screen.Height - Main.Height) / 2
    Main.Show
    Me.Enabled = False
    Me.Hide
  End If
End Sub

Private Sub Timer2_Timer()
  Timer2.Enabled = False
  Cpts = Cpts + 10
  If (Cpts < 255) Then
    ShadowCode.MainShadow Me.hwnd, Cpts
    Timer2.Enabled = True
  Else
    'INIT
    Main.Check1.BackColor = RGB(78, 17, 17)
    Main.Check2.BackColor = Main.Check1.BackColor
    Main.Combo1.AddItem "Windows 3.1"
    Main.Combo1.AddItem "Windows 95"
    Main.Combo1.AddItem "Windows 98"
    Main.Combo1.AddItem "Windows Me"
    Main.Combo1.AddItem "Windows NT 3.1"
    Main.Combo1.AddItem "Windows NT 3.5x"
    Main.Combo1.AddItem "Windows NT 4.0"
    Main.Combo1.AddItem "Windows 2000"
    Main.Combo1.AddItem "Incomplete"
    Main.Combo2.AddItem "None"
    Main.Combo2.AddItem "Curtain"
    Main.Combo2.AddItem "Slide"
    Main.Combo2.AddItem "Still"
    Main.Combo2.AddItem "Diagonal"
    Main.Combo3.AddItem "XP-Style"
    If (vars.WinDesign = 51) Then
      Main.Combo3.ListIndex = 0
    Else
      Main.Combo3.ListIndex = 0
    End If
    vars.TType = GetSetting("LPowerPrompts", "Type", "LogoffWindow")
    vars.ChadShadow = GetSetting("LPowerPrompts", "Type", "ShadowType")
    vars.MoveWin = GetSetting("LPowerPrompts", "Type", "IsMovable")
    vars.RegAnimation = vars.EnableAnimations
    If IsStringInteger(vars.TType) Then
    Else
      vars.TrueOrFalse = False
    End If
    If IsStringInteger(vars.ChadShadow) Then
    Else
      vars.TrueOrFalse = False
    End If
    If (vars.TrueOrFalse = False) Then
      Timer1.Enabled = False
      Timer2.Enabled = False
      Timer3.Enabled = False
      Timer4.Enabled = False
      MsgBox "One or more numeric registry keys were not declared as an integer! LPPConfig Utility will use the default options instead.", vbExclamation, "Error"
      vars.TType = 8
      vars.ChadShadow = 50
      Main.Combo1.ListIndex = 7
      Main.Combo2.ListIndex = 3
      Label1.Caption = "Getting Registry Values..."
      ShadowCode.MainShadow Me.hwnd, 255
      Me.Enabled = False
      Me.Visible = False
      Me.Hide
      If (vars.EnableAnimations = True) Then
        MainFrame.Enabled = True
        MainFrame.Left = (Screen.Width - MainFrame.Width) / 2
        MainFrame.Top = (Screen.Height - MainFrame.Height) / 2
        MainFrame.Show
        Unload Me
      ElseIf (vars.EnableAnimations = False) Then
        Main.Enabled = True
        Main.Left = (Screen.Width - Main.Width) / 2
        Main.Top = (Screen.Height - Main.Height) / 2
        Main.Show
        Unload Me
      End If
    End If
    'Check TTYPE
    If (vars.TType = 1) Then
      Main.Combo1.ListIndex = 0
    ElseIf (vars.TType = 2) Then
      Main.Combo1.ListIndex = 1
    ElseIf (vars.TType = 3) Then
      Main.Combo1.ListIndex = 2
    ElseIf (vars.TType = 4) Then
      Main.Combo1.ListIndex = 3
    ElseIf (vars.TType = 5) Then
      Main.Combo1.ListIndex = 4
    ElseIf (vars.TType = 6) Then
      Main.Combo1.ListIndex = 5
    ElseIf (vars.TType = 7) Then
      Main.Combo1.ListIndex = 6
    ElseIf (vars.TType = 8) Then
      Main.Combo1.ListIndex = 7
    ElseIf (vars.TType = 9) Then
      Main.Combo1.ListIndex = 8
    ElseIf (vars.TType = 16) Then
      Main.Combo1.Text = "Lorem Ipsum"
    Else
      Main.Combo1.ListIndex = 7
    End If
    'Check SHADOW
    If (vars.ChadShadow = 31) Then
      Main.Combo2.ListIndex = 0
    ElseIf (vars.ChadShadow = 41) Then
      Main.Combo2.ListIndex = 1
    ElseIf (vars.ChadShadow = 40) Then
      Main.Combo2.ListIndex = 2
    ElseIf (vars.ChadShadow = 50) Then
      Main.Combo2.ListIndex = 3
    ElseIf (vars.ChadShadow = 51) Then
      Main.Combo2.ListIndex = 4
    Else
      Main.Combo2.ListIndex = 3
    End If
    'Check MOVE
    If (vars.MoveWin = "True") Then
      Main.Check1.Value = 1
    ElseIf (vars.MoveWin = "False") Then
      Main.Check1.Value = 0
    Else
      Main.Check1.Value = 0
    End If
    'JUMP NEXT
    Timer3.Enabled = True
  End If
End Sub

Private Sub Timer3_Timer()
  Timer3.Enabled = False
  Cpts = Cpts - 10
  If (Cpts < 255 And Cpts > 0) Then
    ShadowCode.MainShadow Me.hwnd, Cpts
    Timer3.Enabled = True
  Else
    Me.Hide
  End If
End Sub

Private Sub Timer4_Timer()
  Timer4.Enabled = False
  If (LblPtr = 0) Then
    Label1.Visible = True
    Label1.Caption = "G"
  ElseIf (LblPtr = 1) Then
    Label1.Caption = "Ge"
  ElseIf (LblPtr = 2) Then
    Label1.Caption = "Get"
  ElseIf (LblPtr = 3) Then
    Label1.Caption = "Gett"
  ElseIf (LblPtr = 4) Then
    Label1.Caption = "Getti"
  ElseIf (LblPtr = 5) Then
    Label1.Caption = "Gettin"
  ElseIf (LblPtr = 6) Then
    Label1.Caption = "Getting"
  ElseIf (LblPtr = 7) Then
    Label1.Caption = "Getting "
  ElseIf (LblPtr = 8) Then
    Label1.Caption = "Getting R"
  ElseIf (LblPtr = 9) Then
    Label1.Caption = "Getting Re"
  ElseIf (LblPtr = 10) Then
    Label1.Caption = "Getting Reg"
  ElseIf (LblPtr = 11) Then
    Label1.Caption = "Getting Regi"
  ElseIf (LblPtr = 12) Then
    Label1.Caption = "Getting Regis"
  ElseIf (LblPtr = 13) Then
    Label1.Caption = "Getting Regist"
  ElseIf (LblPtr = 14) Then
    Label1.Caption = "Getting Registr"
  ElseIf (LblPtr = 15) Then
    Label1.Caption = "Getting Registry"
  ElseIf (LblPtr = 16) Then
    Label1.Caption = "Getting Registry "
  ElseIf (LblPtr = 17) Then
    Label1.Caption = "Getting Registry V"
  ElseIf (LblPtr = 18) Then
    Label1.Caption = "Getting Registry Va"
  ElseIf (LblPtr = 19) Then
    Label1.Caption = "Getting Registry Val"
  ElseIf (LblPtr = 20) Then
    Label1.Caption = "Getting Registry Valu"
  ElseIf (LblPtr = 21) Then
    Label1.Caption = "Getting Registry Value"
  ElseIf (LblPtr = 22) Then
    Label1.Caption = "Getting Registry Values"
  ElseIf (LblPtr = 23) Then
    Label1.Caption = "Getting Registry Values."
  ElseIf (LblPtr = 24) Then
    Label1.Caption = "Getting Registry Values.."
  ElseIf (LblPtr = 25) Then
    Label1.Caption = "Getting Registry Values..."
  End If
  LblPtr = LblPtr + 1
  Timer4.Enabled = True
End Sub
