VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Uninstall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uninstall LPP"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   Icon            =   "Welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "Welcome.frx":014A
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
      Picture         =   "Welcome.frx":1DD3C
      ScaleHeight     =   5895
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   840
      Width           =   2535
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   480
         Top             =   4680
      End
      Begin VB.Timer Timer5 
         Interval        =   1
         Left            =   120
         Top             =   4200
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   480
         Top             =   5160
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   120
         Top             =   4680
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   5160
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Uninstall LPP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "   Steps   "
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
      Picture         =   "Welcome.frx":4EB56
      ScaleHeight     =   5895
      ScaleWidth      =   7815
      TabIndex        =   10
      Top             =   840
      Width           =   7815
      Begin VB.TextBox Text2 
         Height          =   3135
         Left            =   1920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1920
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   5055
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7320
         Top             =   5400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Uninstalling..."
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
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Uninstall from:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1695
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
      Picture         =   "Welcome.frx":E4C88
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
         Picture         =   "Welcome.frx":F771A
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
         Picture         =   "Welcome.frx":F8EEC
         ScaleHeight     =   390
         ScaleWidth      =   1155
         TabIndex        =   6
         Top             =   120
         Width           =   1155
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Uninstall"
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
         Picture         =   "Welcome.frx":FA6BE
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
         Caption         =   "LPP uninstaller Version 1.0"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Uninstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TCount As Integer
Dim currentTime As Date
Dim formattedTime As String

Private Sub Form_Load()
  Me.Hide
  Me.Enabled = False
  result = MsgBox("Uninstalling LegacyPowerPrompts will remove all of its files and clears the settings that you set. The icons and uninstaller would remain, which would have to be manually deleted later. Do you wish to uninstall LegacyPowerPrompts?", vbQuestion + vbYesNo, "Uninstall LPP?")
  If result = vbYes Then
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    vars.unInsPath = GetSetting("LPowerPrompts", "Info", "InstalledPath")
    Text1.Text = vars.unInsPath
    TCount = 0
    vars.unInsFinish = False
    vars.DontExit = False
    Text2.Text = "00:00:00> "
    Me.Show
    Me.Enabled = True
  Else
    End
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = 1
  If (vars.unInsFinish = True) Then End
  If Not (vars.DontExit = True) Then
    result = MsgBox("Are you sure that you want to exit LegacyPowerPrompts uninstaller?", vbQuestion + vbYesNo, "Exit uninstaller")
    If result = vbYes Then
      End
    Else
      
    End If
  End If
End Sub

Private Sub Label1_Click()
  Unload Me
End Sub

Private Sub Label25_Click()
  vars.DontExit = True
  Label3.Enabled = False
  Text1.Enabled = False
  Label1.Enabled = False
  Picture5.Enabled = False
  Label25.Enabled = False
  Picture6.Enabled = False
  Label4.Enabled = True
  Text2.Enabled = True
  Label4.Visible = True
  Text2.Visible = True
  Timer1.Enabled = True
  Timer3.Enabled = True
End Sub

Private Sub Picture5_Click()
  Unload Me
End Sub

Private Sub Picture6_Click()
  vars.DontExit = True
  Label3.Enabled = False
  Text1.Enabled = False
  Label1.Enabled = False
  Picture5.Enabled = False
  Label25.Enabled = False
  Picture6.Enabled = False
  Label4.Enabled = True
  Text2.Enabled = True
  Label4.Visible = True
  Text2.Visible = True
  Timer1.Enabled = True
  Timer3.Enabled = True
End Sub

Private Sub Text1_Change()
  Label25.Enabled = True
  Picture6.Enabled = True
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Text2.Text = Text2.Text & "_"
  Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
  Timer2.Enabled = False
  Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
  Timer1.Enabled = True
End Sub

Private Sub Timer3_Timer()
  Timer3.Enabled = False
  TCount = TCount + 1
  If (TCount = 1) Then
    Text2.Text = formattedTime & "> LegacyPowerPrompts Uninstaller "
  End If
  Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
  Timer4.Enabled = False
  TCount = TCount + 1
  If (TCount = 2) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Version 1.0 "
    Timer4.Interval = 100
  ElseIf (TCount = 3) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> -------------------------------------------------------------------------------- "
  ElseIf (TCount = 4) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Made by @Novabits on YouTube. "
    'If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    'Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Includes software from NirSoft. "
  ElseIf (TCount = 5) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> "
  ElseIf (TCount = 6) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Removing file: qscrdlg.exe "
    Kill vars.unInsPath & "\qscrdlg.exe"
    'Shell "CMD.EXE /c ""del " & vars.unInsPath & "\qscrdlg.exe""", vbHide
    Timer4.Interval = 500
  ElseIf (TCount = 7) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Removing file: config.exe "
    Kill vars.unInsPath & "\config.exe"
    'Shell "CMD.EXE /c ""del " & vars.unInsPath & "\config.exe""", vbHide
  ElseIf (TCount = 8) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Removing file: nsp.exe "
    Kill vars.unInsPath & "\nsp.exe"
    'Shell "CMD.EXE /c ""del " & vars.unInsPath & "nsp.exe""", vbHide
  ElseIf (TCount = 9) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Deleting Registry values "
    CreateObject("WScript.Shell").RegDelete "HKCU\Software\VB and VBA Program Settings\LPowerPrompts\Type\LogoffWindow"   'Delete a Registry Value
    CreateObject("WScript.Shell").RegDelete "HKCU\Software\VB and VBA Program Settings\LPowerPrompts\Info\InstalledPath"
    Timer4.Interval = 1000
  'ElseIf (TCount = 10) Then
    'If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    'Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Deleting Registry Keys "
    'CreateObject("WScript.Shell").RegDelete "HKCU\Software\VB and VBA Program Settings\LPowerPrompts\Type"              'Delete a Registry Key (it must have no subkeys)
    'CreateObject("WScript.Shell").RegDelete "HKCU\Software\VB and VBA Program Settings\LPowerPrompts\Info"
    'CreateObject("WScript.Shell").RegDelete "HKCU\Software\VB and VBA Program Settings\LPowerPrompts"
    'Timer4.Interval = 1000
  ElseIf (TCount = 10) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> "
    Timer4.Interval = 100
  ElseIf (TCount = 11) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Done!_"
    Timer1.Enabled = False
    Timer2.Enabled = False
    'Picture6.Enabled = True
    'Label25.Caption = "Next"
    'Label25.Enabled = True
    Label1.Enabled = True
    Picture5.Enabled = True
    vars.unInsFinish = True
    Exit Sub
  End If
  Timer4.Enabled = True
End Sub

Private Sub Timer5_Timer()
  currentTime = Time
  formattedTime = Format(currentTime, "hh:mm:ss")
End Sub
