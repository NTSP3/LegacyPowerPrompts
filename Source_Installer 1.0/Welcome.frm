VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Welcome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Install LPP"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   Icon            =   "Welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
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
      Begin VB.Timer Timer5 
         Interval        =   1
         Left            =   120
         Top             =   3960
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   480
         Top             =   4920
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   4920
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   480
         Top             =   4440
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   120
         Top             =   4440
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Quick Install"
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
         Height          =   3495
         Left            =   1800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   4815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7320
         Top             =   5400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "NirSoft programs are owned by Nir Sofer at www.nirsoft.net"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   5640
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: This program includes software from NirSoft."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   5400
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Installing..."
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
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Install To:"
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
         TabIndex        =   14
         Top             =   840
         Width           =   1455
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
            Caption         =   "Install"
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
         Caption         =   "LPP Version 1.0"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TString As String
Dim formattedTime As String
Dim tempstart As String
Dim TCount As Integer
Dim currentTime As Date
Dim WshShell As Object

Private Sub Command1_Click()
    On Error Resume Next
    CommonDialog1.CancelError = True
    
    ' Set flags to show only folders and drives
    CommonDialog1.Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNFolder
    
    ' Show the open file dialog
    CommonDialog1.ShowOpen
    
    If Err.Number = cdlCancel Then
        ' User clicked Cancel
        'MsgBox "Operation canceled by user"
    ElseIf IsFolder(CommonDialog1.FileName) Then
        ' User selected a folder or drive
        'MsgBox "Selected folder or drive: " & CommonDialog1.FileName
        Text1.Text = CommonDialog1.FileName
    Else
        ' User selected a file
        'MsgBox "Please select a folder."
        TCount = Len(CommonDialog1.FileName) - 1
        Do Until TString = "\"
          TString = Mid(CommonDialog1.FileName, TCount, 1)
          If Not (TString = "\") Then
            TCount = TCount - 1
            Text1.Text = Mid(CommonDialog1.FileName, 1, TCount - 1)
          End If
        Loop
    End If
End Sub

Function IsFolder(path As String) As Boolean
    Dim attr As Integer
    ' Get the attributes of the selected item
    attr = GetAttr(path)
    ' Check if it's a directory
    IsFolder = (attr And vbDirectory) = vbDirectory
End Function

Private Sub Form_Load()
  Set WshShell = Nothing
  Set WshShell = CreateObject("WScript.Shell")
  tempstart = WshShell.SpecialFolders("AllUsersPrograms") & "\LegacyPowerPrompts"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = 1
  Unload BkgdWindow
End Sub

Private Sub Label1_Click()
  Unload Me
End Sub

Private Sub Label25_Click()
  vars.QVal = False
  Label3.Enabled = False
  Text1.Enabled = False
  Command1.Enabled = False
  Label25.Enabled = False
  Picture6.Enabled = False
  Label1.Enabled = False
  Picture5.Enabled = False
  Label4.Visible = True
  Text2.Visible = True
  Text2.Text = "00:00:00>"
  TCount = 0
  Timer1.Enabled = True
  Timer3.Enabled = True
End Sub

Private Sub Picture5_Click()
  Unload Me
End Sub

Private Sub Picture6_Click()
  vars.QVal = False
  Label3.Enabled = False
  Text1.Enabled = False
  Command1.Enabled = False
  Label25.Enabled = False
  Picture6.Enabled = False
  Label1.Enabled = False
  Picture5.Enabled = False
  Label4.Visible = True
  Text2.Visible = True
  Text2.Text = "00:00:00>"
  TCount = 0
  Timer1.Enabled = True
  Timer3.Enabled = True
End Sub

Private Sub Text1_Change()
  vars.InsPath = Text1.Text
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
    Text2.Text = formattedTime & "> LegacyPowerPrompts Installer "
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
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Includes software from NirSoft. "
  ElseIf (TCount = 5) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> "
  ElseIf (TCount = 6) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Extracting files: ltxmain.sup "
    FileCopy App.path & "\ltxmain.sup", vars.InsPath & "\qscrdlg.exe"
    'Shell "CMD.EXE /k ""copy """"" & App.path & "\ltxmain.sup"""" """"" & vars.InsPath & "\qscrdlg.exe""""""", vbNormalFocus
    Timer4.Interval = 500
  ElseIf (TCount = 7) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Extracting files: ltxdom.sup "
    FileCopy App.path & "\ltxdom.sup", vars.InsPath & "\config.exe"
    'Shell "CMD.EXE /c ""copy """"" & App.path & "\ltxdom.sup"""" """"" & vars.InsPath & "\config.exe""""""", vbHide
  ElseIf (TCount = 8) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Extracting files: nac.sup "
    FileCopy App.path & "\nac.sup", vars.InsPath & "\nsp.exe"
    'Shell "CMD.EXE /c ""copy """"" & App.path & "\nac.sup"""" """"" & vars.InsPath & "\nsp.exe""""""", vbHide
  ElseIf (TCount = 9) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Creating Registry values "
    SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "0"
    SaveSetting "LPowerPrompts", "Info", "InstalledPath", "Path"
  ElseIf (TCount = 10) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Setting Registry values "
    SaveSetting "LPowerPrompts", "Type", "LogoffWindow", "8"
    SaveSetting "LPowerPrompts", "Info", "InstalledPath", vars.InsPath
    Timer4.Interval = 1000
  ElseIf (TCount = 11) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Creating Desktop Shortcuts "
    Set WshShell = Nothing
    Set WshShell = CreateObject("WScript.Shell")
    With WshShell.CreateShortcut(WshShell.SpecialFolders("Desktop") & "\LPP Shutdown.lnk")
        .TargetPath = vars.InsPath & "\qscrdlg.exe"
        .WorkingDirectory = vars.InsPath
        .IconLocation = vars.InsPath & "\qscrdlg.exe,0"
        .Description = "Show LegacyPowerPrompts Shutdown dialog."
        .Save
    End With
    Set WshShell = Nothing
    Set WshShell = CreateObject("WScript.Shell")
    With WshShell.CreateShortcut(WshShell.SpecialFolders("Desktop") & "\LPP Configuration Utility.lnk")
        .TargetPath = vars.InsPath & "\config.exe"
        .WorkingDirectory = vars.InsPath
        .IconLocation = vars.InsPath & "\config.exe,0"
        .Description = "Change the settings used in LPP Shutdown dialog."
        .Save
    End With
    Set WshShell = Nothing
    Timer4.Interval = 250
  ElseIf (TCount = 12) Then
    If Dir(tempstart, vbDirectory) = "" Then
      MkDir tempstart
    End If
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Creating Startmenu Shortcuts "
    Set WshShell = Nothing
    Set WshShell = CreateObject("WScript.Shell")
    With WshShell.CreateShortcut(WshShell.SpecialFolders("AllUsersPrograms") & "\LegacyPowerPrompts\Shutdown dialog.lnk")
        .TargetPath = vars.InsPath & "\qscrdlg.exe"
        .WorkingDirectory = vars.InsPath
        .IconLocation = vars.InsPath & "\qscrdlg.exe,0"
        .Description = "Show LegacyPowerPrompts Shutdown dialog."
        .Save
    End With
    Set WshShell = Nothing
    Set WshShell = CreateObject("WScript.Shell")
    With WshShell.CreateShortcut(WshShell.SpecialFolders("AllUsersPrograms") & "\LegacyPowerPrompts\Configuration Utility.lnk")
        .TargetPath = vars.InsPath & "\config.exe"
        .WorkingDirectory = vars.InsPath
        .IconLocation = vars.InsPath & "\config.exe,0"
        .Description = "Change the settings used in LPP Shutdown dialog."
        .Save
    End With
    Set WshShell = Nothing
    Set WshShell = CreateObject("WScript.Shell")
    With WshShell.CreateShortcut(WshShell.SpecialFolders("AllUsersPrograms") & "\LegacyPowerPrompts\Uninstall LPP.lnk")
        .TargetPath = vars.InsPath & "\unins.exe"
        .WorkingDirectory = vars.InsPath
        .IconLocation = vars.InsPath & "\unins.exe,0"
        .Description = "Uninstall LegacyPowerPrompts."
        .Save
    End With
    Set WshShell = Nothing
  ElseIf (TCount = 13) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Creating uninstaller "
    'Shell "CMD.EXE /c ""copy """"" & App.path & "\ltxdel.sup"""" """"" & vars.InsPath & "unins.exe""""""", vbHide
    FileCopy App.path & "\ltxdel.sup", vars.InsPath & "\unins.exe"
  ElseIf (TCount = 14) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> "
    Timer4.Interval = 100
  ElseIf (TCount = 15) Then
    If (Right(Text2.Text, 1) = "_") Then Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
    Text2.Text = Text2.Text & vbCrLf & formattedTime & "> Done!_"
    Timer1.Enabled = False
    Timer2.Enabled = False
    'Picture6.Enabled = True
    'Label25.Caption = "Next"
    'Label25.Enabled = True
    vars.Finish = True
    Label1.Enabled = True
    Picture5.Enabled = True
    Exit Sub
  End If
  Timer4.Enabled = True
End Sub

Private Sub Timer5_Timer()
  currentTime = Time
  formattedTime = Format(currentTime, "hh:mm:ss")
End Sub
