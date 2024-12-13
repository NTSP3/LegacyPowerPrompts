VERSION 5.00
Begin VB.Form WinForm 
   BorderStyle     =   0  'None
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   Enabled         =   0   'False
   Icon            =   "WinForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Checking registry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "WinForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DlgTypeConversion As String
Dim userName As String

Private Sub Form_Load()
  'Load Form properties
  Me.Hide
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
  
  'Check if app is disabled
  DlgTypeConversion = GetSetting("LPowerPrompts", "Type", "Disabled")
  If (DlgTypeConversion = "True") Then End
  DlgTypeConversion = ""
  
  'Load variables
  vars.CurrentUser = "user"
  vars.UseDefault = False
  
  'Get the Shutdown screen type from registry
  DlgTypeConversion = GetSetting("LPowerPrompts", "Type", "LogoffWindow")
  'Check if hidden num is activated
  If (DlgTypeConversion = "16") Then
    'vars.Fatmode = 1
    Fatmode.Enabled = True
    Fatmode.Show
    Unload Me
    Exit Sub
  End If
  'Check if the item is a valid number between 1 and 9
  If Not IsNumeric(DlgTypeConversion) Then
    MsgBox "Legacy Power Prompts was unable to get the type of shutdown dialog from the Windows Registry. Possible fixes include Re-installation and changing configuration using the lppconfig utility.", vbExclamation, "Error (LPP)"
    End
  ElseIf ((DlgTypeConversion < 1) Or (DlgTypeConversion > 9)) Then
    MsgBox "Legacy Power Prompts was unable to get the type of shutdown dialog from the Windows Registry. Possible fixes include Re-installation and changing configuration using the lppconfig utility.", vbExclamation, "Error (LPP)"
    End
  End If
  'Save if the item is a number
  vars.DlgType = DlgTypeConversion
  
  'Put the Shadow type (needs to be done BEFORE screenshadow)
  DlgTypeConversion = ""
  DlgTypeConversion = GetSetting("LPowerPrompts", "Type", "ShadowType")
  If IsNumeric(DlgTypeConversion) Then
    If Not (DlgTypeConversion = "31" Or DlgTypeConversion = "40" Or DlgTypeConversion = "41" Or DlgTypeConversion = "50" Or DlgTypeConversion = "51") Then
      If ((vars.DlgType = "1") Or (vars.DlgType = "5") Or (vars.DlgType = "6")) Then vars.ShadowMode = "1"
      If ((vars.DlgType = "2") Or (vars.DlgType = "3") Or (vars.DlgType = "4")) Then vars.ShadowMode = "2"
      If (vars.DlgType = "7") Then vars.ShadowMode = "3"
      If (vars.DlgType = "8") Then vars.ShadowMode = "4"
      If (vars.DlgType = "9") Then vars.ShadowMode = "5"
    ElseIf (DlgTypeConversion = "31") Then
      vars.ShadowMode = "1"
    ElseIf (DlgTypeConversion = "41") Then
      vars.ShadowMode = "2"
    ElseIf (DlgTypeConversion = "40") Then
      vars.ShadowMode = "3"
    ElseIf (DlgTypeConversion = "50") Then
      vars.ShadowMode = "4"
    ElseIf (DlgTypeConversion = "51") Then
      vars.ShadowMode = "5"
    End If
  Else
    If ((vars.DlgType = "2") Or (vars.DlgType = "3") Or (vars.DlgType = "4")) Then vars.ShadowMode = "2"
    If (vars.DlgType = "7") Then vars.ShadowMode = "3"
    If (vars.DlgType = "8") Then vars.ShadowMode = "4"
    If (vars.DlgType = "9") Then vars.ShadowMode = "5"
  End If
  'Check if the window can be moved
  DlgTypeConversion = ""
  DlgTypeConversion = GetSetting("LPowerPrompts", "Type", "IsMovable")
  If Not (DlgTypeConversion = "True" Or DlgTypeConversion = "False") Then
    vars.UseDefault = True
  ElseIf (DlgTypeConversion = "True") Then
    vars.MoveWindow = True
  ElseIf (DlgTypeConversion = "False") Then
    vars.MoveWindow = False
  End If
  'Load the screenfreeze window
  ScreenFreeze.Enabled = False
  ScreenFreeze.Show
  
  'Load the appropriate form
  If (vars.UseDefault = True) Then
    If (vars.DlgType = "1") Then
      Win3xR.Left = (Screen.Width - Win3x.Width) / 2
      Win3xR.Top = (Screen.Height - Win3x.Height) / 2
      Win3xR.Enabled = True
      Win3xR.Show
    ElseIf (vars.DlgType = "2") Then
      Win95.Left = (Screen.Width - Win95.Width) / 2
      Win95.Top = (Screen.Height - Win95.Height) / 2
      Win95.Enabled = True
      Win95.Show
    ElseIf (vars.DlgType = "3") Then
      Win98.Left = (Screen.Width - Win98.Width) / 2
      Win98.Top = (Screen.Height - Win98.Height) / 2
      Win98.Enabled = True
      Win98.Show
    ElseIf (vars.DlgType = "4") Then
      WinMe.Left = (Screen.Width - WinMe.Width) / 2
      WinMe.Top = (Screen.Height - WinMe.Height) / 2
      WinMe.Enabled = True
      WinMe.Show
    ElseIf (vars.DlgType = "5") Then
      NT31R.Left = (Screen.Width - NT31.Width) / 2
      NT31R.Top = (Screen.Height - NT31.Height) / 3
      NT31R.Enabled = True
      NT31R.Show
    ElseIf (vars.DlgType = "6") Then
      NT35R.Left = (Screen.Width - NT35.Width) / 2
      NT35R.Top = (Screen.Height - NT35.Height) / 3
      NT35R.Enabled = True
      NT35R.Show
    ElseIf (vars.DlgType = "7") Then
      NT4.Left = (Screen.Width - NT4.Width) / 2
      NT4.Top = (Screen.Height - NT4.Height) / 2
      NT4.Enabled = True
      NT4.Show
    ElseIf (vars.DlgType = "8") Then
      vars.CurrentUser = Environ("USERNAME") 'Get username for use in windows 2000 power form
      Win2k.Left = (Screen.Width - Win2k.Width) / 2
      Win2k.Top = (Screen.Height - Win2k.Height) / 3
      Win2k.Enabled = True
      Win2k.Show
    ElseIf (vars.DlgType = "9") Then
      'vars.WLeft = ScreenShadow.Image1.Left
      'vars.WTop = ScreenShadow.Image1.Top
      Whistler.Left = (Screen.Width - Whistler.Width) / 2
      Whistler.Top = (Screen.Height - Whistler.Height) / 2
      Whistler.Enabled = True
      Whistler.Show
    End If
  ElseIf (vars.MoveWindow = False) Then
    If (vars.DlgType = "1") Then
      Win3x.Left = (Screen.Width - Win3x.Width) / 2
      Win3x.Top = (Screen.Height - Win3x.Height) / 2
      Win3x.Enabled = True
      Win3x.Show
    ElseIf (vars.DlgType = "2") Then
      Win95.Left = (Screen.Width - Win95.Width) / 2
      Win95.Top = (Screen.Height - Win95.Height) / 2
      Win95.Enabled = True
      Win95.Show
    ElseIf (vars.DlgType = "3") Then
      Win98.Left = (Screen.Width - Win98.Width) / 2
      Win98.Top = (Screen.Height - Win98.Height) / 2
      Win98.Enabled = True
      Win98.Show
    ElseIf (vars.DlgType = "4") Then
      WinMe.Left = (Screen.Width - WinMe.Width) / 2
      WinMe.Top = (Screen.Height - WinMe.Height) / 2
      WinMe.Enabled = True
      WinMe.Show
    ElseIf (vars.DlgType = "5") Then
      NT31.Left = (Screen.Width - NT31.Width) / 2
      NT31.Top = (Screen.Height - NT31.Height) / 3
      NT31.Enabled = True
      NT31.Show
    ElseIf (vars.DlgType = "6") Then
      NT35.Left = (Screen.Width - NT35.Width) / 2
      NT35.Top = (Screen.Height - NT35.Height) / 3
      NT35.Enabled = True
      NT35.Show
    ElseIf (vars.DlgType = "7") Then
      NT4.Left = (Screen.Width - NT4.Width) / 2
      NT4.Top = (Screen.Height - NT4.Height) / 2
      NT4.Enabled = True
      NT4.Show
    ElseIf (vars.DlgType = "8") Then
      vars.CurrentUser = Environ("USERNAME") 'Get username for use in windows 2000 power form
      Win2k.Left = (Screen.Width - Win2k.Width) / 2
      Win2k.Top = (Screen.Height - Win2k.Height) / 3
      Win2k.Enabled = True
      Win2k.Show
    ElseIf (vars.DlgType = "9") Then
      'vars.WLeft = ScreenShadow.Image1.Left
      'vars.WTop = ScreenShadow.Image1.Top
      Whistler.Left = (Screen.Width - Whistler.Width) / 2
      Whistler.Top = (Screen.Height - Whistler.Height) / 2
      Whistler.Enabled = True
      Whistler.Show
    End If
  ElseIf (vars.MoveWindow = True) Then
    If (vars.DlgType = "1") Then
      Win3xR.Left = (Screen.Width - Win3x.Width) / 2
      Win3xR.Top = (Screen.Height - Win3x.Height) / 2
      Win3xR.Enabled = True
      Win3xR.Show
    ElseIf (vars.DlgType = "2") Then
      Win95R.Left = (Screen.Width - Win95.Width) / 2
      Win95R.Top = (Screen.Height - Win95.Height) / 2
      Win95R.Enabled = True
      Win95R.Show
    ElseIf (vars.DlgType = "3") Then
      Win98R.Left = (Screen.Width - Win98.Width) / 2
      Win98R.Top = (Screen.Height - Win98.Height) / 2
      Win98R.Enabled = True
      Win98R.Show
    ElseIf (vars.DlgType = "4") Then
      WinMeR.Left = (Screen.Width - WinMe.Width) / 2
      WinMeR.Top = (Screen.Height - WinMe.Height) / 2
      WinMeR.Enabled = True
      WinMeR.Show
    ElseIf (vars.DlgType = "5") Then
      NT31R.Left = (Screen.Width - NT31.Width) / 2
      NT31R.Top = (Screen.Height - NT31.Height) / 3
      NT31R.Enabled = True
      NT31R.Show
    ElseIf (vars.DlgType = "6") Then
      NT35R.Left = (Screen.Width - NT35.Width) / 2
      NT35R.Top = (Screen.Height - NT35.Height) / 3
      NT35R.Enabled = True
      NT35R.Show
    ElseIf (vars.DlgType = "7") Then
      NT4R.Left = (Screen.Width - NT4.Width) / 2
      NT4R.Top = (Screen.Height - NT4.Height) / 2
      NT4R.Enabled = True
      NT4R.Show
    ElseIf (vars.DlgType = "8") Then
      vars.CurrentUser = Environ("USERNAME") 'Get username for use in windows 2000 power form
      Win2kR.Left = (Screen.Width - Win2k.Width) / 2
      Win2kR.Top = (Screen.Height - Win2k.Height) / 3
      Win2kR.Enabled = True
      Win2kR.Show
    ElseIf (vars.DlgType = "9") Then
      'vars.WLeft = ScreenShadow.Image1.Left
      'vars.WTop = ScreenShadow.Image1.Top
      WhistlerR.Left = (Screen.Width - Whistler.Width) / 2
      WhistlerR.Top = (Screen.Height - Whistler.Height) / 2
      WhistlerR.Enabled = True
      WhistlerR.Show
    End If
  End If
  
  'Quit
  Unload Me
End Sub