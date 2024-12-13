VERSION 5.00
Begin VB.Form WinMeR 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shut Down Windows"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4905
   Enabled         =   0   'False
   Icon            =   "WinMer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Help"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      Picture         =   "WinMer.frx":014A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ends your session and shuts down Windows so that you can safely turn off power."
      Height          =   735
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "What do you want the computer to do?"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "WinMeR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
  If (Combo1.ListIndex = "0") Then Label2.Caption = "Ends your session and shuts down Windows so that you can safely turn off power."
  If (Combo1.ListIndex = "1") Then Label2.Caption = "Ends your session, shuts down Windows, and starts Windows again."
  If (Combo1.ListIndex = "2") Then Label2.Caption = "Maintains your session, keeping the computer running on low power with data still in memory."
  If (Combo1.ListIndex = "3") Then Label2.Caption = "Saves your session to disk so that you can safely turn off power. Your session is restored the next time you start Windows."
End Sub

Private Sub Command1_Click()
  If (Combo1.ListIndex = "0") Then
    Shell "nsp exitwin poweroff", vbHide
  ElseIf (Combo1.ListIndex = "1") Then
    Shell "nsp exitwin reboot", vbHide
  ElseIf (Combo1.ListIndex = "2") Then
    Shell "nsp standby", vbHide
  ElseIf (Combo1.ListIndex = "3") Then
    Shell "nsp hibernate", vbHide
  End If
  End
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Form_Load()
  Combo1.AddItem "Shut down"
  Combo1.AddItem "Restart"
  Combo1.AddItem "Stand by"
  Combo1.AddItem "Hibernate"
  Combo1.ListIndex = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
