VERSION 5.00
Begin VB.Form Win2kR 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shut Down Windows"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   Enabled         =   0   'False
   Icon            =   "Win2kr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      Picture         =   "Win2kr.frx":014A
      ScaleHeight     =   1215
      ScaleWidth      =   6135
      TabIndex        =   7
      Top             =   0
      Width           =   6135
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Help"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      Picture         =   "Win2kr.frx":17550
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   960
      TabIndex        =   6
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "What do you want the computer to do?"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
End
Attribute VB_Name = "Win2kR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
  If (Combo1.ListIndex = "0") Then Label2.Caption = "Ends your session, leaving the computer running on full power."
  If (Combo1.ListIndex = "1") Then Label2.Caption = "Ends your session and shuts down Windows so that you can safely turn off power."
  If (Combo1.ListIndex = "2") Then Label2.Caption = "Ends your session, shuts down Windows, and starts Windows again."
  If (Combo1.ListIndex = "3") Then Label2.Caption = "Maintains your session, keeping the computer running on low power with data still in memory."
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
  Me.BackColor = RGB(212, 208, 200)
  Command1.BackColor = RGB(212, 208, 200)
  Command2.BackColor = RGB(212, 208, 200)
  Command3.BackColor = RGB(212, 208, 200)
  Picture1.BackColor = RGB(212, 208, 200)
  Picture2.BackColor = RGB(212, 208, 200)
  Combo1.AddItem "Log off " & vars.CurrentUser
  Combo1.AddItem "Shut down"
  Combo1.AddItem "Restart"
  Combo1.AddItem "Stand by"
  Combo1.ListIndex = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
