VERSION 5.00
Begin VB.Form Win98R 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shut Down Windows"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   Enabled         =   0   'False
   Icon            =   "Win98r.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Stand by"
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   720
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Help"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Restart in MS-DOS mode"
      Enabled         =   0   'False
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Restart"
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Shut down"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      Picture         =   "Win98r.frx":014A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "What do you want the computer to do?"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Win98R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If (Option1.Value = True) Then
    Shell "nsp exitwin poweroff", vbHide
  ElseIf (Option2.Value = True) Then
    Shell "nsp exitwin reboot", vbHide
  ElseIf (Option4.Value = True) Then
    Shell "nsp standby", vbHide
  End If
  End
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
