VERSION 5.00
Begin VB.Form Win95 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shut Down Windows"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   Enabled         =   0   'False
   Icon            =   "Win95.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Help"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "No"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Yes"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Restart the computer in MS-DOS mode?"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Restart the computer?"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Shut down the computer?"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      Picture         =   "Win95.frx":014A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure you want to:"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Win95"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If (Option1.Value = True) Then
    Shell "nsp exitwin poweroff", vbHide
  ElseIf (Option2.Value = True) Then
    Shell "nsp exitwin reboot", vbHide
  End If
  End
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
