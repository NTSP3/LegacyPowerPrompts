VERSION 5.00
Begin VB.Form Whistler 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2955
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4635
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "Whistler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      Picture         =   "Whistler.frx":014A
      ScaleHeight     =   735
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Placeholder texts for functions not built yet."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Restart"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Off"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sleep"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   4680
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "Whistler"
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

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub Label1_Click()
  Shell "nsp standby", vbHide
End Sub

Private Sub Label2_Click()
  Shell "nsp exitwin poweroff", vbHide
End Sub

Private Sub Label3_Click()
  Shell "nsp exitwin reboot", vbHide
End Sub

Private Sub Label4_Click()
  End
End Sub
