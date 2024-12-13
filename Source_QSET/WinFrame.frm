VERSION 5.00
Begin VB.Form WinFrame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Settings Panel (LPP)"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   Icon            =   "WinFrame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Saved!"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dialog type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "WinFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  SaveSetting "LPowerPrompts", "Type", "LogoffWindow", Text1.Text
  Label2.Visible = True
  Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Label2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
