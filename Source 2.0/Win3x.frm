VERSION 5.00
Begin VB.Form Win3x 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                    Exit Windows                                   "
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4335
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "Win3x.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         Picture         =   "Win3x.frx":014A
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "This will end your Windows session."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Win3x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
  End
End Sub

Private Sub Command1_Click()
  Me.Hide
  Me.Enabled = False
  Shell "nsp exitwin poweroff", vbHide
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

