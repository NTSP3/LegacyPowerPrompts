VERSION 5.00
Begin VB.MDIForm BkgdWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Legacy Power Prompts Configuration Utility"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "BkgdWindow.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "BkgdWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Me.Height = vars.Mheight
  Me.Width = vars.Mwidth
  Timer1.Enabled = True
End Sub
