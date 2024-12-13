VERSION 5.00
Begin VB.Form ScreenFreeze 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "ScreenFreeze.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "ScreenFreeze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  ShadowCode.MainShadow Me.hwnd, 1
  Me.Left = 0
  Me.Top = 0
  Me.Height = Screen.Height
  Me.Width = Screen.Width
  ScreenShadow.Enabled = False
  ScreenShadow.Show
End Sub
