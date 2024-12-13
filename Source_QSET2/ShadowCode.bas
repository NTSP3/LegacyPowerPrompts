Attribute VB_Name = "ShadowCode"
Option Explicit

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Const GWL_EXSTYLE = -20
Const WS_EX_LAYERED = &H80000
Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2

Sub MainShadow(ByVal hwnd As Long, ByVal transparency As Byte)
    Dim currentExStyle As Long
    currentExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    SetWindowLong hwnd, GWL_EXSTYLE, currentExStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, RGB(128, 128, 128), transparency, LWA_COLORKEY Or LWA_ALPHA
End Sub

