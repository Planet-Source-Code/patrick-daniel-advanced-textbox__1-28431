Attribute VB_Name = "Module1"
Option Explicit

Type TEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
End Type

Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long
Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Public Const GWL_WNDPROC = (-4)
'Public Const WM_PASTE = &H302

Type POINTAPI
   x As Long
   y As Long
End Type


'Type Msg
'   hwnd As Long
'   message As Long
'   wParam As Long
'   lParam As Long
'   time As Long
'   pt As POINTAPI
'End Type

'Dim mlPrevProc As Long

'Public Sub Hook(robjTextbox As TextBox)
'On Error Resume Next
'
'   mlPrevProc = SetWindowLong(robjTextbox.hwnd, GWL_WNDPROC, AddressOf TextProc)
'
'End Sub
'
'
'Public Sub UnHook(robjTextbox As TextBox)
'On Error Resume Next
'
'   SetWindowLong robjTextbox.hwnd, GWL_WNDPROC, mlPrevProc
'
'End Sub
'
'
'Public Function TextProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'On Error Resume Next
'
'   If uMsg = WM_PASTE Then
'      uMsg = 0
'   End If
'
'   TextProc = CallWindowProc(mlPrevProc, hwnd, uMsg, wParam, lParam)
'
'End Function

