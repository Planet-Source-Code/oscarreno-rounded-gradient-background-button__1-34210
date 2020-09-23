Attribute VB_Name = "Util"
Option Explicit

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long 'API for finding the hWnd of the window under the cursor

Public Function IsHot(hWnd As Long) As Boolean
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    IsHot = WindowFromPoint(CursorPosition.X, CursorPosition.Y) = hWnd 'Return     whether the object is hot
End Function
