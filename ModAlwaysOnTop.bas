Attribute VB_Name = "ModWindow"
Option Explicit
'  This module allows you to place a window Always on Top or
' Always Off Top.

Public Sub SetAlwaysOnTop(Value As OnOrOff, WindowHandle As Long)
If Value = AlwaysOff Then
SetWindowPos WindowHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
                     Else
SetWindowPos WindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
End If
End Sub

Public Function GetWinPos(Handle As Long) As RECT
Dim r As RECT
GetWindowRect Handle, r
GetWinPos = r
End Function

Public Sub SetWinPos(Handle As Long, r As RECT)
SetWindowPos Handle, 0&, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, SWP_NOZORDER Or SWP_NOACTIVATE
End Sub
'
' This sub places a form at the bottom of the screen
' the difference is that it takes extra care not to put the window over
' the taskbar. It's long but it is not as complicated as it seems.
'
Public Sub PlaceWindow(Handle As Long)
Dim MyRect As RECT
Dim DeskRect As RECT
Dim TaskBarRect As RECT
Dim Rct As RECT
TaskBarRect = GetWinPos(FindWindow(TaskBar_Name, vbNullString))
DeskRect = GetWinPos(GetDesktopWindow)
MyRect = GetWinPos(Handle)
'
'
' Assume the task bar is on the bottom of the screen
Rct.Left = -2
'If the screen width is 1024 then taskbar has 1026 pixels
Rct.Right = DeskRect.Right + 2
Rct.Top = DeskRect.Bottom - (TaskBarRect.Bottom - TaskBarRect.Top) + 2
'If the screen height is 768 then taskbar has 770 pixels
Rct.Bottom = DeskRect.Bottom + 2
If (CompareRects(Rct, TaskBarRect)) Then
' Taskbar is at the bottom of the screen
MyRect.Top = TaskBarRect.Top - (MyRect.Bottom - MyRect.Top)
MyRect.Bottom = TaskBarRect.Top
MyRect.Left = DeskRect.Right - (MyRect.Right - MyRect.Left)
MyRect.Right = DeskRect.Right
SetWinPos Handle, MyRect
GoTo PositionDetected      ' Skip the other checks
End If
'
' Assume the task bar is on the left part of the screen
'
Rct.Left = -2
Rct.Right = TaskBarRect.Right
Rct.Top = -2
Rct.Bottom = DeskRect.Bottom + 2
If (CompareRects(Rct, TaskBarRect)) Then
'Taskbar is on the left part of the screen
MyRect.Left = DeskRect.Right - (MyRect.Right - MyRect.Left)
MyRect.Right = DeskRect.Right
MyRect.Top = DeskRect.Bottom - (MyRect.Bottom - MyRect.Top) + 1
MyRect.Bottom = DeskRect.Bottom + 1
SetWinPos Handle, MyRect
GoTo PositionDetected      ' Skip the other checks
End If
'
' Assume the task bar is on the right part of the screen
'
Rct.Left = TaskBarRect.Left
Rct.Right = DeskRect.Right + 2
Rct.Top = -2
Rct.Bottom = DeskRect.Bottom + 2
If (CompareRects(Rct, TaskBarRect)) Then
'Taskbar is on the right part of the screen
MyRect.Left = TaskBarRect.Left - (MyRect.Right - MyRect.Left) + 1
MyRect.Right = TaskBarRect.Left + 1
MyRect.Top = DeskRect.Bottom - (MyRect.Bottom - MyRect.Top) + 1
MyRect.Bottom = DeskRect.Bottom + 1
SetWinPos Handle, MyRect
GoTo PositionDetected      ' Skip the other checks
End If
'
' Assume the task bar is on the top of the screen
'
Rct.Left = -2
Rct.Right = DeskRect.Right + 2
Rct.Top = -2
Rct.Bottom = TaskBarRect.Bottom
If (CompareRects(Rct, TaskBarRect)) Then
'Taskbar is on the top of the screen
MyRect.Left = DeskRect.Right - (MyRect.Right - MyRect.Left)
MyRect.Right = DeskRect.Right
MyRect.Top = DeskRect.Bottom - (MyRect.Bottom - MyRect.Top) + 1
MyRect.Bottom = DeskRect.Bottom + 1
SetWinPos Handle, MyRect
End If
'
' If the tskbar is not detected the main form is left where
' I have told Visual Basic to put it, somewhere between the
' the 640x480 and the 800x600 guide lines.
'
PositionDetected:
End Sub

Private Function CompareRects(a1 As RECT, a2 As RECT) As Boolean
If a1.Bottom <> a2.Bottom Or a1.Left <> a2.Left Or a1.Right <> a2.Right Or a1.Top <> a2.Top Then
CompareRects = False
Else
CompareRects = True
End If
End Function

