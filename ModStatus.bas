Attribute VB_Name = "ModStatus"
Option Explicit
'
' This module creates, destroys & sets the text of a status bar.
' This way there is no need to use an Active-X control.
'
Private Const ICC_WIN95_CLASSES = &HFF
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000

Private StatusBarHandle As Long

'
' This sub loads the staus bar
'

Public Sub InitStatusBar()
Dim initcc As InitCommonControlsExType
initcc.dwSize = Len(initcc) ' Set the length of the structure
initcc.dwICC = ICC_WIN95_CLASSES ' Set the ICC
InitCommonControlsEx initcc ' Init the Common Controls
' Set the staus bar on the window
StatusBarHandle = CreateWindowEx(0, "msctls_statusbar32", "", WS_VISIBLE Or WS_CHILD, 0, 0, 1, 1, FrmMain.hWnd, ByVal 0&, ByVal 0&, ByVal 0&)
End Sub

'
' This window unloads the status bar window
'
Public Sub DeInitStatusBar()
DestroyWindow StatusBarHandle
End Sub
'
' This sub draws the text on the status bar
'
Public Sub SetStatusBarText(text As String)
On Error Resume Next
SetWindowText StatusBarHandle, text
End Sub

'
' The next 2 subs get and set the position of the video window
'

