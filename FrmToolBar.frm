VERSION 5.00
Begin VB.Form FrmToolBar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   3255
   ClientTop       =   3450
   ClientWidth     =   13305
   Icon            =   "FrmToolBar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   13305
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmOne 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.Timer TimerShow 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   11880
         Top             =   120
      End
      Begin VB.Timer TimerHide 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   11400
         Top             =   120
      End
      Begin VB.Timer TimerProgress 
         Interval        =   200
         Left            =   10920
         Top             =   120
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "<< Back"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame FrmProgress 
         Height          =   735
         Left            =   6360
         TabIndex        =   5
         Top             =   0
         Width           =   4455
         Begin VB.PictureBox PicProgress 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   3225
            TabIndex        =   6
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label LblProgress 
            Caption         =   "Progress"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame FrmVolume 
         Height          =   735
         Index           =   0
         Left            =   3120
         TabIndex        =   2
         Top             =   0
         Width           =   3975
         Begin VB.PictureBox PicVolume 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   2985
            TabIndex        =   3
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label LblVolume 
            Caption         =   "Volume"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   1080
         TabIndex        =   1
         Top             =   0
         Width           =   2175
         Begin VB.CommandButton CmdRescale 
            Caption         =   "Scale"
            Height          =   375
            Left            =   1080
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmdpause 
            Caption         =   "Pause"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "FrmToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MouseHandle As Long
Dim position As PointApi

Dim count_ As Integer
Const maxcount As Integer = 15
Private WindowTop As Long

Private DesktopRect As RECT
Private Sub CmdBack_Click()
FullScreenOff
Unload Me
End Sub

Private Sub Cmdpause_Click()
If Cmdpause.Caption = TEXT_PAUSE Then
   PauseStream
   Cmdpause.Caption = TEXT_RESUME
Else
   ResumeStream
   Cmdpause.Caption = TEXT_PAUSE
End If
End Sub

Private Sub CmdRescale_Click()
PopupMenu FrmMain.mnu_movieformat, vbPopupMenuLeftAlign, , , FrmMain.mnu_bestfit
End Sub

Private Sub Form_Load()
Dim MyRect As RECT
DesktopRect = GetWinPos(GetDesktopWindow)
MyRect = GetWinPos(Me.hWnd)
MyRect.Left = 0
MyRect.Right = DesktopRect.Right
MyRect.Top = DesktopRect.Bottom - (MyRect.Bottom - MyRect.Top)
MyRect.Bottom = DesktopRect.Bottom
SetWinPos Me.hWnd, MyRect
SetAlwaysOnTop AlwaysOn, Me.hWnd
Me.Show

FrmOne.Width = Me.Width - 1
FrmOne.Left = 1
WindowTop = MyRect.Top
ScaleVideoBest
FrmProgress.Width = Screen.Width - FrmProgress.Left
PicProgress.Width = PicVolume.Width
SetProgressValue Me.PicVolume, GetVolume(0), True
MouseHandle = GetCursor

End Sub

Private Sub PicProgress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveStream (Media_Information.TotalFrames) * (X / (PicProgress.Width))
SetProgressValue Me.PicProgress, GetPercent, True
End Sub

Private Sub Picvolume_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Vol As Long
Vol = Round(X / (PicVolume.Width / 100))
If Vol > 97 Then
Vol = 100
End If

SetVolume Vol, 0
SetProgressValue Me.PicVolume, Vol, True
End Sub


Private Sub TimerProgress_Timer()
Dim Posit As PointApi
Dim ret
ret = GetCursorPos(Posit)
SetProgressValue Me.PicProgress, GetPercent, True
If Posit.X <> position.X Or Posit.Y <> position.Y Then
   If count_ = maxcount Then
      ShowMouse
      TimerHide.Enabled = False
      TimerShow.Enabled = True
   End If
   count_ = 0
   position.X = Posit.X
   position.Y = Posit.Y
Else
If count_ < maxcount Then
   count_ = count_ + 1
End If

If count_ = maxcount Then
   HideMouse
   TimerShow.Enabled = False
   TimerHide.Enabled = True
End If
End If
End Sub
Public Sub StopTimerHide()
TimerHide.Enabled = False
End Sub
Private Sub StopTimerShow()
TimerShow.Enabled = False
End Sub
Private Sub TimerHide_Timer()
Dim MyRect As RECT
MyRect = GetWinPos(Me.hWnd)
If MyRect.Top > DesktopRect.Bottom Then
StopTimerHide
Exit Sub
Else
MyRect.Top = MyRect.Top + 3
MyRect.Bottom = MyRect.Bottom + 3
SetWinPos Me.hWnd, MyRect
SetAlwaysOnTop AlwaysOn, Me.hWnd
End If
End Sub

Private Sub TimerShow_Timer()
Dim MyRect As RECT
MyRect = GetWinPos(Me.hWnd)
If MyRect.Top <= WindowTop Then
StopTimerShow
Exit Sub
Else
MyRect.Top = MyRect.Top - 3
MyRect.Bottom = MyRect.Bottom - 3
SetWinPos Me.hWnd, MyRect
SetAlwaysOnTop AlwaysOn, Me.hWnd
End If
End Sub

Public Sub HideMouse()
SetCursor 0&
End Sub

Public Sub ShowMouse()
SetCursor MouseHandle
End Sub

Public Sub FullScreenOff()
Dim Taskbar_Handle As Long
If Not (IsFullscreen) Then Exit Sub
PauseStream 'Pause the stream so we won't miss anything
' Next we search for the taskbar and get its handle
Taskbar_Handle = FindWindow(TaskBar_Name, vbNullString)
' We'll use the handle to turn back on Always On Top
SetAlwaysOnTop AlwaysOn, Taskbar_Handle
FrmView.Hide 'Hide the form while we're resizing it
'Set back it's caption
FrmView.Caption = FRMVIEW_CAPTION & " - " & ExtractName(FileName) & " " & _
                           "(" & LTrim$(Str$(Media_Information.Width)) & _
                           "x" & LTrim$(Str$(Media_Information.Height)) & _
                           " - " & Str$(Media_Information.FramesPerSec) & _
                           "fps )"
SetWinPos FrmView.hWnd, WindowPos 'Resize to the original state
FrmView.ResizeMe 'Resize the video stream
IsFullscreen = False ' Set our flag
FrmView.Visible = True
If FrmMain.mnu_aotvideo.Checked = False Then
   SetAlwaysOnTop AlwaysOff, FrmView.hWnd
End If
If PlayListVisible = True Then
frmPlayList.Visible = True
FrmMain.mnu_Playlist.Checked = True
End If

FrmMain.Visible = True
ResumeStream 'Resume Playback
'That's it, we're back from FullScreen

End Sub



