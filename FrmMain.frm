VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Media Player"
   ClientHeight    =   1275
   ClientLeft      =   6315
   ClientTop       =   6555
   ClientWidth     =   5820
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   5820
   Begin VB.CommandButton CmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "<"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton CmdOpen 
      Appearance      =   0  'Flat
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Open a media file"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CmdPlayStop 
      Caption         =   "Play"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Play or stop the media file"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CmdPauseResume 
      Caption         =   "Pause"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      ToolTipText     =   "Pause or resume the media file"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      ToolTipText     =   "Close the media file"
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox PicProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawStyle       =   1  'Dash
      DrawWidth       =   32
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   5745
      TabIndex        =   1
      Top             =   720
      Width           =   5775
   End
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   4680
      Top             =   840
   End
   Begin VB.Frame Frmplayback 
      Caption         =   " Playback Controls "
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Timer TimerPos 
      Interval        =   75
      Left            =   4200
      Top             =   840
   End
   Begin VB.Frame FrmVol 
      Caption         =   "Volume Controls"
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   0
      Width           =   1575
      Begin VB.HScrollBar scrVolume 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         Max             =   100
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_open 
         Caption         =   "&Open File"
      End
      Begin VB.Menu mnu_subs 
         Caption         =   "&Subtitles"
         Begin VB.Menu mnu_opensub 
            Caption         =   "Load &Subtitle..."
         End
         Begin VB.Menu mnu_nosub 
            Caption         =   "Clear Subtitle"
         End
      End
      Begin VB.Menu mnu_sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_close 
         Caption         =   "&Close"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnu_view 
      Caption         =   "&View"
      Begin VB.Menu mnu_aot 
         Caption         =   "&Always On Top"
         Begin VB.Menu mnu_aotvideo 
            Caption         =   "&Video Window"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnu_sep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_aotthis 
            Caption         =   "&This Window"
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu mnu_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_movieformat 
         Caption         =   "&Movie Format"
         Begin VB.Menu mnu_bestfit 
            Caption         =   "Best Fit [ Default ]"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnu_sep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_fitwindow 
            Caption         =   "Fit Window"
         End
      End
      Begin VB.Menu mnu_moviesize 
         Caption         =   "Movie &Size"
         Begin VB.Menu mnu_Show100 
            Caption         =   "100% (Default)"
         End
         Begin VB.Menu mnu_sep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_Show50 
            Caption         =   "50% ( Half )"
         End
         Begin VB.Menu mnu_Show200 
            Caption         =   "200% (Double)"
         End
      End
      Begin VB.Menu mnu_sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Fullscreen 
         Caption         =   "&Fullscreen"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu_sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Playlist 
         Caption         =   "&Playlist Editor"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnu_options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_autoplay 
         Caption         =   "Auto &Play"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_about 
         Caption         =   "&About"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub CmdNext_Click()
Dim s As String
s = GetNextFile
If s <> "" Then OpenMediaFile s
End Sub
Private Sub CmdPrevious_Click()
Dim s As String
s = GetPreviousFile
If s <> "" Then OpenMediaFile s

End Sub
Private Sub CmdOpen_Click()
Open_Click
End Sub

Private Sub Open_Click()
On Error Resume Next
Dim temp As String
temp = OpenFile
temp = StripNulls(temp)
If temp = "" Then Exit Sub
OpenMediaFile temp
InitPlayList
AddFile temp
End Sub

Private Sub CmdClose_Click()
Close_Click
End Sub
Private Sub Close_Click()
If IsLoaded Then
If Not (CloseStream) Then
   MsgBox GetLastErrorString, vbOKOnly + vbCritical, ERROR_TITLE
End If
If IsMovie(FileName) Then
   FrmView.Hide
End If
ResetButtons
End If
FrmMain.SetFocus
ClearProgressBar PicProgress
Me.Caption = TEXT_TITLEBAR & " - File not loaded. * "
End Sub
Private Sub cmdPauseResume_Click()
PauseResume_Click
End Sub
Private Sub PauseResume_Click()
Dim Button As String
Button = CmdPauseResume.Caption
If IsLoaded And IsPlaying Then
   If Button = TEXT_PAUSE Then
      StreamPause
      If IsPaused Then
         CmdPauseResume.Caption = TEXT_RESUME
      End If
   Else
      StreamResume
      If Not (IsPaused) Then
         CmdPauseResume.Caption = TEXT_PAUSE
      End If
   End If
End If

End Sub
Private Sub CmdPlayStop_Click()
Play_Click
End Sub

Private Sub Play_Click()

Dim Button As String
Button = CmdPlayStop.Caption

If IsLoaded Then
   If Button = TEXT_PLAY Then
      StreamPlay
      If IsPlaying Then
         CmdPlayStop.Caption = TEXT_STOP
         scrVolume.Value = VolumePercent
      End If
      Else
      StreamStop
      If Not (IsPlaying) Then
         CmdPlayStop.Caption = TEXT_PLAY
         IsFullscreen = False
      End If
      ResetButtons
       
   End If
End If
End Sub

Private Sub StreamResume()
If IsLoaded Then
   If IsPlaying Then
      If IsPaused Then
         If Not (ResumeStream) Then
            MsgBox GetLastErrorString, vbOKOnly + vbCritical, TXTERROR
         End If
      End If
   End If
End If
End Sub

Private Sub StreamPause()
If IsLoaded Then
   If IsPlaying Then
      If Not (IsPaused) Then
         If Not (PauseStream) Then
            MsgBox GetLastErrorString, vbOKOnly + vbCritical, TXTERROR
         End If
      End If
    End If
End If
End Sub


Private Sub StreamPlay()
If IsLoaded Then
   If IsPlaying Then
      If IsPaused Then
         ResumeStream
      End If
      StopStream
   End If
   If Not (PlayStream) Then
      MsgBox GetLastErrorString, vbOKOnly + vbCritical, ERROR_TITLE
   Else
   SetVolume VolumePercent, 0
   End If
End If
End Sub

Private Sub StreamStop()
If IsLoaded Then
   If IsPlaying Then
      If IsPaused Then ResumeStream
      If Not (StopStream) Then MsgBox GetLastErrorString, vbOKCancel + vbCritical, ERROR_TITLE
   End If
End If
End Sub


Private Sub Form_Load()

InitMCI                             ' Init the multimedia system
InitStatusBar                       ' Load the status bar
InitLargeArray                      ' Load the subtitle support
ResetStreamStatus                   ' Set the stream status to default
Load FrmView                        ' Load the video window
WindowPos = GetWinPos(FrmView.hWnd) ' Get the video window position
Load FrmSub                         ' Load the subtitle window
Load frmPlayList                    ' Load the playlist
FrmSub.Visible = False              ' Hide the subtitles
FrmView.Visible = False             ' and the video window
scrVolume.Value = GetVolume(0)      ' Get the volume
VolumePercent = scrVolume.Value     ' Keep the volume value
IsFullscreen = False                ' We are not in Full Screen
Call mnu_bestfit_Click              ' Check the best fit menu option
FrmMain.mnu_Show100.Checked = True  ' Check the 100% default zoom
'
'   Next there are some adjustments of the controls located in the
' form.
'
PicProgress.Top = PicProgress.Top - 5 * Screen.TwipsPerPixelY
Me.Caption = TEXT_TITLEBAR & NO_FILE_LOADED
CmdPlayStop.Height = CmdPlayStop.Height - 4 * Screen.TwipsPerPixelY
CmdPauseResume.Height = CmdPauseResume.Height - 4 * Screen.TwipsPerPixelY
CmdOpen.Height = CmdOpen.Height - 4 * Screen.TwipsPerPixelY
CmdClose.Height = CmdClose.Height - 4 * Screen.TwipsPerPixelY
CmdNext.Height = CmdNext.Height - 4 * Screen.TwipsPerPixelY
CmdPrevious.Height = CmdPrevious.Height - 4 * Screen.TwipsPerPixelY
'
'  Now, the main window is placed on the bottom-right side of the
' screen.We have to be careful not to place our form on top of the
' taskbar because if it is always on top a part of our main form would
' not be visible.
'
PlaceWindow Me.hWnd
End Sub

Private Sub Form_Terminate()
mnu_close_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
If IsLoaded Then
   If IsPlaying Then
      If IsPaused Then
         ResumeStream
      End If
   StopStream
   End If
   If Not (CloseStream) Then
   MsgBox GetLastErrorString, vbOKOnly + vbCritical, ERROR_TITLE
   Cancel = 1
   GoTo EndOfProc
   End If
End If

DeInitStatusBar
FrmView.CloseVideoWindow = True
Unload FrmView
Unload FrmToolBar
Unload FrmSub
ClosePlayList
Unload Me
EndOfProc:
End Sub

Private Sub mnu_close_Click()
If IsLoaded Then
   If IsPlaying Then
      If IsPaused Then ResumeStream
      StopStream
   End If
   If Not (CloseStream) Then
      MsgBox GetLastErrorString, vbOKOnly, ERROR_TITLE
      GoTo EndOfProc
   End If
End If
DeInitStatusBar
FrmView.CloseVideoWindow = True
Unload FrmView
Unload FrmSub
ClosePlayList
Unload Me
EndOfProc:
End Sub

Private Sub mnu_about_Click()
MsgBox TEXT_TITLEBAR & Space$(1) & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & ABOUT_TEXT, vbOKOnly + vbInformation, "About" & Space(1) & TEXT_TITLEBAR
End Sub

Private Sub mnu_aotthis_Click()
If mnu_aotthis.Checked Then
   SetAlwaysOnTop AlwaysOff, FrmMain.hWnd
   mnu_aotthis.Checked = False
   Else
   SetAlwaysOnTop AlwaysOn, FrmMain.hWnd
   mnu_aotthis.Checked = True
End If

End Sub

Private Sub mnu_aotvideo_Click()
If FrmView.Visible = False Then Exit Sub
If IsLoaded Then
   If mnu_aotvideo.Checked Then
      SetAlwaysOnTop AlwaysOff, FrmView.hWnd
      mnu_aotvideo.Checked = False
   Else
      SetAlwaysOnTop AlwaysOn, FrmView.hWnd
      mnu_aotvideo.Checked = True
   End If
End If
End Sub

Private Sub mnu_autoplay_Click()
mnu_autoplay.Checked = Not (mnu_autoplay.Checked)
End Sub


Private Sub mnu_Fullscreen_Click()
If IsLoaded And IsPlaying And Not (IsPaused) And IsMovie(FileName) Then FullScreenOn
End Sub

Private Sub mnu_bestfit_Click()
If mnu_bestfit.Checked = False Then
   mnu_fitwindow.Checked = False
   mnu_bestfit.Checked = True
End If
ScaleVideoBest
End Sub

Private Sub mnu_fitwindow_Click()
If mnu_fitwindow.Checked = False Then
   mnu_fitwindow.Checked = True
   mnu_bestfit.Checked = False
End If
ScaleVideoMax
End Sub

Private Sub mnu_nosub_Click()
InitLargeArray
End Sub

Private Sub mnu_open_Click()
Call CmdOpen_Click
End Sub

Private Sub mnu_opensub_Click()
Dim s As String
Dim SubtitleType As String
s = StripNulls(OpenSubtitle)
If s <> "" Then
SubtitleType = DetectSubtitleType(s)
If SubtitleType = "microDVD" Then
  InitLargeArray
  LoadSUB (s)
  Exit Sub
End If
If SubtitleType = "SubRIP" Then
  If Not (IsLoaded) Or Not (IsMovie(FileName)) Or LCase$(ExtractExtension(FileName)) <> "avi" Then
     MsgBox LOAD_SRT, vbOKOnly + vbCritical, TXTERROR
     Exit Sub
  End If
  InitLargeArray
  LoadSRT (s)
Exit Sub
End If
MsgBox BAD_SUBTITLE, vbOKOnly + vbExclamation, TXTERROR
End If
End Sub

Private Sub mnu_Playlist_Click()
If mnu_Playlist.Checked = True Then
   frmPlayList.Visible = False
   mnu_Playlist.Checked = False
   Exit Sub
Else
   frmPlayList.Visible = True
   mnu_Playlist.Checked = True
   Exit Sub
End If
End Sub

Private Sub mnu_Show100_Click()
Dim MyRect As RECT
Dim Width As Long
Dim Height As Long
If Not (IsLoaded) Or Not (IsMovie(FileName)) Then Exit Sub
If IsFullscreen Then Exit Sub
MyRect = GetWinPos(FrmView.hWnd)
Width = Media_Information.Width
Height = Media_Information.Height
If (Width \ 2) * 2 < Width Then Width = ((Width + 1) \ 2) * 2
If (Height \ 2) * 2 < Height Then Width = ((Height + 1) \ 2) * 2
MyRect.Right = MyRect.Left + Width
MyRect.Bottom = MyRect.Top + Height
SetWinPos FrmView.hWnd, MyRect
FrmView.ResizeMe
mnu_Show50.Checked = False
mnu_Show100.Checked = True
mnu_Show200.Checked = False
End Sub

Private Sub mnu_Show200_Click()
Dim MyRect As RECT
Dim Width As Long
Dim Height As Long

If Not (IsLoaded) Or Not (IsMovie(FileName)) Then Exit Sub
If IsFullscreen Then Exit Sub
MyRect = GetWinPos(FrmView.hWnd)
Width = Media_Information.Width * 2
Height = Media_Information.Height * 2
If (Width \ 2) * 2 < Width Then Width = ((Width + 1) \ 2) * 2
If (Height \ 2) * 2 < Height Then Width = ((Height + 1) \ 2) * 2
MyRect.Right = MyRect.Left + Width
MyRect.Bottom = MyRect.Top + Height
SetWinPos FrmView.hWnd, MyRect
FrmView.ResizeMe
mnu_Show50.Checked = False
mnu_Show100.Checked = False
mnu_Show200.Checked = True
End Sub

Private Sub mnu_Show50_Click()
Dim MyRect As RECT
Dim Width As Long
Dim Height As Long

If Not (IsLoaded) Or Not (IsMovie(FileName)) Then Exit Sub
If IsFullscreen Then Exit Sub
MyRect = GetWinPos(FrmView.hWnd)
Width = Media_Information.Width / 2
Height = Media_Information.Height / 2
If (Width \ 2) * 2 < Width Then Width = ((Width + 1) \ 2) * 2
If (Height \ 2) * 2 < Height Then Width = ((Height + 1) \ 2) * 2
MyRect.Right = MyRect.Left + Width
MyRect.Bottom = MyRect.Top + Height
SetWinPos FrmView.hWnd, MyRect
FrmView.ResizeMe
mnu_Show50.Checked = True
mnu_Show100.Checked = False
mnu_Show200.Checked = False
End Sub

Private Sub PicProgress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim OnePicPercent As Long
Dim CurrentPercent As Long
Dim Frames As Long
Dim OneFramePercent As Long
If IsPlaying Then
CmdPauseResume.Caption = TEXT_PAUSE
OnePicPercent = PicProgress.Width / (100)
CurrentPercent = X / OnePicPercent
Frames = Media_Information.TotalFrames
OneFramePercent = Frames / 100
MoveStream Round(CurrentPercent * OneFramePercent, 0)
End If
End Sub

Private Sub scrVolume_Change()
     SetVolume scrVolume.Value
     VolumePercent = scrVolume.Value
End Sub

Private Sub scrVolume_Scroll()
     SetVolume scrVolume.Value
     VolumePercent = scrVolume.Value
End Sub

Private Sub Timer_Timer()
On Error Resume Next
Dim nfile As String
Dim TotalTime As Long
Dim TotalFrames As Long
Dim CurrentFrame As Long
Dim CurrentTime As Long
Dim NewFile As String
Dim s As String
s = Me.Caption
Me.Caption = Right$(s, Len(s) - 1) & Left$(s, 1)
If IsLoaded Then
   s = ExtractName(FileName)
      If Len(s) >= 56 Then
      nfile = "..." & Right$(s, 56)
      Else
      nfile = s
      End If
      SetStatusBarText GetStreamStatus & " | " & nfile
      
Else 'If isLoaded then
SetStatusBarText "No open files."
End If

If IsLoaded And IsStreamAtEnd(CurrentStreamFrame) Then
   ClearProgressBar PicProgress
   If Not (StopStream) Then MsgBox GetLastErrorString, vbOKOnly + vbCritical, ERROR_TITLE
   ResetButtons
   NewFile = GetNextFile
   If NewFile <> "" Then
      If Not (IsMovie(NewFile)) And IsFullscreen Then
         FrmToolBar.FullScreenOff
         Unload FrmToolBar
         OpenMediaFile NewFile
      Else
         FrmToolBar.FullScreenOff
         OpenMediaFile NewFile
         FullScreenOn
      End If  ' If not(isMovie...
    End If   'If newfile <>"" ..
End If  'If isStreamAtEnd..
   
If IsPlaying Then
   SetProgressValue PicProgress, GetPercent(CurrentStreamFrame), True
   If LCase$(ExtractExtension(FileName)) <> "avi" Then
   CurrentTime = CurrentStreamFrame / Media_Information.FramesPerSec * 1000
   Else
   CurrentTime = CurrentStreamFrame
   End If
   Frmplayback.Caption = FRMPLAYBACK_CAPTION & " - " & ConvertTime2String(CurrentTime) & " "
Else
Frmplayback.Caption = FRMPLAYBACK_CAPTION
End If

If Me.WindowState = vbMinimized Then
frmPlayList.Hide
Me.mnu_Playlist.Checked = False
PlayListVisible = True
End If

End Sub

Public Sub ResetButtons()
CmdPlayStop.Caption = "Play"
CmdPauseResume.Caption = "Pause"
End Sub

'
' I believe that this is the only place where the mci si accessed by the program to
' detect the stream position.The stream position is retrieved every 75 ms or each
' time the timer is activated ( it is 75 ms when i write this )
'
Private Sub TimerPos_Timer()
CurrentStreamFrame = GetCurrentStreamPos
End Sub

Public Sub FullScreenOn()
Dim Taskbar_Handle As Long
Dim DesktopRect As RECT
If Not (IsLoaded) Then GoTo EndMe
If Not (IsMovie(FileName)) Then GoTo EndMe
If Not (IsPlaying) Then GoTo EndMe
WindowPos = GetWinPos(FrmView.hWnd) ' Preserve the window's position
PauseStream 'Pause the stream so we won't miss anything
' Next we search for the taskbar and get its handle.
Taskbar_Handle = FindWindow(TaskBar_Name, vbNullString)
' We'll use the handle to turn off Always On Top
SetAlwaysOnTop AlwaysOff, Taskbar_Handle
FrmView.Hide       'Hide the form while we're resizing it
FrmView.Caption = "" 'Hide the caption
'Get the desktop size
DesktopRect = GetWinPos(GetDesktopWindow)
' Make the window the same size as the desktop, the taskbar is hidden
' because it is no longer Always On Top
SetWinPos FrmView.hWnd, DesktopRect
FrmView.ResizeMe 'Resize the stream
FrmView.Show 'Show back the window
PlayListVisible = frmPlayList.Visible
frmPlayList.Hide
FrmMain.mnu_Playlist.Checked = False
FrmMain.Hide
SetAlwaysOnTop AlwaysOn, FrmView.hWnd 'Set it Always On Top
SetAlwaysOnTop AlwaysOn, FrmSub.hWnd
Load FrmToolBar 'Load the toolbar
SetAlwaysOnTop AlwaysOn, FrmToolBar.hWnd
IsFullscreen = True ' Set our flag
ResumeStream 'Resume Playback
'That's it, we're on FullScreen
EndMe:
End Sub

Public Sub OpenMediaFile(TheFileName As String)
Dim text As String
Dim VideoRect As RECT
If StripNulls(TheFileName) = "" Then Exit Sub

If TheFileName <> "" Then
   If IsLoaded Then
      If IsPlaying Then
         If IsPaused Then ResumeStream
         StopStream
      End If
   CloseStream
   End If
 
   PicProgress.Cls
   If OpenStream(FrmView.hWnd, TheFileName) Then
      text = ExtractName(FileName)
      Me.Caption = TEXT_TITLEBAR & " - " & Left$(text, Len(text) - 4) & " * "
      Media_Information = GetMediaInfo
      If IsMovie(FileName) Then
         FrmView.Show
         FrmView.Caption = FRMVIEW_CAPTION & " - " & text & " " & _
                           "(" & LTrim$(Str$(Media_Information.Width)) & "x" & LTrim$(Str$(Media_Information.Height)) & " - " & Str$(Media_Information.FramesPerSec) & "fps )"
         VideoRect = GetWinPos(FrmView.hWnd)
         VideoRect.Right = VideoRect.Left + Media_Information.Width
         VideoRect.Bottom = VideoRect.Top + Media_Information.Height
         SetWinPos FrmView.Height, VideoRect
      Else
         FrmView.Hide
      End If
      ClearProgressBar PicProgress
      ResetButtons
      If Not (IsFullscreen) Then mnu_Show100_Click
      If mnu_autoplay.Checked Then Play_Click
   Else
   MsgBox GetLastErrorString, vbOKOnly + vbCritical, TXTERROR
   End If
End If
If Not (IsFullscreen) Then FrmMain.SetFocus
End Sub
