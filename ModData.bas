Attribute VB_Name = "ModData"
Option Explicit
'
'  This module mantains variables and constants that should
' be available to any part of the project
'  It's like a resource.
'
'
'
'

Public Type typeStreamStatus
Loaded As Boolean
Playing As Boolean
Paused As Boolean
Fullscreen As Boolean
End Type

Public StreamStatus As typeStreamStatus

Public Const ERROR_TITLE = "MCI Error"
Public Const TEXT_PAUSE = "Pause"
Public Const TEXT_RESUME = "Resume"
Public Const TEXT_PLAY = "Play"
Public Const TEXT_STOP = "Stop"
Public Const TEXT_TITLEBAR = "Media Player"
Public Const FRMVIEW_CAPTION = "Video Window"
Public Const LOAD_SRT = "SRT subtitles require that you have an AVI movie loaded!"
Public Const ERROR_SUB = "Error ocurred while reading subtitles!"
Public Const TXTERROR = "Error"
Public Const BAD_SUBTITLE = "Unrecognized subtitle format!"
Public Const FRMPLAYBACK_CAPTION = "Playback Controls"
Public Const NO_FILE_LOADED = " - File not loaded. * "
Public Const ABOUT_TEXT = " ©2002 Marius Hudea" & vbCrLf & vbCrLf & "Comments, critics & sugestions:" & vbCrLf & vbCrLf & _
                          "Marius_Hudea@Personal.ro" & vbCrLf & vbCrLf & "Made In Romania :-) "

Public WindowPos As RECT

Public VolumePercent As Long
Public CurrentStreamFrame As Long
Public PlayListVisible As Boolean
'
'  This sub resets the status of the stream by setting the flags
' to the default values.
'
Public Sub ResetStreamStatus()
With StreamStatus
.Loaded = False
.Paused = False
.Playing = False
.Fullscreen = False
End With
End Sub

