Attribute VB_Name = "ModMCI"
Option Explicit
'
'   This is the module that does all the decoding and playing.
'   The API functions used by this module are located in
' the modAPI module.
'   The quoted functions below are only the essential functions
' that are required if you want to take this module and use it
' in your project.
'
'Public Declare Function GetTickCount Lib "kernel32" () As Long
'Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
'Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Type Media_Info 'Maintains information about the stream
Width As Long
Height As Long
TotalFrames As Long
TotalTime As Long
Hours As Long
Minutes As Long
Seconds As Long
MilliSeconds As Long
FramesPerSec As Double
End Type

Public Type typeTime
Hours As Long
Minutes As Long
Seconds As Long
MilliSeconds As Long
End Type

Private Type typeExtensions
Extension As String
Device As String
End Type

Public Media_Information As Media_Info ' Media Information

Private StreamAlias As String     ' An unique alias for the stream.
Private StreamDevice As String    ' The device that plays the stream.
Private VideoWindowHandle As Long ' The handle for the video window

Private LastError As Long         'The last error that has ocurred.

'  Constants used to detect the type of media file.
'  It is not a good way to determine the type this way
'  Extensions are detected automatically when this module starts so
' other extensions might appear. I am only using this string constants
' to detect if the file has video. I don't know other method (yet).
'
Public Const Stream_Movies = "*.avi;*.m1v;*.mpg;*.dat;*.mpe;*.mv1;*.mov;*.mpm;*.mpv;*.enc;*.vob;*.m2v;*.asf;*.asx;*.ivf;*.isx;*.lsx;*.wvx;*.wmv;*.mpeg;*.mpv2"
Public Const Stream_Sounds = "*.mp3;*.mp2;*.mp1;*.m3u;*.mpa;*.wav;*.aif;*.snd;*.wax;*.wma;*.aiff;*.aifc"
Public Const Stream_Midi = "*.mid;*.rmi;*.midi"
Public Const Stream_AudioCD = "*.cda"


Dim FrameFrom As Long          ' From where to start playing
Dim FrameTo As Long            ' Where to stop playing

Public IsLoaded As Boolean     ' Stream is loaded
Public IsPlaying As Boolean    ' Stream is playing
Public IsPaused As Boolean     ' Stream is paused
Public IsFullscreen As Boolean ' Stream is in FullScreen mode

Public StreamName As String 'Filename without spaces (MS-DOS 8.3)
Public FileName As String   ' Original FileName

Public ExtensionList(0 To 255) As typeExtensions 'The extensions and their type
Public ExtensionNr As Long  'Total number of extensions

'
'  This sub loads the extensions that MCI recognizes by reading
' them from the win.ini file located in your Windows folder.
'  Type at the Run.. prompt msconfig or open it with notepad.
'  You can use ini file handling functions but in this case
' I believe it would only complicate things.
'  The extension list is provided in the win.ini file only for
' compatibility with old programs.
'  If the list isn't there or if the win.ini file is missing a
' default extension list is loaded.
'
'  Note: You could obtain another list of extension by reading
' the Media Player's list of extensions from the registry located
' at :
'
' HKEY_LOCAL_MACHINE\Software\Microsoft\MediaPlayer\Player\Extensions
'
Public Sub LoadExtensions()
On Error Resume Next
Dim s As String
Dim EqualPos As Long
ExtensionNr = -1

Open Environ$("WINDIR") & "\win.ini" For Input Access Read As 1
'
'
'
If Err.Number <> 0 Then   ' Error opening the file.
LoadDefaultExtensionList  ' Load default list
Exit Sub                  ' Exit
End If

While Not EOF(1)          ' Search for the [mci extensions] tab.
Line Input #1, s
If LCase$(s) = "[mci extensions]" Then GoTo out
Wend
out:                      ' Tab not found or end of file.
If EOF(1) Then
LoadDefaultExtensionList  ' If end of file load the default list,
Close #1                  ' close the file and exit
Exit Sub
End If
While s <> "" And Not (EOF(1)) ' Get the extension list
Line Input #1, s
If s <> "" Then
EqualPos = 0
EqualPos = InStr(1, s, "=", vbTextCompare)
If EqualPos <> 0 Then
   AddExtension Trim$(Mid$(s, 1, EqualPos - 1)), Trim$(Mid$(s, EqualPos + 1, Len(s) - EqualPos))
End If
End If
Wend
Close #1
End Sub
'
'  This sub adds an extension to the extension list.
'
Public Sub AddExtension(Ext As String, Dev As String)
ExtensionNr = ExtensionNr + 1
With ExtensionList(ExtensionNr)
.Extension = Ext
If LCase$(Ext) = "avi" Then
.Device = "MPEGVideo"
Else
.Device = Dev
End If
End With
End Sub
'
'   This sub is called if there was an error obtaining the list of mci extensions located
' in the Windows folder.
'   These are common file extensions that mci recognizes  on my computer.
' I have the standard Windows98 Second Edition and Windows Media Player 6.4 (with DirectX
' Media).
'   I am adding the "avi" extension with the "MPEGVideo" device instead of the "avivideo"
' device because it works better.
'
Public Sub LoadDefaultExtensionList()
Const Sequencer = "Sequencer"
Const MPEGVideo = "MPEGVideo"
Const MPEGVideo2 = "MPEGVideo2"
ExtensionNr = -1
AddExtension "cda", "CDAudio"
AddExtension "wav", "waveaudio"
AddExtension "mid", Sequencer
AddExtension "midi", Sequencer
AddExtension "rmi", Sequencer
AddExtension "dat", MPEGVideo
AddExtension "avi", MPEGVideo
AddExtension "aif", MPEGVideo
AddExtension "aifc", MPEGVideo
AddExtension "aiff", MPEGVideo
AddExtension "au", MPEGVideo
AddExtension "m1v", MPEGVideo
AddExtension "m3u", MPEGVideo
AddExtension "mov", MPEGVideo
AddExtension "mp2", MPEGVideo
AddExtension "mp3", MPEGVideo
AddExtension "mpa", MPEGVideo
AddExtension "mpe", MPEGVideo
AddExtension "mpeg", MPEGVideo
AddExtension "mpg", MPEGVideo
AddExtension "mpv2", MPEGVideo
AddExtension "qt", MPEGVideo
AddExtension "snd", MPEGVideo
AddExtension "mp2v", MPEGVideo
AddExtension "asf", MPEGVideo2
AddExtension "asx", MPEGVideo2
AddExtension "ivf", MPEGVideo2
AddExtension "lsf", MPEGVideo2
AddExtension "lsx", MPEGVideo2
AddExtension "wax", MPEGVideo2
AddExtension "wvx", MPEGVideo2
AddExtension "wm", MPEGVideo2
AddExtension "wma", MPEGVideo2
AddExtension "wmv", MPEGVideo2
End Sub
'
' This sub starts the whole thing by setting the flags to their
'default values and by loading the default extensions.
Public Sub InitMCI()
IsLoaded = False
IsPlaying = False
IsPaused = False
IsFullscreen = False
With Media_Information
.Width = 0
.Height = 0
End With
LoadExtensions ' Load the extensions
AddExtension "divx", "MPEGVideo" 'Load 2 new extensions
AddExtension "mp4", "MPEGVideo"
'  Can't add ogg vorbis files yet because i don't think there
' is a codec for it.
End Sub
'
'  This function retrieves the device that will be used to decode the
' selected media file.
'
Private Function DetectStreamDevice(StreamName As String) As String
Dim txt As String
Dim i As Long
Dim s As String
If StreamName = "" Then
   DetectStreamDevice = ""
   Exit Function
End If
'Get the extension
txt = ""
txt = LCase$(Right$(StreamName, 4))
If Left$(txt, 1) = "." Then txt = Right$(txt, 3)
For i = 0 To ExtensionNr
 If ExtensionList(i).Extension = txt Then
    DetectStreamDevice = ExtensionList(i).Device
    Exit Function
 End If
Next i
'
' If that extension is unknown, we assume it is "MPEGVideo"
'
DetectStreamDevice = "MPEGVideo"
End Function
'
'  This function determines if the stream is a video file.
'  Please tell me if you know a better way to detect if a file
' has video.
'
Public Function IsMovie(TheFile As String) As Boolean
Dim Ext As String
If TheFile = "" Then
   IsMovie = False
End If
'  Get the file extension
Ext = LCase(Right$(TheFile, 4))
' the tmp extension is detected as movie
If Ext = ".tmp" Then IsMovie = False
' QuickTime extension
If Right$(Ext, 2) = "qt" Then IsMovie = True
'  Search it in the list of video file extensions
If InStr(1, Stream_Movies, Ext) <> 0 Then
   IsMovie = True
Else
   IsMovie = False
End If
End Function

'
'  This function determines if the stream is an audio file.
'
Public Function IsAudio(TheFile As String) As Boolean
Dim Ext As String
If TheFile = "" Then
   IsAudio = False
End If
Ext = LCase(Right$(TheFile, 4))
' AU extension
If Right$(Ext, 2) = "au" Then IsAudio = True
If InStr(1, Stream_Sounds, Ext) <> 0 Then
   IsAudio = True
Else
   IsAudio = False
End If
End Function

'
'  This function determines if the stream is a midi file.
'
Public Function IsMidi(TheFile As String) As Boolean
Dim Ext As String
If TheFile = "" Then
   IsMidi = False
End If
Ext = LCase(Right$(TheFile, 4))
If InStr(1, Stream_Midi, Ext) <> 0 Then
   IsMidi = True
Else
   IsMidi = False
End If
End Function
'
'  This function can be used to trim null characters from the end
' of a string. Null characters often appear when working with
' API functions.
'
Public Function StripNulls(ByVal FileWithNulls As String) As String
Dim NewString As String
Dim i As Long
Dim NullPos As Integer
NullPos = InStr(1, FileWithNulls, vbNullChar, 0)
If NullPos <> 0 Then
   StripNulls = Left(FileWithNulls, NullPos - 1)
Else
   StripNulls = FileWithNulls
End If
End Function

Public Function StripNullsTrim(OriginalStr As String) As String
Dim a As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        a = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
StripNullsTrim = Trim$(a)
End Function

'
'  The following function generates an unique alias each time
' a new file is beeing loaded. This prevents the use of the
' same alias, if two instances of your application are running
' at the same time.
'
Private Function GenerateAlias() As String
Dim Generated As String
'
'  Get the number of milliseconds that have elapsed since Windows
'  started and convert it to string.
'
Generated = Str$(GetTickCount)
Generated = Mid$(Generated, Len(Generated) - 5, 5) 'Get only the last 5 charachers
GenerateAlias = "Stream" & Generated & "maertS" ' Return our unique generated alias
End Function

'
'  This function opens a specified file. This function must be used
' before you play the stream.
'  Important!
'  You MUST close the file when you close the program using the
' CloseStream function, otherwise the stream will remain playing in
' memory and you will only be able to close it if you know the Alias
'
Public Function OpenStream(VideoHandle As Long, FileTitle As String) As Boolean
Dim SmallFileName As String
Dim StringCommand As String * 255
Dim TempNr As Long
Dim tmp As String * 255
'
'  If a stream is loaded then resume it and stoop it, if needed,
' then close it.
'
If IsLoaded Then
   If IsPlaying Then
      If IsPaused Then ResumeStream
      StopStream
   End If
   CloseStream
End If
'
'  If the stream is still loaded there was an error closing it.
'
If IsLoaded Then OpenStream = False
'
'  If no stream specified, function will exit, no error ocurred.
'
If FileTitle = "" Then OpenStream = True
'
'  Get the filename in MS-DOS format.We are not allowed to have
' spaces in the command we send to the MCI system. We also have
' to keep the whole command in 255 characters.
'
TempNr = GetShortPathName(FileTitle, tmp, 255)
'
'  Get only what we need form the string returned by the function
'
SmallFileName = Mid$(tmp, 1, TempNr)
'
'  Set the video handle. This handle is required only for video
' files but we have to specify it each time we open a file. It
' tells MCI where to draw the images.It can be a handle from
' windows, menus, pictureboxes and even command buttons.
'
VideoWindowHandle = VideoHandle
'
'  Detect the proper stream device that will handle our stream.
'
StreamDevice = DetectStreamDevice(FileTitle)
'
'  In this project we are using only MPEGVideo because it can handle
' any file. Remove the following line to make OpenStream select the
' proper device.
StreamDevice = "MPEGVideo"
'
'  If there are any null characters (from the open file function),
' delete them. Store the names for later use.
'
FileName = StripNulls(FileTitle)
StreamName = SmallFileName
'
'  We now generate an alias. This alias is used later when playing,
' pausing, resuming , stoping & closing the stream.
'
StreamAlias = GenerateAlias

'
'  We have all the necessary data.We can now build a string command
' that will be sent to the MCI system.
StringCommand = "open " & StreamName & " type " & StreamDevice & " Alias " & StreamAlias & " parent " & VideoWindowHandle & " style child"
LastError = mciSendString(StringCommand, 0&, 0&, 0&)
'
'  All is OK.
'
If LastError = 0 Then
'
' Set our flags to Loaded, NOT Playing, NOT Paused
'
SetFlags True, False, False
OpenStream = True
Else
OpenStream = False
End If
End Function
'
' ----------------------------------------------------------
'
'  The rest of the function are easy to understand.I won't
' comment them.
'
' ----------------------------------------------------------
'
Public Function PlayStream(Optional Startframe As Long = 1, Optional EndFrame As Long = -1) As Boolean
Dim StringCommand As String * 255
If Not (IsLoaded) Then PlayStream = True
If IsPaused Then ResumeStream
If IsPlaying Then StopStream

FrameFrom = Startframe
If EndFrame = -1 Then
   FrameTo = GetTotalFrames
Else
   FrameTo = EndFrame
End If
If FrameTo < 1 Then
   PlayStream = True
   Exit Function
End If
               
StringCommand = "play " & StreamAlias & " from " & FrameFrom & " to " & FrameTo
LastError = mciSendString(StringCommand, 0&, 0&, 0&)
If LastError = 0 Then
SetFlags True, True, False
PlayStream = True
Else
PlayStream = False
End If
End Function

Public Function StopStream() As Boolean
Dim StringCommand As String * 255

If Not (IsLoaded) Then StopStream = True
If Not (IsPlaying) Then StopStream = True
If IsPaused Then ResumeStream

StringCommand = "Stop " & StreamAlias
LastError = mciSendString(StringCommand, 0&, 0&, 0&)
If LastError = 0 Then
SetFlags True, False, False
StopStream = True
Else
StopStream = False
End If
End Function

Public Function CloseStream() As Boolean
Dim StringCommand As String * 255

If Not (IsLoaded) Then CloseStream = True
If IsPlaying Then
   If IsPaused Then ResumeStream
   StopStream
End If

StringCommand = "Close " & StreamAlias
LastError = mciSendString(StringCommand, 0&, 0&, 0&)
If LastError = 0 Then
SetFlags False, False, False
CloseStream = True
Else
CloseStream = False
End If
End Function

Public Function PauseStream() As Boolean
Dim StringCommand As String * 255

If Not (IsLoaded) Then PauseStream = True
If Not (IsPlaying) Then PauseStream = True
If IsPaused Then PauseStream = True

StringCommand = "Pause " & StreamAlias
LastError = mciSendString(StringCommand, 0&, 0&, 0&)
If LastError = 0 Then
SetFlags True, True, True
PauseStream = True
Else
PauseStream = False
End If
End Function

Public Function ResumeStream() As Boolean
Dim StringCommand As String * 255

If Not (IsLoaded) Then ResumeStream = True
If Not (IsPlaying) Then ResumeStream = True
If Not (IsPaused) Then ResumeStream = True

StringCommand = "Resume " & StreamAlias
LastError = mciSendString(StringCommand, 0&, 0&, 0&)
If LastError = 0 Then
SetFlags True, True, False
ResumeStream = True
Else
ResumeStream = False
End If
End Function

Public Function GetTotalFrames() As Long
Dim StringCommand As String
Dim TotalFrames As String * 255

If Not (IsLoaded) Then GetTotalFrames = 0
StringCommand = "set " & StreamAlias & " time format frames" ' Convert time to frames.
LastError = mciSendString(StringCommand, TotalFrames, 255, 0&)
StringCommand = "status " & StreamAlias & " length" ' Get the number of frames the stream has.
LastError = mciSendString(StringCommand, TotalFrames, 255, 0&)
If LastError = 0 Then
GetTotalFrames = CLng(Val(TotalFrames))
Else
GetTotalFrames = 0
End If
End Function

Public Function GetStreamStatus() As String
Dim StringCommand As String
Dim Status As String * 255
Dim ReturnedStatus As String
Dim i As Integer
Dim CharA As String
Dim RChar As String
ReturnedStatus = ""
StringCommand = "status " & StreamAlias & " mode"
LastError = mciSendString(StringCommand, Status, 255, 0&)

If LastError = 0 Then
 GetStreamStatus = "Error."
End If

RChar = Mid$(Status, Len(Status) - 1, 1)
For i = 1 To Len(Status)
    CharA = Mid(Status, i, 1)
    If CharA = RChar Then Exit For
    ReturnedStatus = ReturnedStatus & CharA
Next i
GetStreamStatus = UCase$(Left$(ReturnedStatus, 1)) & Right$(ReturnedStatus, Len(ReturnedStatus) - 1)
End Function

Public Function GetTotalTime() As Long
Dim TotalTime As String * 255
Dim total As String * 255

If Not (IsLoaded) Then GetTotalTime = 0
'
' Convert the time to milliseconds
'
LastError = mciSendString("set " & StreamAlias & " time format ms", 0&, 0&, 0&)
'
' Get the length of the media file in milliseconds
'
LastError = mciSendString("status " & StreamAlias & " length", TotalTime, 255, 0&)
If LastError <> 0 Then  'not success
LastError = mciSendString("set " & StreamAlias & " time format frames", total, 255, 0&) ' convert back to frames
GetTotalTime = 0
Else
LastError = mciSendString("set " & StreamAlias & " time format frames", total, 255, 0&) ' convert back to frames
GetTotalTime = CLng(Val(TotalTime))
End If
End Function

Public Function MoveStream(position As Long) As Boolean
Dim ReturnValue As Long
Dim ret As String * 255
If Not (IsLoaded) Then MoveStream = True
LastError = mciSendString("seek " & StreamAlias & " to " & Str$(position), 0&, 0&, 0&)
If LastError = 0 Then
   LastError = mciSendString("Play " & StreamAlias, 0&, 0&, 0&)
   If LastError = 0 Then
      MoveStream = True
   Else
      MoveStream = False
   End If
Else
MoveStream = False
End If
End Function

Public Function GetCurrentStreamPos() As Long
Dim pos As String * 255
Dim total As String * 255
If Not (IsLoaded) Then GetCurrentStreamPos = 0
LastError = mciSendString("status " & StreamAlias & " position", pos, 255, 0&)
If LastError = 0 Then
GetCurrentStreamPos = CLng(Val(pos))
Else
GetCurrentStreamPos = 0
End If
End Function

'
'  This function places the video image on the specified window.
'  First it determines the rectange of the window and obtains the
' height and the width, then sends the command to the MCI system.
'
Public Function PutStream(Left As Long, Top As Long, Width As Long, Height As Long) As Boolean
On Error Resume Next
Dim rec As RECT
Dim LocalWidth As Long
Dim LocalHeight As Long

If Not (IsLoaded) Then PutStream = True

LocalWidth = Width
LocalHeight = Height

If Width = 0 Or Height = 0 Then
    rec = GetWinPos(VideoWindowHandle)
    LocalWidth = rec.Right - rec.Left
    LocalHeight = rec.Bottom - rec.Top
End If

LastError = mciSendString("put " & StreamAlias & " window at " & Left & " " & Top & " " & LocalWidth & " " & LocalHeight, 0&, 0&, 0&)
If LastError = 0 Then
PutStream = True
Else
PutStream = False
End If
End Function
Public Function GetFramesPerSecond() As Long
On Error Resume Next
Dim TotalFrames As Long
Dim TotalTime As Long
TotalTime = GetTotalTime
TotalFrames = GetTotalFrames
TotalTime = TotalTime / 1000
If TotalFrames = 0 Then GetFramesPerSecond = 0
GetFramesPerSecond = CLng(Round((TotalFrames / TotalTime), 3))
End Function

'
' The following two function request the current stream position from the mci.
' This might cause slowdowns on low-end computers while the position is determined.
' If you use these function and you have already determined the current position
' use that value as a parameter. The functions will use that value instead of asking
' for a new value from mci
'
Public Function GetPercent(Optional SpecifiedPosition As Long = 0) As Long
On Error Resume Next
Dim TotalFrames As Long
Dim CurrentFrame As Long
TotalFrames = GetTotalFrames
If SpecifiedPosition = 0 Then
CurrentFrame = GetCurrentStreamPos
Else
CurrentFrame = SpecifiedPosition
End If

GetPercent = CLng(Round(CurrentFrame * 100 / TotalFrames))
End Function

Public Function IsStreamAtEnd(Optional SpecifiedPosition As Long = 0) As Boolean
Dim CurrentPos As Long
If SpecifiedPosition = 0 Then
CurrentPos = GetCurrentStreamPos
Else
CurrentPos = SpecifiedPosition
End If

If Not (IsLoaded) Then IsStreamAtEnd = True

If FrameTo = CurrentPos Or (FrameTo - 1) < CurrentPos Then
IsStreamAtEnd = True
Else
IsStreamAtEnd = False
End If
End Function
'
'  This function sets the volume of the stream.If the channel is
' specified it will modify only that channel. Set to 0 or don't
' specify the channel in order to change both channels at once.
'
Public Function SetVolume(Volume As Long, Optional Channel As Integer = 0) As Boolean

Dim StringCommand As String * 128
Dim VolumeValue As Long

VolumeValue = Volume

If VolumeValue < 0 Then VolumeValue = 0
If VolumeValue > 100 Then VolumeValue = 100

VolumeValue = VolumeValue * 10
If Channel = 1 Then
StringCommand = "setaudio " & StreamAlias & " left volume to " & Str$(VolumeValue)
Else
   If Channel = 2 Then
   StringCommand = "setaudio " & StreamAlias & " right volume to " & Str$(VolumeValue)
   Else
   StringCommand = "setaudio " & StreamAlias & " volume to " & Str$(VolumeValue)
   End If
End If

LastError = mciSendString(StringCommand, 0&, 0&, 0&)

If LastError = 0 Then
SetVolume = True
Else
SetVolume = False
End If
End Function
'
' Guess what this function does ? :~)
'
Public Function GetVolume(Optional Channel As Integer = 0) As Long
Dim StringCommand As String * 128
Dim Volume As String * 128

If Channel = 1 Then StringCommand = "status " & StreamAlias & " left volume"
If Channel = 2 Then StringCommand = "status " & StreamAlias & " right volume"
If Channel <> 1 And Channel <> 2 Then StringCommand = "status " & StreamAlias & " volume"

LastError = mciSendString(StringCommand, Volume, 128, 0&)
If LastError = 0 Then
GetVolume = CLng(Round(Val(Volume) / 10))
Else
GetVolume = 100
End If

End Function
'
'   This function sets the playback speed of the stream and
' the one after this one returns the playback speed.
'   I don't use this in the program because I don't find it
' (very) useful for a media player.
'
Public Function SetRate(RateValue As Long) As Boolean

Dim StringCommand As String * 128
Dim RateV As Long

If Not (IsLoaded) Then SetRate = True

RateV = RateValue
If RateV < 0 Or RateV > 200 Then RateV = 100

RateV = RateV * 10

StringCommand = "set " & StreamAlias & " speed " & RateV
LastError = mciSendString(StringCommand, 0&, 0&, 0&)
If LastError <> 0 Then
   SetRate = False
Else
SetRate = True
End If
End Function

Public Function GetRate() As Long
Dim StringCommand As String * 128
Dim Rate As String * 128
If Not (IsLoaded) Then GetRate = 1

StringCommand = "status " & StreamAlias & " speed"
LastError = mciSendString(StringCommand, Rate, 128, 0&)

If LastError = 0 Then
GetRate = CDbl(Val(Rate) / 10)
Else
GetRate = 1
End If
End Function
'
'  This function gives us a string that explains an error,
' if it has ocurred.
'
Public Function GetLastErrorString() As String
Dim ErrorString As String * 255
'create a buffer
ErrorString = Space$(255)
'retrieve the error string
mciGetErrorString LastError, ErrorString, Len(ErrorString)
'strip off the trailing spaces
GetLastErrorString = Trim$(ErrorString)
End Function

'
'  This function returns information form the stream.
'
Public Function GetMediaInfo() As Media_Info
On Error Resume Next
Dim Media As Media_Info
Dim StringCommand As String * 128
Dim size As String * 128
Dim var() As String
Dim AllTime As typeTime

With Media
.Height = 0
.Width = 0
.TotalFrames = 0
.TotalTime = 0
.FramesPerSec = 25
End With
If Not (IsLoaded) Then GetMediaInfo = Media

If IsMovie(FileName) Then
StringCommand = "where " & StreamAlias & " destination"
LastError = mciSendString(StringCommand, size, 128, 0&)
 If LastError = 0 Then
    var = Split(size, " ", -1)
    With Media
         .Width = CCur(var(2))
         .Height = CCur(var(3))
    End With
 End If
End If
Media.TotalFrames = GetTotalFrames
Media.TotalTime = GetTotalTime
Media.FramesPerSec = GetFramesPerSecond
If LCase$(ExtractExtension(FileName)) = "avi" Then Media.FramesPerSec = Round(AVI_FramesPerSecond(FileName), 3)
AllTime = ConvertTime(Media.TotalTime)

Media.Hours = AllTime.Hours
Media.Minutes = AllTime.Minutes
Media.Seconds = AllTime.Seconds
Media.MilliSeconds = AllTime.MilliSeconds
GetMediaInfo = Media
End Function

Public Sub SetFlags(Loaded As Boolean, Playing As Boolean, Paused As Boolean)
IsLoaded = Loaded
IsPlaying = Playing
IsPaused = Paused
End Sub
'
'  This sub zooms to the original size mantaining the aspect ratio
' of the video stream.
'
Public Sub ScaleVideoBest()
On Error GoTo endofsub
Dim rc As RECT
Dim WindowWidth As Double
Dim windowHeight As Double
Dim NewWidth As Double
Dim NewHeight As Double
Dim raport As Double
GetWindowRect VideoWindowHandle, rc
WindowWidth = (rc.Right - rc.Left)
windowHeight = (rc.Bottom - rc.Top)
raport = Media_Information.Height / Media_Information.Width

NewWidth = WindowWidth
NewHeight = WindowWidth * raport
If NewHeight > windowHeight Then
   NewHeight = windowHeight
   NewWidth = windowHeight / raport
   PutStream (WindowWidth - NewWidth) / 2, 0, CLng(NewWidth), CLng(NewHeight)
Else
PutStream 0, (windowHeight - NewHeight) / 2 - 10, CLng(NewWidth), CLng(NewHeight)
End If
If windowHeight < NewHeight Then
SetWindowPos VideoWindowHandle, 0&, rc.Left, rc.Top, NewWidth, NewHeight, SWP_NOZORDER Or SWP_NOACTIVATE
End If
endofsub:
End Sub
'
'  This sub maximizes the video to the window. It causes the movie
' to lose it's aspect ratio but sometimes it looks good on FullScreen
' mode. ( at 1600x1200 looks very nice :~) )
'
Public Sub ScaleVideoMax()
On Error GoTo endofsub

Dim rc As RECT
Dim WindowWidth As Long
Dim windowHeight As Long
Dim raport As Long

GetWindowRect VideoWindowHandle, rc
WindowWidth = (rc.Right - rc.Left)
windowHeight = (rc.Bottom - rc.Top)
PutStream 0, 0, WindowWidth, windowHeight
endofsub:
End Sub
'
'  You can use this sub to force the video file to be 4:3
' I have used this sub in an earlier version of this project.
'
Public Sub ScaleVideoPAL()
On Error GoTo endofsub
Dim rc As RECT
Dim WindowWidth As Long
Dim windowHeight As Long
Dim NewWidth As Long
Dim NewHeight As Long

Dim raport As Long
GetWindowRect VideoWindowHandle, rc
WindowWidth = (rc.Right - rc.Left)
windowHeight = (rc.Bottom - rc.Top)
NewWidth = WindowWidth
NewHeight = NewWidth / 4 * 3
PutStream 0, (windowHeight - NewHeight) / 2 - 10, NewWidth, NewHeight
If windowHeight < NewHeight Then
SetWindowPos VideoWindowHandle, 0&, rc.Left, rc.Top, NewWidth, NewHeight, SWP_NOZORDER Or SWP_NOACTIVATE
End If
endofsub:
End Sub

'
'  You can use this sub to force the video file to be 16:9
' I have used this sub in an earlier version of this project.
'
Public Sub ScaleVideoNTSC()
On Error GoTo endofsub

Dim rc As RECT
Dim WindowWidth As Long
Dim windowHeight As Long
Dim NewWidth As Long
Dim NewHeight As Long

Dim raport As Long
GetWindowRect VideoWindowHandle, rc
WindowWidth = (rc.Right - rc.Left)
windowHeight = (rc.Bottom - rc.Top)
NewWidth = WindowWidth
NewHeight = NewWidth / 16 * 9
PutStream 0, (windowHeight - NewHeight) / 2, NewWidth, NewHeight
If windowHeight < NewHeight Then
SetWindowPos VideoWindowHandle, 0&, rc.Left, rc.Top, NewWidth, NewHeight, SWP_NOZORDER Or SWP_NOACTIVATE
End If
endofsub:
End Sub


Public Function ConvertTime2String(Value As Long) As String
Dim ms As Long
Dim s As Long
Dim m As Long
Dim h As Long
Dim t As Long
Dim sms As String
Dim ss As String
Dim sm As String
Dim sh As String

Dim converted As String
ms = Value - (Value \ 1000) * 1000
s = 0
m = 0
h = 0
t = Value \ 1000
While t > 59
m = m + 1
If m = 60 Then
   m = 0
   h = h + 1
End If
t = t - 60
Wend
s = t
sh = Trim$(CStr(h))
If Len(sh) = 1 Then sh = "0" & sh
sm = Trim$(CStr(m))
If Len(sm) = 1 Then sm = "0" & sm
ss = Trim$(CStr(s))
If Len(ss) = 1 Then ss = "0" & ss

sms = Trim$(CStr(ms))
If Len(sms) = 1 Then sms = "0" & sms
If Len(sms) = 2 Then sms = "0" & sms
converted = sm & ":" & ss & "." & sms
If h > 0 Then converted = sh & ":" & converted
ConvertTime2String = converted
End Function

Public Function ConvertTime(Value As Long) As typeTime
Dim ElapsedTime As typeTime
Dim t As Long

With ElapsedTime
.Hours = 0
.Minutes = 0
.Seconds = 0
.MilliSeconds = Value - (Value \ 1000) * 1000
t = Value \ 1000
While t > 59
.Minutes = .Minutes + 1
If .Minutes = 60 Then
   .Minutes = 0
   .Hours = .Hours + 1
End If
t = t - 60
Wend
.Seconds = t
End With
ConvertTime = ElapsedTime
End Function
