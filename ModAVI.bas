Attribute VB_Name = "ModAVI"
Option Explicit
'
'  This module is used to determine the number of frames/second of
' an *.avi format movie.
'  See the documentation included for aknowledjements.
'

Public Function AVI_FramesPerSecond(FileName As String) As Double

Dim FileHandle As Long         ' Handle to the video file
Dim AviInfo As AVIFileInfo     ' Avi Information Structure
Dim fps As Double              ' The number of frames/second

fps = 0     ' Set the initial frames/second to 0
AVIFileInit ' Initialize the AVIFile library

' The next step is to create a handle to the AVI file
If AVIFileOpen(FileHandle, FileName, OF_SHARE_DENY_WRITE, ByVal 0&) = 0 Then
' We now retrieve the AVI information..
 If AVIFileInfo(FileHandle, AviInfo, Len(AviInfo)) = 0 Then ' All is Ok
    fps = AviInfo.dwRate / AviInfo.dwScale                  ' Get the value
 End If
' We got the value needed, we can close the file now
 AVIFileRelease FileHandle
Else  'There was a problem opening the video file or it wasn't an *.avi file
fps = 0
End If

' Exit the AVIFile library and decrement the reference count for the library
AVIFileExit
' Return the value we wanted.
AVI_FramesPerSecond = fps
End Function

