Attribute VB_Name = "ModMediaInfo"
'
'  -=[ AVI File Information ]=-
'
'  This module is used to extract additional information from
' AVI files like horizontal and vertical pixel count.
'  This way we can resize the window when we start the movie
' and it will look good.
'
Option Explicit
Private Const OF_SHARE_DENY_WRITE As Long = &H20

Private Type AVIFileInfo
    dwMaxBytesPerSec As Long
    dwFlags As Long
    dwCaps As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwLength As Long
    dwEditCount As Long
    szFileType As String * 64
End Type
Private Declare Function AVIFileOpen Lib "avifil32" Alias "AVIFileOpenA" (ppfile As Long, ByVal szFile As String, ByVal mode As Long, pclsidHandler As Any) As Long
Private Declare Function AVIFileRelease Lib "avifil32" (ByVal pfile As Long) As Long
Private Declare Function AVIFileInfo Lib "avifil32" Alias "AVIFileInfoA" (ByVal pfile As Long, pfi As AVIFileInfo, ByVal lSize As Long) As Long
Private Declare Sub AVIFileInit Lib "avifil32" ()
Private Declare Sub AVIFileExit Lib "avifil32" ()



Public Sub GetAVI_Information(FileName As String)
' This sub retrieves information about the avi file FileName and stores
' the width and height in the two dims AVI_Width and AVI_Height
' If there will be an error the values will be set to 320 & 240
Dim FileHandle As Long, AviInfo As AVIFileInfo
' Initialize the AVIFile library
AVIFileInit
' Create a handle to the AVI file
If AVIFileOpen(FileHandle, FileName, OF_SHARE_DENY_WRITE, ByVal 0&) = 0 Then
        'retrieve the AVI information
        If AVIFileInfo(FileHandle, AviInfo, Len(AviInfo)) = 0 Then
            ResetMediaInformation
            Media_Information.Width = AviInfo.dwWidth
            Media_Information.Height = AviInfo.dwHeight
        Else
            'Unable to retrieve information
        ResetMediaInformation
        End If
        'release the file handle
        AVIFileRelease FileHandle
    Else
        ' Error while opening the AVI file...
    ResetMediaInformation
    End If
    'exit the AVIFile library and decrement the reference count for the library
    AVIFileExit
End Sub

Public Sub ResetMediaInformation()

End Sub
