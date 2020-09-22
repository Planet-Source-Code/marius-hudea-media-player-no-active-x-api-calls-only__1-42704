Attribute VB_Name = "ModPlayList"
Option Explicit
'
'  This module manages the playlist.
'  The most complex sub here is the RecurseSearch, that's why
' it is the only one commented.
'  I hope you will understand it...
'
Public Playlist() As String
Dim PlayListItems As Long

Public Sub InitPlayList()
ReDim Playlist(0 To 32767)
PlayListItems = -1
frmPlayList.LstPlay.Clear
frmPlayList.LstPlay.ListIndex = -1
End Sub

Public Sub AddFile(Optional FileName As String = "")
Dim s As String
If FileName = "" Then
s = StripNulls(OpenFile)
Else
s = StripNulls(FileName)
End If

If s <> "" And PlayListItems <> 32767 Then
frmPlayList.LstPlay.AddItem ExtractName(s)
PlayListItems = PlayListItems + 1
Playlist(PlayListItems) = s
End If
DoEvents
End Sub

Public Sub DeleteFile()
Dim i As Long
Dim j As Long
Dim ListCount As Long
Dim s As String
If frmPlayList.LstPlay.ListCount = 0 Then Exit Sub
If frmPlayList.LstPlay.SelCount = 0 Then Exit Sub
ListCount = frmPlayList.LstPlay.ListCount
i = 0
While i <= ListCount - 1

If frmPlayList.LstPlay.Selected(i) Then
   s = frmPlayList.LstPlay.List(i)
   frmPlayList.LstPlay.RemoveItem (i)
   ListCount = ListCount - 1
   For j = 0 To PlayListItems
     If ExtractName(Playlist(j)) = s Then
        Playlist(j) = Playlist(PlayListItems)
        PlayListItems = PlayListItems - 1
        GoTo OutOfFor
     End If
   Next j
OutOfFor:
DoEvents
Else
i = i + 1
End If
Wend
If frmPlayList.LstPlay.ListCount = 0 Then frmPlayList.LstPlay.ListIndex = -1
End Sub

Public Sub ClosePlayList()
frmPlayList.KeepPlayList = False
Unload frmPlayList
End Sub

Public Sub AddFolder()
Dim FolderPath As String
'
'  Show the user the Browse For Folder dialog..
'
FolderPath = StripNulls(OpenFolder)
If FolderPath = "" Then Exit Sub
'
'  Our search string must be something like C:\*.* or C:\Music\*.*
'
If Right$(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"

ScanFolder (FolderPath)
End Sub

Public Sub ScanFolder(Path As String, Optional Recursive As Boolean = True)
Dim FindHandle As Long
Dim FindData As Win32_Find_Data
Dim Attr As Long
Dim ret As Long
Dim CurrentFile As String
If Path = "" Then Exit Sub
'Search for all files
'Find the first file..
FindHandle = FindFirstFile(Path & "*", FindData)
If FindHandle = INVALID_HANDLE_VALUE Then
   Exit Sub
End If
ret = 1
While ret <> 0
If ret <> 0 Then
CurrentFile = StripNullsTrim(FindData.cFileName)
Attr = GetFileAttributes(Path & CurrentFile)
If Attr And FILE_ATTRIBUTE_DIRECTORY Then
   If CurrentFile <> "." And CurrentFile <> ".." Then
      If Recursive Then ScanFolder Path & CurrentFile & "\"
   End If
Else
   If IsMovie(CurrentFile) Or IsAudio(CurrentFile) Or IsMidi(CurrentFile) Or LCase$(ExtractExtension(CurrentFile)) = "cda" Then
   
   If LCase$(ExtractExtension(CurrentFile)) <> "dat" Then AddFile Path & CurrentFile
   DoEvents
   End If
   
End If
ret = FindNextFile(FindHandle, FindData)
End If
Wend
FindClose FindHandle
End Sub

Public Function GetExtendedName(s As String) As String
Dim i As Long
If PlayListItems = -1 Then
   GetExtendedName = ""
   Exit Function
End If
For i = 0 To PlayListItems
   If ExtractName(Playlist(i)) = s Then
                                   GetExtendedName = Playlist(i)
                                   Exit Function
   End If
Next i
GetExtendedName = ""
End Function

Public Function GetNextFile() As String
If frmPlayList.LstPlay.ListCount = 0 Then
   GetNextFile = ""
   Exit Function
End If
If frmPlayList.LstPlay.ListIndex < 0 Then
   frmPlayList.LstPlay.ListIndex = 0
   GetNextFile = GetExtendedName(frmPlayList.LstPlay.List(frmPlayList.LstPlay.ListIndex))
   Exit Function
Else
   If frmPlayList.LstPlay.ListIndex = frmPlayList.LstPlay.ListCount - 1 Then
      frmPlayList.LstPlay.ListIndex = 0
      GetNextFile = GetExtendedName(frmPlayList.LstPlay.List(frmPlayList.LstPlay.ListIndex))
      Exit Function
   Else
      frmPlayList.LstPlay.ListIndex = frmPlayList.LstPlay.ListIndex + 1
      GetNextFile = GetExtendedName(frmPlayList.LstPlay.List(frmPlayList.LstPlay.ListIndex))
      Exit Function
   End If
End If
GetNextFile = ""
End Function

Public Function GetPreviousFile() As String
If frmPlayList.LstPlay.ListCount = 0 Then
   GetPreviousFile = ""
   Exit Function
End If
If frmPlayList.LstPlay.ListIndex < 0 Then
   frmPlayList.LstPlay.ListIndex = 0
   GetPreviousFile = GetExtendedName(frmPlayList.LstPlay.List(frmPlayList.LstPlay.ListIndex))
   Exit Function
Else
   If frmPlayList.LstPlay.ListIndex = 0 Then
      frmPlayList.LstPlay.ListIndex = frmPlayList.LstPlay.ListCount - 1
      GetPreviousFile = GetExtendedName(frmPlayList.LstPlay.List(frmPlayList.LstPlay.ListIndex))
      Exit Function
   Else
      frmPlayList.LstPlay.ListIndex = frmPlayList.LstPlay.ListIndex - 1
      GetPreviousFile = GetExtendedName(frmPlayList.LstPlay.List(frmPlayList.LstPlay.ListIndex))
      Exit Function
   End If
End If
GetPreviousFile = ""
End Function

