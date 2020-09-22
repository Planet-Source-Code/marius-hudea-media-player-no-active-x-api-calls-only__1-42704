Attribute VB_Name = "ModOpenFile"
Option Explicit
' This module selects a file using the common dialog control
' but using API functions , not using the ocx control.
Private FolderPath As String

Public Function OpenFile() As String
Dim s As String
Dim i As Long
'
'  This function allows the one who runs this application to select
' a media file.
'
Dim OFName As OpenFileName
OFName.lStructSize = Len(OFName) 'Set the length of the structure
OFName.hWndOwner = FrmMain.hWnd  'Set the parent window
OFName.hInstance = App.hInstance 'Set the application's instance
' Select a filter
s = ""
For i = 0 To ExtensionNr
s = s & "*." & ExtensionList(i).Extension & ";"
Next i
s = Left$(s, Len(s) - 1)
OFName.lpstrFilter = "All Media Files " & Chr$(0) & s & Chr(0) & _
                     "Only Video Files" & Chr(0) & Stream_Movies & Chr(0) & _
                     "Only Audio Files" & Chr(0) & Stream_Sounds & Chr(0) & _
                     "Only MIDI Files" & Chr(0) & Stream_Midi & Chr(0) & _
                     "Audio CD" & Chr(0) & Stream_AudioCD & Chr(0) & _
                     "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
'
OFName.lpstrFile = Space$(254) 'create a buffer for the file
OFName.nMaxFile = 255 'set the maximum length of a returned file
OFName.lpstrFileTitle = Space$(254) 'Create a buffer for the file title
OFName.nMaxFileTitle = 255 'Set the maximum length of a returned file title

If FolderPath = "" Then
FolderPath = "C:\"
End If

OFName.lpstrInitialDir = FolderPath 'Set the initial directory
OFName.lpstrTitle = "Select a media file" 'Set the title
'Next we set the flags that will modify the way our window looks
OFName.flags = OFN_NONETWORKBUTTON Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
'Show the 'Open File'-dialog
If GetOpenFileName(OFName) Then ' Everything is Ok
   FolderPath = ExtractFolder(OFName.lpstrFile)
   OpenFile = Trim$(OFName.lpstrFile)
Else 'There was an error or the user pressed Cancel
OpenFile = ""
End If
End Function
'
' This function will extract the folder from a file name
' Ex. C:\My Documents\Document.doc --> C:\My Documents
'
Private Function ExtractFolder(FileName As String) As String
Dim i As Integer
Dim LastPos As Integer

For i = 1 To Len(FileName)
    If Mid$(FileName, i, 1) = "\" Then
       LastPos = i
    End If
Next i
ExtractFolder = Mid$(FileName, 1, LastPos)
End Function
'
' This function will extract the file title from a file name
' Ex. C:\My Documents\Document.doc --> Document.doc
'
Public Function ExtractName(FileName As String) As String
Dim i As Integer
Dim LastPos As Integer
Dim s As String

For i = 1 To Len(FileName)
    If Mid$(FileName, i, 1) = "\" Then
       LastPos = i
    End If
Next i
s = Mid$(FileName, LastPos + 1, Len(FileName) - LastPos)
s = Right(s, Len(s))
ExtractName = s

End Function
'
' This function will extract the extension from a file name
' Ex. C:\My Documents\Document.doc --> doc
'

Public Function ExtractExtension(FileName As String) As String
Dim i As Integer
Dim LastPos As Integer
Dim s As String

For i = 1 To Len(FileName)
    If Mid$(FileName, i, 1) = "." Then
       LastPos = i
    End If
Next i
s = Mid$(FileName, LastPos + 1, Len(FileName) - LastPos)
ExtractExtension = LCase$(s)
End Function

'
' This function will select a subtitle. It is almost the same
' function like the previous OpenFile function so i won't explain
' it again
'
Public Function OpenSubtitle() As String
Dim OFName As OpenFileName
OFName.lStructSize = Len(OFName)
OFName.hWndOwner = FrmMain.hWnd
OFName.hInstance = App.hInstance
OFName.lpstrFilter = "All Subtitles " & Chr$(0) & "*.sub;*.srt;*.txt" + Chr$(0) & _
                     "*.SUB" & Chr(0) & "*.sub" & Chr(0) & _
                     "*.SRT" & Chr(0) & "*.srt" & Chr(0) & _
                     "Text Files ( Auto Detect )" & Chr$(0) & "*.txt" & Chr$(0)

OFName.lpstrFile = Space$(254)
OFName.nMaxFile = 255
OFName.lpstrFileTitle = Space$(254)
OFName.nMaxFileTitle = 255

If FolderPath = "" Then
FolderPath = "C:\"
End If
OFName.lpstrInitialDir = FolderPath
OFName.lpstrTitle = "Select a subtitle"
OFName.flags = OFN_NONETWORKBUTTON Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
If GetOpenFileName(OFName) Then
   FolderPath = ExtractFolder(OFName.lpstrFile)
   OpenSubtitle = Trim$(OFName.lpstrFile)
Else
OpenSubtitle = ""
End If

End Function

Public Function OpenFolder() As String
Dim iNull As Integer
Dim lpIDList As Long
Dim lResult As Long
Dim sPath As String
Dim BInfo As BrowseInfo

With BInfo
 'Set the owner window
 .hWndOwner = FrmMain.hWnd
 'lstrcat appends the two strings and returns the memory address
 .lpszTitle = lstrcat("Please select a folder:", "")
 'Return only if the user selected a directory
 .ulFlags = BIF_RETURNONLYFSDIRS
 End With
 'Show the 'Browse for folder' dialog
 lpIDList = SHBrowseForFolder(BInfo)
 If lpIDList Then
    sPath = String$(MAX_PATH, 0)
    'Get the path from the IDList
    SHGetPathFromIDList lpIDList, sPath
    'free the block of memory
    CoTaskMemFree lpIDList
    iNull = InStr(sPath, vbNullChar)
    If iNull Then
       sPath = Left$(sPath, iNull - 1)
    End If
Else
sPath = ""
End If
OpenFolder = sPath
End Function
