Attribute VB_Name = "ModSubtitles"
Option Explicit
'
'  This module is used to open and to load subtitles.
'  The array where the subtitle is kept is dynamic, it can hold
' about 1000 millions of lines ( perhaps it is limited by the
' memory on your computer )
'  At this moment only *.SUB and *.SRT subtitles are accepted.
'  Due to the fact that the video is shown in Overlay Mode, I
' can't put text directly on the video image. ( At least I don't
' know how, if somebody knows... ) I have to use a new form to
' show maximum 3 lines of text.

Public Type SubRecord
Startframe As Long
EndFrame As Long
text As String
End Type

Private LargeArray() As SubRecord

Dim TotalArrays As Long
Dim TotalRecords As Long

Dim CurrentRecord As Long
Dim CurrentArray As Long

Public SubtitlesLoaded As Boolean

Public Sub InitLargeArray()
Dim StartRecord As SubRecord
ReDim LargeArray(0 To 0, 0 To 32767)
CurrentRecord = 0
CurrentArray = 0
TotalRecords = 0
TotalArrays = 0
With StartRecord
.Startframe = 0
.EndFrame = 0
.text = ""
End With
LargeArray(0, 0) = StartRecord
SubtitlesLoaded = False
End Sub

Public Sub AddRecord(Record As SubRecord)
If CurrentRecord = 32767 Then
   TotalArrays = TotalArrays + 1
   CurrentArray = CurrentArray + 1
   CurrentRecord = 0
   ReDim Preserve LargeArray(0 To TotalArrays, 0 To 32767)
Else
CurrentRecord = CurrentRecord + 1
End If
LargeArray(CurrentArray, CurrentRecord) = Record
TotalRecords = TotalRecords + 1
End Sub

Public Function GetRecord(RecordPosition As Long) As SubRecord
Dim RequestArray As Long
Dim RequestPosition As Long

RequestArray = Round((RecordPosition / 32767), 0)
RequestPosition = RecordPosition - (RequestArray) * 32768
GetRecord = LargeArray(RequestArray, RequestPosition)
End Function

Public Function GetTotalRecords() As Long
GetTotalRecords = TotalRecords
End Function

Public Sub LoadSUB(FileName As String)
On Error GoTo error_ocurred
Dim s As String
Dim NewRecord As SubRecord
Dim i As Long
InitLargeArray
Open FileName For Input Access Read As #1
While Not (EOF(1))
Line Input #1, s
If CanBeSUBRecord(s) Then
AddRecord DecodeLine(s)
End If
Wend
Close #1
GoTo success
error_ocurred:
Close #1
MsgBox ERROR_SUB, vbOKOnly + vbCritical, TXTERROR
SubtitlesLoaded = False
Exit Sub
success:
SubtitlesLoaded = True
End Sub

' This function returns True is the text line can be
' a valid line from a *.sub file
'
Private Function CanBeSUBRecord(s As String) As Boolean
Dim Data1(0 To 16) As Long
Dim Data2(0 To 16) As Long
Dim NrData1 As Long
Dim NrData2 As Long
Dim t1 As String
Dim t2 As String
Dim i As Long
If s = "" Then
CanBeSUBRecord = False
Exit Function
End If
If Len(s) < 6 Then
CanBeSUBRecord = False
Exit Function
End If
If Len(s) > 500 Then
CanBeSUBRecord = False
Exit Function
End If
NrData1 = -1
NrData2 = -1
For i = 1 To Len(s)
If Mid$(s, i, 1) = "{" Then
                       NrData1 = NrData1 + 1
                       Data1(NrData1) = i
ElseIf Mid$(s, i, 1) = "}" Then
                       NrData2 = NrData2 + 1
                       Data2(NrData2) = i
End If
Next i
t1 = Mid$(s, Data1(0) + 1, Data2(0) - Data1(0) - 1)
t2 = Mid$(s, Data1(1) + 1, Data2(1) - Data1(1) - 1)
If Not (IsNumeric(t1)) Then
CanBeSUBRecord = False
Exit Function
End If
If Not (IsNumeric(t2)) Then
CanBeSUBRecord = False
Exit Function
End If
CanBeSUBRecord = True
End Function

Private Function DecodeLine(s As String) As SubRecord
Dim Data1(0 To 16) As Long
Dim Data2(0 To 16) As Long
Dim NrData1 As Long
Dim NrData2 As Long
Dim text As String
Dim rec As SubRecord
Dim i As Integer
Dim t1 As String
Dim t2 As String

NrData1 = -1
NrData2 = -1

For i = 1 To Len(s)
If Mid$(s, i, 1) = "{" Then
                       NrData1 = NrData1 + 1
                       Data1(NrData1) = i
ElseIf Mid$(s, i, 1) = "}" Then
                       NrData2 = NrData2 + 1
                       Data2(NrData2) = i
End If
Next i
rec.Startframe = 0
rec.EndFrame = 0
rec.text = ""
t1 = Mid$(s, Data1(0) + 1, Data2(0) - Data1(0) - 1)
t2 = Mid$(s, Data1(1) + 1, Data2(1) - Data1(1) - 1)
If IsNumeric(t1) Then rec.Startframe = CLng(t1)
If IsNumeric(t2) Then rec.EndFrame = CLng(t2)
rec.text = Mid$(s, Data2(1) + 1, Len(s) - Data2(1))
DecodeLine = rec
End Function

Public Function FindRecord(Frame As Long) As Long
Dim RecordNumber As Long
Dim rec As SubRecord
Dim i As Long
RecordNumber = 0
For i = 1 To GetTotalRecords
   rec = GetRecord(i)
   If (rec.Startframe <= Frame) And (rec.EndFrame >= Frame) Then
      FindRecord = i
      Exit Function
      Exit For
   End If
Next i
FindRecord = 0
End Function

Public Sub LoadSRT(FileName As String)
On Error GoTo error_ocurred

Dim fps As Long
Dim s As String
Dim NewRecord As SubRecord
fps = Media_Information.FramesPerSec
InitLargeArray
Open FileName For Input Access Read As #1
While Not (EOF(1))
Line Input #1, s
If CanBeSrtRecord(s) Then
 NewRecord.Startframe = GetTime(s, 1) * fps
 NewRecord.EndFrame = GetTime(s, 2) * fps
 Line Input #1, s
 NewRecord.text = s
 Line Input #1, s
 While s <> ""
 NewRecord.text = NewRecord.text & "|" & s
 Line Input #1, s
 Wend
 AddRecord NewRecord
End If
Wend
Close #1
GoTo success
error_ocurred:
Close #1
MsgBox ERROR_SUB, vbOKOnly + vbCritical, TXTERROR
SubtitlesLoaded = True
Exit Sub
success:
SubtitlesLoaded = True
End Sub

' This function returns True is the text line can be
' a valid line from a *.srt file
'
Private Function CanBeSrtRecord(s As String) As Boolean
If s = "" Then
CanBeSrtRecord = False
End If
If InStr(1, s, "-->", vbTextCompare) = 0 Then
CanBeSrtRecord = False
Else
CanBeSrtRecord = True
End If
End Function

Private Function GetTime(s As String, Nr As Byte) As Long
Dim Hours As Long
Dim Minutes As Long
Dim Seconds As Long
Dim MilliSeconds As Long
Dim Value As Byte
If Nr = 1 Then
   Value = 0
Else
   Value = 17
End If
Hours = CLng(Mid$(s, 1 + Value, 2))
Minutes = CLng(Mid$(s, 4 + Value, 2))
Seconds = CLng(Mid$(s, 7 + Value, 2))
MilliSeconds = CLng(Mid$(s, 10 + Value, 3))
GetTime = CLng(Hours * 3600 + Minutes * 60 + Seconds + CLng(MilliSeconds / 1000))
End Function

Public Function DetectSubtitleType(FileName As String) As String
On Error GoTo detected
Dim microDVD As Boolean
Dim SubRIP As Boolean
Dim s As String
microDVD = False
SubRIP = False
If DetectFileSize(FileName) = 0 Then
   DetectSubtitleType = "Other"
   Exit Function
End If
Open FileName For Input Access Read As #1
While Not EOF(1)
Line Input #1, s
If CanBeSUBRecord(Trim$(s)) Then
   microDVD = True
   GoTo detected
End If

Wend
detected:
Close #1
SubRIP = IsSRT(FileName)
If (SubRIP = False) And (microDVD = False) Then
    DetectSubtitleType = "Other"
    Exit Function
End If
If (SubRIP = True) Then
  DetectSubtitleType = "SubRIP"
  Exit Function
End If

If (microDVD = True) Then
 DetectSubtitleType = "microDVD"
 Exit Function
End If

End Function


Public Function GetSubLine(text As String, LineNumber As Integer) As String
Dim StringLines(1 To 16) As String
Dim i As Integer
Dim j As Integer
Dim s As String
For i = 1 To 3
StringLines(i) = ""
Next
j = 1
For i = 1 To Len(text)
s = Mid$(text, i, 1)
If s = "|" Then
j = j + 1
Else
StringLines(j) = StringLines(j) + s
End If
Next
GetSubLine = StringLines(LineNumber)
End Function

Private Function DetectFileSize(FileName As String) As Long
Dim filesizehigh As Long
Dim fs As Long
Dim Handle As Long
Handle = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
If Handle = 0 Then
DetectFileSize = 0
Exit Function
Else
fs = GetFileSize(Handle, filesizehigh)
CloseHandle Handle
DetectFileSize = fs
End If
End Function

Public Function IsSRT(FileName As String) As Boolean
On Error Resume Next
Dim isSubtitle As Boolean
Dim s As String

isSubtitle = False

Open FileName For Input Access Read As #1
If Err.Number <> 0 Then
   IsSRT = False
   Exit Function
End If
While Not EOF(1)
Line Input #1, s
If InStr(1, s, "-->") <> 0 Then isSubtitle = True
Wend
Close #1
IsSRT = isSubtitle
End Function



