VERSION 5.00
Begin VB.Form FrmSub 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Subtitle Window"
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   Icon            =   "FrmSub.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer TimerPos 
      Interval        =   250
      Left            =   6840
      Top             =   480
   End
   Begin VB.TextBox txtSub 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   360
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmSub.frx":000C
      Top             =   0
      Width           =   5775
   End
   Begin VB.Timer TimerSub 
      Interval        =   150
      Left            =   6840
      Top             =   0
   End
End
Attribute VB_Name = "FrmSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OldText  As String
Private Const Lines1 = 25
Private Const Lines2 = 49
Private Const Lines3 = 73
Dim Lines As Long
Public Sub HideSub()
Me.Top = Screen.Height + 10
End Sub

Private Sub Form_Load()
HideSub
End Sub

Private Sub TimerPos_Timer()
Dim Height As Long
Dim Width As Long
Dim Top As Long
Dim Left As Long
If Not (IsLoaded) Then Exit Sub
If Not (IsPlaying) Then Exit Sub
If IsPaused Then Exit Sub
If Me.Top >= Screen.Height Then Exit Sub
If GetForegroundWindow <> FrmView.hWnd And Not (IsFullscreen) Then Exit Sub
If IsFullscreen And FrmToolBar.Top < Screen.Height Then Exit Sub
If Lines = 3 Then Height = Lines3 * Screen.TwipsPerPixelY
If Lines = 2 Then Height = Lines2 * Screen.TwipsPerPixelY
If Lines = 1 Then Height = Lines1 * Screen.TwipsPerPixelY
Left = FrmView.Left + 5 * Screen.TwipsPerPixelX
Width = FrmView.Width - 10 * Screen.TwipsPerPixelX
Top = FrmView.Top + FrmView.Height - Height - 5 * Screen.TwipsPerPixelY
If Left <> Me.Left Or Top <> Me.Top Or Width <> Me.Width Or Height <> Me.Height Then
Me.Left = Left
Me.Top = Top
Me.Height = Height
Me.Width = Width
End If
End Sub

Public Sub TimerSub_Timer()
Dim Record As SubRecord
Dim RecordNr As Long
Dim i As Long
Dim s As String
Dim text As String
If SubtitlesLoaded = False Then GoTo HideAndClose
If IsLoaded = False Then GoTo HideAndClose
If Not (IsPlaying) Then GoTo HideAndClose
If IsPaused Then GoTo HideAndClose
If IsMovie(FileName) = False Then GoTo HideAndClose
If LCase$(ExtractExtension(FileName)) <> "avi" Then GoTo HideAndClose
If GetForegroundWindow <> FrmView.hWnd And Not (IsFullscreen) Then GoTo HideAndClose
If IsFullscreen And FrmToolBar.Top < Screen.Height Then GoTo HideAndClose
RecordNr = FindRecord(Round((CurrentStreamFrame / 1000) * Media_Information.FramesPerSec, 3))
Record = GetRecord(RecordNr)

text = Record.text
If text = "" Then GoTo HideAndClose
If text = OldText Then Exit Sub
OldText = text
txtSub.text = ""
txtSub.Left = 0
txtSub.Top = 0
Lines = 1
For i = 1 To Len(text)
 s = Mid$(text, i, 1)
 If s = "|" Then
   txtSub.text = txtSub.text & vbCrLf
   Lines = Lines + 1
 Else
   txtSub.text = txtSub.text & s
 End If
Next i

If Lines = 3 Then Me.Height = Lines3 * Screen.TwipsPerPixelY
If Lines = 2 Then Me.Height = Lines2 * Screen.TwipsPerPixelY
If Lines = 1 Then Me.Height = Lines1 * Screen.TwipsPerPixelY
Me.Left = FrmView.Left + 5 * Screen.TwipsPerPixelX
Me.Width = FrmView.Width - 10 * Screen.TwipsPerPixelX
Me.Top = FrmView.Top + FrmView.Height - Me.Height - 5 * Screen.TwipsPerPixelY
txtSub.Width = Me.Width
SetAlwaysOnTop AlwaysOn, Me.hWnd
Me.Visible = True
GoTo AllOk
HideAndClose:
HideSub
Exit Sub
AllOk:
 End Sub

Private Sub txtSub_GotFocus()
If IsFullscreen Then
FrmView.SetFocus
Else
Me.SetFocus
End If


End Sub
