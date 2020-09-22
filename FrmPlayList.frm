VERSION 5.00
Begin VB.Form frmPlayList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Playlist Editor"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "FrmPlayList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton CmdAddFolder 
      Caption         =   "Add Folder"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton CmdAddFile 
      Caption         =   "Add File"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.ListBox LstPlay 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmPlayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public KeepPlayList As Boolean

Private Sub CmdAddFile_Click()
Dim s As String
CmdAddFile.Enabled = False
AddFile
CmdAddFile.Enabled = True
End Sub

Private Sub CmdAddFolder_Click()
AddFolder
End Sub

Private Sub CmdClear_Click()
InitPlayList
End Sub

Private Sub CmdDelete_Click()
DeleteFile
End Sub

Private Sub Form_Load()
InitPlayList
KeepPlayList = True
Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If KeepPlayList Then
frmPlayList.Visible = False
FrmMain.mnu_Playlist.Checked = False
Cancel = -1
Else
Unload Me
End If

End Sub

Private Sub LstPlay_DblClick()
Dim s As String
If LstPlay.ListCount = 0 Then Exit Sub
If LstPlay.ListIndex < 0 Then Exit Sub
s = GetExtendedName(LstPlay.List(LstPlay.ListIndex))
If s = "" Then Exit Sub
FrmMain.OpenMediaFile s
End Sub
