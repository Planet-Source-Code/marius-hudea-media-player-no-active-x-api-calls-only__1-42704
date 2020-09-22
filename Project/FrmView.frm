VERSION 5.00
Begin VB.Form FrmView 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Video Window"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   3060
   ControlBox      =   0   'False
   Icon            =   "FrmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FrmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CloseVideoWindow As Boolean

Private Sub Form_Load()
CloseVideoWindow = False
End Sub

Private Sub Form_Resize()
ResizeMe
End Sub

Public Sub ResizeMe()
If FrmMain.mnu_bestfit.Checked Then
ScaleVideoBest
ElseIf FrmMain.mnu_fitwindow.Checked Then
ScaleVideoMax
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
If CloseVideoWindow = True Then
Unload Me
Else
Cancel = 1
End If
End Sub

