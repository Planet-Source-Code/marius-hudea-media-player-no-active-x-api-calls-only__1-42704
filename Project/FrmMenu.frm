VERSION 5.00
Begin VB.Form FrmMenu 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1500
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   3945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mnu_size 
      Caption         =   "Size"
      Begin VB.Menu mnu_best 
         Caption         =   "&Best Fit ( Default )"
      End
      Begin VB.Menu mnu_sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_full 
         Caption         =   "&Full Size"
      End
      Begin VB.Menu mnu_43 
         Caption         =   "4:3"
      End
      Begin VB.Menu mnu_169 
         Caption         =   "16:9"
      End
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnu_169_Click()
ScaleVideoNTSC
End Sub

Private Sub mnu_43_Click()
ScaleVideoPAL
End Sub

Private Sub mnu_best_Click()
ScaleVideoBest
End Sub

Private Sub mnu_full_Click()
ScaleVideoMax
End Sub
