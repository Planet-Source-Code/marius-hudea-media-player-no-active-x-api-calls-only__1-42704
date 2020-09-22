Attribute VB_Name = "ModProgressBar"
'
'  This module is used to draw the progress bars. They're
' using the GDI (Graphics Device Interface), "a dynamic-link
' library that processes graphics function calls from a
' Windows-based application and passes those calls to the
' appropriate device driver."
'
'  In short, everything is drawn very fast, due to the hardware
' acceleration almoust any newer graphics card has.
'
Option Explicit

Private Brush As Long

Private Sub DrawBar(Handle As Long, StartX As Long, StartY As Long, EndX As Long, EndY As Long, Color As Long)
Dim rct As RECT
'
' First, we create the brush that will be used to fill the bar.
'
Brush = CreateSolidBrush(Color)
'
' We "tell" the PictureBox to use the new brush
'
SelectObject Handle, Brush
'
' Next step is to define the rectangular region we want to fill.
'
With rct
.Left = StartX
.Top = StartY
.Right = EndX - StartX
.Bottom = EndY - StartY
End With
'
' Fill the rectangle with the brush we have created.
'
FillRect Handle, rct, Brush
'
' Delete the brush from the memory
'
DeleteObject Brush

End Sub
'
'  This sub will draw the text on a picture box or on anything
' that has a hDC using the GDI
Private Sub DrawText(Handle As Long, StartX As Long, StartY As Long, text As String, Color As Long)
'
' Set the text color.
'
SetTextColor Handle, Color
'
' Print the text on the PictureBox.
'
TextOut Handle, StartX, StartY, text, Len(text)
End Sub
'
' This sub draws the progress bar
'
Public Sub SetProgressValue(Picture As PictureBox, Percent As Long, Optional ShowPercent As Boolean = True, Optional FillColor As Long = vbBlue, Optional BackColor As Long = vbWhite)
Dim PicWidth As Long
Dim PicHeight As Long
Dim RealPercent As Long
Dim StringPercent As String
RealPercent = Percent

If RealPercent < 0 Or RealPercent > 100 Then RealPercent = 0

StringPercent = Str$(RealPercent) & "%"
'
' Get the width and height of the picture in pixels
PicWidth = Round(Picture.Width / Screen.TwipsPerPixelX)
PicHeight = Round(Picture.Height / Screen.TwipsPerPixelY)
'
' Clear the picture
'
Picture.Cls

If RealPercent = 0 Then
Else
If RealPercent = 100 Then
DrawBar Picture.hdc, 0, 0, CLng(PicWidth), PicHeight, FillColor
Else
DrawBar Picture.hdc, 0, 0, CLng(Round(PicWidth / 100 * RealPercent)), PicHeight, FillColor
End If
End If
'
' Draw the text with 3 different text colors, depending on the
' current percent of the stream.
'
If ShowPercent Then
 If RealPercent < 51 Then
 DrawText Picture.hdc, Round(PicWidth / 2 - 5), 1, StringPercent, FillColor
 Else
  If RealPercent > 50 And RealPercent < 56 Then
  DrawText Picture.hdc, Round(PicWidth / 2 - 5), 1, StringPercent, vbGreen
  Else
  DrawText Picture.hdc, Round(PicWidth / 2 - 5), 1, StringPercent, BackColor
  End If
 End If
End If
End Sub
'
'  This sub clears the progress bar.
'

Public Sub ClearProgressBar(pic As PictureBox, Optional BackColor As Long = vbWhite)
Dim PicWidth As Long
Dim PicHeight As Long
PicWidth = Round(pic.Width / Screen.TwipsPerPixelX)
PicHeight = Round(pic.Height / Screen.TwipsPerPixelY)
DrawBar pic.hdc, 0, 0, CLng(PicWidth), CLng(PicHeight), BackColor
pic.Refresh
End Sub
