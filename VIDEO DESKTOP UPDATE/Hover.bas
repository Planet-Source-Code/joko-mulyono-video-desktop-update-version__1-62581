Attribute VB_Name = "Hover"
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Public Sub HoverBut(ByVal X As Single, ByVal Y As Single)

'<:-) :WARNING: Unused Sub 'HoverBut'
'Tested:OK
Dim Nama As String
Static Hover As Boolean
'<:-) :SUGGESTION: Static is very memory hungry; try using a Private Module level variable instead
On Error Resume Next
With Form1
If (X < 0) Or (Y < 0) Or (X > 768) Or (Y > 0) Then
ReleaseCapture
'Hover = False

   Nama = GetFileName(.filmname, True)
   PopINFO Nama, "VIDEO DESKTOP", Information, False
Else
SetCapture .hwnd
If Not Hover Then
Hover = True
Nama = GetFileName(.filmname, True)
PopINFO Nama, "VIDEO DESKTOP", Information, True
End If
End If
End With
Hover = False
On Error GoTo 0
End Sub


