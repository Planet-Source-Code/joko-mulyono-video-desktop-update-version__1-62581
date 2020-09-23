Attribute VB_Name = "HOTKEYmod"
Option Explicit
Private isHide As Boolean
Public Enum ModConst
    MOD_ALT = &H1
    MOD_CONTROL = &H2
    MOD_SHIFT = &H4
End Enum
#If False Then
Private MOD_ALT, MOD_CONTROL, MOD_SHIFT
#End If

Private px                As Long
Private Const WM_HOTKEY   As Long = &H312
Private hot_counter       As Long
Public oldProc            As Long
Public MM                 As Object
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                      ByVal lpWindowName As String) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, _
                                                       ByVal hwnd As Long, _
                                                       ByVal Msg As Long, _
                                                       ByVal wParam As Long, _
                                                       ByVal lParam As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal id As Long, _
                                                      ByVal fsModifiers As Long, _
                                                      ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, _
                                                        ByVal id As Long) As Long
Private Sub Main()
    On Error Resume Next

    Set MM = New VIDEODESKTOPCLASS
    Load Form1
    On Error GoTo 0

End Sub
Public Sub ReleasHotKey(ByVal lngHwnd As Long)

Dim I As Integer
    For I = 1 To hot_counter
        UnregisterHotKey lngHwnd, I
    Next I
    hot_counter = 0
End Sub
Public Function SetHotKey(ByVal lngHwnd As Long, _
                          Modifier As ModConst, _
                          Optional KeyCode As Integer) As Long

    hot_counter = hot_counter + 1
    SetHotKey = RegisterHotKey(lngHwnd, hot_counter, Modifier, KeyCode)
End Function
Public Function WndProc(ByVal lngHwnd As Long, _
                        ByVal uMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

    WndProc = 0
    If uMsg = WM_HOTKEY Then
        Select Case wParam
        Case 1
            Form1.OPENVIDEO
        Case 2
            isHide = IIf(isHide = True, False, True)
            If isHide Then
               Form1.HDesktop
               isHide = False
            Else
               Form1.SDesktop
               isHide = True
            End If
        Case 3
            Form1.SDesktop
        Case 4
            px = px - 10
            MM.FForward px
            px = 0
        Case 5
            px = px + 10
            MM.FForward px
            px = 0
        Case 6
            Form1.Play
        Case 7
            Form1.StopPlay
        Case 8
            Form1.Mute
        Case 9
            Form1.Paused
        End Select
    Else
        WndProc = CallWindowProcA(oldProc, lngHwnd, uMsg, wParam, lParam)
    End If
End Function

