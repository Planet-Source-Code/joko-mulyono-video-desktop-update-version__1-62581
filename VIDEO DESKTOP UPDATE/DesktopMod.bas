Attribute VB_Name = "DesktopMod"
Option Explicit
Private Const SPI_SETSCREENSAVEACTIVE     As Integer = 17
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                                                          ByVal hWnd2 As Long, _
                                                                          ByVal lpsz1 As String, _
                                                                          ByVal lpsz2 As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          ByVal lpvParam As Any, _
                                                                                          ByVal fuWinIni As Long) As Long
Public Sub HideDesktop(ByVal lngHwnd As Long, ByVal DeskShow As Boolean)
'If you use this modul without VIDEODESKTOPCLASS then remove these comments below
'-----------------------
'Dim lnghwnd As Long
'lnghwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
'-----------------------
    If DeskShow Then
        ShowWindow lngHwnd, 0
    End If
    If Not DeskShow Then
        ShowWindow lngHwnd, 5
    End If
End Sub
Public Sub ScreenSaverActive(ByVal Active As Boolean)
Dim Enabled As Long
    Enabled = IIf(Active, 1, 0)
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, Enabled, 0&, 0
End Sub


