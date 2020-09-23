Attribute VB_Name = "Systray"
Option Explicit
Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type
Public Enum IconMessenger
    none = &H10
    Warning = &H12
    Critical = &H13
    Information = &H11
End Enum

Private Const NIM_ADD         As Long = &H0
Private Const NIM_DELETE      As Long = &H2
Private Const NIM_MODIFY      As Long = &H1

Private Const NIF_MESSAGE     As Long = &H1
Private Const NIF_ICON        As Long = &H2
Private Const NIF_INFO        As Long = &H10
Private Const NIF_TIP         As Long = &H4

Public Const WM_MOUSEMOVE     As Long = &H200
Public Const WM_RBUTTONUP     As Long = &H205
Public Const WM_LBUTTONUP     As Long = &H202
'Public Const WM_MBUTTONDBLCLK = &H209
'Public Const WM_LBUTTONDOWN = &H201
'Public Const WM_LBUTTONUP = &H202
'Public Const WM_LBUTTONDBLCLK = &H203
'Public Const WM_RBUTTONDOWN = &H204
'Public Const WM_RBUTTONDBLCLK = &H206
'Public Const WM_MBUTTONDOWN = &H207
'Public Const WM_MBUTTONUP = &H208
'Public Const WM_MBUTTONDBLCLK = &H209

Private iData                 As NOTIFYICONDATA
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                                                          ByVal hWnd2 As Long, _
                                                                          ByVal lpsz1 As String, _
                                                                          ByVal lpsz2 As String) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                                                                       lpData As NOTIFYICONDATA) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Sub IconTray(DataIcon As PictureBox, _
                          ByVal zTip As String, remove As Boolean)
    With iData
        .cbSize = Len(iData)
        .hwnd = Form1.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = DataIcon
        .szTip = zTip & vbNullChar
    End With
    If Not remove Then
       Shell_NotifyIcon NIM_ADD, iData
       
    Else
       Shell_NotifyIcon NIM_DELETE, iData
    End If
End Sub

Public Sub PopINFO(Message As String, Title As String, IconMessenger As IconMessenger, ShowPop As Boolean, Optional TimeOut As Long)
With iData
    .hwnd = Form1.hwnd
    .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
    
    If ShowPop Then
        .szInfo = Message & Chr(0)
        .szInfoTitle = Title & Chr(0)
        .dwInfoFlags = IconMessenger
    Else
        .szInfo = Chr(0)
        .szInfoTitle = Chr(0)
        .dwInfoFlags = &H0
    End If
End With
Shell_NotifyIcon NIM_MODIFY, iData
If TimeOut > 0 Then
   HideBaloon (TimeOut)
End If
End Sub
Public Sub HideBaloon(mSecond As Long)
Sleep mSecond
PopINFO "", "", none, False
End Sub
