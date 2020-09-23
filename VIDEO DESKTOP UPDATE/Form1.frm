VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3990
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3285
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   120
      Picture         =   "Form1.frx":34CA
      ScaleHeight     =   525
      ScaleWidth      =   6105
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.Timer tmrLoop 
      Left            =   1680
      Top             =   2400
   End
   Begin VB.PictureBox imgIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5160
      Picture         =   "Form1.frx":DC64
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Tag             =   "pic"
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   120
      Picture         =   "Form1.frx":1112E
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Menu mnuOp 
      Caption         =   "Option"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuFName 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuDT 
         Caption         =   "Tool"
         Begin VB.Menu mnuHD 
            Caption         =   "Hide Desktop"
            Shortcut        =   +{DEL}
         End
         Begin VB.Menu mnuSD 
            Caption         =   "Show Desktop"
            Enabled         =   0   'False
            Shortcut        =   +{INSERT}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuArrow 
            Caption         =   "Hide Arrow Shortcut"
         End
         Begin VB.Menu mnuloop 
            Caption         =   "Loop"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSt 
            Caption         =   "StartUp"
         End
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCon 
         Caption         =   "Control"
         Enabled         =   0   'False
      End
      Begin VB.Menu spc3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuFF 
         Caption         =   "FForward"
      End
      Begin VB.Menu mnuFR 
         Caption         =   "FRewind"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuM 
         Caption         =   "Mute"
      End
      Begin VB.Menu spc2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HKEY_LOCAL_MACHINE    As Long = &H80000002
Private Const CCM_FIRST             As Long = &H2000
Private Const CCM_SETCOLORSCHEME    As Long = (CCM_FIRST + 2)
Private Const GWL_WNDPROC           As Long = -4
Private Const MOD_CONTROL           As Long = &H2
Private MM                          As Object
Private isPlaying                   As Boolean
Private isPaused                    As Boolean
Private isMute                      As Boolean
Public filmname                     As String
Private lngMsg                      As Long
Private blnFlag                     As Boolean
Private StarUP                      As Boolean
Private nPesan                      As Long
Private HwndDesktop                 As Long
Private isDesktopHide               As Boolean
Private exFName     As Long
Private spath  As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
                                                      lpRect As Any, _
                                                      ByVal bErase As Long) As Long
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                 ByVal hWndNewParent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                      ByVal lpWindowName As String) As Long

Private Sub CLOSEALLPLAYER()
    If isPlaying Then
        With MM
            .setCommand StopCD
            .setCommand CloseCD
        End With
        WALLPAPERTHEME (RestoreWall)
        isChange = False
        tmrLoop.Enabled = False
    End If
End Sub
Private Sub Form_Initialize()
    If App.PrevInstance Then
        End
    End If
    isPaused = False
    isMute = False
    isChange = False
    ScreenSaverActive False
    
End Sub
Private Sub Form_Load()
'Const vbKeySpace = 32 (&H20)
    On Error Resume Next
    Set MM = New VIDEODESKTOPCLASS
    
    HwndDesktop = FindWindow(vbNullString, "Program Manager")
    MM.hwndParent = HwndDesktop
    SetCustomMenus
    IconTray Form1.imgIcon, "VIDEO DESKTOP", False
    PopINFO "Nice and Cool", "VIDEO DESKTOP", Information, True, 1000
    INSERT Picture2.hwnd, Me.hwnd, 6, 247
    SetWidth = 307

     
    With Me
        OldProc = SetWindowLongA(.hwnd, GWL_WNDPROC, AddressOf WndProc)
        SetHotKey .hwnd, MOD_CONTROL, Asc("O")
        SetHotKey .hwnd, MOD_SHIFT, vbKeyDelete  'Del
        SetHotKey .hwnd, MOD_SHIFT, vbKeyInsert  'Ins
        SetHotKey .hwnd, MOD_CONTROL, vbKeyLeft  'left arrow
        SetHotKey .hwnd, MOD_CONTROL, vbKeyRight 'right arrow
        SetHotKey .hwnd, MOD_CONTROL, Asc("P")
        SetHotKey .hwnd, MOD_CONTROL, Asc("S")
        SetHotKey .hwnd, MOD_CONTROL, Asc("M")
        SetHotKey .hwnd, MOD_CONTROL, vbKeySpace 'spacebar
    End With 'Me
    getArrowData
    If isArrow Then
       mnuArrow.Checked = True
    Else
       mnuArrow.Checked = False
    End If
    If CheckStart Then
       mnuSt.Checked = True
    Else
       mnuSt.Checked = False
    End If
    On Error GoTo 0

End Sub
Private Sub SetCustomMenus()
    startODMenus Me, True
    With CustomMenu
        .Texture = True
        Set .Picture = Image1.Picture
        .UseCustomFonts = False
'        .FontUnderline = False
'        .FontName = "Lucida Sans" '"Comic Sans MS" '
'        .FontItalic = False
'        .FontStrikeOut = False
         .PosX = 28
    End With
    With CustomColor
        .ForeColor = RGB(16, 0, 16) 'the magic color
        .DefTextColor = vbRed ' vbBlack
        .HilightColor = RGB(182, 189, 210)
        .NormalColor = RGB(186, 186, 204)
        .BackColor = RGB(58, 110, 165)
        .SelectedTextColor = RGB(0, 0, 255)
        .MenuTextColor = vbBlack
        '.BorderColor = RGB(240, 72, 72)
        .RECTColor = RGB(10, 36, 106)
    End With
    MenuMode = XPlook
    With CustomMenu
        
        .Icon.Add Array(101, 102), "Open"
        .Icon.Add Array(105, 106), "Mute"
        .Icon.Add Array(107, 108), "Play"
        .Icon.Add Array(113, 114), "Pause"
        .Icon.Add Array(109, 110), "FForward"
        .Icon.Add Array(111, 112), "FRewind"
        .Icon.Add Array(115, 116), "Tool"
        '.Icon.Add Array(119, 120), "Hide Desktop"
        '.Icon.Add Array(117, 118), "Show Desktop"
        .Icon.Add Array(121, 122), "Stop"
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)
    Dim Str  As String
    Dim nama As String
    lngMsg = x / Screen.TwipsPerPixelX
    If Not blnFlag Then
        blnFlag = True
        Select Case lngMsg
        Case WM_RBUTTONUP
            SetForegroundWindow Me.hwnd
            Me.PopupMenu mnuOp
        Case WM_LBUTTONUP
             nPesan = nPesan + 1
             Str = LoadResString(100 + nPesan)
             PopINFO Str, "VIDEO DESKTOP", Information, True, 1500
             If nPesan + 100 = 108 Then nPesan = 0
'        Case WM_MOUSEMOVE
'             nama = GetFileName(filmname, True)
'             PopINFO nama, "VIDEO DESKTOP", Information, True, 1500
'

        End Select
        blnFlag = False
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    SetWindowLongA Me.hwnd, GWL_WNDPROC, OldProc
    stopODMenus Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CLOSEALLPLAYER
    IconTray Form1.imgIcon, "VIDEO DESKTOP", True
    ReleasHotKey Me.hwnd
End Sub
Public Sub HDesktop()
    mnuHD_Click
End Sub

Private Sub mnuArrow_Click()
mnuArrow.Checked = Not mnuArrow.Checked
isArrow = IIf(isArrow = True, False, True)
setArrow
End Sub

Private Sub mnuExit_Click()
    
    Unload Me
End Sub
Private Sub mnuFF_Click()
    MM.FForward 5 '5 seconds
End Sub
Private Sub mnuFName_Click()
tmrLoop.Enabled = False
    If isPlaying Then
        CLOSEALLPLAYER
        isPlaying = False
        isChange = False
    End If
    Filter = "Movie (*.dat;*.mpg;*.avi;*.asf;*.wmv)" & Chr$(0) & "*.dat;*.mpg;*.avi;*.wmv" & Chr$(0) & "Other Mov(*.mov)" & Chr$(0) & "*.mov" & Chr$(0)
    exFName = OpenDialog(hwnd, vbNullString, 0, "", vbNullString, "  OPEN MEDIA FILE", spath)
    filmname = ExFilename
    If LenB(filmname) Then
       mnuPlay_Click
    End If
End Sub
Private Sub mnuFR_Click()
    MM.FRewind 5
End Sub
Private Sub mnuHD_Click()
mnuHD.Checked = Not mnuHD.Checked
isDesktopHide = IIf(isDesktopHide = True, False, True)
If Not isDesktopHide Then
   HideDesktop HwndDesktop, False
Else
   HideDesktop HwndDesktop, True
End If
    'mnuSD.Enabled = True
    'mnuHD.Enabled = False
End Sub

Private Sub mnuloop_Click()
mnuloop.Checked = Not mnuloop.Checked
End Sub

Private Sub mnuM_Click()
    isMute = IIf(isMute, False, True)

    If isMute Then
        MM.SetAudioState Chan_All, vd_Off
    Else
        MM.SetAudioState Chan_All, vd_On
    End If
End Sub
Public Sub Mute()
mnuM_Click
End Sub
Private Sub mnuPause_Click()
    MM.setCommand (PauseCD)
    isPaused = True
End Sub
Public Sub Paused()
 mnuPause_Click
End Sub

Private Sub mnuPlay_Click()
'Thank to Daniel camacho for his info about flickering : > flickering is disable now
Dim nama As String
On Error Resume Next
    If isPaused Then
        MM.setCommand (ResumeCD)
        isPaused = False
    Else
        If LenB(filmname) Then
           isPaused = False
            With MM
                .Filename = filmname
                    If Not isChange Then
                       WALLPAPERTHEME (ClearWall) 'remove wallpaper and change color background (RGB:16,0,16)
                       InvalidateRect 0&, ByVal 0, 1& 'refresh desktop :> this can make a flickering on display
                    End If
                .PlayMEDIAFILE
            End With
            nama = GetFileName(filmname, True)
            PopINFO nama, "VIDEO DESKTOP", Information, True, 1500
            isPlaying = True
            tmrLoop.Enabled = True
            tmrLoop.Interval = 500 '
            
        End If
    End If
    On Error GoTo 0
End Sub
Private Sub mnuSD_Click()
    HideDesktop HwndDesktop, False
    InvalidateRect 0&, ByVal 0, 1&
    mnuHD.Enabled = True
    mnuSD.Enabled = False
    
End Sub
Private Sub mnuSt_Click()
mnuSt.Checked = Not mnuSt.Checked
StarUP = IIf(StarUP = True, False, True)
If StarUP Then
   SetStart False
Else
   SetStart True
End If
End Sub
Private Sub mnuStop_Click()
'Thank to Roger Gilchrist for his suggestion
If isPlaying Then
    CLOSEALLPLAYER
End If
End Sub
Public Sub OPENVIDEO()
    mnuFName_Click
End Sub
Public Sub Play()
    mnuPlay_Click
End Sub
Public Sub SDesktop()
    mnuSD_Click
End Sub
Public Sub StopPlay()
    mnuStop_Click
End Sub
Private Sub tmrLoop_Timer()
If MM.THE_ENDOFSONG(ByMS) Then
   If mnuloop.Checked Then
      mnuPlay_Click
   Else
      mnuStop_Click
   End If
End If
End Sub
