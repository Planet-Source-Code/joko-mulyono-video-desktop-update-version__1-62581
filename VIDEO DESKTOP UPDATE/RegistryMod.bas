Attribute VB_Name = "RegistryMod"
Option Explicit

    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Boolean
    End Type
    
    Public Enum Theme
        ClearWall
        StoreWall
        RestoreWall
    End Enum

    #If False Then
    Private ClearWall, StoreWall, RestoreWall
    #End If
    Private Const STANDARD_RIGHTS_ALL        As Long = &H1F0000
    Private Const KEY_QUERY_VALUE            As Long = &H1
    Private Const KEY_SET_VALUE              As Long = &H2
    Private Const KEY_CREATE_SUB_KEY         As Long = &H4
    Private Const KEY_ENUMERATE_SUB_KEYS     As Long = &H8
    Private Const KEY_NOTIFY                 As Long = &H10
    Private Const KEY_CREATE_LINK            As Long = &H20
    Private Const SYNCHRONIZE                As Long = &H100000
    Private Const KEY_ALL_ACCESS             As Double = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
   
    Private Const COLOR_BACKGROUND            As Integer = 1
    Private Const ERROR_SUCCESS = 0
    Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

    Private Const HKEY_CLASSES_ROOT = &H80000000
    Private Const HKEY_CURRENT_CONFIG = &H80000005
    Private Const HKEY_CURRENT_USER = &H80000001
    Private Const HKEY_DYN_DATA = &H80000006
    Private Const HKEY_LOCAL_MACHINE = &H80000002
    Private Const HKEY_PERFORMANCE_DATA = &H80000004
    Private Const HKEY_USERS = &H80000003
    Private Const REG_SZ                      As Long = 1 ' Unicode nul terminated string
    Private Const REG_DWORD                   As Long = 4 ' 32-bit number
    Private Const REG_EXPAND_SZ               As Long = 2 'Unicode nul terminated string
    Private Const ERROR_NONE                  As Integer = 0
    Private Const SPI_SETDESKWALLPAPER        As Integer = 20
    Private Const SPIF_SENDWININICHANGE       As Long = &H2
    Private Const SPIF_UPDATEINIFILE          As Long = &H1
    Private ppt_retWall                       As String
    Private ppt_retStyle                      As String
    Private retOrig                           As String
    Public isChange                           As Boolean
    Public isArrow                            As Boolean
    Private retval                            As Long
    Private lngType                           As Long
    Private lngData                           As String * 255
    Private lngResult                         As Long
    Private lngKey                            As String
'---------------------------------------------------------------
'-Registry API Declarations...
'---------------------------------------------------------------

Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal Reserved As Long, _
                                                                                ByVal lpClass As String, _
                                                                                ByVal dwOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                                                                ByRef phkResult As Long, _
                                                                                ByRef lpdwDisposition As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal lpReserved As Long, _
                                                                                  ByRef lpType As Long, _
                                                                                  ByVal lpData As String, _
                                                                                  ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                            ByVal lpValueName As String, _
                                                                                            ByVal lpReserved As Long, _
                                                                                            lpType As Long, _
                                                                                            ByVal lpData As String, _
                                                                                            lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                          ByVal lpValueName As String, _
                                                                                          ByVal lpReserved As Long, _
                                                                                          lpType As Long, _
                                                                                          lpData As Long, _
                                                                                          lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                          ByVal lpValueName As String, _
                                                                                          ByVal lpReserved As Long, _
                                                                                          lpType As Long, _
                                                                                          ByVal lpData As Long, _
                                                                                          lpcbData As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          ByVal lpvParam As Any, _
                                                                                          ByVal fuWinIni As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, _
                                                                            ByVal lpSubKey As String, _
                                                                            phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal Reserved As Long, _
                                                                                  ByVal dwType As Long, _
                                                                                  ByVal lpData As String, _
                                                                                  ByVal cbData As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                                                                    ByVal lpValueName As String) As Long
Private Declare Function SetSysColors Lib "user32.dll" (ByVal nChanges As Long, _
                                                        lpSysColor As Long, _
                                                        lpColorValues As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub ClearDesktop()
'TESTED:OK
    SetSysColors 1, COLOR_BACKGROUND, RGB(16, 0, 16) 'set magic color transparent
    UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", "" 'remove wallpaper
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, "", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE 'instan Update wallpaper
    isChange = True
End Sub
Public Function getString(Str As String) As String
'Purpose : to extract string from buffer string
Dim a As Integer
    For a = 1 To Len(Str)
        If Mid$(Str, a, 1) = vbNullChar Then
            Exit For
        End If
    Next a
    getString = RTrim$(left(Str, a - 1))
End Function

Public Sub SetStart(ByVal StartUp As Boolean)
'TESTED:OK
    If StartUp Then
    
        UpdateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe C:\WINDOWS\VIDEODESKTOP.exe"
    Else
        UpdateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", ""
    End If
End Sub
Private Sub SetWallPaper(ByVal Display As Integer, _
                         ByVal sdir As String)
'TESTED:OK
Dim NewPaper As String
    NewPaper = sdir
    Select Case Display
    Case 0
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0"
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0"

    Case 1
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "1"
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0"

    Case 2
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0"
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "2"

    End Select
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, NewPaper, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE

End Sub
Public Sub StoreWallpaper()
'TESTED:OK
    On Error Resume Next
    ppt_retWall = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\Desktop\", "Wallpaper") 'get default wallpaper
    ppt_retStyle = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\Desktop\", "WallpaperStyle") 'get default WallpaperStyle
    retOrig = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\Desktop\", "OriginalWallpaper") 'get default Original Wallpaper
    If LenB(ppt_retWall) = 0 Then 'If wallpaper is nothing then ...
        ppt_retWall = retOrig     ' set wallpaper with original wallpaper
    ElseIf ppt_retStyle = Not IsNumeric(ppt_retStyle) Then
        ppt_retStyle = 2 'as default [ stretch ]
    End If
    On Error GoTo 0
End Sub
Public Sub WALLPAPERTHEME(thmOpt As Theme)
'TESTED:OK
    Select Case thmOpt
    Case ClearWall
        StoreWallpaper
        ClearDesktop
    Case StoreWall
        StoreWallpaper
    Case RestoreWall
        If Not IsNumeric(ppt_retStyle) Then
            ppt_retStyle = 2 'as default [ stretch ]
        End If
        SetWallPaper ppt_retStyle, ppt_retWall
'SetSysColors 1, COLOR_BACKGROUND, Form1.Image1.BackColor 'RGB(16, 0, 16)
    End Select
End Sub

Public Function CheckStart() As Boolean
'TESTED:OK
Dim nCek    As String
Dim sBuffer As String * 256
Dim WinDir  As String
Dim xstr    As Long
Dim sKey    As String
xstr = GetWindowsDirectory(sBuffer, Len(sBuffer))
WinDir = getString(left(sBuffer, xstr))
nCek = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell")
sKey = "explorer.exe " & WinDir & "\" & App.EXEName & ".exe" '& Chr(34)

If LenB(nCek) Then
   If nCek = sKey Then '"explorer.exe " & WinDir & "\" & App.EXEName & ".EXE" Then
      CheckStart = True
   Else
      CheckStart = False
   End If
Else
   CheckStart = False
End If
End Function

''Public Function GetWindowsColor() As OLE_COLOR
''GetWindowsColor = GetSysColor(COLOR_BACKGROUND)
''End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
    Dim I As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
      
    tmpVal = left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' Search Data Types...
    Case REG_SZ, REG_EXPAND_SZ                              ' String Registry Key Data Type
        sKeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = sKeyVal                                   ' Return Value
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' Set Return Val To Empty String
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Function UpdateKey(KeyRoot As Long, _
                           KeyName As String, _
                           SubKeyName As String, _
                           SubKeyValue As String) As Boolean
Dim rc     As Long
Dim hKey   As Long
Dim hDepth As Long
Dim lpAttr As SECURITY_ATTRIBUTES

    With lpAttr
        .nLength = 50
        .lpSecurityDescriptor = 0
        .bInheritHandle = True
    End With
    rc = RegCreateKeyEx(KeyRoot, KeyName, 0, REG_SZ, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, hKey, hDepth)
    If rc <> ERROR_SUCCESS Then
        GoTo CreateKeyError
    End If
    If LenB(SubKeyValue) = 0 Then
        SubKeyValue = " "
    End If
    rc = RegSetValueEx(hKey, SubKeyName, 0, REG_SZ, SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
    If rc <> ERROR_SUCCESS Then
        GoTo CreateKeyError
    End If
    rc = RegCloseKey(hKey)
    UpdateKey = True
Exit Function
CreateKeyError:
    UpdateKey = False
    rc = RegCloseKey(hKey)
End Function

Public Sub getArrowData()
'TESTED:OK
lngKey = "lnkfile"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, lngKey, 0, KEY_ALL_ACCESS, lngResult)
    retval = RegQueryValueEx(lngResult, "IsShortcut", 0, lngType, ByVal lngData, 255)
    If retval <> 0 Then
        isArrow = True
    End If
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, "piffile", 0, KEY_ALL_ACCESS, lngResult)
    retval = RegQueryValueEx(lngResult, "IsShortcut", 0, lngType, ByVal lngData, 255)
    If retval <> 0 Then
       isArrow = True
    Else
       isArrow = False
    End If
End Sub

Public Sub setArrow()
'TESTED:OK
lngKey = "lnkfile"
retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, lngKey, 0, KEY_ALL_ACCESS, lngResult)
    If isArrow Then
        retval = RegDeleteValue(lngResult, "IsShortcut")
        retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, "piffile", 0, KEY_ALL_ACCESS, lngResult)
        retval = RegDeleteValue(lngResult, "IsShortcut")
    Else
        retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, "lnkfile", 0, KEY_ALL_ACCESS, lngResult)
        retval = RegSetValueEx(lngResult, "IsShortcut", 0, 1, ByVal "", 1)
        retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, "piffile", 0, KEY_ALL_ACCESS, lngResult)
        retval = RegSetValueEx(lngResult, "IsShortcut", 0, 1, ByVal "", 1)
    End If
End Sub
