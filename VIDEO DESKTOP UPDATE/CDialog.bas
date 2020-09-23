Attribute VB_Name = "CDIALOG"
Option Explicit
Private Type OPENFILENAME
    lStructSize                        As Long
    hWndOwner                          As Long
    hInstance                          As Long
    lpstrFilter                        As String
    lpstrCustomFilter                  As String
    nMaxCustFilter                     As Long
    nFilterIndex                       As Long
    lpstrFile                          As String
    nMaxFile                           As Long
    lpstrFileTitle                     As String
    nMaxFileTitle                      As Long
    lpstrInitialDir                    As String
    lpstrTitle                         As String
    flags                              As Long
    nFileOffset                        As Integer
    nFileExtension                     As Integer
    lpstrDefExt                        As String
    lCustData                          As Long
    lpfnHook                           As Long
    lpTemplateName                     As String
End Type
Private Const WM_USER              As Long = &H400
Private Const CDM_FIRST            As Double = WM_USER + 100
Private Const OFN_EXPLORER         As Long = &H80000
Private sInitDir                   As String
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Function OpenDialog(ByVal aHwnd As Long) As String
Dim OFName      As OPENFILENAME
Dim sTemp       As String
    If LenB(sInitDir) = 0 Then
       sInitDir = App.Path
    End If
    With OFName
        .lStructSize = Len(OFName)
        .hWndOwner = aHwnd
        .hInstance = App.hInstance
        .lpstrFilter = "Movie (*.dat;*.mpg;*.avi;*.wmv)" & Chr$(0) & "*.dat;*.mpg;*.avi;*.wmv" & Chr$(0) & "Other Mov(*.mov)" & Chr$(0) & "*.mov" & Chr$(0)
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = sInitDir
        .lpstrTitle = "Open File"
        .flags = OFN_EXPLORER ' Or OFN_HIDEREADONLY
    End With 'OFName
    If GetOpenFileName(OFName) Then
        sTemp = Trim$(OFName.lpstrFile)
        If (Asc(Mid$(sTemp, Len(sTemp), 1))) = 0 Then
            sTemp = Mid$(sTemp, 1, Len(sTemp) - 1)
            OpenDialog = sTemp
        Else
            OpenDialog = sTemp
        End If
    Else
        OpenDialog = ""
    End If
End Function

