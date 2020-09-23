Attribute VB_Name = "CDialogEX"
Option Explicit
Private Const HH_DISPLAY_TEXT_POPUP        As Long = &HE
Private Const TPM_LEFTALIGN                As Long = &H0
Private xhwnd                              As Long
Private PXhwnd                             As Long
Private Xpos                               As Long
Private Ypos                               As Long
Private prevID                             As Long
Private hwnWIDTH                           As Long
Private OldProc                            As Long
Private CLL                                As Long
Private DIALOGHWND                         As Long
Private Const GWL_ID                       As Long = (-12)
Public ExFilename                            As String
Private mFilter                            As String
Private DlgHwnd                            As Long
Type RECT
    Left                                   As Long
    Top                                    As Long
    Right                                  As Long
    Bottom                                 As Long
End Type
Type POINTAPI
    x                                      As Long
    y                                      As Long
End Type
Type CREATESTRUCT
    lpCreateParams                         As Long
    hInstance                              As Long
    hMenu                                  As Long
    hwndParent                             As Long
    cy                                     As Long
    cx                                     As Long
    y                                      As Long
    x                                      As Long
    style                                  As Long
    lpszname                               As String
    lpszClass                              As String
    ExStyle                                As Long
End Type
Type DRAWITEMSTRUCT
    CtlType                                As Long
    CtlID                                  As Long
    itemID                                 As Long
    itemAction                             As Long
    itemState                              As Long
    hwndItem                               As Long
    hdc                                    As Long
    rcItem                                 As RECT
    itemData                               As Long
End Type
Private ParOld                             As Long
Private Const WM_COMMAND                   As Long = &H111
Private Const WM_CONTEXTMENU               As Long = &H7B
Private Const WM_SETTEXT                   As Long = &HC
Private Const WM_DESTROY                   As Long = &H2
Private Const GW_CHILD                     As Integer = 5
Private Const GW_HWNDNEXT                  As Integer = 2
Private Const GWL_STYLE                    As Long = (-16)
Private Const GWL_EXSTYLE                  As Long = (-20)
Private Const WS_EX_TOOLWINDOW             As Long = &H80
' GetOpen/SaveFileName
Private Const WM_INITDIALOG                 As Long = &H110
Private Const SWP_NOSIZE                    As Long = &H1
Private Const SWP_NOZORDER                  As Long = &H4
Private Const SWP_NOACTIVATE                As Long = &H10
Private Const SWP_FRAMECHANGED              As Long = &H20
Private Const SWP_NOMOVE                    As Long = &H2
Private Const MAX_PATH                      As Integer = 260
Public Type OPENFILENAME  '  OFName
    lStructSize                             As Long
    hWndOwner                               As Long
    hInstance                               As Long
    lpstrFilter                             As String
    lpstrCustomFilter                       As String
    nMaxCustFilter                          As Long
    nFilterIndex                            As Long
    lpstrFile                               As String
    nMaxFile                                As Long
    lpstrFileTitle                          As String
    nMaxFileTitle                           As Long
    lpstrInitialDir                         As String
    lpstrTitle                              As String
    Flags                                   As OFN_Flags
    nFileOffset                             As Integer
    nFileExtension                          As Integer
    lpstrDefExt                             As String
    lCustData                               As Long
    lpfnHook                                As Long
    lpTemplateName                          As String
End Type
Private OFName                              As OPENFILENAME
Public Enum OFN_Flags
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000               ' force no long names for 4.x modules
    OFN_EXPLORER = &H80000                       ' new look commdlg
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000                 ' force long names for 3.x modules
    OFN_ENABLEINCLUDENOTIFY = &H400000           ' send include message to callback
    OFN_ENABLESIZING = &H800000
End Enum
#If False Then
Private OFN_READONLY, OFN_OVERWRITEPROMPT, OFN_HIDEREADONLY, OFN_NOCHANGEDIR, OFN_SHOWHELP, OFN_ENABLEHOOK, OFN_ENABLETEMPLATE
Private OFN_ENABLETEMPLATEHANDLE, OFN_NOVALIDATE, OFN_ALLOWMULTISELECT, OFN_EXTENSIONDIFFERENT, OFN_PATHMUSTEXIST, OFN_FILEMUSTEXIST
Private OFN_CREATEPROMPT, OFN_SHAREAWARE, OFN_NOREADONLYRETURN, OFN_NOTESTFILECREATE, OFN_NONETWORKBUTTON, OFN_NOLONGNAMES
Private OFN_EXPLORER, OFN_NODEREFERENCELINKS, OFN_LONGNAMES, OFN_ENABLEINCLUDENOTIFY, OFN_ENABLESIZING
#End If
Private Type HH_POPUP
    cbStruct                                As Long
    hinst                                   As Long
    idString                                As Long
    pszText                                 As Long    ' pointer na string
    pt                                      As POINTAPI
    clrForeground                           As Long
    clrBackground                           As Long
    rcMargins                               As RECT
    pzsFont                                 As Long    ' pointer na string
End Type
Private Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, _
                                                                      ByVal pszFile As String, _
                                                                      ByVal uCommand As Long, _
                                                                      dwData As Any) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, _
                                                                      ByVal wFlags As Long, _
                                                                      ByVal wIDNewItem As Long, _
                                                                      ByVal lpNewItem As Any) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, _
                                                      ByVal wFlags As Long, _
                                                      ByVal x As Long, _
                                                      ByVal y As Long, _
                                                      ByVal nReserved As Long, _
                                                      ByVal hwnd As Long, _
                                                      lprc As RECT) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal Msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, _
                                                  ByVal nIDDlgItem As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                 ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, _
                                                                            ByVal lpString As String) As Long

Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, _
                                                                              ByVal nIDDlgItem As Long, _
                                                                              ByVal lpString As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Public Function CDHOOK(ByVal hDlg As Long, _
                       ByVal uiMsg As Long, _
                       ByVal wParam As Long, _
                       ByVal lParam As Long) As Long


Dim ParRECT As RECT
Dim ctrlX   As Long
Dim newHWND As Long
Dim XY      As POINTAPI
Dim RC1     As RECT
Dim x       As Long
Dim y       As Long
    Select Case uiMsg
    Case WM_INITDIALOG
        newHWND = GetParent(hDlg)
        DIALOGHWND = newHWND
        ctrlX = GetDlgItem(newHWND, &H1)
        CLL = ctrlX
        OldProc = SetWindowLong(ctrlX, -4, AddressOf SETTING)
        SetWindowText CLL, "Kerjakan"
        SetParent xhwnd, newHWND
        prevID = GetDlgCtrlID(xhwnd)
        SetWindowLong xhwnd, GWL_ID, &H6000
        GetWindowRect newHWND, RC1
        XY.x = RC1.Left
        XY.y = RC1.Top
        ScreenToClient newHWND, XY
        GetWindowRect xhwnd, RC1
        MoveWindow xhwnd, XY.x + Xpos, XY.y + Ypos, RC1.Right - RC1.Left, RC1.Bottom - RC1.Top, 1
        ShowWindow xhwnd, 1
        SetDlgItemText newHWND, &H2, "Batalkan"
        SetDlgItemText newHWND, &H443, "Lokasi:"
        SetDlgItemText newHWND, &H442, "Nama file:"
        SetDlgItemText newHWND, &H441, "Tipe file:"
        GetWindowRect newHWND, RC1
        MoveWindow newHWND, 0, 0, RC1.Right - RC1.Left, hwnWIDTH, 1 'Promjeni velicinu prozora
        GetWindowRect newHWND, ParRECT
        x = (Screen.Width / 15 - (ParRECT.Right - ParRECT.Left)) / 2
        y = (Screen.Height / 15 - (ParRECT.Bottom - ParRECT.Top)) / 2
        SetWindowPos newHWND, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        CustimizeDlg hDlg
    Case WM_DESTROY
        ShowWindow xhwnd, 0
        SetParent xhwnd, newHWND
        SetWindowLong xhwnd, GWL_ID, prevID
        SetWindowLong CLL, -4, OldProc
        SetWindowLong newHWND, -4, ParOld
    End Select

End Function

Public Property Get Filter() As String
Dim stemp As String
Dim i As Long
    stemp = mFilter
    For i = 1 To Len(stemp)
        If Mid$(stemp, i, 1) = "|" Then
            Mid$(stemp, i, 1) = vbNullChar
        End If
    Next i
    stemp = stemp & String$(2, 0)
    Filter = stemp

End Property

Public Property Let Filter(ByVal sFilter As String)
mFilter = sFilter
End Property

Public Function GetAddressHOOK(ByVal address As Long) As Long

    GetAddressHOOK = address

End Function

Public Sub INSERT(ByVal nHWND As Long, _
                      ByVal newHWND As Long, _
                      ByVal x As Long, _
                      ByVal y As Long)

    xhwnd = nHWND
    PXhwnd = newHWND
    Xpos = x + 4
    Ypos = y + 23

End Sub

Public Sub CustimizeDlg(hDlg)
Dim sClass As String
Dim h      As Long
Dim k      As Long
Dim rc     As RECT
Dim pt     As POINTAPI
Dim rEdge  As Long

    DlgHwnd = GetParent(hDlg)
    h = GetWindow(DlgHwnd, GW_CHILD)
    WindowStyle DlgHwnd, True, WS_EX_TOOLWINDOW, True, False
 
    Do
        sClass = Space$(128)
        k = GetClassName(h, ByVal sClass, 128)
        sClass = Left$(sClass, k)
        Select Case sClass
        Case "ListBox"
             GetWindowRect h, rc
             pt.x = rc.Left - 5
             pt.y = rc.Top - 32
             rEdge = rc.Right
             ScreenToClient DlgHwnd, pt
             'MoveWindow h, pt.x, pt.y, rEdge - rc.Left + 9, Form1.ScaleHeight - 8, 1
        Case "Edit" ', "Static", "Button", "ToolbarWindow32","ComboBox" ',
             GetWindowRect h, rc
             pt.x = rc.Left ' - 20
             pt.y = rc.Top '- 32
             rEdge = rc.Right - 1
             ScreenToClient DlgHwnd, pt
            'MoveWindow h, pt.x, pt.y, rEdge - rc.Left, rc.Bottom - rc.Top, 1
            '           SendMessage DlgHwnd, CDM_HIDECONTROL, ID_OPEN, ByVal CDlg.OKText 'CDM_SETCONTROLTEXT
            '           SendMessage DlgHwnd, CDM_HIDECONTROL, ID_CANCEL, ByVal CDlg.CancelText 'CDM_SETCONTROLTEXT
            '           SendMessage DlgHwnd, CDM_HIDECONTROL, ID_HELP, ByVal CDlg.HelpText 'CDM_SETCONTROLTEXT
          Case "Button" ', "ToolbarWindow32","ComboBox" ',
             GetWindowRect h, rc
             pt.x = rc.Left + 7
             pt.y = rc.Top '- 32
             rEdge = rc.Right - 1
             ScreenToClient DlgHwnd, pt
             MoveWindow h, pt.x, pt.y, rEdge - rc.Left, rc.Bottom - rc.Top, 1
          
        
        End Select
        
        h = GetWindow(h, GW_HWNDNEXT)
    Loop While h <> 0
   
End Sub

Public Function OpenDialog(ByVal lngHWnd As Long, _
                           ByVal sFilter As String, _
                           ByVal iFilter As Integer, _
                           ByVal sFile As String, _
                           ByVal sInitDir As String, _
                           ByVal sTitle As String, _
                           sRtnPath As String) As Boolean


    With OFName
        .lStructSize = Len(OFName)
        .hWndOwner = lngHWnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrFile = sFile & String$(MAX_PATH - Len(sFile), 0)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = sInitDir
        .lpstrTitle = sTitle
        .Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_EXPLORER Or OFN_ENABLEHOOK
        .lpfnHook = GetAddressHOOK(AddressOf CDHOOK)
    End With
    If GetOpenFileName(OFName) Then
        sRtnPath = Trim$(OFName.lpstrFile)
        If (Asc(Mid$(sRtnPath, Len(sRtnPath), 1))) = 0 Then
            sRtnPath = Mid$(sRtnPath, 1, Len(sRtnPath) - 1)
            ExFilename = sRtnPath
        Else
            ExFilename = sRtnPath
        End If
    Else
        OpenDialog = False
    End If

End Function

Public Function ParentSubclass(ByVal lngHWnd As Long, _
                               ByVal uMsg As Long, _
                               ByVal wParam As Long, _
                               ByVal lParam As Long) As Long


ParentSubclass = CallWindowProc(ParOld, lngHWnd, uMsg, wParam, lParam)

End Function

Public Function SETTING(ByVal lngHWnd As Long, _
                          ByVal uMsg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long) As Long

Dim TTX  As String
Dim TX() As Byte
    
    Select Case uMsg
    Case WM_SETTEXT
        TTX = "Kerjakan" & Chr$(CByte(0))
        TX = StrConv(TTX, vbFromUnicode)
        lParam = VarPtr(TX(0))
    End Select
    SETTING = CallWindowProc(OldProc, lngHWnd, uMsg, wParam, lParam)

End Function

Public Property Let SetWidth(ByVal new_Width As Long)
hwnWIDTH = new_Width
End Property

Public Function WindowStyle(ByVal lngHWnd As Long, _
                            ByVal extended_style As Boolean, _
                            ByVal style_value As Long, _
                            Optional ByVal New_Value As Boolean, _
                            Optional ByVal GetValue As Boolean) As Boolean

Dim style_type As Long
Dim style      As Long

    If extended_style Then
        style_type = GWL_EXSTYLE
    Else
        style_type = GWL_STYLE
    End If
    style = GetWindowLong(lngHWnd, style_type)
    If GetValue Then
        If New_Value Then
            WindowStyle = (style Or style_value)
        Else '
            WindowStyle = (style And Not style_value)
        End If
        Exit Function
    End If
    If New_Value Then
        style = style Or style_value
    Else
        style = style And Not style_value
    End If
    SetWindowLong lngHWnd, style_type, style
    SetWindowPos lngHWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER

End Function


