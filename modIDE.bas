Attribute VB_Name = "modIDE"
Option Explicit

' =====================================================================
'  Shared state, Win32 declarations and small helpers.
' =====================================================================

Public gVBE As VBIDE.VBE

' --- structures -------------------------------------------------------

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type WINDOWPOS
    hwnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    flags As Long
End Type

Public Type MSGSTRUCT
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Type SIZEAPI
    cx As Long
    cy As Long
End Type

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Public Type SCROLLBARINFO
    cbSize As Long
    rcScrollBar As RECT
    dxyLineButton As Long
    xyThumbTop As Long
    xyThumbBottom As Long
    reserved As Long
    rgstate(0 To 5) As Long
End Type

' --- user32 / kernel32 ------------------------------------------------

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Public Declare Function GetClassNameA Lib "user32" _
    (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowTextA Lib "user32" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" _
    (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, _
     ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ScreenToClient Lib "user32" _
    (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SendMessageA Lib "user32" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowLongA Lib "user32" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLongA Lib "user32" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProcA Lib "user32" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" _
    (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function InvalidateRect Lib "user32" _
    (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" _
    (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function SetWindowsHookExA Lib "user32" _
    (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, _
     ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" _
    (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SystemParametersInfoA Lib "user32" _
    (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_HIDE As Long = 0
Public Const SW_SHOWNOACTIVATE As Long = 4

Public Declare Function PostMessageA Lib "user32" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" _
    (ByVal hwnd As Long) As Long

Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
     ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
     ByVal wFlags As Long) As Long
Public Declare Function SetScrollInfo Lib "user32" _
    (ByVal hwnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO, _
     ByVal fRedraw As Long) As Long
Public Declare Function GetScrollInfo Lib "user32" _
    (ByVal hwnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO) As Long
Public Declare Function GetScrollBarInfo Lib "user32" _
    (ByVal hwnd As Long, ByVal idObject As Long, _
     psbi As SCROLLBARINFO) As Long
Public Declare Function DrawIconEx Lib "user32" _
    (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long, _
     ByVal cx As Long, ByVal cy As Long, ByVal istepIfAniCur As Long, _
     ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

' --- shell32 (file icons) ---------------------------------------------

Public Declare Function SHGetFileInfoA Lib "shell32" _
    (ByVal pszPath As String, ByVal dwFileAttributes As Long, _
     psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Public Const SHGFI_ICON As Long = &H100&
Public Const SHGFI_SMALLICON As Long = &H1&
Public Const SHGFI_USEFILEATTRIBUTES As Long = &H10&
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80&
Public Const DI_NORMAL As Long = &H3&

' --- gdi32 ------------------------------------------------------------

Public Declare Function CreatePen Lib "gdi32" _
    (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" _
    (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function Rectangle Lib "gdi32" _
    (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
     ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function CreateFontA Lib "gdi32" _
    (ByVal H As Long, ByVal w As Long, ByVal E As Long, ByVal o As Long, _
     ByVal w2 As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, _
     ByVal c As Long, ByVal op As Long, ByVal cp As Long, ByVal q As Long, _
     ByVal pf As Long, ByVal Face As String) As Long
Public Declare Function GetTextExtentPoint32A Lib "gdi32" _
    (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, _
     lpSize As SIZEAPI) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" _
    (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
     lpPoint As Any) As Long
Public Declare Function LineTo Lib "gdi32" _
    (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" _
    (ByVal hdc As Long, ByVal nBkMode As Long) As Long

' --- advapi32 (read the IDE editor font from the registry) ------------

Private Declare Function RegOpenKeyExA Lib "advapi32" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
     ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
     lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' --- constants ---------------------------------------------------------

Public Const GWL_STYLE As Long = -16
Public Const GWL_WNDPROC As Long = -4
Public Const GWL_HWNDPARENT As Long = -8
Public Const WS_CHILD As Long = &H40000000
Public Const WS_POPUP As Long = &H80000000

Public Const WM_NULL As Long = &H0
Public Const WM_DESTROY As Long = &H2
Public Const WM_PAINT As Long = &HF
Public Const WM_KEYDOWN As Long = &H100
Public Const WM_KEYUP As Long = &H101
Public Const WM_SYSKEYDOWN As Long = &H104
Public Const WM_VSCROLL As Long = &H115
Public Const WM_HSCROLL As Long = &H114
Public Const WM_MOUSEWHEEL As Long = &H20A
Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_TIMER As Long = &H113

' private message (WM_APP range): redraw the scrollbar tick overlay.
' Posted during thumb tracking, where the scrollbar's modal loop eats
' mouse messages but still dispatches posted ones.
Public Const WM_APP_SBTICKS As Long = &H8055&
Public Const WM_WINDOWPOSCHANGING As Long = &H46
Public Const WM_WINDOWPOSCHANGED As Long = &H47

Public Const SB_LINEUP As Long = 0
Public Const SB_LINEDOWN As Long = 1
Public Const SB_PAGEUP As Long = 2
Public Const SB_PAGEDOWN As Long = 3
Public Const SB_THUMBPOSITION As Long = 4
Public Const SB_THUMBTRACK As Long = 5
Public Const SB_TOP As Long = 6
Public Const SB_BOTTOM As Long = 7
Public Const SB_VERT As Long = 1

Public Const SIF_RANGE As Long = &H1
Public Const SIF_PAGE As Long = &H2
Public Const SIF_POS As Long = &H4
Public Const SIF_TRACKPOS As Long = &H10
Public Const SIF_ALL As Long = &H17

Public Const WS_VSCROLL As Long = &H200000
Public Const OBJID_CLIENT As Long = &HFFFFFFFC

Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_FRAMECHANGED As Long = &H20

Public Const WH_GETMESSAGE As Long = 3
Public Const HC_ACTION As Long = 0
Public Const PM_REMOVE As Long = 1

Public Const SM_CYVSCROLL As Long = 20
Public Const SPI_GETWHEELSCROLLLINES As Long = 104
Public Const LOGPIXELSY As Long = 90
Public Const NULL_BRUSH As Long = 5
Public Const PS_SOLID As Long = 0
Public Const PS_DOT As Long = 2
Public Const BKMODE_TRANSPARENT As Long = 1

Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const KEY_READ As Long = &H20019

' Window class of VB6/VBA code editor panes
Public Const CLS_CODEPANE As String = "VbaWindow"
Public Const CLS_MDICLIENT As String = "MDIClient"

' =====================================================================
'  Helpers
' =====================================================================

Public Function WndClass(ByVal hwnd As Long) As String
    Dim s As String, n As Long
    s = Space$(128)
    n = GetClassNameA(hwnd, s, 128)
    If n > 0 Then WndClass = Left$(s, n)
End Function

Public Function WndText(ByVal hwnd As Long) As String
    Dim s As String, n As Long
    s = Space$(512)
    n = GetWindowTextA(hwnd, s, 512)
    If n > 0 Then WndText = Left$(s, n)
End Function

Public Function MainHwnd() As Long
    On Error Resume Next
    MainHwnd = gVBE.MainWindow.hwnd
End Function

Public Function MDIClientHwnd() As Long
    MDIClientHwnd = FindWindowEx(MainHwnd(), 0, CLS_MDICLIENT, vbNullString)
End Function

' Depth-first search for a descendant window of a given class.
Public Function FindDescendantByClass(ByVal hParent As Long, _
        ByVal sClass As String) As Long
    Dim H As Long, r As Long
    H = FindWindowEx(hParent, 0, vbNullString, vbNullString)
    Do While H <> 0
        If StrComp(WndClass(H), sClass, vbTextCompare) = 0 Then
            FindDescendantByClass = H
            Exit Function
        End If
        r = FindDescendantByClass(H, sClass)
        If r <> 0 Then
            FindDescendantByClass = r
            Exit Function
        End If
        H = FindWindowEx(hParent, H, vbNullString, vbNullString)
    Loop
End Function

' IDE window captions carry a "ProjectName - " prefix that disappears
' while the window is active; strip it so captions compare stably.
Public Function NormalizeCaption(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, " - ")
    If p > 0 Then s = Mid$(s, p + 3)
    NormalizeCaption = s
End Function

' Find the MDI child whose title matches the given caption.
Public Function FindMDIChildByCaption(ByVal sCaption As String) As Long
    Dim H As Long, hMDI As Long, want As String
    want = NormalizeCaption(sCaption)
    hMDI = MDIClientHwnd()
    H = FindWindowEx(hMDI, 0, vbNullString, vbNullString)
    Do While H <> 0
        If InStr(1, NormalizeCaption(WndText(H)), want, vbTextCompare) > 0 Then
            FindMDIChildByCaption = H
            Exit Function
        End If
        H = FindWindowEx(hMDI, H, vbNullString, vbNullString)
    Loop
End Function

' Editor window ("VbaWindow") hosting this VBIDE.CodePane. In the VB6
' IDE the MDI child itself is the VbaWindow.
Public Function CodePaneHwnd(ByVal cp As VBIDE.CodePane) As Long
    On Error Resume Next
    Dim hChild As Long
    hChild = FindMDIChildByCaption(cp.Window.Caption)
    If hChild = 0 Then Exit Function
    If StrComp(WndClass(hChild), CLS_CODEPANE, vbTextCompare) = 0 Then
        CodePaneHwnd = hChild
    Else
        CodePaneHwnd = FindDescendantByClass(hChild, CLS_CODEPANE)
    End If
End Function

' The editor's scrollbars are child "ScrollBar" controls; tell them
' apart by shape.
Public Function FindVScrollBarChild(ByVal hParent As Long) As Long
    Dim H As Long, rc As RECT
    H = FindWindowEx(hParent, 0, "ScrollBar", vbNullString)
    Do While H <> 0
        GetWindowRect H, rc
        If (rc.Bottom - rc.Top) > (rc.Right - rc.Left) Then
            FindVScrollBarChild = H
            Exit Function
        End If
        H = FindWindowEx(hParent, H, "ScrollBar", vbNullString)
    Loop
End Function

Public Function FindHScrollBarChild(ByVal hParent As Long) As Long
    Dim H As Long, rc As RECT
    H = FindWindowEx(hParent, 0, "ScrollBar", vbNullString)
    Do While H <> 0
        GetWindowRect H, rc
        If (rc.Right - rc.Left) > (rc.Bottom - rc.Top) Then
            FindHScrollBarChild = H
            Exit Function
        End If
        H = FindWindowEx(hParent, H, "ScrollBar", vbNullString)
    Loop
End Function

' --- registry helpers (IDE editor font) --------------------------------

Public Function RegReadStringHKCU(ByVal sKey As String, ByVal sValue As String, _
        ByVal sDefault As String) As String
    Dim hKey As Long, buf As String, cb As Long, typ As Long
    RegReadStringHKCU = sDefault
    If RegOpenKeyExA(HKEY_CURRENT_USER, sKey, 0, KEY_READ, hKey) = 0 Then
        cb = 256: buf = String$(cb, 0)
        If RegQueryValueExA(hKey, sValue, 0, typ, ByVal buf, cb) = 0 Then
            If cb > 1 Then RegReadStringHKCU = Left$(buf, cb - 1)
        End If
        RegCloseKey hKey
    End If
End Function

Public Function RegReadDwordHKCU(ByVal sKey As String, ByVal sValue As String, _
        ByVal lDefault As Long) As Long
    Dim hKey As Long, dat As Long, cb As Long, typ As Long
    RegReadDwordHKCU = lDefault
    If RegOpenKeyExA(HKEY_CURRENT_USER, sKey, 0, KEY_READ, hKey) = 0 Then
        cb = 4
        If RegQueryValueExA(hKey, sValue, 0, typ, dat, cb) = 0 Then
            RegReadDwordHKCU = dat
        End If
        RegCloseKey hKey
    End If
End Function

Public Function HiWordSigned(ByVal lVal As Long) As Long
    Dim H As Long
    H = (lVal And &HFFFF0000) \ &H10000
    If (H And &H8000&) <> 0 Then H = H Or &HFFFF0000
    HiWordSigned = H
End Function
