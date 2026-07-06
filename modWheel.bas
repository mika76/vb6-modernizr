Attribute VB_Name = "modWheel"
Option Explicit

' =====================================================================
'  Thread-local WH_GETMESSAGE hook, two jobs:
'
'  1. Mouse wheel scrolling: the VB6 editor ignores WM_MOUSEWHEEL, so
'     wheel messages over a code pane become WM_VSCROLL / WM_HSCROLL
'     (Shift held) sent to that pane.
'  2. Find-bar shortcuts while typing in code: Ctrl+F opens the bar,
'     F3 / Shift+F3 = find next/previous, Esc closes the bar.
' =====================================================================

Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_ESCAPE As Long = &H1B
Private Const VK_F3 As Long = &H72
Private Const VK_F As Long = &H46

Private mHook As Long

Public Sub Wheel_Init()
    On Error Resume Next
    If mHook <> 0 Then Exit Sub
    mHook = SetWindowsHookExA(WH_GETMESSAGE, AddressOf WheelGetMsgProc, _
                              0, GetCurrentThreadId())
End Sub

Public Sub Wheel_Term()
    On Error Resume Next
    If mHook <> 0 Then UnhookWindowsHookEx mHook
    mHook = 0
End Sub

Private Function WheelGetMsgProc(ByVal nCode As Long, ByVal wParam As Long, _
        ByVal lParam As Long) As Long
    On Error Resume Next
    If nCode = HC_ACTION And wParam = PM_REMOVE Then
        Dim m As MSGSTRUCT, swallow As Boolean
        CopyMemory m, ByVal lParam, Len(m)

        Select Case m.message
        Case WM_MOUSEWHEEL
            Dim H As Long
            H = WindowFromPoint(m.pt.x, m.pt.y)
            If StrComp(WndClass(H), CLS_CODEPANE, vbTextCompare) = 0 Then
                ScrollPane H, HiWordSigned(m.wParam), (m.wParam And 4) <> 0 ' MK_SHIFT
                swallow = True
            End If
        Case WM_KEYDOWN
            swallow = HandleKeyDown(m)
        End Select

        If swallow Then
            m.message = WM_NULL
            CopyMemory ByVal lParam, m, Len(m)
        End If
    End If
    WheelGetMsgProc = CallNextHookEx(mHook, nCode, wParam, lParam)
End Function

Private Function HandleKeyDown(m As MSGSTRUCT) As Boolean
    On Error Resume Next
    Select Case m.wParam
    Case VK_F
        If GetKeyState(VK_CONTROL) < 0 And FocusInCodePane() Then
            frmFind.ShowBar
            HandleKeyDown = True
        End If
    Case VK_F3
        If gFindBarVisible Then
            frmFind.DoFindPublic (GetKeyState(VK_SHIFT) >= 0)
            HandleKeyDown = True
        End If
    Case VK_ESCAPE
        If gFindBarVisible And FocusInCodePane() Then
            frmFind.HideBar
            HandleKeyDown = True
        End If
    End Select
End Function

Private Function FocusInCodePane() As Boolean
    FocusInCodePane = _
        (StrComp(WndClass(GetFocus()), CLS_CODEPANE, vbTextCompare) = 0)
End Function

Private Sub ScrollPane(ByVal hwnd As Long, ByVal delta As Long, _
        ByVal horizontal As Boolean)
    On Error Resume Next
    Dim lines As Long, n As Long, i As Long, cmd As Long, msg As Long
    SystemParametersInfoA SPI_GETWHEELSCROLLLINES, 0, lines, 0
    If lines < 1 Or lines > 40 Then lines = 3

    n = lines * (Abs(delta) \ 120)
    If n < 1 Then n = 1
    cmd = IIf(delta > 0, SB_LINEUP, SB_LINEDOWN)
    msg = IIf(horizontal, WM_HSCROLL, WM_VSCROLL)

    ' The editor only honors WM_V/HSCROLL when lParam carries its own
    ' scrollbar control's hwnd (same trick as Microsoft's wheel fix).
    Dim hSB As Long
    If horizontal Then
        hSB = FindHScrollBarChild(hwnd)
    Else
        hSB = FindVScrollBarChild(hwnd)
    End If

    For i = 1 To n
        SendMessageA hwnd, msg, cmd, hSB
    Next
End Sub
