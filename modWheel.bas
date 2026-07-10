Attribute VB_Name = "modWheel"
Option Explicit

' =====================================================================
'  Thread-local WH_GETMESSAGE hook. Handles:
'
'  - Mouse wheel scrolling over code panes (the editor ignores
'    WM_MOUSEWHEEL); Shift+wheel scrolls horizontally.
'  - Find bar keys: Ctrl+F toggle, F3/Shift+F3 next/prev, Esc close.
'  - Ctrl+Tab / Ctrl+Shift+Tab MRU window switcher (frmSwitcher);
'    releasing Ctrl commits, Esc cancels.
'  - Ctrl+P Quick Open fuzzy file/module palette (frmQuickOpen).
'  - Editing shortcuts (modEditOps): Ctrl+D duplicate, Alt+Up/Down
'    move lines, Ctrl+Shift+K delete lines, Ctrl+/ comment toggle,
'    Shift+F12 find all references.
' =====================================================================

Private Const VK_TAB As Long = &H9
Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_ESCAPE As Long = &H1B
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28
Private Const VK_B As Long = &H42
Private Const VK_D As Long = &H44
Private Const VK_F As Long = &H46
Private Const VK_G As Long = &H47
Private Const VK_K As Long = &H4B
Private Const VK_L As Long = &H4C
Private Const VK_O As Long = &H4F
Private Const VK_P As Long = &H50
Private Const VK_F2 As Long = &H71
Private Const VK_F3 As Long = &H72
Private Const VK_F12 As Long = &H7B
Private Const VK_OEM_2 As Long = &HBF     ' the '/' key (US layouts)

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
            ElseIf Scroll_IsHost(H) Then
                ' wheel over a custom list scrolls it even unfocused
                Scroll_Wheel H, HiWordSigned(m.wParam)
                swallow = True
            End If
        Case WM_MOUSEMOVE
            ' scrollbar thumb tracking: the bar's modal loop retrieves
            ' these moves without dispatching them, repainting over our
            ' tick marks. Post a marker; the loop DOES dispatch posted
            ' messages, and it arrives after the thumb repaint.
            If GetCapture() = m.hwnd And Hook_IsHooked(m.hwnd) Then
                If StrComp(WndClass(m.hwnd), "ScrollBar", vbTextCompare) = 0 Then
                    PostMessageA m.hwnd, WM_APP_SBTICKS, 0, 0
                End If
            End If

        Case WM_KEYDOWN
            swallow = HandleKeyDown(m)
        Case WM_SYSKEYDOWN
            swallow = HandleSysKeyDown(m)
        Case WM_KEYUP
            If m.wParam = VK_CONTROL And gSwitcherActive Then
                frmSwitcher.CommitSwitch
            End If
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
    Dim ctrl As Boolean, shift As Boolean
    ctrl = (GetKeyState(VK_CONTROL) < 0)
    shift = (GetKeyState(VK_SHIFT) < 0)

    Select Case m.wParam

    Case VK_TAB
        If gSwitcherActive Then
            frmSwitcher.StepSwitch Not shift
            HandleKeyDown = True
        ElseIf ctrl And FocusInCodePane() Then
            HandleKeyDown = frmSwitcher.BeginSwitch(Not shift)
        End If

    Case VK_ESCAPE
        If gSwitcherActive Then
            frmSwitcher.CancelSwitch
            HandleKeyDown = True
        ElseIf QuickOpenVisible() Then
            frmQuickOpen.Hide
            HandleKeyDown = True
        ElseIf gFindBarVisible And FocusInCodePane() Then
            frmFind.HideBar
            HandleKeyDown = True
        End If

    Case VK_F
        If ctrl Then
            ' claim Ctrl+F from the code pane AND from the bar itself,
            ' or the IDE's accelerator opens its native Find dialog
            If FocusInCodePane() Or FocusInFindBar() Then
                If gFindBarVisible Then
                    frmFind.HideBar
                Else
                    frmFind.ShowBar
                End If
                HandleKeyDown = True
            End If
        End If

    Case VK_F3
        If ctrl And FocusInCodePane() Then
            Edit_HighlightWord            ' Ctrl+F3 = highlight word
            HandleKeyDown = True
        ElseIf gFindBarVisible Then
            frmFind.DoFindPublic Not shift
            HandleKeyDown = True
        End If

    Case VK_F12
        If FocusInCodePane() Then
            If shift Then
                Edit_FindAllReferences
            Else
                Edit_GoToDefinition
            End If
            HandleKeyDown = True
        End If

    Case VK_F2
        If FocusInCodePane() And Not shift Then   ' Shift+F2 stays native
            If ctrl Then
                BM_Toggle
            Else
                BM_NextBookmark
            End If
            HandleKeyDown = True
        End If

    Case VK_O
        If ctrl And shift And FocusInCodePane() Then
            frmBrowser.ShowBrowser
            HandleKeyDown = True
        End If

    Case VK_P
        ' claim Ctrl+P from the IDE (normally Print) for Quick Open.
        ' Unlike the other shortcuts this works anywhere in the IDE
        ' (Project Explorer included), but never in a running program.
        If ctrl And Not shift And FocusInIDE() Then
            If QuickOpenVisible() Then
                frmQuickOpen.Hide
            Else
                frmQuickOpen.ShowQuickOpen
            End If
            HandleKeyDown = True
        End If

    Case VK_G
        If ctrl And shift And FocusInCodePane() Then
            frmChanges.ShowChanges
            HandleKeyDown = True
        End If

    Case VK_B
        If ctrl And shift And FocusInCodePane() Then
            Git_BlameCurrentLine
            HandleKeyDown = True
        End If

    Case VK_L
        If ctrl And shift And FocusInCodePane() Then
            frmGitLog.ShowLog
            HandleKeyDown = True
        End If

    Case VK_D
        If ctrl And Not shift And FocusInCodePane() Then
            Edit_DuplicateLines
            HandleKeyDown = True
        End If

    Case VK_K
        If ctrl And shift And FocusInCodePane() Then
            Edit_DeleteLines
            HandleKeyDown = True
        End If

    Case VK_OEM_2
        If ctrl And FocusInCodePane() Then
            If shift Then
                frmShortcuts.ShowSheet     ' Ctrl+Shift+/ = cheat sheet
            Else
                Edit_ToggleComment
            End If
            HandleKeyDown = True
        End If

    End Select
End Function

' Alt combos arrive as WM_SYSKEYDOWN (context bit 29 of lParam set)
Private Function HandleSysKeyDown(m As MSGSTRUCT) As Boolean
    On Error Resume Next
    If (m.lParam And &H20000000) = 0 Then Exit Function
    If Not FocusInCodePane() Then Exit Function
    Select Case m.wParam
    Case VK_UP
        Edit_MoveLinesUp
        HandleSysKeyDown = True
    Case VK_DOWN
        Edit_MoveLinesDown
        HandleSysKeyDown = True
    End Select
End Function

Private Function QuickOpenVisible() As Boolean
    On Error Resume Next
    QuickOpenVisible = frmQuickOpen.Visible
End Function

' focus anywhere in the IDE main window hierarchy (code panes, docked
' tool windows, owned add-in dialogs) - GetParent also climbs from an
' owned top-level window to its owner. Windows of a program being run
' in the IDE are unowned, so they never satisfy this.
Private Function FocusInIDE() As Boolean
    On Error Resume Next
    Dim H As Long, hMain As Long
    hMain = MainHwnd()
    If hMain = 0 Then Exit Function
    H = GetFocus()
    Do While H <> 0
        If H = hMain Then FocusInIDE = True: Exit Function
        H = GetParent(H)
    Loop
End Function

Private Function FocusInCodePane() As Boolean
    FocusInCodePane = _
        (StrComp(WndClass(GetFocus()), CLS_CODEPANE, vbTextCompare) = 0)
End Function

Private Function FocusInFindBar() As Boolean
    On Error Resume Next
    If Not gFindBarVisible Then Exit Function
    Dim H As Long, hBar As Long
    hBar = frmFind.hwnd
    H = GetFocus()
    Do While H <> 0
        If H = hBar Then FocusInFindBar = True: Exit Function
        H = GetParent(H)
    Loop
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
