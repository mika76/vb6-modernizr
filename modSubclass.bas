Attribute VB_Name = "modSubclass"
Option Explicit

' =====================================================================
'  Minimal, purpose-tagged window subclassing.
'
'  Every hooked window shares one WndProc; behavior is selected by the
'  purpose recorded at hook time. All handlers run with On Error
'  Resume Next: an unhandled error inside a WndProc kills the IDE.
' =====================================================================

Public Enum HookPurpose
    hpMDIClient = 1     ' reserve space for the docked bars
    hpCodePane = 2      ' paint match highlights after WM_PAINT
    hpScrollBar = 3     ' paint match marks on the vertical scrollbar
    hpScrollHost = 4    ' native WS_VSCROLL host (see modScroll)
End Enum

Private Const MAX_HOOKS As Long = 128

Private mHwnd(1 To MAX_HOOKS) As Long
Private mOldProc(1 To MAX_HOOKS) As Long
Private mPurpose(1 To MAX_HOOKS) As Long

' Reserved-strip geometry shared with the MDIClient handler
Public gReserveActive As Boolean
Public gReservePx As Long       ' strip above the MDI client (bars)
Public gReserveLeftPx As Long   ' strip left of the MDI client (gutter)
Private mAdjY As Long           ' last rect we produced, to avoid
Private mAdjCY As Long          ' shrinking an already-shrunk rect twice
Private mAdjX As Long
Private mAdjCX As Long

' ---------------------------------------------------------------------

Public Function Hook_Window(ByVal hwnd As Long, ByVal purpose As HookPurpose) As Boolean
    Dim i As Long, slot As Long
    If hwnd = 0 Then Exit Function
    If IsWindow(hwnd) = 0 Then Exit Function
    For i = 1 To MAX_HOOKS
        If mHwnd(i) = hwnd Then Hook_Window = True: Exit Function ' already hooked
        If slot = 0 And mHwnd(i) = 0 Then slot = i
    Next
    If slot = 0 Then Exit Function
    mHwnd(slot) = hwnd
    mPurpose(slot) = purpose
    mOldProc(slot) = SetWindowLongA(hwnd, GWL_WNDPROC, AddressOf SubWndProc)
    Hook_Window = (mOldProc(slot) <> 0)
    If Not Hook_Window Then mHwnd(slot) = 0
End Function

Public Sub Unhook_Window(ByVal hwnd As Long)
    Dim i As Long
    For i = 1 To MAX_HOOKS
        If mHwnd(i) = hwnd Then
            If IsWindow(hwnd) Then SetWindowLongA hwnd, GWL_WNDPROC, mOldProc(i)
            mHwnd(i) = 0: mOldProc(i) = 0: mPurpose(i) = 0
            Exit Sub
        End If
    Next
End Sub

Public Sub Unhook_All()
    Dim i As Long
    For i = 1 To MAX_HOOKS
        If mHwnd(i) <> 0 Then
            If IsWindow(mHwnd(i)) Then SetWindowLongA mHwnd(i), GWL_WNDPROC, mOldProc(i)
            mHwnd(i) = 0: mOldProc(i) = 0: mPurpose(i) = 0
        End If
    Next
End Sub

Public Function Hook_IsHooked(ByVal hwnd As Long) As Boolean
    Dim i As Long
    For i = 1 To MAX_HOOKS
        If mHwnd(i) = hwnd Then Hook_IsHooked = True: Exit Function
    Next
End Function

' ---------------------------------------------------------------------

Private Function SubWndProc(ByVal hwnd As Long, ByVal uMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim i As Long, idx As Long
    For i = 1 To MAX_HOOKS
        If mHwnd(i) = hwnd Then idx = i: Exit For
    Next
    If idx = 0 Then
        SubWndProc = 0
        Exit Function
    End If

    Dim oldProc As Long
    oldProc = mOldProc(idx)

    Select Case mPurpose(idx)

    Case hpMDIClient
        If uMsg = WM_WINDOWPOSCHANGING And gReserveActive Then
            AdjustMDIPos lParam
        End If
        SubWndProc = CallWindowProcA(oldProc, hwnd, uMsg, wParam, lParam)
        If uMsg = WM_WINDOWPOSCHANGED And gReserveActive Then
            Layout_Reposition
        ElseIf uMsg = WM_DESTROY Then
            Unhook_Window hwnd
        End If

    Case hpCodePane
        SubWndProc = CallWindowProcA(oldProc, hwnd, uMsg, wParam, lParam)
        Select Case uMsg
        Case WM_PAINT
            Highlight_PaintPane hwnd
            If gLineNumsEnabled Then frmGutter.Poll
        Case WM_VSCROLL, WM_MOUSEWHEEL
            ' vertical scrolling copies pixels, and everything we draw
            ' is line-anchored, so it rides along correctly; drawing
            ' the overlays again on top is idempotent and avoids the
            ' erase+repaint flicker a full invalidate would cause.
            ' The scrollbar repaints its thumb directly (no WM_PAINT),
            ' erasing the tick marks - redraw those directly too.
            If Highlight_Active() Then
                Highlight_PaintPane hwnd
                Dim hSB As Long
                hSB = FindVScrollBarChild(hwnd)
                If hSB <> 0 Then Highlight_PaintScrollbar hSB
            End If
            If gLineNumsEnabled Then frmGutter.Poll
        Case WM_HSCROLL
            ' horizontal scrolling breaks the column-1 assumption the
            ' boxes are placed with - here a full repaint is needed
            If Highlight_Active() Then InvalidateRect hwnd, 0, 0
        Case WM_DESTROY
            Unhook_Window hwnd
        End Select

    Case hpScrollBar
        SubWndProc = CallWindowProcA(oldProc, hwnd, uMsg, wParam, lParam)
        Select Case uMsg
        Case WM_PAINT
            Highlight_PaintScrollbar hwnd
        Case WM_MOUSEMOVE, WM_LBUTTONUP, WM_TIMER, WM_APP_SBTICKS
            ' these repaint the bar directly (no WM_PAINT) - put the
            ' tick marks back each time. WM_APP_SBTICKS is posted by
            ' the modWheel hook during thumb tracking, whose modal
            ' loop swallows mouse moves but dispatches posted messages
            Highlight_PaintScrollbar hwnd
        Case WM_DESTROY
            Unhook_Window hwnd
        End Select

    Case hpScrollHost
        If uMsg = WM_VSCROLL Then
            Scroll_OnVScroll hwnd, wParam And &HFFFF&
            SubWndProc = 0
        ElseIf uMsg = WM_MOUSEWHEEL Then
            Scroll_Wheel hwnd, HiWordSigned(wParam)
            SubWndProc = 0
        Else
            SubWndProc = CallWindowProcA(oldProc, hwnd, uMsg, wParam, lParam)
            If uMsg = WM_DESTROY Then Scroll_Detach hwnd
        End If

    Case Else
        SubWndProc = CallWindowProcA(oldProc, hwnd, uMsg, wParam, lParam)
    End Select
End Function

' Shift the MDI client down by the reserved height whenever the IDE
' lays it out, so the docked bars own a strip at the top.
Private Sub AdjustMDIPos(ByVal lParam As Long)
    On Error Resume Next
    Dim wp As WINDOWPOS
    CopyMemory wp, ByVal lParam, Len(wp)
    If (wp.flags And SWP_NOSIZE) <> 0 Then Exit Sub
    If (wp.flags And SWP_NOMOVE) <> 0 Then Exit Sub
    ' Skip if this is the rect we already adjusted (someone re-applied
    ' the current position) - prevents cumulative shrinking.
    If wp.y = mAdjY And wp.cy = mAdjCY And _
       wp.x = mAdjX And wp.cx = mAdjCX Then Exit Sub
    If wp.cy < gReservePx + 120 Then Exit Sub
    If wp.cx < gReserveLeftPx + 120 Then Exit Sub

    wp.y = wp.y + gReservePx
    wp.cy = wp.cy - gReservePx
    wp.x = wp.x + gReserveLeftPx
    wp.cx = wp.cx - gReserveLeftPx
    mAdjY = wp.y
    mAdjCY = wp.cy
    mAdjX = wp.x
    mAdjCX = wp.cx
    CopyMemory ByVal lParam, wp, Len(wp)
End Sub

Public Sub ResetMDIAdjustGuard()
    mAdjY = -99999
    mAdjCY = -99999
    mAdjX = -99999
    mAdjCX = -99999
End Sub
