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
End Enum

Private Const MAX_HOOKS As Long = 128

Private mHwnd(1 To MAX_HOOKS) As Long
Private mOldProc(1 To MAX_HOOKS) As Long
Private mPurpose(1 To MAX_HOOKS) As Long

' Reserved-strip geometry shared with the MDIClient handler
Public gReserveActive As Boolean
Public gReservePx As Long
Private mAdjY As Long           ' last y/cy we produced, to avoid
Private mAdjCY As Long          ' shrinking an already-shrunk rect twice

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
        Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL
            ' scrolling copies pixels, which drags stale highlight
            ' boxes along - repaint the whole pane
            If Highlight_Active() Then InvalidateRect hwnd, 0, 0
        Case WM_DESTROY
            Unhook_Window hwnd
        End Select

    Case hpScrollBar
        SubWndProc = CallWindowProcA(oldProc, hwnd, uMsg, wParam, lParam)
        If uMsg = WM_PAINT Then
            Highlight_PaintScrollbar hwnd
        ElseIf uMsg = WM_DESTROY Then
            Unhook_Window hwnd
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
    If wp.y = mAdjY And wp.cy = mAdjCY Then Exit Sub
    If wp.cy < gReservePx + 120 Then Exit Sub

    wp.y = wp.y + gReservePx
    wp.cy = wp.cy - gReservePx
    mAdjY = wp.y
    mAdjCY = wp.cy
    CopyMemory ByVal lParam, wp, Len(wp)
End Sub

Public Sub ResetMDIAdjustGuard()
    mAdjY = -99999
    mAdjCY = -99999
End Sub
