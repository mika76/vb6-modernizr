Attribute VB_Name = "modScroll"
Option Explicit

' =====================================================================
'  Native vertical scrollbar for custom-drawn surfaces (ucList, the
'  git log graph). Instead of a VB.VScrollBar control, the host window
'  gets WS_VSCROLL and is driven through SetScrollInfo/GetScrollInfo,
'  so it looks and behaves like every other Windows scrollbar (themed,
'  proportional thumb, auto-hides when everything fits).
'
'  A host registers with Scroll_Attach(hwnd, Me) and must expose:
'      Public Sub ScrollTo(ByVal pos As Long)
'  WM_VSCROLL / WM_MOUSEWHEEL arrive via the hpScrollHost subclass in
'  modSubclass; hover-wheel is routed here by the modWheel hook.
' =====================================================================

Private mHosts As Collection    ' host objects keyed by "h" & hwnd

Public Sub Scroll_Attach(ByVal hwnd As Long, host As Object)
    On Error Resume Next
    If hwnd = 0 Then Exit Sub
    If mHosts Is Nothing Then Set mHosts = New Collection

    Err.Clear
    mHosts.Add host, "h" & hwnd
    If Err.Number <> 0 Then Exit Sub          ' already attached

    Dim style As Long
    style = GetWindowLongA(hwnd, GWL_STYLE)
    SetWindowLongA hwnd, GWL_STYLE, style Or WS_VSCROLL
    SetWindowPos hwnd, 0, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE _
        Or SWP_FRAMECHANGED
    Hook_Window hwnd, hpScrollHost
End Sub

Public Sub Scroll_Detach(ByVal hwnd As Long)
    On Error Resume Next
    Unhook_Window hwnd
    If mHosts Is Nothing Then Exit Sub
    mHosts.Remove "h" & hwnd
End Sub

Public Function Scroll_IsHost(ByVal hwnd As Long) As Boolean
    On Error Resume Next
    If mHosts Is Nothing Then Exit Function
    Dim o As Object
    Err.Clear
    Set o = mHosts("h" & hwnd)
    Scroll_IsHost = (Err.Number = 0)
End Function

' Publish range/page/position. nMax = last item index, nPage = visible
' rows; Windows hides the bar itself when nPage covers the range.
Public Sub Scroll_Update(ByVal hwnd As Long, ByVal nMax As Long, _
        ByVal nPage As Long, ByVal nPos As Long)
    On Error Resume Next
    Dim si As SCROLLINFO
    si.cbSize = Len(si)
    si.fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS
    si.nMin = 0
    si.nMax = nMax
    si.nPage = nPage
    si.nPos = nPos
    SetScrollInfo hwnd, SB_VERT, si, 1
End Sub

' WM_VSCROLL from the subclass; code = LOWORD(wParam)
Public Sub Scroll_OnVScroll(ByVal hwnd As Long, ByVal code As Long)
    On Error Resume Next
    Dim si As SCROLLINFO, pos As Long
    si.cbSize = Len(si)
    si.fMask = SIF_ALL
    GetScrollInfo hwnd, SB_VERT, si

    pos = si.nPos
    Select Case code
        Case SB_LINEUP:      pos = pos - 1
        Case SB_LINEDOWN:    pos = pos + 1
        Case SB_PAGEUP:      pos = pos - si.nPage
        Case SB_PAGEDOWN:    pos = pos + si.nPage
        Case SB_THUMBTRACK, SB_THUMBPOSITION: pos = si.nTrackPos
        Case SB_TOP:         pos = 0
        Case SB_BOTTOM:      pos = si.nMax
        Case Else: Exit Sub                   ' SB_ENDSCROLL etc.
    End Select
    DispatchScrollTo hwnd, ClampPos(pos, si)
End Sub

Public Sub Scroll_Wheel(ByVal hwnd As Long, ByVal delta As Long)
    On Error Resume Next
    Dim si As SCROLLINFO, lines As Long, n As Long
    si.cbSize = Len(si)
    si.fMask = SIF_ALL
    GetScrollInfo hwnd, SB_VERT, si
    If si.nPage = 0 Then Exit Sub             ' no scrollbar shown

    SystemParametersInfoA SPI_GETWHEELSCROLLLINES, 0, lines, 0
    If lines < 1 Or lines > 40 Then lines = 3
    n = lines * (Abs(delta) \ 120)
    If n < 1 Then n = 1
    If delta > 0 Then n = -n

    DispatchScrollTo hwnd, ClampPos(si.nPos + n, si)
End Sub

Private Function ClampPos(ByVal pos As Long, si As SCROLLINFO) As Long
    Dim maxPos As Long
    maxPos = si.nMax - si.nPage + 1
    If maxPos < 0 Then maxPos = 0
    If pos < 0 Then pos = 0
    If pos > maxPos Then pos = maxPos
    ClampPos = pos
End Function

Private Sub DispatchScrollTo(ByVal hwnd As Long, ByVal pos As Long)
    On Error Resume Next
    If mHosts Is Nothing Then Exit Sub
    Dim o As Object
    Set o = mHosts("h" & hwnd)
    If o Is Nothing Then Exit Sub
    o.ScrollTo pos
End Sub
