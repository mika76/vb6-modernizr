Attribute VB_Name = "modTabs"
Option Explicit

' =====================================================================
'  Layout manager for the strips docked above the MDI client:
'  the tab bar and the find bar. Space is reserved by adjusting the
'  MDI client during the IDE's own layout passes (see modSubclass).
' =====================================================================

Public Const TAB_BAR_HEIGHT As Long = 26     ' px at 96 dpi
Public Const FIND_BAR_HEIGHT As Long = 60    ' px at 96 dpi

Public gTabBarVisible As Boolean
Public gFindBarVisible As Boolean

Private mOldReserve As Long
Private mHooked As Boolean

' ---------------------------------------------------------------------

Public Sub TabBar_Init()
    gTabBarVisible = True
    Layout_Update
End Sub

Public Sub TabBar_Term()
    On Error Resume Next
    gTabBarVisible = False
    gFindBarVisible = False
    Layout_Update
    If mHooked Then
        Unhook_Window MDIClientHwnd()
        mHooked = False
    End If
    Unload frmTabs
End Sub

Public Sub TabBar_Toggle()
    gTabBarVisible = Not gTabBarVisible
    Layout_Update
End Sub

Public Sub FindBar_Show()
    gFindBarVisible = True
    Layout_Update
End Sub

Public Sub FindBar_Hide()
    gFindBarVisible = False
    Layout_Update
End Sub

' ---------------------------------------------------------------------

Public Sub Layout_Update()
    On Error Resume Next
    Dim hMDI As Long, newReserve As Long
    hMDI = MDIClientHwnd()
    If hMDI = 0 Then Exit Sub

    If Not mHooked Then
        If Not Hook_Window(hMDI, hpMDIClient) Then Exit Sub
        mHooked = True
    End If

    If gTabBarVisible Then newReserve = ScaleForDpi(TAB_BAR_HEIGHT)
    If gFindBarVisible Then newReserve = newReserve + ScaleForDpi(FIND_BAR_HEIGHT)

    ' Give the old strip back to the MDI client, then reserve the new one.
    gReserveActive = False
    If mOldReserve > 0 Then ExpandMDI hMDI, mOldReserve
    gReservePx = newReserve
    mOldReserve = newReserve
    If newReserve > 0 Then
        gReserveActive = True
        ResetMDIAdjustGuard
        NudgeMDI hMDI
    End If

    If gTabBarVisible Then
        frmTabs.Attach
        frmTabs.Reposition
        frmTabs.Visible = True
    ElseIf IsFormLoaded("frmTabs") Then
        frmTabs.Visible = False
    End If

    If gFindBarVisible Then
        frmFind.Attach
        frmFind.Reposition
        frmFind.Visible = True
    ElseIf IsFormLoaded("frmFind") Then
        frmFind.Visible = False
    End If
End Sub

' Called from the MDIClient subclass whenever the IDE re-lays-out.
Public Sub Layout_Reposition()
    On Error Resume Next
    If gTabBarVisible Then frmTabs.Reposition
    If gFindBarVisible Then frmFind.Reposition
End Sub

' Top of the reserved strip, in main-window client coordinates.
' Returns False if the MDI client can't be found.
Public Function Layout_StripOrigin(x As Long, y As Long, w As Long) As Boolean
    On Error Resume Next
    Dim hMDI As Long, rc As RECT, pt As POINTAPI
    hMDI = MDIClientHwnd()
    If hMDI = 0 Then Exit Function
    GetWindowRect hMDI, rc
    pt.x = rc.Left: pt.y = rc.Top
    ScreenToClient MainHwnd(), pt
    x = pt.x
    y = pt.y - gReservePx
    w = rc.Right - rc.Left
    Layout_StripOrigin = True
End Function

' ---------------------------------------------------------------------

Private Sub ExpandMDI(ByVal hMDI As Long, ByVal reserve As Long)
    On Error Resume Next
    Dim rc As RECT, pt As POINTAPI
    GetWindowRect hMDI, rc
    pt.x = rc.Left: pt.y = rc.Top
    ScreenToClient GetParent(hMDI), pt
    MoveWindow hMDI, pt.x, pt.y - reserve, rc.Right - rc.Left, _
               rc.Bottom - rc.Top + reserve, 1
End Sub

' Re-apply the MDI client's current rect so the subclass proc sees a
' WM_WINDOWPOSCHANGING it can adjust.
Private Sub NudgeMDI(ByVal hMDI As Long)
    On Error Resume Next
    Dim rc As RECT, pt As POINTAPI
    GetWindowRect hMDI, rc
    pt.x = rc.Left: pt.y = rc.Top
    ScreenToClient GetParent(hMDI), pt
    MoveWindow hMDI, pt.x, pt.y, rc.Right - rc.Left, rc.Bottom - rc.Top, 1
End Sub

Private Function IsFormLoaded(ByVal nm As String) As Boolean
    Dim f As Form
    For Each f In Forms
        If f.Name = nm Then IsFormLoaded = True: Exit Function
    Next
End Function

Public Function ScaleForDpi(ByVal px As Long) As Long
    Dim hdc As Long, dpi As Long
    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
    If dpi <= 0 Then dpi = 96
    ScaleForDpi = px * dpi \ 96
End Function
