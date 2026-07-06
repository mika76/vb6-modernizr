Attribute VB_Name = "modHighlight"
Option Explicit

' =====================================================================
'  "Highlight All" overlay.
'
'  The VB6 editor exposes no highlight API, so after the code pane
'  paints we draw outline boxes over every match that is currently
'  visible, and tick marks on the vertical scrollbar at each match's
'  relative position.
'
'  Geometry comes from the extensibility model (TopLine /
'  CountOfVisibleLines) plus the editor font read from the registry.
'  Horizontal scrolling is not detectable, so boxes assume column 1
'  is at the left edge (fine for the overwhelmingly common case).
' =====================================================================

Private Const HL_COLOR As Long = &H157DE9      ' RGB(233,125,21) - orange
Private Const MARK_COLOR As Long = &H157DE9

Private mHL() As MatchInfo
Private mHLCount As Long
Private mFindLen As Long

' Cached editor font cell size (pixels)
Private mCharW As Long
Private mLineHFont As Long

Public Sub Highlight_SetFromSearch()
    On Error Resume Next
    Highlight_ClearInternal
    If gMatchCount = 0 Then Exit Sub
    mHL = gMatches
    mHLCount = gMatchCount
    Highlight_EnsureHooks
    Highlight_InvalidateAll
End Sub

Public Sub Highlight_Clear()
    On Error Resume Next
    Highlight_ClearInternal
    Highlight_InvalidateAll
End Sub

Private Sub Highlight_ClearInternal()
    mHLCount = 0
    ReDim mHL(0 To 0)
End Sub

Public Sub Highlight_Terminate()
    On Error Resume Next
    Highlight_ClearInternal
    ' Hooks are torn down by Unhook_All at disconnect.
End Sub

Public Function Highlight_Active() As Boolean
    Highlight_Active = (mHLCount > 0)
End Function

' Hook every open code pane (and its vertical scrollbar) that has
' matches. Called after Highlight All and from the tab-bar timer so
' newly opened panes get hooked too.
Public Sub Highlight_EnsureHooks()
    On Error Resume Next
    If mHLCount = 0 Then Exit Sub
    Dim cp As VBIDE.CodePane, H As Long
    For Each cp In gVBE.CodePanes
        H = CodePaneHwnd(cp)
        If H <> 0 Then
            If Not Hook_IsHooked(H) Then
                Hook_Window H, hpCodePane
                HookPaneScrollbar H
            End If
        End If
    Next
End Sub

Private Sub HookPaneScrollbar(ByVal hPane As Long)
    On Error Resume Next
    ' The editor's scrollbars are child "ScrollBar" controls of the
    ' VbaWindow itself.
    Dim hSB As Long
    hSB = FindVScrollBar(hPane)
    If hSB <> 0 Then
        If Not Hook_IsHooked(hSB) Then Hook_Window hSB, hpScrollBar
    End If
End Sub

Private Function FindVScrollBar(ByVal hParent As Long) As Long
    Dim H As Long, rc As RECT
    H = FindWindowEx(hParent, 0, "ScrollBar", vbNullString)
    Do While H <> 0
        GetWindowRect H, rc
        If (rc.Bottom - rc.Top) > (rc.Right - rc.Left) Then  ' vertical
            FindVScrollBar = H
            Exit Function
        End If
        H = FindWindowEx(hParent, H, "ScrollBar", vbNullString)
    Loop
End Function

Public Sub Highlight_InvalidateAll()
    On Error Resume Next
    Dim cp As VBIDE.CodePane, H As Long
    For Each cp In gVBE.CodePanes
        H = CodePaneHwnd(cp)
        If H <> 0 Then
            InvalidateRect H, 0, 1
            Dim hSB As Long
            hSB = FindVScrollBar(H)
            If hSB <> 0 Then InvalidateRect hSB, 0, 0
        End If
    Next
End Sub

' ---------------------------------------------------------------------
'  Painting (called from the subclass proc after the default paint)
' ---------------------------------------------------------------------

Public Sub Highlight_PaintPane(ByVal hwnd As Long)
    On Error Resume Next
    If mHLCount = 0 Then Exit Sub

    Dim cp As VBIDE.CodePane
    Set cp = PaneFromHwnd(hwnd)
    If cp Is Nothing Then Exit Sub

    Dim compName As String, projName As String
    compName = cp.CodeModule.Parent.Name
    projName = cp.CodeModule.Parent.Collection.Parent.Name

    Dim rc As RECT
    GetClientRect hwnd, rc

    ' the object/procedure combo header sits inside the client area
    Dim yTop As Long
    yTop = EditorTopOffset(hwnd)

    Dim visLines As Long, topLine As Long, lineH As Long
    visLines = cp.CountOfVisibleLines
    topLine = cp.topLine
    If visLines < 1 Then Exit Sub
    lineH = (rc.Bottom - rc.Top - yTop) \ visLines
    If lineH < 4 Then Exit Sub

    Dim hdc As Long
    hdc = GetDC(hwnd)
    If hdc = 0 Then Exit Sub

    EnsureFontMetrics hdc, lineH
    Dim marginPx As Long
    marginPx = lineH + ScaleForDpi(6)   ' approx. indicator margin width

    Dim hPen As Long, hOldPen As Long, hOldBr As Long
    hPen = CreatePen(PS_SOLID, 1, HL_COLOR)
    hOldPen = SelectObject(hdc, hPen)
    hOldBr = SelectObject(hdc, GetStockObject(NULL_BRUSH))

    Dim i As Long, x As Long, y As Long
    For i = 0 To mHLCount - 1
        If mHL(i).Comp = compName And mHL(i).Proj = projName Then
            If mHL(i).LineNum >= topLine And _
               mHL(i).LineNum < topLine + visLines Then
                y = yTop + (mHL(i).LineNum - topLine) * lineH
                x = marginPx + (mHL(i).Col - 1) * mCharW
                Rectangle hdc, x - 1, y, _
                          x + mHL(i).MatchLen * mCharW + 1, y + lineH
            End If
        End If
    Next

    SelectObject hdc, hOldPen
    SelectObject hdc, hOldBr
    DeleteObject hPen
    ReleaseDC hwnd, hdc
End Sub

Public Sub Highlight_PaintScrollbar(ByVal hwnd As Long)
    On Error Resume Next
    If mHLCount = 0 Then Exit Sub

    ' The scrollbar is a child of the VbaWindow editor.
    Dim hPane As Long
    hPane = GetParent(hwnd)
    If StrComp(WndClass(hPane), CLS_CODEPANE, vbTextCompare) <> 0 Then Exit Sub

    Dim cp As VBIDE.CodePane
    Set cp = PaneFromHwnd(hPane)
    If cp Is Nothing Then Exit Sub

    Dim compName As String, projName As String, totalLines As Long
    compName = cp.CodeModule.Parent.Name
    projName = cp.CodeModule.Parent.Collection.Parent.Name
    totalLines = cp.CodeModule.CountOfLines
    If totalLines < 1 Then Exit Sub

    Dim rc As RECT
    GetClientRect hwnd, rc
    Dim arrowH As Long, trackH As Long
    arrowH = GetSystemMetrics(SM_CYVSCROLL)
    trackH = (rc.Bottom - rc.Top) - 2 * arrowH
    If trackH < 10 Then Exit Sub

    Dim hdc As Long
    hdc = GetDC(hwnd)
    If hdc = 0 Then Exit Sub

    Dim hBr As Long
    hBr = CreateSolidBrush(MARK_COLOR)
    Dim hPen As Long, hOldPen As Long, hOldBr As Long
    hPen = CreatePen(PS_SOLID, 1, MARK_COLOR)
    hOldPen = SelectObject(hdc, hPen)
    hOldBr = SelectObject(hdc, hBr)

    Dim i As Long, y As Long
    For i = 0 To mHLCount - 1
        If mHL(i).Comp = compName And mHL(i).Proj = projName Then
            y = arrowH + CLng((mHL(i).LineNum - 1) / totalLines * trackH)
            Rectangle hdc, 2, y, (rc.Right - rc.Left) - 2, y + 3
        End If
    Next

    SelectObject hdc, hOldPen
    SelectObject hdc, hOldBr
    DeleteObject hPen
    DeleteObject hBr
    ReleaseDC hwnd, hdc
End Sub

' ---------------------------------------------------------------------

' Bottom edge of the object/procedure combos, in pane client coords.
Private Function EditorTopOffset(ByVal hPane As Long) As Long
    On Error Resume Next
    Dim H As Long, rc As RECT, pt As POINTAPI, bottomMax As Long
    H = FindWindowEx(hPane, 0, "ComboBox", vbNullString)
    Do While H <> 0
        If IsWindowVisible(H) Then
            GetWindowRect H, rc
            pt.x = rc.Left: pt.y = rc.Bottom
            ScreenToClient hPane, pt
            If pt.y > bottomMax Then bottomMax = pt.y
        End If
        H = FindWindowEx(hPane, H, "ComboBox", vbNullString)
    Loop
    If bottomMax > 0 Then EditorTopOffset = bottomMax + 2
End Function

Private Function PaneFromHwnd(ByVal hwnd As Long) As VBIDE.CodePane
    On Error Resume Next
    Dim cp As VBIDE.CodePane
    For Each cp In gVBE.CodePanes
        If CodePaneHwnd(cp) = hwnd Then
            Set PaneFromHwnd = cp
            Exit Function
        End If
    Next
End Function

' Measure the editor font (from the VBA/VB6 registry settings) once
' per line-height so column -> pixel math is possible.
Private Sub EnsureFontMetrics(ByVal hdc As Long, ByVal lineH As Long)
    On Error Resume Next
    If mCharW > 0 And mLineHFont = lineH Then Exit Sub

    Dim face As String, pts As Long
    face = RegReadStringHKCU("Software\Microsoft\VBA\Microsoft Visual Basic", _
                             "FontFace", "Courier New")
    pts = RegReadDwordHKCU("Software\Microsoft\VBA\Microsoft Visual Basic", _
                           "FontHeight", 10)
    If pts < 6 Or pts > 72 Then pts = 10

    Dim hFont As Long, hOld As Long, sz As SIZEAPI
    hFont = CreateFontA(-MulDivPts(pts, hdc), 0, 0, 0, 400, 0, 0, 0, _
                        0, 0, 0, 0, 0, face)
    If hFont <> 0 Then
        hOld = SelectObject(hdc, hFont)
        GetTextExtentPoint32A hdc, "M", 1, sz
        SelectObject hdc, hOld
        DeleteObject hFont
        If sz.cx > 0 Then
            mCharW = sz.cx
            mLineHFont = lineH
        End If
    End If
    If mCharW <= 0 Then mCharW = lineH * 6 \ 10   ' rough fallback
End Sub

Private Function MulDivPts(ByVal pts As Long, ByVal hdc As Long) As Long
    Dim dpi As Long
    dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    If dpi <= 0 Then dpi = 96
    MulDivPts = pts * dpi \ 72
End Function
