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
Private Const BM_COLOR As Long = &HD77800      ' RGB(0,120,215) - blue
Private Const GUIDE_COLOR As Long = &HD8D8D8   ' light gray

' indentation guides (toggle persists via SaveSetting)
Public gGuidesEnabled As Boolean
Private mTabW As Long

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
    Highlight_Active = (mHLCount > 0 Or BM_Count() > 0 Or _
                        Git_MarkCount() > 0 Or gGuidesEnabled)
End Function

Public Sub Guides_Init()
    On Error Resume Next
    gGuidesEnabled = (GetSetting("VB6Modernizr", "Options", "Guides", "0") = "1")
    mTabW = RegReadDwordHKCU("Software\Microsoft\VBA\Microsoft Visual Basic", _
                             "TabWidth", 4)
    If mTabW < 2 Or mTabW > 16 Then mTabW = 4
End Sub

Public Sub Guides_Toggle()
    On Error Resume Next
    gGuidesEnabled = Not gGuidesEnabled
    SaveSetting "VB6Modernizr", "Options", "Guides", _
                IIf(gGuidesEnabled, "1", "0")
    Highlight_EnsureHooks
    Highlight_InvalidateAll
End Sub

' Hook every open code pane (and its vertical scrollbar) that has
' matches. Called after Highlight All and from the tab-bar timer so
' newly opened panes get hooked too.
Public Sub Highlight_EnsureHooks()
    On Error Resume Next
    If Not Highlight_Active() Then Exit Sub
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
    hSB = FindVScrollBarChild(hPane)
    If hSB <> 0 Then
        If Not Hook_IsHooked(hSB) Then Hook_Window hSB, hpScrollBar
    End If
End Sub

Public Sub Highlight_InvalidateAll()
    On Error Resume Next
    Dim cp As VBIDE.CodePane, H As Long
    For Each cp In gVBE.CodePanes
        H = CodePaneHwnd(cp)
        If H <> 0 Then
            InvalidateRect H, 0, 1
            Dim hSB As Long
            hSB = FindVScrollBarChild(H)
            If hSB <> 0 Then InvalidateRect hSB, 0, 0
        End If
    Next
End Sub

' ---------------------------------------------------------------------
'  Painting (called from the subclass proc after the default paint)
' ---------------------------------------------------------------------

Public Sub Highlight_PaintPane(ByVal hwnd As Long)
    On Error Resume Next
    If Not Highlight_Active() Then Exit Sub

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

    ' in Procedure View only the current proc's lines are displayed
    Dim loLine As Long, hiLine As Long
    GetDisplayedRange cp, loLine, hiLine

    If gGuidesEnabled Then
        PaintGuides hdc, cp, topLine, visLines, loLine, hiLine, _
                    lineH, yTop, marginPx
    End If

    Dim i As Long, y As Long
    If mHLCount > 0 Then
        Dim hPen As Long, hOldPen As Long, hOldBr As Long
        hPen = CreatePen(PS_SOLID, 1, HL_COLOR)
        hOldPen = SelectObject(hdc, hPen)
        hOldBr = SelectObject(hdc, GetStockObject(NULL_BRUSH))

        Dim x1 As Long, x2 As Long, halfW As Long
        halfW = mCharW \ 2
        For i = 0 To mHLCount - 1
            If mHL(i).Comp = compName And mHL(i).Proj = projName Then
                If mHL(i).LineNum >= topLine And _
                   mHL(i).LineNum < topLine + visLines And _
                   mHL(i).LineNum >= loLine And mHL(i).LineNum <= hiLine Then
                    y = yTop + (mHL(i).LineNum - topLine) * lineH + ScaleForDpi(3)
                    ' margin estimate runs one cell short in practice,
                    ' hence Col instead of Col - 1; the right edge gets
                    ' an extra half cell so the last char isn't bisected
                    x1 = marginPx + mHL(i).Col * mCharW
                    x2 = x1 + mHL(i).MatchLen * mCharW + halfW
                    Rectangle hdc, x1 - 1, y, x2 + 1, y + lineH
                End If
            End If
        Next

        SelectObject hdc, hOldPen
        SelectObject hdc, hOldBr
        DeleteObject hPen
    End If

    ' bookmarks: filled blue squares in the left margin
    Dim bmLines() As Long, nBM As Long
    nBM = BM_LinesForComp(compName, bmLines)
    If nBM > 0 Then
        Dim hBmBr As Long, hBmPen As Long, hOldP2 As Long, hOldB2 As Long
        hBmBr = CreateSolidBrush(BM_COLOR)
        hBmPen = CreatePen(PS_SOLID, 1, BM_COLOR)
        hOldP2 = SelectObject(hdc, hBmPen)
        hOldB2 = SelectObject(hdc, hBmBr)
        For i = 0 To nBM - 1
            If bmLines(i) >= topLine And bmLines(i) < topLine + visLines And _
               bmLines(i) >= loLine And bmLines(i) <= hiLine Then
                y = yTop + (bmLines(i) - topLine) * lineH
                Rectangle hdc, 2, y + 3, ScaleForDpi(10), y + lineH - 1
            End If
        Next
        SelectObject hdc, hOldP2
        SelectObject hdc, hOldB2
        DeleteObject hBmPen
        DeleteObject hBmBr
    End If

    ' git changed-line bars at the right edge of the margin:
    ' green = added, blue = modified, red = deletion below this line
    Dim gitLines() As Long, gitKinds() As Long, nGit As Long
    nGit = Git_MarksForComp(compName, gitLines, gitKinds)
    If nGit > 0 Then
        For i = 0 To nGit - 1
            If gitLines(i) >= topLine And gitLines(i) < topLine + visLines And _
               gitLines(i) >= loLine And gitLines(i) <= hiLine Then
                y = yTop + (gitLines(i) - topLine) * lineH
                DrawGitBar hdc, marginPx - 5, y, marginPx - 1, y + lineH, _
                           gitKinds(i)
            End If
        Next
    End If

    ReleaseDC hwnd, hdc
End Sub

' Dotted vertical lines at each indent step, drawn only inside the
' leading whitespace; blank lines bridge between their neighbors.
Private Sub PaintGuides(ByVal hdc As Long, ByVal cp As VBIDE.CodePane, _
        ByVal topLine As Long, ByVal visLines As Long, _
        ByVal loLine As Long, ByVal hiLine As Long, _
        ByVal lineH As Long, ByVal yTop As Long, ByVal marginPx As Long)
    On Error Resume Next
    Dim cm As VBIDE.CodeModule
    Set cm = cp.CodeModule
    If cm Is Nothing Then Exit Sub
    If mTabW < 2 Then mTabW = 4
    If mCharW < 2 Then Exit Sub

    Dim hPen As Long, hOld As Long, oldBk As Long
    hPen = CreatePen(PS_DOT, 1, GUIDE_COLOR)
    hOld = SelectObject(hdc, hPen)
    oldBk = SetBkMode(hdc, BKMODE_TRANSPARENT)

    Dim row As Long, ln As Long, ind As Long, c As Long
    Dim x As Long, y As Long
    For row = 0 To visLines - 1
        ln = topLine + row
        If ln > cm.CountOfLines Then Exit For
        If ln >= loLine And ln <= hiLine Then
            ind = EffectiveIndent(cm, ln, loLine, hiLine)
            y = yTop + row * lineH
            For c = mTabW To ind - 1 Step mTabW
                ' same one-cell margin calibration as the match boxes
                x = marginPx + (c + 2) * mCharW - mCharW \ 2
                MoveToEx hdc, x, y, ByVal 0&
                LineTo hdc, x, y + lineH
            Next
        End If
    Next

    SetBkMode hdc, oldBk
    SelectObject hdc, hOld
    DeleteObject hPen
End Sub

Private Function EffectiveIndent(cm As VBIDE.CodeModule, ByVal ln As Long, _
        ByVal loLine As Long, ByVal hiLine As Long) As Long
    Dim ind As Long, up As Long, dn As Long, i As Long
    ind = LineIndent(cm, ln)
    If ind >= 0 Then EffectiveIndent = ind: Exit Function

    ' blank line: min indent of the nearest non-blank neighbors
    up = 0
    For i = ln - 1 To loLine Step -1
        up = LineIndent(cm, i)
        If up >= 0 Then Exit For
        If ln - i > 60 Then up = 0: Exit For
    Next
    If up < 0 Then up = 0
    dn = 0
    For i = ln + 1 To hiLine
        dn = LineIndent(cm, i)
        If dn >= 0 Then Exit For
        If i - ln > 60 Then dn = 0: Exit For
    Next
    If dn < 0 Then dn = 0
    EffectiveIndent = IIf(up < dn, up, dn)
End Function

' leading whitespace width in columns; -1 for blank lines
Private Function LineIndent(cm As VBIDE.CodeModule, ByVal ln As Long) As Long
    On Error Resume Next
    Dim s As String, i As Long, n As Long, ch As String
    s = cm.lines(ln, 1)
    If Len(Trim$(s)) = 0 Then LineIndent = -1: Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Then
            n = n + 1
        ElseIf ch = vbTab Then
            n = n + mTabW - (n Mod mTabW)
        Else
            Exit For
        End If
    Next
    LineIndent = n
End Function

Private Sub DrawGitBar(ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, _
        ByVal x2 As Long, ByVal Y2 As Long, ByVal kind As Long)
    Dim clr As Long, hBr As Long, hPn As Long, hOP As Long, hOB As Long
    Select Case kind
    Case GITK_ADD: clr = &H3CA03C          ' green
    Case GITK_MOD: clr = &HFF901E          ' dodger blue
    Case Else:     clr = &H3C3CC8          ' red (deletion marker)
    End Select
    hBr = CreateSolidBrush(clr)
    hPn = CreatePen(PS_SOLID, 1, clr)
    hOP = SelectObject(hdc, hPn)
    hOB = SelectObject(hdc, hBr)
    Rectangle hdc, x1, Y1, x2, Y2
    SelectObject hdc, hOP
    SelectObject hdc, hOB
    DeleteObject hPn
    DeleteObject hBr
End Sub

Public Sub Highlight_PaintScrollbar(ByVal hwnd As Long)
    On Error Resume Next
    If Not Highlight_Active() Then Exit Sub

    ' The scrollbar is a child of the VbaWindow editor.
    Dim hPane As Long
    hPane = GetParent(hwnd)
    If StrComp(WndClass(hPane), CLS_CODEPANE, vbTextCompare) <> 0 Then Exit Sub

    Dim cp As VBIDE.CodePane
    Set cp = PaneFromHwnd(hPane)
    If cp Is Nothing Then Exit Sub

    Dim compName As String, projName As String
    compName = cp.CodeModule.Parent.Name
    projName = cp.CodeModule.Parent.Collection.Parent.Name

    ' the scrollbar ranges over the displayed lines only: the whole
    ' module in Full Module View, the current proc in Procedure View
    Dim loLine As Long, hiLine As Long, totalLines As Long
    GetDisplayedRange cp, loLine, hiLine
    totalLines = hiLine - loLine + 1
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
            If mHL(i).LineNum >= loLine And mHL(i).LineNum <= hiLine Then
                y = arrowH + CLng((mHL(i).LineNum - loLine) / totalLines * trackH)
                Rectangle hdc, 2, y, (rc.Right - rc.Left) - 2, y + 3
            End If
        End If
    Next

    SelectObject hdc, hOldPen
    SelectObject hdc, hOldBr
    DeleteObject hPen
    DeleteObject hBr

    ' bookmark marks in blue
    Dim bmLines() As Long, nBM As Long
    nBM = BM_LinesForComp(compName, bmLines)
    If nBM > 0 Then
        Dim hBmBr As Long, hBmPen As Long
        hBmBr = CreateSolidBrush(BM_COLOR)
        hBmPen = CreatePen(PS_SOLID, 1, BM_COLOR)
        hOldPen = SelectObject(hdc, hBmPen)
        hOldBr = SelectObject(hdc, hBmBr)
        For i = 0 To nBM - 1
            If bmLines(i) >= loLine And bmLines(i) <= hiLine Then
                y = arrowH + CLng((bmLines(i) - loLine) / totalLines * trackH)
                Rectangle hdc, 2, y, (rc.Right - rc.Left) - 2, y + 3
            End If
        Next
        SelectObject hdc, hOldPen
        SelectObject hdc, hOldBr
        DeleteObject hBmPen
        DeleteObject hBmBr
    End If

    ' git marks on the scrollbar, colored by kind
    Dim gitLines() As Long, gitKinds() As Long, nGit As Long
    nGit = Git_MarksForComp(compName, gitLines, gitKinds)
    If nGit > 0 Then
        For i = 0 To nGit - 1
            If gitLines(i) >= loLine And gitLines(i) <= hiLine Then
                y = arrowH + CLng((gitLines(i) - loLine) / totalLines * trackH)
                DrawGitBar hdc, 2, y, (rc.Right - rc.Left) \ 2, y + 3, _
                           gitKinds(i)
            End If
        Next
    End If

    ReleaseDC hwnd, hdc
End Sub

' ---------------------------------------------------------------------

' Range of module lines the pane is currently displaying: the whole
' module in Full Module View, else the proc (or the declarations
' section) that the top visible line belongs to.
Private Sub GetDisplayedRange(ByVal cp As VBIDE.CodePane, _
        loLine As Long, hiLine As Long)
    On Error Resume Next
    Dim cm As VBIDE.CodeModule
    Set cm = cp.CodeModule
    loLine = 1
    hiLine = cm.CountOfLines
    If cp.CodePaneView = vbext_cv_ProcedureView Then
        Dim pk As vbext_ProcKind, nm As String
        nm = cm.ProcOfLine(cp.topLine, pk)
        If Len(nm) > 0 Then
            loLine = cm.ProcStartLine(nm, pk)
            hiLine = loLine + cm.ProcCountLines(nm, pk) - 1
        Else
            hiLine = cm.CountOfDeclarationLines
            If hiLine < 1 Then hiLine = 1
        End If
    End If
End Sub

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
