VERSION 5.00
Begin VB.Form frmSwitcher
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   ""
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   333
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmSwitcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Ctrl+Tab MRU window switcher (Alt-Tab style). Shown WITHOUT taking
'  focus; while Ctrl is held, the message hook in modWheel drives it:
'  Ctrl+Tab / Ctrl+Shift+Tab step through the list, releasing Ctrl
'  activates the selection, Esc cancels.
'
'  Custom drawn: shell file icon + caption per row, highlight bar on
'  the selection, orange dot on files modified vs git HEAD.
' =====================================================================

Private Const MAX_ROWS As Long = 15

Private mItemKeys() As String
Private mCaps() As String
Private mCount As Long
Private mSel As Long             ' 0-based selected row
Private mScroll As Long          ' first visible row

Private mRowH As Long            ' px
Private mHdrH As Long
Private mIconS As Long

Public Function BeginSwitch(ByVal forward As Boolean) As Boolean
    On Error Resume Next
    BuildList
    If mCount < 2 Then Exit Function

    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    mSel = IIf(forward, 1, mCount - 1)
    mScroll = 0
    EnsureVisible
    SizeAndPosition
    Repaint
    gSwitcherActive = True
    ShowWindow Me.hwnd, SW_SHOWNOACTIVATE
    BeginSwitch = True
End Function

Public Sub StepSwitch(ByVal forward As Boolean)
    On Error Resume Next
    mSel = mSel + IIf(forward, 1, -1)
    If mSel < 0 Then mSel = mCount - 1
    If mSel >= mCount Then mSel = 0
    EnsureVisible
    Repaint
End Sub

Public Sub CommitSwitch()
    On Error Resume Next
    Dim i As Long
    i = mSel
    gSwitcherActive = False
    ShowWindow Me.hwnd, SW_HIDE
    If i >= 0 And i < mCount Then ActivateByKey mItemKeys(i)
End Sub

Public Sub CancelSwitch()
    On Error Resume Next
    gSwitcherActive = False
    ShowWindow Me.hwnd, SW_HIDE
End Sub

' hovering tracks the selection, clicking an entry commits
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
        x As Single, y As Single)
    On Error Resume Next
    If Not gSwitcherActive Then Exit Sub
    Dim i As Long
    i = RowAt(y)
    If i >= 0 And i <> mSel Then
        mSel = i
        Repaint
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, _
        x As Single, y As Single)
    On Error Resume Next
    If Not gSwitcherActive Then Exit Sub
    Dim i As Long
    i = RowAt(y)
    If i >= 0 Then mSel = i
    CommitSwitch
End Sub

' ---------------------------------------------------------------------

Private Sub BuildList()
    On Error Resume Next
    Dim w As VBIDE.Window, k As String, cap As String, i As Long

    mCount = 0
    ReDim mItemKeys(0 To 32)
    ReDim mCaps(0 To 32)

    ' snapshot the current code/designer windows
    Dim keys As New Collection, caps As New Collection
    For Each w In gVBE.Windows
        If w.Visible Then
            If w.Type = vbext_wt_CodeWindow Or w.Type = vbext_wt_Designer Then
                k = NormalizeCaption(w.Caption) & "|" & w.Type
                Err.Clear
                keys.Add k, k
                If Err.Number = 0 Then caps.Add NormalizeCaption(w.Caption), k
            End If
        End If
    Next

    ' MRU-ranked windows first...
    For i = 1 To MRU_Count()
        k = MRU_Key(i)
        Err.Clear
        cap = caps(k)
        If Err.Number = 0 Then
            AddEntry k, cap
            caps.Remove k
            keys.Remove k
        End If
    Next

    ' ...then any windows never activated yet, in collection order
    For i = 1 To keys.Count
        k = keys(i)
        Err.Clear
        cap = caps(k)
        If Err.Number = 0 Then AddEntry k, cap
    Next
End Sub

Private Sub AddEntry(ByVal k As String, ByVal cap As String)
    If mCount > UBound(mItemKeys) Then
        ReDim Preserve mItemKeys(0 To mCount * 2)
        ReDim Preserve mCaps(0 To mCount * 2)
    End If
    mItemKeys(mCount) = k
    mCaps(mCount) = cap
    mCount = mCount + 1
End Sub

' ---------------------------------------------------------------------
'  Layout / drawing
' ---------------------------------------------------------------------

Private Sub Metrics()
    mRowH = ScaleForDpi(22)
    mHdrH = ScaleForDpi(19)
    mIconS = ScaleForDpi(16)
End Sub

Private Function VisRows() As Long
    VisRows = mCount
    If VisRows > MAX_ROWS Then VisRows = MAX_ROWS
End Function

Private Sub EnsureVisible()
    Dim vr As Long
    vr = VisRows()
    If vr <= 0 Then Exit Sub
    If mSel < mScroll Then mScroll = mSel
    If mSel >= mScroll + vr Then mScroll = mSel - vr + 1
    If mScroll < 0 Then mScroll = 0
End Sub

Private Function RowAt(ByVal y As Single) As Long
    Dim i As Long
    RowAt = -1
    If y < mHdrH Then Exit Function
    i = mScroll + (CLng(y) - mHdrH) \ mRowH
    If i >= mScroll + VisRows() Then Exit Function
    If i >= 0 And i < mCount Then RowAt = i
End Function

Private Sub SizeAndPosition()
    On Error Resume Next
    Metrics

    ' width from the longest caption, clamped
    Dim i As Long, wpx As Long, tw As Long
    For i = 0 To mCount - 1
        tw = Me.TextWidth(mCaps(i))
        If tw > wpx Then wpx = tw
    Next
    wpx = wpx + ScaleForDpi(8) + mIconS + ScaleForDpi(6) + ScaleForDpi(24)
    If wpx < ScaleForDpi(260) Then wpx = ScaleForDpi(260)
    If wpx > ScaleForDpi(440) Then wpx = ScaleForDpi(440)

    Dim hpx As Long
    hpx = mHdrH + VisRows() * mRowH + 2

    Dim rc As RECT
    GetWindowRect MDIClientHwnd(), rc
    Me.Move ((rc.Left + rc.Right) \ 2 - wpx \ 2) * Screen.TwipsPerPixelX, _
            ((rc.Top + rc.Bottom) \ 2 - hpx \ 2) * Screen.TwipsPerPixelY, _
            wpx * Screen.TwipsPerPixelX, hpx * Screen.TwipsPerPixelY
End Sub

Private Sub Repaint()
    On Error Resume Next
    Dim w As Long, H As Long, i As Long, y As Long, vr As Long

    Me.Cls
    w = Me.ScaleWidth
    H = Me.ScaleHeight

    ' header band
    Me.Line (1, 1)-(w - 2, mHdrH - 1), vbButtonFace, BF
    Me.ForeColor = vbGrayText
    Me.CurrentX = ScaleForDpi(8)
    Me.CurrentY = (mHdrH - Me.TextHeight("X")) \ 2 + 1
    Me.Print "Switch to"

    vr = VisRows()
    For i = mScroll To mScroll + vr - 1
        If i >= mCount Then Exit For
        y = mHdrH + (i - mScroll) * mRowH
        DrawRow i, y, w
    Next

    ' scroll hints when the list is longer than the popup
    Me.ForeColor = vbGrayText
    If mScroll > 0 Then
        Me.CurrentX = w - ScaleForDpi(14)
        Me.CurrentY = mHdrH + 1
        Me.Print "^"
    End If
    If mScroll + vr < mCount Then
        Me.CurrentX = w - ScaleForDpi(14)
        Me.CurrentY = H - ScaleForDpi(14)
        Me.Print "v"
    End If

    ' frame
    Me.Line (0, 0)-(w - 1, H - 1), THEME_BORDER, B

    Me.Refresh
End Sub

Private Sub DrawRow(ByVal i As Long, ByVal y As Long, ByVal w As Long)
    On Error Resume Next
    Dim xText As Long, cap As String, full As String

    If i = mSel Then
        Me.Line (1, y)-(w - 2, y + mRowH - 1), vbHighlight, BF
        Me.ForeColor = vbHighlightText
    Else
        Me.ForeColor = vbWindowText
    End If

    ' shell icon for the component behind this window
    DrawIcon16 Me.hdc, ScaleForDpi(8), y + (mRowH - mIconS) \ 2, _
        IconForCaption(mCaps(i))

    xText = ScaleForDpi(8) + mIconS + ScaleForDpi(6)

    ' orange dot = file modified vs git HEAD
    Dim nm As String, pp As Long
    nm = mCaps(i)
    pp = InStrRev(nm, " (")
    If pp > 0 Then nm = Left$(nm, pp - 1)
    If Git_IsCompChanged(nm) Then
        Me.Line (w - ScaleForDpi(14), y + mRowH \ 2 - 2)- _
                (w - ScaleForDpi(14) + 4, y + mRowH \ 2 + 2), THEME_ACCENT, BF
    End If

    ' caption, clipped with an ellipsis
    full = mCaps(i)
    cap = full
    Do While Me.TextWidth(cap) > (w - xText - ScaleForDpi(20)) And Len(cap) > 1
        cap = Left$(cap, Len(cap) - 1)
    Loop
    If cap <> full Then cap = Left$(cap, Len(cap) - 1) & Chr$(133)

    Me.CurrentX = xText
    Me.CurrentY = y + (mRowH - Me.TextHeight("X")) \ 2
    Me.Print cap
End Sub

' ---------------------------------------------------------------------

Private Sub ActivateByKey(ByVal k As String)
    On Error Resume Next
    Dim w As VBIDE.Window
    For Each w In gVBE.Windows
        If NormalizeCaption(w.Caption) & "|" & w.Type = k Then
            w.SetFocus
            If gTabBarVisible And w.WindowState <> vbext_ws_Maximize Then
                w.WindowState = vbext_ws_Maximize
            End If
            MRU_Touch k
            Exit Sub
        End If
    Next
End Sub
