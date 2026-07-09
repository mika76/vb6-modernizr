VERSION 5.00
Begin VB.Form frmTabs
   Appearance      =   0  'Flat
   BackColor       =   &H8000000F&
   BorderStyle     =   0  'None
   Caption         =   ""
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   AutoRedraw      =   -1  'True
   Begin VB.Timer tmrRefresh
      Interval        =   400
      Left            =   120
      Top             =   0
   End
   Begin VB.Menu mnuCtx
      Caption         =   "ctx"
      Visible         =   0   'False
      Begin VB.Menu mnuCtxClose
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuCtxCloseOthers
         Caption         =   "Close &Others"
      End
      Begin VB.Menu mnuCtxCloseAll
         Caption         =   "Close &All"
      End
      Begin VB.Menu mnuCtxSep1
         Caption         =   "-"
      End
      Begin VB.Menu mnuCtxCopyPath
         Caption         =   "Copy &Full Path"
      End
      Begin VB.Menu mnuCtxOpenFolder
         Caption         =   "Open Containin&g Folder"
      End
   End
   Begin VB.Menu mnuList
      Caption         =   "list"
      Visible         =   0   'False
      Begin VB.Menu mnuListItem
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Tab bar for the IDE's MDI windows (code + designer windows).
'  Owner drawn in classic VB6 3D style. Left click activates,
'  middle click closes, right click shows a context menu, and the
'  drop-down button at the right lists every window.
' =====================================================================

Private Const DROPBTN_W As Long = 20

Private mWins As Collection      ' of VBIDE.Window, tab order
Private mSig As String           ' change-detection signature
Private mActiveIdx As Long       ' 1-based index into mWins, 0 = none

Private mTabL() As Long          ' hit-test rects for drawn tabs
Private mTabR() As Long
Private mTabIdx() As Long        ' mWins index for each drawn tab
Private mTabCount As Long        ' number of drawn (visible) tabs
Private mScroll As Long          ' first visible tab (1-based)

Private mCtxIdx As Long          ' tab targeted by the context menu
Private mLastAct As String       ' last active-window key fed to MRU

Private mOrder As Collection     ' user tab order (keys), survives refresh
Private mDragIdx As Long         ' mWins index being dragged, 0 = none
Private mDragKey As String
Private mDownX As Single
Private mDragging As Boolean

' ---------------------------------------------------------------------

Public Sub Attach()
    On Error Resume Next
    If GetParent(Me.hwnd) = MainHwnd() Then Exit Sub
    Dim style As Long
    style = GetWindowLongA(Me.hwnd, GWL_STYLE)
    style = (style Or WS_CHILD) And (Not WS_POPUP)
    SetWindowLongA Me.hwnd, GWL_STYLE, style
    SetParent Me.hwnd, MainHwnd()
    RefreshTabs True
End Sub

Public Sub Reposition()
    On Error Resume Next
    If Not gTabBarVisible Then Exit Sub
    Dim x As Long, y As Long, w As Long
    If Not Layout_StripOrigin(x, y, w) Then Exit Sub
    MoveWindow Me.hwnd, x, y, w, ScaleForDpi(TAB_BAR_HEIGHT), 1
End Sub

' ---------------------------------------------------------------------

Private Sub tmrRefresh_Timer()
    On Error Resume Next
    If gVBE Is Nothing Then Exit Sub
    If gTabBarVisible Then RefreshTabs False
    BM_Poll
    Git_Poll
    Backup_Poll
    Highlight_EnsureHooks
    If gLineNumsEnabled Then frmGutter.Poll
End Sub

' external state (git) changed: repaint on the next tick
Public Sub ForceRedraw()
    mSig = ""
End Sub

Private Sub RefreshTabs(ByVal force As Boolean)
    On Error Resume Next
    Dim w As VBIDE.Window, j As Long, k As String
    Dim raw As New Collection      ' windows in IDE order
    Dim byKey As New Collection    ' keyed, unclaimed windows
    Dim ordered As New Collection

    For Each w In gVBE.Windows
        If w.Visible Then
            If w.Type = vbext_wt_CodeWindow Or w.Type = vbext_wt_Designer Then
                raw.Add w
                Err.Clear
                byKey.Add w, WinKey(w)
                Err.Clear
            End If
        End If
    Next

    ' user-defined order first, new windows appended at the end
    If mOrder Is Nothing Then Set mOrder = New Collection
    For j = 1 To mOrder.Count
        k = mOrder(j)
        Err.Clear
        Set w = byKey(k)
        If Err.Number = 0 Then
            ordered.Add w
            byKey.Remove k
        End If
    Next
    For j = 1 To raw.Count
        k = WinKey(raw(j))
        Err.Clear
        Set w = byKey(k)
        If Err.Number = 0 Then
            ordered.Add w
            byKey.Remove k
            mOrder.Add k, k
        End If
    Next

    Dim sig As String
    For j = 1 To ordered.Count
        sig = sig & WinKey(ordered(j)) & _
              IIf(TabDirty(ordered(j)), "*", "") & ";"
    Next

    Dim aw As VBIDE.Window, actCap As String
    Set aw = gVBE.ActiveWindow
    If Not aw Is Nothing Then
        If aw.Type = vbext_wt_CodeWindow Or aw.Type = vbext_wt_Designer Then
            actCap = WinKey(aw)
            ' keep MDI children maximized while the tab bar is on
            If gTabBarVisible And aw.WindowState <> vbext_ws_Maximize Then
                aw.WindowState = vbext_ws_Maximize
            End If
        End If
    End If
    sig = sig & "#" & actCap
    If Len(actCap) > 0 And actCap <> mLastAct Then
        mLastAct = actCap
        MRU_Touch actCap
    End If

    If sig = mSig And Not force Then Exit Sub
    mSig = sig
    Set mWins = ordered

    mActiveIdx = 0
    If Len(actCap) > 0 Then
        For j = 1 To mWins.Count
            If WinKey(mWins(j)) = actCap Then
                mActiveIdx = j
                Exit For
            End If
        Next
    End If

    Redraw
End Sub

Private Function WinKey(ByVal w As VBIDE.Window) As String
    On Error Resume Next
    WinKey = NormalizeCaption(w.Caption) & "|" & w.Type
End Function

Private Function CompForWindow(ByVal w As VBIDE.Window) As VBIDE.VBComponent
    On Error Resume Next
    Dim nm As String, pp As Long
    nm = NormalizeCaption(w.Caption)
    pp = InStrRev(nm, " (")
    If pp > 0 Then nm = Left$(nm, pp - 1)
    Dim proj As VBIDE.VBProject
    For Each proj In gVBE.VBProjects
        Set CompForWindow = Nothing
        Err.Clear
        Set CompForWindow = proj.VBComponents(nm)
        If Not CompForWindow Is Nothing Then Exit Function
    Next
End Function

Private Function TabDirty(ByVal w As VBIDE.Window) As Boolean
    On Error Resume Next
    Dim comp As VBIDE.VBComponent
    Set comp = CompForWindow(w)
    If Not comp Is Nothing Then TabDirty = comp.IsDirty
End Function

Private Function TabDispCaption(ByVal w As VBIDE.Window) As String
    TabDispCaption = TabCaption(w) & IIf(TabDirty(w), " *", "")
End Function

' ---------------------------------------------------------------------

Private Sub Redraw()
    On Error Resume Next
    Dim H As Long, w As Long
    H = Me.ScaleHeight
    w = Me.ScaleWidth

    Me.Cls
    Me.Line (0, 0)-(w, H), vbButtonFace, BF
    ' bottom shadow line, classic 3D edge
    Me.Line (0, H - 1)-(w, H - 1), vb3DShadow

    If mWins Is Nothing Then Exit Sub
    Dim n As Long
    n = mWins.Count
    ReDim mTabL(1 To n + 1)
    ReDim mTabR(1 To n + 1)
    ReDim mTabIdx(1 To n + 1)
    mTabCount = 0
    If n = 0 Then Exit Sub

    ' git branch label, right-aligned before the drop button
    Dim gitLbl As String, gitW As Long
    gitLbl = Git_Branch()
    If Len(gitLbl) > 0 Then
        If Git_RepoDirty() Then gitLbl = gitLbl & " *"
        gitW = Me.TextWidth(gitLbl) + 12
    End If

    Dim avail As Long
    avail = w - DROPBTN_W - gitW - 4

    If mScroll < 1 Then mScroll = 1
    If mScroll > n Then mScroll = n
    ' make sure the active tab is within the visible run
    If mActiveIdx > 0 And mActiveIdx < mScroll Then mScroll = mActiveIdx

    Dim x As Long, i As Long, tw As Long, cap As String
    Dim fits As Boolean

TryLayout:
    x = 2
    mTabCount = 0
    fits = False
    ' measure with the bold font: the active tab draws bold, and using
    ' one width for both states keeps tabs from resizing on activation
    Me.FontBold = True
    For i = mScroll To n
        cap = TabDispCaption(mWins(i))
        tw = Me.TextWidth(cap) + ScaleForDpi(35)
        If tw < ScaleForDpi(60) Then tw = ScaleForDpi(60)
        If tw > ScaleForDpi(200) Then tw = ScaleForDpi(200)
        If x + tw > avail Then Exit For
        mTabCount = mTabCount + 1
        mTabL(mTabCount) = x
        mTabR(mTabCount) = x + tw
        mTabIdx(mTabCount) = i
        If i = mActiveIdx Then fits = True
        x = x + tw + 1
    Next
    Me.FontBold = False
    ' active tab off the right edge: scroll right and retry
    If mActiveIdx > 0 And Not fits And mScroll < mActiveIdx Then
        mScroll = mScroll + 1
        GoTo TryLayout
    End If

    For i = 1 To mTabCount
        DrawTab i, (mTabIdx(i) = mActiveIdx)
    Next

    DrawDropButton

    If gitW > 0 Then
        Me.CurrentX = w - DROPBTN_W - gitW + 6
        Me.CurrentY = (H - Me.TextHeight("X")) \ 2 + 1
        Me.ForeColor = vbGrayText
        Me.Print gitLbl
        Me.ForeColor = vbButtonText
    End If

    ' icons are drawn via the API, so VB must be told to flush
    Me.Refresh
End Sub

Private Function TabCaption(ByVal w As VBIDE.Window) As String
    On Error Resume Next
    TabCaption = NormalizeCaption(w.Caption)
End Function

Private Sub DrawTab(ByVal slot As Long, ByVal active As Boolean)
    On Error Resume Next
    Dim l As Long, r As Long, H As Long, w As VBIDE.Window
    l = mTabL(slot): r = mTabR(slot)
    H = Me.ScaleHeight
    Set w = mWins(mTabIdx(slot))

    If active Then
        ' raised tab: highlight top/left, shadow right, no bottom edge
        Me.Line (l, 2)-(r, H - 1), vbButtonFace, BF
        Me.Line (l, 2)-(r - 1, 2), vb3DHighlight            ' top
        Me.Line (l, 2)-(l, H - 1), vb3DHighlight            ' left
        Me.Line (r - 1, 2)-(r - 1, H - 1), vb3DShadow       ' right
        Me.Line (l + 1, H - 1)-(r - 1, H - 1), vbButtonFace ' erase bottom
    Else
        ' subtle separator on the right side
        Me.Line (r, 6)-(r, H - 6), vb3DShadow
    End If

    ' shell file icon for the component behind this window
    Dim gy As Long
    gy = H \ 2
    DrawIcon16 Me.hdc, l + ScaleForDpi(5), gy - ScaleForDpi(8), _
        IconForCaption(TabCaption(w))

    ' orange badge on the icon corner = file modified vs git HEAD
    Dim nm As String, pp As Long
    nm = TabCaption(w)
    pp = InStrRev(nm, " (")
    If pp > 0 Then nm = Left$(nm, pp - 1)
    If Git_IsCompChanged(nm) Then
        Me.Line (l + ScaleForDpi(16), gy - ScaleForDpi(9))- _
                (l + ScaleForDpi(21), gy - ScaleForDpi(4)), THEME_ACCENT, BF
    End If

    ' caption, clipped to the tab; measure in the style it draws in
    Dim full As String, cap As String
    Me.FontBold = active
    full = TabDispCaption(w)
    cap = full
    Do While Me.TextWidth(cap) > (r - l - ScaleForDpi(33)) And Len(cap) > 1
        cap = Left$(cap, Len(cap) - 1)
    Loop
    If cap <> full Then cap = Left$(cap, Len(cap) - 1) & Chr$(133)

    Me.CurrentX = l + ScaleForDpi(25)
    Me.CurrentY = (H - Me.TextHeight("X")) \ 2 + 1
    Me.ForeColor = vbButtonText
    Me.Print cap
    Me.FontBold = False
End Sub

Private Sub DrawDropButton()
    On Error Resume Next
    Dim x As Long, H As Long, cy As Long
    x = Me.ScaleWidth - DROPBTN_W
    H = Me.ScaleHeight
    cy = H \ 2 - 1
    ' triangle pointing down
    Me.Line (x + 4, cy - 2)-(x + 12, cy - 2), vbButtonText
    Me.Line (x + 5, cy - 1)-(x + 11, cy - 1), vbButtonText
    Me.Line (x + 6, cy)-(x + 10, cy), vbButtonText
    Me.Line (x + 7, cy + 1)-(x + 9, cy + 1), vbButtonText
    Me.PSet (x + 8, cy + 2), vbButtonText
End Sub

' ---------------------------------------------------------------------
'  Interaction
' ---------------------------------------------------------------------

Private Function HitTest(ByVal x As Single) As Long
    Dim i As Long
    For i = 1 To mTabCount
        If x >= mTabL(i) And x < mTabR(i) Then
            HitTest = mTabIdx(i)
            Exit Function
        End If
    Next
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, _
        x As Single, y As Single)
    On Error Resume Next
    If x >= Me.ScaleWidth - DROPBTN_W Then
        If Button = vbLeftButton Then ShowWindowList
        Exit Sub
    End If

    Dim idx As Long
    idx = HitTest(x)
    If idx = 0 Then Exit Sub

    Select Case Button
    Case vbLeftButton
        ActivateTab idx
        ' arm a possible drag-reorder
        mDragIdx = idx
        mDragKey = WinKey(mWins(idx))
        mDownX = x
        mDragging = False
    Case vbMiddleButton
        CloseTab idx
    Case vbRightButton
        ActivateTab idx
        mCtxIdx = idx
        PopupMenu mnuCtx
    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
        x As Single, y As Single)
    On Error Resume Next
    If (Button And vbLeftButton) = 0 Then Exit Sub
    If mDragIdx = 0 Then Exit Sub
    If Not mDragging Then
        If Abs(x - mDownX) > 6 Then mDragging = True
    End If
    If mDragging Then
        Dim tgt As Long
        tgt = HitTest(x)
        If tgt > 0 Then
            If WinKey(mWins(tgt)) <> mDragKey Then DragReorder tgt
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, _
        x As Single, y As Single)
    mDragIdx = 0
    mDragging = False
End Sub

' Move the dragged tab's key so it lands before/after the target
' (after when dragging rightwards, before when dragging leftwards).
Private Sub DragReorder(ByVal targetIdx As Long)
    On Error Resume Next
    Dim tKey As String, posD As Long, posT As Long
    tKey = WinKey(mWins(targetIdx))
    posD = OrderPos(mDragKey)
    posT = OrderPos(tKey)
    If posD = 0 Or posT = 0 Then Exit Sub

    mOrder.Remove mDragKey
    If posT > mOrder.Count Then
        mOrder.Add mDragKey, mDragKey
    Else
        mOrder.Add mDragKey, mDragKey, posT
    End If
    RefreshTabs True
End Sub

Private Function OrderPos(ByVal k As String) As Long
    Dim i As Long
    For i = 1 To mOrder.Count
        If mOrder(i) = k Then OrderPos = i: Exit Function
    Next
End Function

Private Sub ActivateTab(ByVal idx As Long)
    On Error Resume Next
    mWins(idx).SetFocus
    If mWins(idx).WindowState <> vbext_ws_Maximize Then
        mWins(idx).WindowState = vbext_ws_Maximize
    End If
    RefreshTabs True
End Sub

Private Sub CloseTab(ByVal idx As Long)
    On Error Resume Next
    mWins(idx).Close
    RefreshTabs True
End Sub

Private Sub mnuCtxClose_Click()
    CloseTab mCtxIdx
End Sub

Private Sub mnuCtxCloseOthers_Click()
    On Error Resume Next
    Dim i As Long
    For i = mWins.Count To 1 Step -1
        If i <> mCtxIdx Then mWins(i).Close
    Next
    RefreshTabs True
End Sub

Private Sub mnuCtxCloseAll_Click()
    On Error Resume Next
    Dim i As Long
    For i = mWins.Count To 1 Step -1
        mWins(i).Close
    Next
    RefreshTabs True
End Sub

Private Sub mnuCtxCopyPath_Click()
    On Error Resume Next
    Dim comp As VBIDE.VBComponent
    Set comp = CompForWindow(mWins(mCtxIdx))
    If comp Is Nothing Then Exit Sub
    If comp.FileCount = 0 Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText comp.FileNames(1)
End Sub

Private Sub mnuCtxOpenFolder_Click()
    On Error Resume Next
    Dim comp As VBIDE.VBComponent
    Set comp = CompForWindow(mWins(mCtxIdx))
    If comp Is Nothing Then Exit Sub
    If comp.FileCount = 0 Then Exit Sub
    Shell "explorer.exe /select,""" & comp.FileNames(1) & """", vbNormalFocus
End Sub

' Drop-down list of all windows (for when tabs overflow)
Private Sub ShowWindowList()
    On Error Resume Next
    If mWins Is Nothing Then Exit Sub
    If mWins.Count = 0 Then Exit Sub

    Dim i As Long
    ' clear previous dynamic items (keep index 0)
    For i = mnuListItem.UBound To 1 Step -1
        Unload mnuListItem(i)
    Next

    For i = 1 To mWins.Count
        If i > 1 Then Load mnuListItem(i - 1)
        mnuListItem(i - 1).Caption = Replace$(TabCaption(mWins(i)), "&", "&&")
        mnuListItem(i - 1).Checked = (i = mActiveIdx)
    Next

    PopupMenu mnuList, , Me.ScaleWidth - DROPBTN_W, Me.ScaleHeight
End Sub

Private Sub mnuListItem_Click(Index As Integer)
    ActivateTab Index + 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Redraw
End Sub
