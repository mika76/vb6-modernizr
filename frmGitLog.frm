VERSION 5.00
Begin VB.Form frmGitLog
   Caption         =   "Git Log - Modernizr"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAll
      Caption         =   "All &branches"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1500
   End
   Begin VB.CommandButton cmdRefresh
      Caption         =   "&Refresh"
      Height          =   345
      Left            =   9020
      TabIndex        =   1
      Top             =   135
      Width           =   1000
   End
   Begin VB.PictureBox picLog
      AutoRedraw      =   -1  'True
      Height          =   3780
      Left            =   180
      ScaleHeight     =   248
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   652
      TabIndex        =   2
      Top             =   630
      Width           =   9840
   End
   Begin VB.TextBox txtDetails
      BeginProperty Font
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   4570
      Width           =   9840
   End
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   240
      Left            =   1860
      TabIndex        =   5
      Top             =   195
      Width           =   7000
   End
End
Attribute VB_Name = "frmGitLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Git Log (Ctrl+Shift+L). The commit graph is drawn by us: commits
'  come in with parent hashes (%h %P), lanes are laid out top-down
'  with the classic active-lanes algorithm, then painted as colored
'  lane lines, merge/branch diagonals and commit dots (ring = merge).
'  Branch/tag decorations render as colored chips. Click a commit to
'  see its details/diffstat below.
' =====================================================================

Private Const MAX_COMMITS As Long = 300
Private Const MAX_DRAW_LANES As Long = 12

' per-commit data (parallel arrays)
Private mHash() As String
Private mParents() As String    ' space separated
Private mDate() As String
Private mAuthor() As String
Private mRefs() As String       ' %D decorations, comma separated
Private mSubject() As String
Private mLane() As Long         ' lane of the commit dot
Private mMerge() As Boolean
' drawing segments per row: comma list of "Va:b" (lane pass-through,
' top lane a to bottom lane b), "Ia" (top lane a into the dot),
' "Oa" (dot to bottom lane a)
Private mSegs() As String
Private mCount As Long
Private mMaxLane As Long

Private mSelRow As Long         ' selected row, -1 = none
Private mTop As Long            ' first visible row
Private mRowH As Long
Private mLaneW As Long
Private mAttached As Boolean

Public Sub ShowLog()
    On Error Resume Next
    Load Me
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    If Not mAttached Then
        mAttached = True
        Scroll_Attach picLog.hwnd, Me
    End If
    Me.Show vbModeless
    LoadLog
End Sub

' modScroll callback: scrollbar drag, arrows, page clicks, wheel
Public Sub ScrollTo(ByVal pos As Long)
    If pos < 0 Then pos = 0
    If pos > MaxTop() Then pos = MaxTop()
    If pos = mTop Then Exit Sub
    mTop = pos
    UpdateScrollbar
    DrawGraph
End Sub

Private Sub cmdRefresh_Click()
    LoadLog
End Sub

Private Sub chkAll_Click()
    LoadLog
End Sub

' ---------------------------------------------------------------------
'  Load + lane layout
' ---------------------------------------------------------------------

Private Sub LoadLog()
    On Error Resume Next
    mCount = 0
    mSelRow = -1
    mTop = 0
    mMaxLane = 0
    txtDetails.Text = ""

    If Not Git_HasRepo() Then
        lblStatus.Caption = "No git repository found for the active project."
        DrawGraph
        Exit Sub
    End If

    lblStatus.Caption = "Loading..."
    DoEvents

    ' full %H hashes: %P lists FULL parent hashes, so lane matching
    ' only works when the commit hash is full-length too
    Dim args As String
    args = "log -n " & MAX_COMMITS & " --topo-order --date=short " & _
           "--pretty=format:" & Chr$(34) & _
           "%H%x09%P%x09%ad%x09%an%x09%D%x09%s" & Chr$(34)
    If chkAll.Value = vbChecked Then args = args & " --all"

    Dim res As String
    res = Git_RunSync(args, 15000)
    If Len(res) = 0 Then
        lblStatus.Caption = "No output from git log (empty repo?)."
        DrawGraph
        Exit Sub
    End If

    Dim lines() As String, i As Long, s As String, p() As String
    lines = Split(res, vbLf)
    ReDim mHash(0 To UBound(lines) + 1)
    ReDim mParents(0 To UBound(lines) + 1)
    ReDim mDate(0 To UBound(lines) + 1)
    ReDim mAuthor(0 To UBound(lines) + 1)
    ReDim mRefs(0 To UBound(lines) + 1)
    ReDim mSubject(0 To UBound(lines) + 1)
    ReDim mLane(0 To UBound(lines) + 1)
    ReDim mMerge(0 To UBound(lines) + 1)
    ReDim mSegs(0 To UBound(lines) + 1)

    For i = 0 To UBound(lines)
        s = Replace$(lines(i), vbCr, "")
        If Len(s) > 0 Then
            p = Split(s, vbTab)
            If UBound(p) >= 5 Then
                mHash(mCount) = p(0)
                mParents(mCount) = Trim$(p(1))
                mDate(mCount) = p(2)
                mAuthor(mCount) = p(3)
                mRefs(mCount) = Trim$(p(4))
                ' a subject may itself contain tabs - rejoin the tail
                Dim j As Long
                For j = 6 To UBound(p)
                    p(5) = p(5) & " " & p(j)
                Next
                mSubject(mCount) = p(5)
                mCount = mCount + 1
            End If
        End If
    Next

    LayoutLanes
    UpdateScrollbar
    DrawGraph

    lblStatus.Caption = mCount & " commit(s)" & _
        IIf(mCount >= MAX_COMMITS, " (showing latest " & MAX_COMMITS & ")", "") & _
        ".  Click a commit for details."
End Sub

' Classic active-lanes layout: each lane holds the hash it expects
' next. A commit lands on the first lane expecting it (collapsing any
' duplicates = branch point), then the lane carries parent 1; extra
' parents (merge) join an existing lane or open a new one.
Private Sub LayoutLanes()
    On Error Resume Next
    Dim lanes() As String, laneCount As Long
    Dim newLanes() As String, newCount As Long
    Dim row As Long, i As Long, k As Long, c As Long, f As Long
    Dim h As String, pars() As String, np As Long, segs As String

    ReDim lanes(0 To 32)
    laneCount = 0

    For row = 0 To mCount - 1
        h = mHash(row)
        If Len(mParents(row)) > 0 Then
            pars = Split(mParents(row), " ")
            np = UBound(pars) + 1
        Else
            np = 0
        End If
        mMerge(row) = (np >= 2)
        segs = ""

        ' the commit's lane = first lane expecting this hash
        Dim wasNew As Boolean
        wasNew = False
        c = -1
        For i = 0 To laneCount - 1
            If lanes(i) = h Then c = i: Exit For
        Next
        If c = -1 Then
            ' branch tip never referenced yet: opens a new lane
            c = laneCount
            If c > UBound(lanes) Then ReDim Preserve lanes(0 To c * 2)
            lanes(c) = h
            laneCount = laneCount + 1
            wasNew = True
        End If

        ' incoming edges: every lane expecting this commit
        ' (a fresh tip has no line coming in from above)
        For i = 0 To laneCount - 1
            If lanes(i) = h And Not (wasNew And i = c) Then
                AddSeg segs, "I" & i
            End If
        Next

        ' rebuild the lane list below this row
        ReDim newLanes(0 To laneCount + np + 1)
        newCount = 0
        For i = 0 To laneCount - 1
            If lanes(i) = h Then
                If i = c And np > 0 Then
                    newLanes(newCount) = pars(0)
                    AddSeg segs, "O" & newCount
                    newCount = newCount + 1
                End If
                ' duplicates collapse, roots just end
            Else
                newLanes(newCount) = lanes(i)
                AddSeg segs, "V" & i & ":" & newCount
                newCount = newCount + 1
            End If
        Next

        ' extra parents of a merge
        For k = 1 To np - 1
            f = -1
            For i = 0 To newCount - 1
                If newLanes(i) = pars(k) Then f = i: Exit For
            Next
            If f = -1 Then
                newLanes(newCount) = pars(k)
                f = newCount
                newCount = newCount + 1
            End If
            AddSeg segs, "O" & f
        Next

        mLane(row) = c
        mSegs(row) = segs
        If laneCount > mMaxLane Then mMaxLane = laneCount
        If newCount > mMaxLane Then mMaxLane = newCount

        ReDim lanes(0 To newCount + 8)
        For i = 0 To newCount - 1
            lanes(i) = newLanes(i)
        Next
        laneCount = newCount
    Next

    If mMaxLane > MAX_DRAW_LANES Then mMaxLane = MAX_DRAW_LANES
    If mMaxLane < 1 Then mMaxLane = 1
End Sub

Private Sub AddSeg(segs As String, ByVal s As String)
    If Len(segs) > 0 Then segs = segs & ","
    segs = segs & s
End Sub

' ---------------------------------------------------------------------
'  Drawing
' ---------------------------------------------------------------------

Private Sub Metrics()
    mRowH = ScaleForDpi(20)
    mLaneW = ScaleForDpi(12)
End Sub

Private Function VisRows() As Long
    If mRowH <= 0 Then Metrics
    VisRows = picLog.ScaleHeight \ mRowH
    If VisRows < 1 Then VisRows = 1
End Function

Private Function LaneX(ByVal lane As Long) As Long
    If lane > MAX_DRAW_LANES - 1 Then lane = MAX_DRAW_LANES - 1
    LaneX = ScaleForDpi(10) + lane * mLaneW
End Function

Private Sub DrawGraph()
    On Error Resume Next
    Dim w As Long, i As Long, y As Long, lastRow As Long

    Metrics
    picLog.Cls
    w = picLog.ScaleWidth

    lastRow = mTop + VisRows() - 1
    If lastRow > mCount - 1 Then lastRow = mCount - 1

    For i = mTop To lastRow
        y = (i - mTop) * mRowH
        DrawRow i, y, w
    Next

    picLog.Refresh
End Sub

Private Sub DrawRow(ByVal row As Long, ByVal y As Long, ByVal w As Long)
    On Error Resume Next
    Dim graphW As Long, xMid As Long, yMid As Long, yBot As Long
    Dim parts() As String, i As Long, s As String, a As Long, b As Long

    graphW = ScaleForDpi(10) + mMaxLane * mLaneW + ScaleForDpi(4)
    yMid = y + mRowH \ 2
    yBot = y + mRowH

    ' selection band behind everything
    If row = mSelRow Then
        picLog.Line (0, y)-(w - 1, yBot - 1), vb3DLight, BF
        picLog.Line (0, y)-(w - 1, yBot - 1), vbHighlight, B
    End If

    ' graph segments
    picLog.DrawWidth = 2
    If Len(mSegs(row)) > 0 Then
        parts = Split(mSegs(row), ",")
        For i = 0 To UBound(parts)
            s = parts(i)
            Select Case Left$(s, 1)
                Case "V"
                    a = CLng(Mid$(s, 2, InStr(s, ":") - 2))
                    b = CLng(Mid$(s, InStr(s, ":") + 1))
                    picLog.Line (LaneX(a), y)-(LaneX(b), yBot), LaneColor(b)
                Case "I"
                    a = CLng(Mid$(s, 2))
                    picLog.Line (LaneX(a), y)-(LaneX(mLane(row)), yMid), _
                        LaneColor(a)
                Case "O"
                    a = CLng(Mid$(s, 2))
                    picLog.Line (LaneX(mLane(row)), yMid)-(LaneX(a), yBot), _
                        LaneColor(a)
            End Select
        Next
    End If

    ' commit dot (ring = merge commit)
    xMid = LaneX(mLane(row))
    picLog.DrawWidth = 1
    Dim r As Long
    r = ScaleForDpi(3)
    If mMerge(row) Then
        picLog.FillStyle = 1                ' transparent
        picLog.Circle (xMid, yMid), r, LaneColor(mLane(row))
    Else
        picLog.FillStyle = 0                ' solid
        picLog.FillColor = LaneColor(mLane(row))
        picLog.Circle (xMid, yMid), r, LaneColor(mLane(row))
        picLog.FillStyle = 1
    End If

    ' text columns: hash, date, author, [chips] subject
    Dim x As Long, colGray As Long
    colGray = vbGrayText
    x = graphW + ScaleForDpi(6)

    x = PrintCol(x, y, Left$(mHash(row), 7), colGray, _
                 picLog.TextWidth("0000000"))
    x = PrintCol(x, y, mDate(row), colGray, picLog.TextWidth("0000-00-00"))
    x = PrintCol(x, y, ClipText(mAuthor(row), ScaleForDpi(84)), colGray, _
                 ScaleForDpi(84))

    ' decoration chips
    If Len(mRefs(row)) > 0 Then x = DrawChips(x, y, mRefs(row))

    ' subject
    picLog.ForeColor = vbWindowText
    picLog.CurrentX = x
    picLog.CurrentY = y + (mRowH - picLog.TextHeight("X")) \ 2
    picLog.Print ClipText(mSubject(row), w - x - ScaleForDpi(6))
End Sub

Private Function PrintCol(ByVal x As Long, ByVal y As Long, ByVal s As String, _
        ByVal clr As Long, ByVal colW As Long) As Long
    picLog.ForeColor = clr
    picLog.CurrentX = x
    picLog.CurrentY = y + (mRowH - picLog.TextHeight("X")) \ 2
    picLog.Print s
    PrintCol = x + colW + ScaleForDpi(8)
End Function

Private Function ClipText(ByVal s As String, ByVal maxW As Long) As String
    Dim cap As String
    cap = s
    Do While picLog.TextWidth(cap) > maxW And Len(cap) > 1
        cap = Left$(cap, Len(cap) - 1)
    Loop
    If cap <> s Then cap = Left$(cap, Len(cap) - 1) & Chr$(133)
    ClipText = cap
End Function

' Branch / tag / HEAD decorations as colored chips. Returns next x.
Private Function DrawChips(ByVal x As Long, ByVal y As Long, _
        ByVal refs As String) As Long
    On Error Resume Next
    Dim parts() As String, i As Long, nm As String, clr As Long
    Dim tw As Long, chipH As Long, cy As Long

    chipH = mRowH - ScaleForDpi(6)
    cy = y + ScaleForDpi(3)
    parts = Split(refs, ", ")

    For i = 0 To UBound(parts)
        nm = Trim$(parts(i))
        If Len(nm) > 0 Then
            If Left$(nm, 5) = "HEAD " Or nm = "HEAD" Then
                clr = THEME_ACCENT
                nm = Replace$(nm, "HEAD -> ", "")
            ElseIf Left$(nm, 5) = "tag: " Then
                clr = THEME_CODE
                nm = Mid$(nm, 6)
            Else
                clr = THEME_DESIGN
            End If
            nm = ClipText(nm, ScaleForDpi(110))
            tw = picLog.TextWidth(nm)
            picLog.Line (x, cy)-(x + tw + ScaleForDpi(8), cy + chipH), clr, BF
            picLog.ForeColor = vbWhite
            picLog.CurrentX = x + ScaleForDpi(4)
            picLog.CurrentY = cy + (chipH - picLog.TextHeight("X")) \ 2
            picLog.Print nm
            x = x + tw + ScaleForDpi(8) + ScaleForDpi(4)
        End If
    Next
    DrawChips = x + ScaleForDpi(2)
End Function

' ---------------------------------------------------------------------
'  Scrolling / selection
' ---------------------------------------------------------------------

Private Function MaxTop() As Long
    MaxTop = mCount - VisRows()
    If MaxTop < 0 Then MaxTop = 0
End Function

Private Sub UpdateScrollbar()
    On Error Resume Next
    If MaxTop() = 0 Then mTop = 0
    If mAttached Then _
        Scroll_Update picLog.hwnd, mCount - 1, VisRows(), mTop
End Sub

Private Sub EnsureVisible(ByVal i As Long)
    If i < mTop Then mTop = i
    If i >= mTop + VisRows() Then mTop = i - VisRows() + 1
    If mTop < 0 Then mTop = 0
    UpdateScrollbar
End Sub

Private Sub picLog_MouseDown(Button As Integer, Shift As Integer, _
        x As Single, y As Single)
    On Error Resume Next
    Dim i As Long
    i = mTop + CLng(y) \ mRowH
    If i < 0 Or i >= mCount Or i >= mTop + VisRows() Then Exit Sub
    SelectRow i
End Sub

Private Sub picLog_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i As Long
    i = mSelRow
    Select Case KeyCode
        Case vbKeyUp:       i = i - 1
        Case vbKeyDown:     i = i + 1
        Case vbKeyPageUp:   i = i - (VisRows() - 1)
        Case vbKeyPageDown: i = i + (VisRows() - 1)
        Case vbKeyHome:     i = 0
        Case vbKeyEnd:      i = mCount - 1
        Case Else: Exit Sub
    End Select
    KeyCode = 0
    If mCount = 0 Then Exit Sub
    If i < 0 Then i = 0
    If i >= mCount Then i = mCount - 1
    If i <> mSelRow Then
        EnsureVisible i
        SelectRow i
    End If
End Sub

Private Sub SelectRow(ByVal i As Long)
    On Error Resume Next
    mSelRow = i
    DrawGraph

    Dim res As String
    res = Git_RunSync("show --stat --date=short " & mHash(i), 10000)
    If Len(res) > 28000 Then res = Left$(res, 28000) & vbCrLf & "[...]"
    txtDetails.Text = Replace$(Replace$(res, vbCrLf, vbLf), vbLf, vbCrLf)
End Sub

' ---------------------------------------------------------------------

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Dim w As Long, listH As Long
    w = Me.ScaleWidth - MARGIN_STD * 2
    listH = (Me.ScaleHeight - 1020) * 2 \ 3
    If listH < 600 Then listH = 600

    picLog.Width = w
    picLog.Height = listH
    txtDetails.Top = picLog.Top + listH + 160
    txtDetails.Width = w
    txtDetails.Height = Me.ScaleHeight - txtDetails.Top - MARGIN_STD
    cmdRefresh.Left = Me.ScaleWidth - 1180
    lblStatus.Width = cmdRefresh.Left - lblStatus.Left - 120
    UpdateScrollbar
    DrawGraph
End Sub

Private Sub Form_Load()
    Theme_ApplyIcon Me
End Sub
