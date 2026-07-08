VERSION 5.00
Begin VB.Form frmChanges
   Caption         =   "Git Changes - Modernizr"
   ClientHeight    =   7550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7550
   ScaleWidth      =   8000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh
      Caption         =   "&Refresh"
      Height          =   345
      Left            =   6820
      TabIndex        =   1
      Top             =   135
      Width           =   1000
   End
   Begin VB6Modernizr.ucList lstUnstaged
      Height          =   1815
      Left            =   180
      TabIndex        =   3
      Top             =   870
      Width           =   7640
      _ExtentX        =   13476
      _ExtentY        =   3202
   End
   Begin VB.CommandButton cmdStage
      Caption         =   "Stage &Selected"
      Height          =   345
      Left            =   180
      TabIndex        =   4
      Top             =   2805
      Width           =   1500
   End
   Begin VB.CommandButton cmdStageAll
      Caption         =   "Stage &All"
      Height          =   345
      Left            =   1760
      TabIndex        =   5
      Top             =   2805
      Width           =   1200
   End
   Begin VB.CommandButton cmdUnstage
      Caption         =   "&Unstage Selected"
      Height          =   345
      Left            =   3260
      TabIndex        =   6
      Top             =   2805
      Width           =   1700
   End
   Begin VB.CommandButton cmdUnstageAll
      Caption         =   "U&nstage All"
      Height          =   345
      Left            =   5040
      TabIndex        =   7
      Top             =   2805
      Width           =   1300
   End
   Begin VB6Modernizr.ucList lstStaged
      Height          =   1815
      Left            =   180
      TabIndex        =   9
      Top             =   3585
      Width           =   7640
      _ExtentX        =   13476
      _ExtentY        =   3202
   End
   Begin VB.TextBox txtMsg
      Height          =   630
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   5560
      Width           =   6440
   End
   Begin VB.CommandButton cmdCommit
      Caption         =   "&Commit"
      Height          =   630
      Left            =   6820
      TabIndex        =   11
      Top             =   5560
      Width           =   1000
   End
   Begin VB.Label lblBranch
      Caption         =   ""
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   6600
   End
   Begin VB.Label lblUnstaged
      Caption         =   "Unstaged / untracked  (double-click opens the file):"
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   600
      Width           =   7640
   End
   Begin VB.Label lblStaged
      Caption         =   "Staged  (will be committed):"
      Height          =   240
      Left            =   180
      TabIndex        =   8
      Top             =   3315
      Width           =   7640
   End
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   615
      Left            =   180
      TabIndex        =   12
      Top             =   6340
      Width           =   7640
   End
End
Attribute VB_Name = "frmChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Git Changes (Ctrl+Shift+G): staged and unstaged lists built from
'  the porcelain XY codes (X = index, Y = worktree). Stage/unstage
'  selected or all; Commit commits the staged set. A partially staged
'  file appears in both lists. Double-click opens the file.
' =====================================================================

Private mUnPaths() As String
Private mUnCount As Long
Private mStPaths() As String
Private mStCount As Long
Private mLastSig As String     ' last rendered status, to skip no-op rebuilds

Public Sub ShowChanges()
    On Error Resume Next
    Load Me
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    RefreshList
    Me.Show vbModeless
End Sub

Private Sub Form_Load()
    lstUnstaged.MultiSelect = True
    lstStaged.MultiSelect = True
End Sub

Public Sub RefreshList()
    On Error Resume Next
    If Not Git_HasRepo() Then
        lstUnstaged.Clear
        lstStaged.Clear
        mUnCount = 0
        mStCount = 0
        mLastSig = ""
        lblBranch.Caption = "No git repository found for the active project."
        lblStatus.Caption = ""
        Exit Sub
    End If

    lblBranch.Caption = "Branch: " & Git_Branch() & _
        IIf(Git_RepoDirty(), "   (uncommitted changes)", "   (clean)")

    ' build both lists into temp arrays first, so an unchanged status
    ' (the ~5s background poll) never resets selection or scrolling
    Dim st() As String, pth() As String, n As Long, i As Long
    n = Git_ChangedList(st, pth)

    Dim unDisp() As String, unPath() As String, unN As Long
    Dim stDisp() As String, stPath() As String, stN As Long
    ReDim unDisp(0 To n + 1)
    ReDim unPath(0 To n + 1)
    ReDim stDisp(0 To n + 1)
    ReDim stPath(0 To n + 1)

    Dim x As String, y As String, rel As String, sig As String
    For i = 0 To n - 1
        x = Left$(st(i), 1)
        y = Mid$(st(i) & " ", 2, 1)
        rel = Mid$(pth(i), Len(Git_RepoRoot()) + 2)
        If x <> " " And x <> "?" Then
            stDisp(stN) = "[" & x & "]  " & rel
            stPath(stN) = pth(i)
            stN = stN + 1
        End If
        If y <> " " Then
            unDisp(unN) = "[" & y & "]  " & rel
            unPath(unN) = pth(i)
            unN = unN + 1
        End If
        sig = sig & st(i) & "|" & pth(i) & ";"
    Next

    If sig <> mLastSig Then
        mLastSig = sig
        RebuildPreserving lstUnstaged, unDisp, unPath, unN
        RebuildPreserving lstStaged, stDisp, stPath, stN
        mUnPaths = unPath
        mUnCount = unN
        mStPaths = stPath
        mStCount = stN
    End If

    lblStatus.Caption = mUnCount & " unstaged, " & mStCount & _
        " staged.  M=modified, A=added, D=deleted, R=renamed, ?=untracked."
End Sub

' Reload a list but keep what the user had: selected paths, the
' focused row and the scroll position survive a content change.
Private Sub RebuildPreserving(lst As ucList, disp() As String, _
        pths() As String, ByVal cnt As Long)
    On Error Resume Next
    Dim selKeys As New Collection, focusKey As String, topIdx As Long
    Dim i As Long, k As String

    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then selKeys.Add lst.ItemKey(i), lst.ItemKey(i)
    Next
    If lst.ListIndex >= 0 Then focusKey = lst.ItemKey(lst.ListIndex)
    topIdx = lst.TopIndex

    lst.Clear
    For i = 0 To cnt - 1
        lst.AddItem disp(i), pths(i), pths(i)
    Next

    For i = 0 To cnt - 1
        k = lst.ItemKey(i)
        Err.Clear
        selKeys.Item k
        If Err.Number = 0 Then lst.Selected(i) = True
        If Len(focusKey) > 0 And k = focusKey Then lst.ListIndex = i
    Next
    lst.TopIndex = topIdx
End Sub

' ---------------------------------------------------------------------
'  Stage / unstage
' ---------------------------------------------------------------------

Private Sub cmdStage_Click()
    On Error Resume Next
    Dim sel As Collection
    Set sel = SelectedOf(lstUnstaged, mUnPaths)
    If sel.Count = 0 Then
        lblStatus.Caption = "Select file(s) in the unstaged list first."
        Exit Sub
    End If
    Git_StageFiles sel
    RefreshList
End Sub

Private Sub cmdStageAll_Click()
    On Error Resume Next
    Git_StageAll
    RefreshList
End Sub

Private Sub cmdUnstage_Click()
    On Error Resume Next
    Dim sel As Collection
    Set sel = SelectedOf(lstStaged, mStPaths)
    If sel.Count = 0 Then
        lblStatus.Caption = "Select file(s) in the staged list first."
        Exit Sub
    End If
    Git_UnstageFiles sel
    RefreshList
End Sub

Private Sub cmdUnstageAll_Click()
    On Error Resume Next
    Git_UnstageAll
    RefreshList
End Sub

Private Function SelectedOf(lst As ucList, paths() As String) As Collection
    Dim c As New Collection, i As Long
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then c.Add paths(i)
    Next
    Set SelectedOf = c
End Function

' ---------------------------------------------------------------------
'  Commit
' ---------------------------------------------------------------------

Private Sub cmdCommit_Click()
    On Error Resume Next
    Dim msg As String
    msg = Trim$(txtMsg.Text)
    If Len(msg) = 0 Then
        lblStatus.Caption = "Enter a commit message first."
        Exit Sub
    End If
    If mStCount = 0 Then
        lblStatus.Caption = "Nothing staged. Stage files first (or Stage All)."
        Exit Sub
    End If

    cmdCommit.Enabled = False
    lblStatus.Caption = "Committing..."
    Dim res As String
    res = Git_CommitStaged(msg)
    cmdCommit.Enabled = True

    Dim p As Long
    p = InStr(res, vbLf)
    If p > 0 Then p = InStr(p + 1, res, vbLf)
    If p > 0 Then res = Left$(res, p - 1)
    lblStatus.Caption = Replace$(Replace$(res, vbCr, ""), vbLf, "  |  ")
    txtMsg.Text = ""
    RefreshList
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    lblStatus.Caption = "Refreshing..."
    Git_StatusRefreshSync
    RefreshList
End Sub

' git background status refresh completed
Public Sub NotifyGitChanged()
    On Error Resume Next
    If Me.Visible Then RefreshList
End Sub

' ---------------------------------------------------------------------
'  Open file / window plumbing
' ---------------------------------------------------------------------

Private Sub lstUnstaged_DblClick()
    OpenFromList lstUnstaged, mUnPaths, mUnCount
End Sub

Private Sub lstStaged_DblClick()
    OpenFromList lstStaged, mStPaths, mStCount
End Sub

Private Sub OpenFromList(lst As ucList, paths() As String, ByVal cnt As Long)
    On Error Resume Next
    Dim i As Long
    i = lst.ListIndex
    If i < 0 Or i >= cnt Then Exit Sub

    Dim comp As VBIDE.VBComponent
    Set comp = CompByPath(paths(i))
    If comp Is Nothing Then
        Shell "notepad.exe """ & paths(i) & """", vbNormalFocus
    Else
        Dim cp As VBIDE.CodePane
        Set cp = comp.CodeModule.CodePane
        cp.Show
        cp.Window.SetFocus
    End If
End Sub

Private Function CompByPath(ByVal path As String) As VBIDE.VBComponent
    On Error Resume Next
    Dim proj As VBIDE.VBProject, comp As VBIDE.VBComponent, j As Long
    For Each proj In gVBE.VBProjects
        For Each comp In proj.VBComponents
            For j = 1 To comp.FileCount
                If StrComp(comp.FileNames(j), path, vbTextCompare) = 0 Then
                    Set CompByPath = comp
                    Exit Function
                End If
            Next
        Next
    Next
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Dim w As Long, extra As Long
    w = Me.ScaleWidth - MARGIN_STD * 2
    ' distribute extra height between the two lists
    extra = (Me.ScaleHeight - 7550) \ 2
    If extra < -600 Then extra = -600

    lstUnstaged.Width = w
    lstUnstaged.Height = 1815 + extra

    cmdStage.Top = lstUnstaged.Top + lstUnstaged.Height + 120
    cmdStageAll.Top = cmdStage.Top
    cmdUnstage.Top = cmdStage.Top
    cmdUnstageAll.Top = cmdStage.Top

    lblStaged.Top = cmdStage.Top + 510
    lstStaged.Top = lblStaged.Top + 300
    lstStaged.Width = w
    lstStaged.Height = 1815 + extra

    txtMsg.Top = lstStaged.Top + lstStaged.Height + 160
    txtMsg.Width = Me.ScaleWidth - 1480
    cmdCommit.Top = txtMsg.Top
    cmdCommit.Left = Me.ScaleWidth - 1180
    cmdRefresh.Left = Me.ScaleWidth - 1180
    lblStatus.Top = txtMsg.Top + 780
    lblStatus.Width = w
End Sub
