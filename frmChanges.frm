VERSION 5.00
Begin VB.Form frmChanges
   Caption         =   "Git Changes - Modernizr"
   ClientHeight    =   7300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7300
   ScaleWidth      =   8000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh
      Caption         =   "&Refresh"
      Height          =   345
      Left            =   6860
      TabIndex        =   1
      Top             =   105
      Width           =   1000
   End
   Begin VB.ListBox lstUnstaged
      Height          =   1815
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   810
      Width           =   7760
   End
   Begin VB.CommandButton cmdStage
      Caption         =   "Stage &Selected"
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   2730
      Width           =   1500
   End
   Begin VB.CommandButton cmdStageAll
      Caption         =   "Stage &All"
      Height          =   345
      Left            =   1700
      TabIndex        =   5
      Top             =   2730
      Width           =   1200
   End
   Begin VB.CommandButton cmdUnstage
      Caption         =   "&Unstage Selected"
      Height          =   345
      Left            =   3200
      TabIndex        =   6
      Top             =   2730
      Width           =   1700
   End
   Begin VB.CommandButton cmdUnstageAll
      Caption         =   "U&nstage All"
      Height          =   345
      Left            =   4980
      TabIndex        =   7
      Top             =   2730
      Width           =   1300
   End
   Begin VB.ListBox lstStaged
      Height          =   1815
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      Top             =   3480
      Width           =   7760
   End
   Begin VB.TextBox txtMsg
      Height          =   630
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   5430
      Width           =   6600
   End
   Begin VB.CommandButton cmdCommit
      Caption         =   "&Commit"
      Height          =   630
      Left            =   6860
      TabIndex        =   11
      Top             =   5430
      Width           =   1000
   End
   Begin VB.Label lblBranch
      Caption         =   ""
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   6600
   End
   Begin VB.Label lblUnstaged
      Caption         =   "Unstaged / untracked  (double-click opens the file):"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   7700
   End
   Begin VB.Label lblStaged
      Caption         =   "Staged  (will be committed):"
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   3210
      Width           =   7700
   End
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   6180
      Width           =   7760
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

Private Const LB_SETHORIZONTALEXTENT As Long = &H194

Private mUnPaths() As String
Private mUnCount As Long
Private mStPaths() As String
Private mStCount As Long

Public Sub ShowChanges()
    On Error Resume Next
    Load Me
    SendMessageA lstUnstaged.hwnd, LB_SETHORIZONTALEXTENT, 3000, 0
    SendMessageA lstStaged.hwnd, LB_SETHORIZONTALEXTENT, 3000, 0
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    RefreshList
    Me.Show vbModeless
End Sub

Public Sub RefreshList()
    On Error Resume Next
    lstUnstaged.Clear
    lstStaged.Clear
    mUnCount = 0
    mStCount = 0

    If Not Git_HasRepo() Then
        lblBranch.Caption = "No git repository found for the active project."
        lblStatus.Caption = ""
        Exit Sub
    End If

    lblBranch.Caption = "Branch: " & Git_Branch() & _
        IIf(Git_RepoDirty(), "   (uncommitted changes)", "   (clean)")

    Dim st() As String, pth() As String, n As Long, i As Long
    n = Git_ChangedList(st, pth)
    ReDim mUnPaths(0 To n + 1)
    ReDim mStPaths(0 To n + 1)

    Dim x As String, y As String, rel As String
    For i = 0 To n - 1
        x = Left$(st(i), 1)
        y = Mid$(st(i) & " ", 2, 1)
        rel = Mid$(pth(i), Len(Git_RepoRoot()) + 2)
        If x <> " " And x <> "?" Then
            mStPaths(mStCount) = pth(i)
            mStCount = mStCount + 1
            lstStaged.AddItem "[" & x & "]  " & rel
        End If
        If y <> " " Then
            mUnPaths(mUnCount) = pth(i)
            mUnCount = mUnCount + 1
            lstUnstaged.AddItem "[" & y & "]  " & rel
        End If
    Next

    lblStatus.Caption = mUnCount & " unstaged, " & mStCount & _
        " staged.  M=modified, A=added, D=deleted, R=renamed, ?=untracked."
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

Private Function SelectedOf(lst As ListBox, paths() As String) As Collection
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

Private Sub OpenFromList(lst As ListBox, paths() As String, ByVal cnt As Long)
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
    w = Me.ScaleWidth - 240
    ' distribute extra height between the two lists
    extra = (Me.ScaleHeight - 7300) \ 2
    If extra < -600 Then extra = -600

    lstUnstaged.Width = w
    lstUnstaged.Height = 1815 + extra

    cmdStage.Top = lstUnstaged.Top + lstUnstaged.Height + 100
    cmdStageAll.Top = cmdStage.Top
    cmdUnstage.Top = cmdStage.Top
    cmdUnstageAll.Top = cmdStage.Top

    lblStaged.Top = cmdStage.Top + 480
    lstStaged.Top = lblStaged.Top + 270
    lstStaged.Width = w
    lstStaged.Height = 1815 + extra

    txtMsg.Top = lstStaged.Top + lstStaged.Height + 130
    txtMsg.Width = Me.ScaleWidth - 1400
    cmdCommit.Top = txtMsg.Top
    cmdCommit.Left = Me.ScaleWidth - 1140
    cmdRefresh.Left = Me.ScaleWidth - 1140
    lblStatus.Top = txtMsg.Top + 750
    lblStatus.Width = w
End Sub
