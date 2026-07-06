VERSION 5.00
Begin VB.Form frmChanges
   Caption         =   "Git Changes - Modernizr"
   ClientHeight    =   5800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5800
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
   Begin VB.ListBox lstFiles
      Height          =   3400
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   7760
   End
   Begin VB.TextBox txtMsg
      Height          =   630
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4180
      Width           =   6600
   End
   Begin VB.CommandButton cmdCommit
      Caption         =   "&Commit All"
      Height          =   630
      Left            =   6860
      TabIndex        =   4
      Top             =   4180
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
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4980
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
'  Git Changes (Ctrl+Shift+G): modified files from the status cache;
'  double-click opens the file in the IDE. Type a message and Commit
'  All to stage everything and commit.
' =====================================================================

Private Const LB_SETHORIZONTALEXTENT As Long = &H194

Private mPaths() As String
Private mCount As Long

Public Sub ShowChanges()
    On Error Resume Next
    Load Me
    SendMessageA lstFiles.hwnd, LB_SETHORIZONTALEXTENT, 3000, 0
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    RefreshList
    Me.Show vbModeless
End Sub

Public Sub RefreshList()
    On Error Resume Next
    lstFiles.Clear
    mCount = 0

    If Not Git_HasRepo() Then
        lblBranch.Caption = "No git repository found for the active project."
        lblStatus.Caption = ""
        Exit Sub
    End If

    lblBranch.Caption = "Branch: " & Git_Branch() & _
                        IIf(Git_RepoDirty(), "   (uncommitted changes)", "   (clean)")

    Dim st() As String, pth() As String, n As Long, i As Long
    n = Git_ChangedList(st, pth)
    ReDim mPaths(0 To n + 1)
    For i = 0 To n - 1
        mPaths(mCount) = pth(i)
        mCount = mCount + 1
        lstFiles.AddItem "[" & st(i) & "]  " & _
                         Mid$(pth(i), Len(Git_RepoRoot()) + 2)
    Next
    lblStatus.Caption = mCount & " changed file(s). " & _
        "Status: M=modified, A=added, D=deleted, ?=untracked."
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    Git_RefreshNow
    lblStatus.Caption = "Refreshing... (updates within a few seconds)"
End Sub

Private Sub cmdCommit_Click()
    On Error Resume Next
    Dim msg As String
    msg = Trim$(txtMsg.Text)
    If Len(msg) = 0 Then
        lblStatus.Caption = "Enter a commit message first."
        Exit Sub
    End If
    If mCount = 0 Then
        lblStatus.Caption = "Nothing to commit."
        Exit Sub
    End If

    cmdCommit.Enabled = False
    lblStatus.Caption = "Committing..."
    Dim res As String
    res = Git_CommitAll(msg)
    cmdCommit.Enabled = True

    ' first ~2 lines of git's answer as feedback
    Dim p As Long
    p = InStr(res, vbLf)
    If p > 0 Then p = InStr(p + 1, res, vbLf)
    If p > 0 Then res = Left$(res, p - 1)
    lblStatus.Caption = Replace$(Replace$(res, vbCr, ""), vbLf, "  |  ")
    txtMsg.Text = ""
End Sub

' git status refresh completed (called from modGit via frmTabs)
Public Sub NotifyGitChanged()
    On Error Resume Next
    If Me.Visible Then RefreshList
End Sub

Private Sub lstFiles_DblClick()
    On Error Resume Next
    Dim i As Long
    i = lstFiles.ListIndex
    If i < 0 Or i >= mCount Then Exit Sub

    Dim comp As VBIDE.VBComponent
    Set comp = CompByPath(mPaths(i))
    If comp Is Nothing Then
        Shell "notepad.exe """ & mPaths(i) & """", vbNormalFocus
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
    lstFiles.Width = Me.ScaleWidth - 240
    lstFiles.Height = Me.ScaleHeight - lstFiles.Top - 1750
    txtMsg.Top = Me.ScaleHeight - 1620
    txtMsg.Width = Me.ScaleWidth - 1400
    cmdCommit.Top = txtMsg.Top
    cmdCommit.Left = Me.ScaleWidth - 1140
    cmdRefresh.Left = Me.ScaleWidth - 1140
    lblStatus.Top = Me.ScaleHeight - 820
    lblStatus.Width = Me.ScaleWidth - 240
End Sub
