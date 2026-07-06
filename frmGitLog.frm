VERSION 5.00
Begin VB.Form frmGitLog
   Caption         =   "Git Log - Modernizr"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAll
      Caption         =   "All &branches"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1500
   End
   Begin VB.CommandButton cmdRefresh
      Caption         =   "&Refresh"
      Height          =   345
      Left            =   9060
      TabIndex        =   1
      Top             =   105
      Width           =   1000
   End
   Begin VB.ListBox lstLog
      BeginProperty Font
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   120
      TabIndex        =   2
      Top             =   570
      Width           =   9960
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
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   4530
      Width           =   9960
   End
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   240
      Left            =   1800
      TabIndex        =   4
      Top             =   165
      Width           =   7100
   End
End
Attribute VB_Name = "frmGitLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Git Log (Ctrl+Shift+L). git renders the graph (--graph ASCII art,
'  monospace list keeps it aligned); fields are tab-separated via a
'  custom pretty format. Click a commit to see its details/diffstat
'  below; graph-only continuation lines are not clickable.
' =====================================================================

Private Const LB_SETHORIZONTALEXTENT As Long = &H194
Private Const MAX_COMMITS As Long = 300

Private mHash() As String     ' parallel to list rows, "" = graph line

Public Sub ShowLog()
    On Error Resume Next
    Load Me
    SendMessageA lstLog.hwnd, LB_SETHORIZONTALEXTENT, 5000, 0
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    Me.Show vbModeless
    LoadLog
End Sub

Private Sub cmdRefresh_Click()
    LoadLog
End Sub

Private Sub chkAll_Click()
    LoadLog
End Sub

Private Sub LoadLog()
    On Error Resume Next
    lstLog.Clear
    txtDetails.Text = ""
    ReDim mHash(0 To 64)

    If Not Git_HasRepo() Then
        lblStatus.Caption = "No git repository found for the active project."
        Exit Sub
    End If

    lblStatus.Caption = "Loading..."
    DoEvents

    Dim args As String
    args = "log --graph -n " & MAX_COMMITS & " --date=short " & _
           "--pretty=format:" & Chr$(34) & _
           "%x09%h%x09%ad%x09%an%x09%d %s" & Chr$(34)
    If chkAll.Value = vbChecked Then args = args & " --all"

    Dim res As String
    res = Git_RunSync(args, 15000)
    If Len(res) = 0 Then
        lblStatus.Caption = "No output from git log (empty repo?)."
        Exit Sub
    End If

    Dim lines() As String, i As Long, s As String, commits As Long
    lines = Split(res, vbLf)
    For i = 0 To UBound(lines)
        s = Replace$(lines(i), vbCr, "")
        If Len(s) > 0 Then
            Dim p() As String
            p = Split(s, vbTab)
            If lstLog.ListCount > UBound(mHash) Then
                ReDim Preserve mHash(0 To lstLog.ListCount * 2)
            End If
            If UBound(p) >= 4 Then
                ' graph | hash | date | author | decorations+subject
                mHash(lstLog.ListCount) = p(1)
                lstLog.AddItem p(0) & " " & p(1) & "  " & p(2) & "  " & _
                    PadR(p(3), 14) & " " & Trim$(p(4))
                commits = commits + 1
            Else
                ' pure graph continuation line ( |/  |\  etc.)
                mHash(lstLog.ListCount) = ""
                lstLog.AddItem p(0)
            End If
        End If
    Next

    lblStatus.Caption = commits & " commit(s)" & _
        IIf(commits >= MAX_COMMITS, " (showing latest " & MAX_COMMITS & ")", "") & _
        ".  Click a commit for details."
End Sub

Private Function PadR(ByVal s As String, ByVal n As Long) As String
    If Len(s) > n Then s = Left$(s, n - 1) & Chr$(133)
    PadR = Left$(s & Space$(n), n)
End Function

Private Sub lstLog_Click()
    On Error Resume Next
    Dim i As Long
    i = lstLog.ListIndex
    If i < 0 Then Exit Sub
    If Len(mHash(i)) = 0 Then Exit Sub

    Dim res As String
    res = Git_RunSync("show --stat --date=short " & mHash(i), 10000)
    If Len(res) > 28000 Then res = Left$(res, 28000) & vbCrLf & "[...]"
    txtDetails.Text = Replace$(Replace$(res, vbCrLf, vbLf), vbLf, vbCrLf)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Dim w As Long, listH As Long
    w = Me.ScaleWidth - 240
    listH = (Me.ScaleHeight - 900) * 2 \ 3
    If listH < 600 Then listH = 600

    lstLog.Width = w
    lstLog.Height = listH
    txtDetails.Top = lstLog.Top + listH + 120
    txtDetails.Width = w
    txtDetails.Height = Me.ScaleHeight - txtDetails.Top - 120
    cmdRefresh.Left = Me.ScaleWidth - 1140
    lblStatus.Width = cmdRefresh.Left - lblStatus.Left - 120
End Sub
