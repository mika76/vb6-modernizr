VERSION 5.00
Begin VB.Form frmBranches
   Caption         =   "Switch Branch - Modernizr"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFilter
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5640
   End
   Begin VB6Modernizr.ucList lstBranches
      Height          =   3600
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   6350
   End
   Begin VB.CommandButton cmdSwitch
      Caption         =   "&Switch"
      Height          =   345
      Left            =   3700
      TabIndex        =   3
      Top             =   4920
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4780
      TabIndex        =   4
      Top             =   4920
      Width           =   1040
   End
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   540
      Left            =   180
      TabIndex        =   2
      Top             =   4320
      Width           =   5640
   End
End
Attribute VB_Name = "frmBranches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Branch picker. Lists the repository's local branches (most recently
'  committed first) and then its remote-tracking ones; type to filter,
'  Enter or double-click to switch. The current branch is listed with
'  a green dot and cannot be picked.
'
'  Everything that makes a switch safe - and the whole business of the
'  IDE holding stale copies of files the checkout just rewrote - lives
'  in modBranch. This form only chooses a name.
' =====================================================================

' snapshot of the branch list, taken each time the picker opens
Private mName() As String
Private mRemote() As Boolean
Private mCur() As Boolean
Private mTotal As Long

Private mIdx() As Long          ' visible row -> snapshot index
Private mRows As Long

Public Sub ShowPicker()
    On Error Resume Next
    Load Me
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    Snapshot
    If Len(txtFilter.Text) > 0 Then
        txtFilter.Text = ""              ' fires Change -> RebuildList
    Else
        RebuildList
    End If
    Me.Show vbModeless
    txtFilter.SetFocus
End Sub

Private Sub Snapshot()
    On Error Resume Next
    Dim cur As String
    mTotal = Git_BranchList(mName, mRemote, mCur)
    cur = Git_Branch()
    If Len(cur) = 0 Then cur = "(unknown)"
    lblStatus.Caption = "On " & cur & ".  Switching checks out the branch " & _
        "and reloads the project from disk - open windows are reopened. " & _
        "Unsaved IDE changes block the switch."
End Sub

' ---------------------------------------------------------------------
'  List
' ---------------------------------------------------------------------

Private Sub RebuildList()
    On Error Resume Next
    Dim q As String, i As Long, sel As Long, cap As String
    q = Trim$(txtFilter.Text)

    lstBranches.Clear
    ReDim mIdx(0 To mTotal + 1)
    mRows = 0
    sel = -1

    For i = 0 To mTotal - 1
        If Len(q) = 0 Or InStr(1, mName(i), q, vbTextCompare) > 0 Then
            cap = mName(i)
            If mCur(i) Then
                cap = cap & "   - current"
            ElseIf mRemote(i) Then
                cap = cap & "   - remote"
            End If
            lstBranches.AddItem cap, mName(i), , BranchDot(i)
            mIdx(mRows) = i
            ' land on something switchable, so Enter does the obvious
            If sel < 0 And Not mCur(i) Then sel = mRows
            mRows = mRows + 1
        End If
    Next

    If mRows > 0 Then
        If sel < 0 Then sel = 0
        lstBranches.ListIndex = sel
    End If
    cmdSwitch.Enabled = (mRows > 0)
End Sub

' Green = checked out, orange = a remote-tracking ref that would get a
' local branch created for it, gray = an ordinary local branch.
' Literal colors, not system ones: ucList takes a negative dot color
' to mean "no dot", and every vb* system constant is negative.
Private Function BranchDot(ByVal i As Long) As Long
    If mCur(i) Then
        BranchDot = &H208020
    ElseIf mRemote(i) Then
        BranchDot = THEME_ACCENT
    Else
        BranchDot = &H909090
    End If
End Function

' ---------------------------------------------------------------------
'  Input
' ---------------------------------------------------------------------

Private Sub txtFilter_Change()
    RebuildList
End Sub

Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i As Long
    Select Case KeyCode
    Case vbKeyDown, vbKeyUp, vbKeyPageDown, vbKeyPageUp
        i = lstBranches.ListIndex
        Select Case KeyCode
            Case vbKeyDown:     i = i + 1
            Case vbKeyUp:       i = i - 1
            Case vbKeyPageDown: i = i + 10
            Case vbKeyPageUp:   i = i - 10
        End Select
        If i < 0 Then i = 0
        If i >= lstBranches.ListCount Then i = lstBranches.ListCount - 1
        If lstBranches.ListCount > 0 Then lstBranches.ListIndex = i
        KeyCode = 0
    Case vbKeyReturn
        KeyCode = 0
        DoSwitch
    End Select
End Sub

' swallow the ding for keys KeyDown already handled
Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then KeyAscii = 0
End Sub

Private Sub lstBranches_Click()
    On Error Resume Next
    txtFilter.SetFocus                   ' click selects, typing continues
End Sub

Private Sub lstBranches_DblClick()
    DoSwitch
End Sub

Private Sub lstBranches_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        DoSwitch
    End If
End Sub

Private Sub cmdSwitch_Click()
    DoSwitch
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub DoSwitch()
    On Error Resume Next
    Dim r As Long, i As Long
    r = lstBranches.ListIndex
    If r < 0 Or r >= mRows Then Beep: Exit Sub
    i = mIdx(r)

    If mCur(i) Then
        MsgBox "'" & mName(i) & "' is already checked out.", _
               vbInformation, "Switch Branch"
        Exit Sub
    End If

    ' out of the way before the project closes and reopens
    Me.Hide
    Branch_SwitchTo mName(i), mRemote(i)
End Sub

' ---------------------------------------------------------------------
'  Form
' ---------------------------------------------------------------------

Private Sub Form_Load()
    Theme_ApplyIcon Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Me.Hide
    End If
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
    Dim w As Long
    w = Me.ScaleWidth - MARGIN_STD * 2
    If w < 1500 Then w = 1500

    txtFilter.Width = w
    lstBranches.Width = w
    lstBranches.Height = Me.ScaleHeight - lstBranches.Top - 1245
    If lstBranches.Height < 600 Then lstBranches.Height = 600

    lblStatus.Top = lstBranches.Top + lstBranches.Height + 120
    lblStatus.Width = w

    cmdCancel.Top = Me.ScaleHeight - 525
    cmdCancel.Left = Me.ScaleWidth - MARGIN_STD - cmdCancel.Width
    cmdSwitch.Top = cmdCancel.Top
    cmdSwitch.Left = cmdCancel.Left - cmdSwitch.Width - 80
End Sub
