VERSION 5.00
Begin VB.Form frmFind 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   60
      Width           =   2900
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   315
      Left            =   3560
      TabIndex        =   2
      Top             =   60
      Width           =   350
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   315
      Left            =   3930
      TabIndex        =   3
      Top             =   60
      Width           =   350
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Aa"
      Height          =   285
      Left            =   4360
      TabIndex        =   4
      ToolTipText     =   "Match case"
      Top             =   75
      Width           =   560
   End
   Begin VB.CheckBox chkWhole 
      Caption         =   "Word"
      Height          =   285
      Left            =   4940
      TabIndex        =   5
      ToolTipText     =   "Whole word"
      Top             =   75
      Width           =   760
   End
   Begin VB.CheckBox chkRegex 
      Caption         =   ".*"
      Height          =   285
      Left            =   5720
      TabIndex        =   6
      ToolTipText     =   "Regular expression"
      Top             =   75
      Width           =   540
   End
   Begin VB.ComboBox cboScope 
      Height          =   315
      Left            =   6320
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   60
      Width           =   1900
   End
   Begin VB.CommandButton cmdHighlight 
      Caption         =   "Highlight"
      Height          =   315
      Left            =   8280
      TabIndex        =   8
      ToolTipText     =   "Highlight all matches"
      Top             =   60
      Width           =   1000
   End
   Begin VB.CommandButton cmdCloseBar 
      Caption         =   "X"
      Height          =   315
      Left            =   11580
      TabIndex        =   9
      ToolTipText     =   "Close (Esc)"
      Top             =   60
      Width           =   350
   End
   Begin VB.ComboBox cboRepl 
      Height          =   315
      Left            =   600
      TabIndex        =   11
      Top             =   450
      Width           =   2900
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   315
      Left            =   3560
      TabIndex        =   12
      Top             =   450
      Width           =   900
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Height          =   315
      Left            =   4490
      TabIndex        =   13
      Top             =   450
      Width           =   1100
   End
   Begin VB.CommandButton cmdClearHL 
      Caption         =   "Clear"
      Height          =   315
      Left            =   5720
      TabIndex        =   14
      ToolTipText     =   "Clear highlights"
      Top             =   450
      Width           =   800
   End
   Begin VB.Label lblFind 
      BackStyle       =   0  'Transparent
      Caption         =   "Find:"
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   500
   End
   Begin VB.Label lblRepl 
      BackStyle       =   0  'Transparent
      Caption         =   "With:"
      Height          =   240
      Left            =   60
      TabIndex        =   10
      Top             =   495
      Width           =   500
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   6660
      TabIndex        =   15
      Top             =   495
      Width           =   4000
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Docked find/replace bar, VS-style: lives in a strip above the code
'  area (below the tab bar). Ctrl+F opens it, Enter = find next,
'  Shift+Enter = find previous, F3/Shift+F3 work from the code window,
'  Esc closes it. Regex via VBScript.RegExp; "Highlight" outlines all
'  matches in the editor + scrollbar (see modHighlight).
' =====================================================================

Private mLoadedUI As Boolean

' ---------------------------------------------------------------------
'  Docking plumbing
' ---------------------------------------------------------------------

Public Sub Attach()
    On Error Resume Next
    EnsureUI
    If GetParent(Me.hwnd) = MainHwnd() Then Exit Sub
    Dim style As Long
    style = GetWindowLongA(Me.hwnd, GWL_STYLE)
    style = (style Or WS_CHILD) And (Not WS_POPUP)
    SetWindowLongA Me.hwnd, GWL_STYLE, style
    SetParent Me.hwnd, MainHwnd()
End Sub

Public Sub Reposition()
    On Error Resume Next
    If Not gFindBarVisible Then Exit Sub
    Dim x As Long, y As Long, w As Long
    If Not Layout_StripOrigin(x, y, w) Then Exit Sub
    If gTabBarVisible Then y = y + ScaleForDpi(TAB_BAR_HEIGHT)
    MoveWindow Me.hwnd, x, y, w, ScaleForDpi(FIND_BAR_HEIGHT), 1
End Sub

Private Sub EnsureUI()
    If mLoadedUI Then Exit Sub
    mLoadedUI = True
    cboScope.AddItem "Current module"    ' scCurrentDoc = 0
    cboScope.AddItem "Selection"         ' scSelection  = 1
    cboScope.AddItem "Open modules"      ' scOpenDocs   = 2
    cboScope.AddItem "Whole project"     ' scProject    = 3
    cboScope.ListIndex = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cmdCloseBar.Left = Me.ScaleWidth - cmdCloseBar.Width - 60
    lblStatus.Width = Me.ScaleWidth - lblStatus.Left - 500
End Sub

' ---------------------------------------------------------------------
'  Show / hide
' ---------------------------------------------------------------------

Public Sub ShowBar()
    On Error Resume Next
    EnsureUI
    PrefillFromSelection
    FindBar_Show
    SetFocusAPI cboFind.hwnd
End Sub

Public Sub HideBar()
    On Error Resume Next
    FindBar_Hide
    gVBE.ActiveCodePane.Window.SetFocus
End Sub

Private Sub cmdCloseBar_Click()
    HideBar
End Sub

Private Sub PrefillFromSelection()
    On Error Resume Next
    Dim cp As VBIDE.CodePane
    Set cp = gVBE.ActiveCodePane
    If cp Is Nothing Then Exit Sub
    Dim sl As Long, sc As Long, el As Long, ec As Long
    cp.GetSelection sl, sc, el, ec
    If sl = el And ec > sc Then
        cboFind.Text = Mid$(cp.CodeModule.lines(sl, 1), sc, ec - sc)
    End If
End Sub

' ---------------------------------------------------------------------
'  Options / scope
' ---------------------------------------------------------------------

Private Sub cboScope_Click()
    On Error Resume Next
    If cboScope.ListIndex = scSelection Then CaptureSelectionRange
End Sub

Private Sub CaptureSelectionRange()
    On Error Resume Next
    Dim cp As VBIDE.CodePane
    Set cp = gVBE.ActiveCodePane
    If cp Is Nothing Then Exit Sub
    cp.GetSelection gSelSL, gSelSC, gSelEL, gSelEC
    gSelComp = cp.CodeModule.Parent.Name
End Sub

Private Function CurScope() As SearchScope
    If cboScope.ListIndex >= 0 Then CurScope = cboScope.ListIndex
End Function

Private Sub ApplyOptions()
    gOptCase = (chkCase.Value = vbChecked)
    gOptWhole = (chkWhole.Value = vbChecked)
    gOptRegex = (chkRegex.Value = vbChecked)
End Sub

Private Function Ready() As Boolean
    On Error Resume Next
    lblStatus.Caption = ""
    If Len(cboFind.Text) = 0 Then
        lblStatus.Caption = "Enter search text."
        Exit Function
    End If
    If gVBE.ActiveCodePane Is Nothing And CurScope() < scOpenDocs Then
        lblStatus.Caption = "No active code window."
        Exit Function
    End If
    ApplyOptions
    If Not PrepareSearch(cboFind.Text) Then
        lblStatus.Caption = "Invalid regular expression."
        Exit Function
    End If
    AddHistory cboFind
    AddHistory cboRepl
    Ready = True
End Function

Private Sub AddHistory(cbo As ComboBox)
    On Error Resume Next
    Dim i As Long, s As String
    s = cbo.Text
    If Len(s) = 0 Then Exit Sub
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = s Then Exit Sub
    Next
    cbo.AddItem s, 0
    If cbo.ListCount > 20 Then cbo.RemoveItem 20
End Sub

' ---------------------------------------------------------------------
'  Find
' ---------------------------------------------------------------------

Private Sub cmdNext_Click()
    DoFind True
End Sub

Private Sub cmdPrev_Click()
    DoFind False
End Sub

Private Sub DoFind(ByVal forward As Boolean)
    On Error Resume Next
    If Not Ready() Then Exit Sub

    If CollectMatches(CurScope(), cboFind.Text) = 0 Then
        lblStatus.Caption = "Not found."
        Beep
        Exit Sub
    End If

    SelectMatch PickMatch(forward)
End Sub

' Choose the next/previous match relative to the current caret.
Private Function PickMatch(ByVal forward As Boolean) As Long
    On Error Resume Next
    Dim actComp As String, actProj As String
    Dim sl As Long, sc As Long, el As Long, ec As Long
    Dim cp As VBIDE.CodePane
    Set cp = gVBE.ActiveCodePane
    If Not cp Is Nothing Then
        cp.GetSelection sl, sc, el, ec
        actComp = cp.CodeModule.Parent.Name
        actProj = cp.CodeModule.Parent.Collection.Parent.Name
    End If

    Dim i As Long
    If forward Then
        For i = 0 To gMatchCount - 1
            If gMatches(i).comp = actComp And gMatches(i).proj = actProj Then
                If gMatches(i).lineNum > el Or _
                   (gMatches(i).lineNum = el And gMatches(i).col >= ec) Then
                    PickMatch = i
                    Exit Function
                End If
            End If
        Next
        For i = 0 To gMatchCount - 1
            If gMatches(i).comp <> actComp Or gMatches(i).proj <> actProj Then
                PickMatch = i
                Exit Function
            End If
        Next
        PickMatch = 0
    Else
        For i = gMatchCount - 1 To 0 Step -1
            If gMatches(i).comp = actComp And gMatches(i).proj = actProj Then
                If gMatches(i).lineNum < sl Or _
                   (gMatches(i).lineNum = sl And gMatches(i).col < sc) Then
                    PickMatch = i
                    Exit Function
                End If
            End If
        Next
        For i = gMatchCount - 1 To 0 Step -1
            If gMatches(i).comp <> actComp Or gMatches(i).proj <> actProj Then
                PickMatch = i
                Exit Function
            End If
        Next
        PickMatch = gMatchCount - 1
    End If
End Function

Private Sub SelectMatch(ByVal idx As Long)
    On Error Resume Next
    If idx < 0 Or idx >= gMatchCount Then Exit Sub
    Dim m As MatchInfo
    m = gMatches(idx)

    Dim cm As VBIDE.CodeModule, cp As VBIDE.CodePane
    Set cm = FindModule(m.proj, m.comp)
    If cm Is Nothing Then Exit Sub
    Set cp = cm.CodePane                 ' opens the pane if needed
    cp.Show
    cp.SetSelection m.lineNum, m.col, m.lineNum, m.col + m.MatchLen

    If m.lineNum < cp.topLine Or _
       m.lineNum >= cp.topLine + cp.CountOfVisibleLines Then
        Dim t As Long
        t = m.lineNum - cp.CountOfVisibleLines \ 2
        If t < 1 Then t = 1
        cp.topLine = t
    End If

    lblStatus.Caption = "Match " & (idx + 1) & " of " & gMatchCount & _
                        "  (" & m.comp & ", line " & m.lineNum & ")"

    ' keep typing in the bar
    If Me.Visible Then SetFocusAPI cboFind.hwnd
End Sub

' ---------------------------------------------------------------------
'  Replace
' ---------------------------------------------------------------------

Private Sub cmdReplace_Click()
    On Error Resume Next
    If Not Ready() Then Exit Sub

    ' If the current selection is a match, replace it, then find next.
    Dim cp As VBIDE.CodePane
    Set cp = gVBE.ActiveCodePane
    If Not cp Is Nothing Then
        Dim sl As Long, sc As Long, el As Long, ec As Long
        cp.GetSelection sl, sc, el, ec
        If sl = el And ec > sc Then
            Dim sLine As String, sel As String
            sLine = cp.CodeModule.lines(sl, 1)
            sel = Mid$(sLine, sc, ec - sc)
            If SelectionIsMatch(sel) Then
                Dim newSel As String, n As Long
                newSel = ReplaceInLine(sel, cboFind.Text, cboRepl.Text, n)
                cp.CodeModule.ReplaceLine sl, _
                    Left$(sLine, sc - 1) & newSel & Mid$(sLine, ec)
                cp.SetSelection sl, sc, sl, sc + Len(newSel)
            End If
        End If
    End If

    DoFind True
End Sub

Private Function SelectionIsMatch(ByVal sel As String) As Boolean
    Dim cols() As Long, lens() As Long
    If FindInLine(sel, cboFind.Text, cols, lens) > 0 Then
        SelectionIsMatch = (cols(0) = 1 And lens(0) = Len(sel))
    End If
End Function

Private Sub cmdReplaceAll_Click()
    On Error Resume Next
    If Not Ready() Then Exit Sub
    If CurScope() = scSelection And gSelComp = "" Then CaptureSelectionRange

    Dim n As Long
    n = ReplaceAllInScope(CurScope(), cboFind.Text, cboRepl.Text)
    lblStatus.Caption = n & " occurrence(s) replaced."
    Highlight_Clear          ' any previous highlight is now stale
End Sub

' ---------------------------------------------------------------------
'  Keyboard
' ---------------------------------------------------------------------

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        HideBar
    End If
End Sub

Private Sub cboFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        DoFind Shift And vbShiftMask
    End If
End Sub

Private Sub cboRepl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        cmdReplace_Click
    End If
End Sub

' F3 / Shift+F3 from the code window (routed via modWheel's hook)
Public Sub DoFindPublic(ByVal forward As Boolean)
    DoFind forward
End Sub

' ---------------------------------------------------------------------
'  Highlight
' ---------------------------------------------------------------------

Private Sub cmdHighlight_Click()
    On Error Resume Next
    If Not Ready() Then Exit Sub
    Dim n As Long
    n = CollectMatches(CurScope(), cboFind.Text)
    Highlight_SetFromSearch
    lblStatus.Caption = n & " match(es) highlighted."
End Sub

Private Sub cmdClearHL_Click()
    Highlight_Clear
    lblStatus.Caption = "Highlights cleared."
End Sub
