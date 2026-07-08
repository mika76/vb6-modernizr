VERSION 5.00
Begin VB.Form frmBrowser
   Caption         =   "Code Browser - Modernizr"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFilter
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   2800
   End
   Begin VB.OptionButton optKind
      Caption         =   "&Procedures"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      Top             =   180
      Value           =   -1  'True
      Width           =   1400
   End
   Begin VB.OptionButton optKind
      Caption         =   "&TODOs"
      Height          =   255
      Index           =   1
      Left            =   4620
      TabIndex        =   2
      Top             =   180
      Width           =   1100
   End
   Begin VB.CommandButton cmdRefresh
      Caption         =   "&Refresh"
      Height          =   345
      Left            =   6020
      TabIndex        =   3
      Top             =   140
      Width           =   1000
   End
   Begin VB6Modernizr.ucList lstItems
      Height          =   4095
      Left            =   180
      TabIndex        =   4
      Top             =   630
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   7223
   End
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   4980
      Width           =   6840
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Code Browser (Ctrl+Shift+O): every procedure or TODO comment in the
'  active project, with live text filtering. Double-click / Enter
'  jumps to the item.
' =====================================================================

Private mComp() As String
Private mLine() As Long
Private mCount As Long
Private mRX As Object      ' proc-definition regex, built once

Public Sub ShowBrowser()
    On Error Resume Next
    Load Me
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    RebuildList
    Me.Show vbModeless
    txtFilter.SetFocus
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
    lstItems.Width = Me.ScaleWidth - MARGIN_STD * 2
    lstItems.Height = Me.ScaleHeight - lstItems.Top - 560
    lblStatus.Top = Me.ScaleHeight - 420
    lblStatus.Width = Me.ScaleWidth - MARGIN_STD * 2
End Sub

Private Sub txtFilter_Change()
    RebuildList
End Sub

Private Sub optKind_Click(Index As Integer)
    RebuildList
End Sub

Private Sub cmdRefresh_Click()
    RebuildList
End Sub

' ---------------------------------------------------------------------

Private Sub RebuildList()
    On Error Resume Next
    lstItems.Clear
    mCount = 0
    ReDim mComp(0 To 64)
    ReDim mLine(0 To 64)

    Dim proj As VBIDE.VBProject
    Set proj = gVBE.ActiveVBProject
    If proj Is Nothing Then Exit Sub

    Dim flt As String
    flt = LCase$(Trim$(txtFilter.Text))

    Dim comp As VBIDE.VBComponent, cm As VBIDE.CodeModule
    Dim i As Long, s As String
    For Each comp In proj.VBComponents
        Set cm = comp.CodeModule
        For i = 1 To cm.CountOfLines
            s = cm.lines(i, 1)
            If optKind(0).Value Then
                ScanProc comp.Name, i, s, flt
            Else
                ScanTodo comp.Name, i, s, flt
            End If
        Next
    Next

    lblStatus.Caption = mCount & " item(s) in project '" & proj.Name & "'."
End Sub

Private Sub ScanProc(ByVal comp As String, ByVal lineNo As Long, _
        ByVal s As String, ByVal flt As String)
    On Error Resume Next
    If mRX Is Nothing Then
        Set mRX = CreateObject("VBScript.RegExp")
        mRX.IgnoreCase = True
        mRX.Pattern = "^[ \t]*(?:(?:public|private|friend|static)[ \t]+)*" & _
            "(sub|function|property[ \t]+(?:get|let|set))[ \t]+([a-z0-9_]+)"
    End If

    Dim mc As Object
    Set mc = mRX.Execute(s)
    If mc.Count = 0 Then Exit Sub

    Dim nm As String, kind As String
    kind = mc(0).SubMatches(0)
    nm = mc(0).SubMatches(1)
    AddEntry comp, lineNo, comp & "." & nm & "   [" & LCase$(kind) & _
             ", line " & lineNo & "]", flt
End Sub

Private Sub ScanTodo(ByVal comp As String, ByVal lineNo As Long, _
        ByVal s As String, ByVal flt As String)
    Dim apos As Long, p As Long
    apos = InStr(s, "'")
    If apos = 0 Then Exit Sub
    p = InStr(apos, LCase$(s), "todo")
    If p = 0 Then Exit Sub
    AddEntry comp, lineNo, comp & "(" & lineNo & "): " & _
             Trim$(Mid$(s, apos)), flt
End Sub

Private Sub AddEntry(ByVal comp As String, ByVal lineNo As Long, _
        ByVal display As String, ByVal flt As String)
    If Len(flt) > 0 Then
        If InStr(1, display, flt, vbTextCompare) = 0 Then Exit Sub
    End If
    If mCount > UBound(mComp) Then
        ReDim Preserve mComp(0 To mCount * 2)
        ReDim Preserve mLine(0 To mCount * 2)
    End If
    mComp(mCount) = comp
    mLine(mCount) = lineNo
    mCount = mCount + 1
    ' colored dot: blue = procedure, orange = TODO
    lstItems.AddItem display, , , _
        IIf(optKind(0).Value, THEME_CODE, THEME_ACCENT)
End Sub

' ---------------------------------------------------------------------

Private Sub lstItems_DblClick()
    JumpToSelection
End Sub

Private Sub lstItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        JumpToSelection
    End If
End Sub

Private Sub JumpToSelection()
    On Error Resume Next
    Dim i As Long
    i = lstItems.ListIndex
    If i < 0 Or i >= mCount Then Exit Sub

    Dim m As MatchInfo
    m.Proj = gVBE.ActiveVBProject.Name
    m.Comp = mComp(i)
    m.LineNum = mLine(i)
    m.Col = 1
    m.MatchLen = 0
    GoToMatch m
End Sub
