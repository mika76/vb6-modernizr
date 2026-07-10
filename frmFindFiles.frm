VERSION 5.00
Begin VB.Form frmFindFiles
   Caption         =   "Find in Files - Modernizr"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFind
      Height          =   315
      Left            =   1020
      TabIndex        =   1
      Top             =   150
      Width           =   4600
   End
   Begin VB.CheckBox chkCase
      Caption         =   "Match &case"
      Height          =   255
      Left            =   5820
      TabIndex        =   2
      Top             =   180
      Width           =   1500
   End
   Begin VB.CheckBox chkWhole
      Caption         =   "&Whole word"
      Height          =   255
      Left            =   7380
      TabIndex        =   3
      Top             =   180
      Width           =   1500
   End
   Begin VB.CheckBox chkRegex
      Caption         =   "Rege&x"
      Height          =   255
      Left            =   5820
      TabIndex        =   4
      Top             =   630
      Width           =   1500
   End
   Begin VB.TextBox txtPatterns
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Text            =   "*.bas;*.cls;*.frm;*.ctl;*.dob;*.pag;*.dsr;*.vbp;*.vbg"
      Top             =   600
      Width           =   4600
   End
   Begin VB.OptionButton optRoot
      Caption         =   "Active &project folder"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   1110
      Value           =   -1  'True
      Width           =   2200
   End
   Begin VB.OptionButton optRoot
      Caption         =   "F&older:"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Top             =   1110
      Width           =   900
   End
   Begin VB.TextBox txtFolder
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Top             =   1070
      Width           =   4340
   End
   Begin VB.CommandButton cmdBrowse
      Caption         =   "..."
      Height          =   315
      Left            =   7880
      TabIndex        =   10
      Top             =   1070
      Width           =   400
   End
   Begin VB.CommandButton cmdSearch
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   360
      Left            =   8420
      TabIndex        =   11
      Top             =   1050
      Width           =   1000
   End
   Begin VB6Modernizr.ucList lstResults
      Height          =   4300
      Left            =   180
      TabIndex        =   12
      Top             =   1560
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   7585
   End
   Begin VB.Label lblFind
      Caption         =   "Fi&nd:"
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   800
   End
   Begin VB.Label lblPat
      Caption         =   "&Files:"
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   645
      Width           =   800
   End
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   255
      Left            =   180
      TabIndex        =   13
      Top             =   6060
      Width           =   9240
   End
End
Attribute VB_Name = "frmFindFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Find in Files: scans source files on disk (code AND designer text)
'  under the project folder or any folder. Double-click a result to
'  jump to it; files that belong to the open project group open in the
'  IDE at the right line, anything else opens in Notepad.
' =====================================================================

' --- folder browse dialog ---
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Function SHBrowseForFolderA Lib "shell32" _
    (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Const BIF_RETURNONLYFSDIRS As Long = 1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40

' result -> location mapping (parallel to lstResults)
Private mFile() As String
Private mLine() As Long
Private mResCount As Long

Private mSearching As Boolean
Private mCancel As Boolean

' ---------------------------------------------------------------------

Public Sub ShowDialog()
    On Error Resume Next
    Load Me
    ' owned by the IDE main window, so it stays above it
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    Me.Show vbModeless
    txtFind.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        mCancel = True
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        mCancel = True           ' stop a running search too
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    lstResults.Width = Me.ScaleWidth - MARGIN_STD * 2
    lstResults.Height = Me.ScaleHeight - lstResults.Top - 560
    lblStatus.Top = Me.ScaleHeight - 420
    lblStatus.Width = Me.ScaleWidth - MARGIN_STD * 2
End Sub

Private Sub cmdBrowse_Click()
    On Error Resume Next
    Dim s As String
    s = BrowseFolder()
    If Len(s) > 0 Then
        txtFolder.Text = s
        optRoot(1).Value = True
    End If
End Sub

' ---------------------------------------------------------------------
'  Search
' ---------------------------------------------------------------------

Private Sub cmdSearch_Click()
    On Error Resume Next
    If mSearching Then
        mCancel = True
        Exit Sub
    End If

    Dim root As String
    root = SearchRoot()
    If Len(root) = 0 Then
        lblStatus.Caption = "Save the project first, or pick a folder."
        Exit Sub
    End If
    If Len(txtFind.Text) = 0 Then
        lblStatus.Caption = "Enter search text."
        Exit Sub
    End If

    gOptCase = (chkCase.Value = vbChecked)
    gOptWhole = (chkWhole.Value = vbChecked)
    gOptRegex = (chkRegex.Value = vbChecked)
    If Not PrepareSearch(txtFind.Text) Then
        lblStatus.Caption = "Invalid regular expression."
        Exit Sub
    End If

    lstResults.Clear
    mResCount = 0
    ReDim mFile(0 To 256)
    ReDim mLine(0 To 256)

    mSearching = True
    mCancel = False
    cmdSearch.Caption = "S&top"

    Dim files As New Collection
    CollectFiles root, LCase$(txtPatterns.Text), files

    Dim i As Long, hitFiles As Long, lastFile As String
    For i = 1 To files.Count
        If mCancel Then Exit For
        lblStatus.Caption = "Scanning " & i & "/" & files.Count & _
                            ": " & MidPath(CStr(files(i)), root)
        If SearchFile(CStr(files(i)), root) Then hitFiles = hitFiles + 1
        DoEvents
    Next

    cmdSearch.Caption = "&Search"
    mSearching = False
    lblStatus.Caption = mResCount & " match(es) in " & hitFiles & _
                        " file(s), " & files.Count & " file(s) scanned" & _
                        IIf(mCancel, " (stopped)", "") & "."
End Sub

Private Function SearchRoot() As String
    On Error Resume Next
    If optRoot(1).Value Then
        If Len(Dir$(txtFolder.Text, vbDirectory)) > 0 Then _
            SearchRoot = txtFolder.Text
    Else
        Dim f As String
        f = gVBE.ActiveVBProject.FileName
        If Len(f) > 0 Then SearchRoot = Left$(f, InStrRev(f, "\") - 1)
    End If
    If Right$(SearchRoot, 1) = "\" Then _
        SearchRoot = Left$(SearchRoot, Len(SearchRoot) - 1)
End Function

' Recursive file collection. Dir$() is not reentrant, so each
' directory pass finishes before recursing.
Private Sub CollectFiles(ByVal path As String, ByVal patterns As String, _
        files As Collection)
    On Error Resume Next
    Dim f As String, subdirs As New Collection

    f = Dir$(path & "\*", vbDirectory Or vbHidden Or vbSystem)
    Do While Len(f) > 0
        If f <> "." And f <> ".." Then
            If (GetAttr(path & "\" & f) And vbDirectory) <> 0 Then
                subdirs.Add path & "\" & f
            ElseIf MatchesPattern(LCase$(f), patterns) Then
                files.Add path & "\" & f
            End If
        End If
        f = Dir$()
    Loop

    Dim i As Long
    For i = 1 To subdirs.Count
        If mCancel Then Exit Sub
        CollectFiles CStr(subdirs(i)), patterns, files
    Next
End Sub

Private Function MatchesPattern(ByVal fileName As String, _
        ByVal patterns As String) As Boolean
    Dim p() As String, i As Long
    p = Split(patterns, ";")
    For i = 0 To UBound(p)
        If Len(Trim$(p(i))) > 0 Then
            If fileName Like Trim$(p(i)) Then
                MatchesPattern = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function SearchFile(ByVal path As String, ByVal root As String) As Boolean
    On Error GoTo Fail
    Dim ff As Integer, s As String, lineNo As Long
    Dim cols() As Long, lens() As Long
    Dim rel As String
    rel = MidPath(path, root)

    ff = FreeFile
    Open path For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, s
        lineNo = lineNo + 1
        If FindInLine(s, txtFind.Text, cols, lens) > 0 Then
            AddResult path, lineNo, rel & "(" & lineNo & "): " & _
                      Left$(Trim$(s), 250)
            SearchFile = True
        End If
    Loop
    Close #ff
    Exit Function
Fail:
    Close #ff
End Function

Private Sub AddResult(ByVal path As String, ByVal lineNo As Long, _
        ByVal display As String)
    If mResCount > UBound(mFile) Then
        ReDim Preserve mFile(0 To mResCount * 2)
        ReDim Preserve mLine(0 To mResCount * 2)
    End If
    mFile(mResCount) = path
    mLine(mResCount) = lineNo
    mResCount = mResCount + 1
    lstResults.AddItem display, , path
End Sub

Private Function MidPath(ByVal path As String, ByVal root As String) As String
    If StrComp(Left$(path, Len(root)), root, vbTextCompare) = 0 Then
        MidPath = Mid$(path, Len(root) + 2)
    Else
        MidPath = path
    End If
End Function

' ---------------------------------------------------------------------
'  Open a result
' ---------------------------------------------------------------------

Private Sub lstResults_DblClick()
    On Error Resume Next
    Dim i As Long
    i = lstResults.ListIndex
    If i < 0 Or i >= mResCount Then Exit Sub
    OpenResult mFile(i), mLine(i)
End Sub

Private Sub OpenResult(ByVal path As String, ByVal fileLine As Long)
    On Error Resume Next
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponentByFile(path)

    If comp Is Nothing Then
        Shell "notepad.exe """ & path & """", vbNormalFocus
        Exit Sub
    End If

    ' File line -> module line: the code module is the tail of the
    ' file, after the designer block and Attribute header lines.
    Dim offset As Long, modLine As Long
    offset = CountFileLines(path) - comp.CodeModule.CountOfLines
    modLine = fileLine - offset

    If modLine >= 1 Then
        Dim cp As VBIDE.CodePane, s As String
        Set cp = comp.CodeModule.CodePane
        s = comp.CodeModule.lines(modLine, 1)
        cp.SetSelection modLine, 1, modLine, Len(s) + 1
        Dim t As Long
        t = modLine - cp.CountOfVisibleLines \ 2
        If t < 1 Then t = 1
        cp.topLine = t
        cp.Show
        cp.Window.SetFocus
    Else
        ' hit is in the designer/header part: open the designer
        Dim w As VBIDE.Window
        Set w = comp.DesignerWindow
        If Not w Is Nothing Then
            w.Visible = True
            w.SetFocus
        Else
            comp.CodeModule.CodePane.Show
        End If
    End If
End Sub

Private Function FindComponentByFile(ByVal path As String) As VBIDE.VBComponent
    On Error Resume Next
    Dim proj As VBIDE.VBProject, comp As VBIDE.VBComponent, j As Long
    For Each proj In gVBE.VBProjects
        For Each comp In proj.VBComponents
            For j = 1 To comp.FileCount
                If StrComp(comp.FileNames(j), path, vbTextCompare) = 0 Then
                    Set FindComponentByFile = comp
                    Exit Function
                End If
            Next
        Next
    Next
End Function

Private Function CountFileLines(ByVal path As String) As Long
    On Error GoTo Fail
    Dim ff As Integer, s As String, n As Long
    ff = FreeFile
    Open path For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, s
        n = n + 1
    Loop
    Close #ff
    CountFileLines = n
    Exit Function
Fail:
    Close #ff
End Function

' ---------------------------------------------------------------------

Private Function BrowseFolder() As String
    On Error Resume Next
    Dim bi As BROWSEINFO, pidl As Long, s As String
    bi.hOwner = Me.hwnd
    bi.lpszTitle = "Choose the folder to search"
    bi.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    pidl = SHBrowseForFolderA(bi)
    If pidl <> 0 Then
        s = String$(260, 0)
        If SHGetPathFromIDListA(pidl, s) <> 0 Then
            BrowseFolder = Left$(s, InStr(s, Chr$(0)) - 1)
        End If
        CoTaskMemFree pidl
    End If
End Function

Private Sub Form_Load()
    Theme_ApplyIcon Me
End Sub
