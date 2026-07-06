Attribute VB_Name = "modSearch"
Option Explicit

' =====================================================================
'  Search engine shared by frmFind (in-IDE search) and frmFindFiles.
'
'  Plain search uses InStr; regex uses VBScript.RegExp (vbscript.dll,
'  present on every Windows install - late bound, no reference needed).
' =====================================================================

Public Type MatchInfo
    Proj As String
    Comp As String
    LineNum As Long       ' 1-based module line
    Col As Long           ' 1-based column
    MatchLen As Long
    LineText As String
End Type

Public Enum SearchScope
    scCurrentDoc = 0
    scSelection = 1
    scOpenDocs = 2
    scProject = 3
End Enum

' Options set by the Find dialog before every operation
Public gOptCase As Boolean
Public gOptWhole As Boolean
Public gOptRegex As Boolean

' Selection range captured when the "Selection" scope is armed
Public gSelComp As String
Public gSelSL As Long, gSelSC As Long
Public gSelEL As Long, gSelEC As Long

' Result of the last CollectMatches call
Public gMatches() As MatchInfo
Public gMatchCount As Long

Private mRX As Object   ' VBScript.RegExp

' ---------------------------------------------------------------------

' Compile options + pattern. Returns False for an invalid regex.
Public Function PrepareSearch(ByVal sFind As String) As Boolean
    On Error Resume Next
    Set mRX = Nothing
    If Not gOptRegex Then PrepareSearch = (Len(sFind) > 0): Exit Function

    Set mRX = CreateObject("VBScript.RegExp")
    If mRX Is Nothing Then Exit Function
    Dim pat As String
    pat = sFind
    If gOptWhole Then pat = "\b(?:" & pat & ")\b"
    mRX.Pattern = pat
    mRX.IgnoreCase = Not gOptCase
    mRX.Global = True
    Err.Clear
    mRX.Test ""                       ' forces pattern compilation
    PrepareSearch = (Err.Number = 0)
End Function

' All hits within one line. Returns count; fills 0-based cols()/lens()
' with 1-based columns.
Public Function FindInLine(ByVal sLine As String, ByVal sFind As String, _
        cols() As Long, lens() As Long) As Long
    Dim n As Long
    ReDim cols(0 To 8): ReDim lens(0 To 8)

    If gOptRegex Then
        If mRX Is Nothing Then Exit Function
        On Error Resume Next
        Dim mc As Object, m As Object
        Set mc = mRX.Execute(sLine)
        If mc Is Nothing Then Exit Function
        For Each m In mc
            If m.Length > 0 Then AddHit cols, lens, n, m.FirstIndex + 1, m.Length
        Next
    Else
        Dim p As Long, cmp As VbCompareMethod
        If Len(sFind) = 0 Then Exit Function
        cmp = IIf(gOptCase, vbBinaryCompare, vbTextCompare)
        p = InStr(1, sLine, sFind, cmp)
        Do While p > 0
            If (Not gOptWhole) Or IsWholeWordAt(sLine, p, Len(sFind)) Then
                AddHit cols, lens, n, p, Len(sFind)
            End If
            p = InStr(p + 1, sLine, sFind, cmp)
        Loop
    End If
    FindInLine = n
End Function

Private Sub AddHit(cols() As Long, lens() As Long, n As Long, _
        ByVal c As Long, ByVal l As Long)
    If n > UBound(cols) Then
        ReDim Preserve cols(0 To n * 2)
        ReDim Preserve lens(0 To n * 2)
    End If
    cols(n) = c
    lens(n) = l
    n = n + 1
End Sub

Public Function IsWholeWordAt(ByVal s As String, ByVal pos As Long, _
        ByVal length As Long) As Boolean
    Dim ok As Boolean
    ok = True
    If pos > 1 Then
        If IsWordChar(Mid$(s, pos - 1, 1)) Then ok = False
    End If
    If ok And pos + length <= Len(s) Then
        If IsWordChar(Mid$(s, pos + length, 1)) Then ok = False
    End If
    IsWholeWordAt = ok
End Function

Public Function IsWordChar(ByVal ch As String) As Boolean
    IsWordChar = (ch Like "[A-Za-z0-9_]")
End Function

' Open the match's module, select it and scroll it into view.
Public Sub GoToMatch(m As MatchInfo)
    On Error Resume Next
    Dim cm As VBIDE.CodeModule, cp As VBIDE.CodePane
    Set cm = FindModule(m.Proj, m.Comp)
    If cm Is Nothing Then Exit Sub
    Set cp = cm.CodePane                 ' opens the pane if needed
    cp.Show
    cp.SetSelection m.LineNum, m.Col, m.LineNum, m.Col + m.MatchLen
    If m.LineNum < cp.topLine Or _
       m.LineNum >= cp.topLine + cp.CountOfVisibleLines Then
        Dim t As Long
        t = m.LineNum - cp.CountOfVisibleLines \ 2
        If t < 1 Then t = 1
        cp.topLine = t
    End If
    cp.Window.SetFocus
End Sub

' ---------------------------------------------------------------------
'  Scope enumeration + collection
' ---------------------------------------------------------------------

' Fills gMatches()/gMatchCount for the given scope. Returns the count.
Public Function CollectMatches(ByVal scope As SearchScope, _
        ByVal sFind As String) As Long
    On Error Resume Next
    gMatchCount = 0
    ReDim gMatches(0 To 64)

    Select Case scope
    Case scCurrentDoc
        If Not gVBE.ActiveCodePane Is Nothing Then
            ScanModule gVBE.ActiveCodePane.CodeModule, sFind, False
        End If
    Case scSelection
        If Not gVBE.ActiveCodePane Is Nothing Then
            ScanModule gVBE.ActiveCodePane.CodeModule, sFind, True
        End If
    Case scOpenDocs
        Dim cp As VBIDE.CodePane
        Dim done As New Collection
        For Each cp In gVBE.CodePanes
            Dim key As String
            key = ModuleKey(cp.CodeModule)
            If Not KeyExists(done, key) Then
                done.Add True, key
                ScanModule cp.CodeModule, sFind, False
            End If
        Next
    Case scProject
        Dim comp As VBIDE.VBComponent
        If Not gVBE.ActiveVBProject Is Nothing Then
            For Each comp In gVBE.ActiveVBProject.VBComponents
                ScanModule comp.CodeModule, sFind, False
            Next
        End If
    End Select

    CollectMatches = gMatchCount
End Function

Private Function KeyExists(col As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = col(key)
    KeyExists = (Err.Number = 0)
    Err.Clear
End Function

Public Function ModuleKey(ByVal cm As VBIDE.CodeModule) As String
    On Error Resume Next
    ModuleKey = cm.Parent.Collection.Parent.Name & "|" & cm.Parent.Name
End Function

Private Sub ScanModule(ByVal cm As VBIDE.CodeModule, ByVal sFind As String, _
        ByVal selectionOnly As Boolean)
    On Error Resume Next
    If cm Is Nothing Then Exit Sub

    Dim projName As String, compName As String
    compName = cm.Parent.Name
    projName = cm.Parent.Collection.Parent.Name

    If selectionOnly Then
        If compName <> gSelComp Then Exit Sub
    End If

    Dim i As Long, s As String
    Dim cols() As Long, lens() As Long, n As Long, k As Long
    Dim lo As Long, hi As Long
    lo = 1: hi = cm.CountOfLines
    If selectionOnly Then
        lo = gSelSL: hi = gSelEL
        If hi > cm.CountOfLines Then hi = cm.CountOfLines
    End If

    For i = lo To hi
        s = cm.lines(i, 1)
        n = FindInLine(s, sFind, cols, lens)
        For k = 0 To n - 1
            If selectionOnly Then
                If i = gSelSL And cols(k) < gSelSC Then GoTo NextHit
                If i = gSelEL And cols(k) + lens(k) > gSelEC Then GoTo NextHit
            End If
            If gMatchCount > UBound(gMatches) Then
                ReDim Preserve gMatches(0 To gMatchCount * 2)
            End If
            With gMatches(gMatchCount)
                .Proj = projName
                .Comp = compName
                .LineNum = i
                .Col = cols(k)
                .MatchLen = lens(k)
                .LineText = s
            End With
            gMatchCount = gMatchCount + 1
NextHit:
        Next
    Next
End Sub

' ---------------------------------------------------------------------
'  Replace
' ---------------------------------------------------------------------

' Replace every hit in one line; right-to-left so columns stay valid.
' Returns the new line; nCount is incremented per replacement.
Public Function ReplaceInLine(ByVal sLine As String, ByVal sFind As String, _
        ByVal sRepl As String, nCount As Long) As String
    Dim cols() As Long, lens() As Long, n As Long, k As Long
    Dim s As String
    s = sLine
    n = FindInLine(sLine, sFind, cols, lens)
    If n = 0 Then ReplaceInLine = sLine: Exit Function

    If gOptRegex Then
        On Error Resume Next
        s = mRX.Replace(sLine, sRepl)   ' supports $1..$9 group refs
        nCount = nCount + n
    Else
        For k = n - 1 To 0 Step -1
            s = Left$(s, cols(k) - 1) & sRepl & Mid$(s, cols(k) + lens(k))
            nCount = nCount + 1
        Next
    End If
    ReplaceInLine = s
End Function

' Replace across a whole scope. Returns number of replacements.
Public Function ReplaceAllInScope(ByVal scope As SearchScope, _
        ByVal sFind As String, ByVal sRepl As String) As Long
    On Error Resume Next
    Dim total As Long

    ' Collect first so we know exactly which module lines are affected.
    CollectMatches scope, sFind
    If gMatchCount = 0 Then Exit Function

    ' Process distinct (module, line) pairs; matches are in order.
    Dim i As Long
    Dim cm As VBIDE.CodeModule
    Dim lastKey As String, lineDone As String
    For i = 0 To gMatchCount - 1
        Dim key As String
        key = gMatches(i).Proj & "|" & gMatches(i).Comp
        If key <> lastKey Then
            Set cm = FindModule(gMatches(i).Proj, gMatches(i).Comp)
            lastKey = key
            lineDone = ""
        End If
        If Not cm Is Nothing Then
            Dim tag As String
            tag = "|" & gMatches(i).LineNum & "|"
            If InStr(lineDone, tag) = 0 Then
                lineDone = lineDone & tag
                Dim oldLine As String, newLine As String
                oldLine = cm.lines(gMatches(i).LineNum, 1)
                If scope = scSelection Then
                    newLine = ReplaceInLineRange(oldLine, sFind, sRepl, _
                              gMatches(i).LineNum, total)
                Else
                    newLine = ReplaceInLine(oldLine, sFind, sRepl, total)
                End If
                If newLine <> oldLine Then
                    cm.ReplaceLine gMatches(i).LineNum, newLine
                End If
            End If
        End If
    Next

    ReplaceAllInScope = total
End Function

' Selection scope: only replace hits inside the stored range.
Private Function ReplaceInLineRange(ByVal sLine As String, ByVal sFind As String, _
        ByVal sRepl As String, ByVal lineNum As Long, nCount As Long) As String
    Dim cols() As Long, lens() As Long, n As Long, k As Long
    Dim s As String
    s = sLine
    n = FindInLine(sLine, sFind, cols, lens)
    For k = n - 1 To 0 Step -1
        If Not (lineNum = gSelSL And cols(k) < gSelSC) Then
            If Not (lineNum = gSelEL And cols(k) + lens(k) > gSelEC) Then
                Dim piece As String, newPiece As String, dummy As Long
                piece = Mid$(s, cols(k), lens(k))
                If gOptRegex Then
                    newPiece = mRX.Replace(piece, sRepl)
                Else
                    newPiece = sRepl
                End If
                s = Left$(s, cols(k) - 1) & newPiece & Mid$(s, cols(k) + lens(k))
                nCount = nCount + 1
            End If
        End If
    Next
    ReplaceInLineRange = s
End Function

Public Function FindModule(ByVal projName As String, _
        ByVal compName As String) As VBIDE.CodeModule
    On Error Resume Next
    Dim proj As VBIDE.VBProject
    For Each proj In gVBE.VBProjects
        If proj.Name = projName Then
            Set FindModule = proj.VBComponents(compName).CodeModule
            Exit Function
        End If
    Next
End Function
