Attribute VB_Name = "modEditOps"
Option Explicit

' =====================================================================
'  Editing operations bound to keyboard shortcuts (see modWheel):
'    Ctrl+D        duplicate line/selection
'    Alt+Up/Down   move line/selection up/down
'    Ctrl+Shift+K  delete line/selection
'    Ctrl+/        comment/uncomment selection
'    Shift+F12     find all references to the word under the cursor
' =====================================================================

' Active pane's selected line block. If a multi-line selection ends at
' column 1, that last line is excluded (VS convention).
Private Function GetBlock(cp As VBIDE.CodePane, cm As VBIDE.CodeModule, _
        sl As Long, sc As Long, el As Long, ec As Long, _
        bs As Long, be As Long) As Boolean
    On Error Resume Next
    Set cp = gVBE.ActiveCodePane
    If cp Is Nothing Then Exit Function
    Set cm = cp.CodeModule
    cp.GetSelection sl, sc, el, ec
    bs = sl: be = el
    If be > bs And ec = 1 Then be = be - 1
    If bs < 1 Then Exit Function
    If be > cm.CountOfLines Then Exit Function
    GetBlock = True
End Function

Public Sub Edit_DuplicateLines()
    On Error Resume Next
    Dim cp As VBIDE.CodePane, cm As VBIDE.CodeModule
    Dim sl As Long, sc As Long, el As Long, ec As Long, bs As Long, be As Long
    If Not GetBlock(cp, cm, sl, sc, el, ec, bs, be) Then Exit Sub
    cm.InsertLines be + 1, cm.lines(bs, be - bs + 1)
    Dim n As Long
    n = be - bs + 1
    cp.SetSelection sl + n, sc, el + n, ec   ' caret follows the copy
End Sub

Public Sub Edit_MoveLinesUp()
    On Error Resume Next
    Dim cp As VBIDE.CodePane, cm As VBIDE.CodeModule
    Dim sl As Long, sc As Long, el As Long, ec As Long, bs As Long, be As Long
    If Not GetBlock(cp, cm, sl, sc, el, ec, bs, be) Then Exit Sub
    If bs <= 1 Then Exit Sub
    Dim moved As String
    moved = cm.lines(bs - 1, 1)
    cm.DeleteLines bs - 1, 1
    cm.InsertLines be, moved         ' block now sits at bs-1..be-1
    cp.SetSelection sl - 1, sc, el - 1, ec
End Sub

Public Sub Edit_MoveLinesDown()
    On Error Resume Next
    Dim cp As VBIDE.CodePane, cm As VBIDE.CodeModule
    Dim sl As Long, sc As Long, el As Long, ec As Long, bs As Long, be As Long
    If Not GetBlock(cp, cm, sl, sc, el, ec, bs, be) Then Exit Sub
    If be >= cm.CountOfLines Then Exit Sub
    Dim moved As String
    moved = cm.lines(be + 1, 1)
    cm.DeleteLines be + 1, 1
    cm.InsertLines bs, moved         ' pushes the block down one line
    cp.SetSelection sl + 1, sc, el + 1, ec
End Sub

Public Sub Edit_DeleteLines()
    On Error Resume Next
    Dim cp As VBIDE.CodePane, cm As VBIDE.CodeModule
    Dim sl As Long, sc As Long, el As Long, ec As Long, bs As Long, be As Long
    If Not GetBlock(cp, cm, sl, sc, el, ec, bs, be) Then Exit Sub
    cm.DeleteLines bs, be - bs + 1
    If bs > cm.CountOfLines Then bs = cm.CountOfLines
    If bs < 1 Then bs = 1
    cp.SetSelection bs, 1, bs, 1
End Sub

Public Sub Edit_ToggleComment()
    On Error Resume Next
    Dim cp As VBIDE.CodePane, cm As VBIDE.CodeModule
    Dim sl As Long, sc As Long, el As Long, ec As Long, bs As Long, be As Long
    If Not GetBlock(cp, cm, sl, sc, el, ec, bs, be) Then Exit Sub

    ' uncomment only if every non-blank line is already commented
    Dim i As Long, s As String, t As String
    Dim allC As Boolean, anyText As Boolean
    allC = True
    For i = bs To be
        t = Trim$(cm.lines(i, 1))
        If Len(t) > 0 Then
            anyText = True
            If Left$(t, 1) <> "'" Then allC = False: Exit For
        End If
    Next
    If Not anyText Then Exit Sub

    Dim p As Long, ind As Long
    For i = bs To be
        s = cm.lines(i, 1)
        If Len(Trim$(s)) > 0 Then
            If allC Then
                p = InStr(s, "'")
                cm.ReplaceLine i, Left$(s, p - 1) & Mid$(s, p + 1)
            Else
                ind = Len(s) - Len(LTrim$(s))
                cm.ReplaceLine i, Left$(s, ind) & "'" & Mid$(s, ind + 1)
            End If
        End If
    Next
    cp.SetSelection bs, 1, be, Len(cm.lines(be, 1)) + 1
End Sub

' ---------------------------------------------------------------------
'  Find all references
' ---------------------------------------------------------------------

Public Function WordUnderCursor() As String
    On Error Resume Next
    Dim cp As VBIDE.CodePane
    Set cp = gVBE.ActiveCodePane
    If cp Is Nothing Then Exit Function

    Dim sl As Long, sc As Long, el As Long, ec As Long, s As String
    cp.GetSelection sl, sc, el, ec
    s = cp.CodeModule.lines(sl, 1)

    If sl = el And ec > sc Then
        WordUnderCursor = Mid$(s, sc, ec - sc)
        Exit Function
    End If

    ' expand outward from the caret column
    Dim i As Long, st As Long, en As Long
    i = sc
    If i > Len(s) + 1 Then i = Len(s) + 1
    st = i
    Do While st > 1
        If Not IsWordChar(Mid$(s, st - 1, 1)) Then Exit Do
        st = st - 1
    Loop
    en = i
    Do While en <= Len(s)
        If Not IsWordChar(Mid$(s, en, 1)) Then Exit Do
        en = en + 1
    Loop
    WordUnderCursor = Mid$(s, st, en - st)
End Function

Public Sub Edit_FindAllReferences()
    On Error Resume Next
    Dim w As String
    w = WordUnderCursor()
    If Len(w) = 0 Then Beep: Exit Sub

    gOptRegex = False
    gOptWhole = True
    gOptCase = False
    If Not PrepareSearch(w) Then Exit Sub
    If CollectMatches(scProject, w) = 0 Then Beep: Exit Sub

    frmRefs.ShowRefs w
End Sub
