Attribute VB_Name = "modBookmarks"
Option Explicit

' =====================================================================
'  Persistent bookmarks. Ctrl+F2 toggles, F2 jumps to the next one.
'  Stored per project in "<project>.vbp.bookmarks" beside the .vbp,
'  as "component|line|line text". The text snapshot lets a bookmark
'  re-find its line after edits move it (nearest matching line wins).
'  Painted by modHighlight: blue margin squares + scrollbar marks.
' =====================================================================

Private Type BookmarkInfo
    Comp As String
    LineNum As Long
    LineText As String
End Type

Private mBM() As BookmarkInfo
Private mCount As Long
Private mProjFile As String
Private mInit As Boolean

' Called from the tab-bar timer: (re)load when the project changes.
Public Sub BM_Poll()
    On Error Resume Next
    Dim f As String
    f = ActiveProjFile()
    If mInit And f = mProjFile Then Exit Sub
    mInit = True
    mProjFile = f
    LoadFromDisk
    Highlight_InvalidateAll
End Sub

Public Function BM_Count() As Long
    BM_Count = mCount
End Function

' Bookmark lines for one component (for the painters).
Public Function BM_LinesForComp(ByVal comp As String, outLines() As Long) As Long
    Dim i As Long, n As Long
    ReDim outLines(0 To 8)
    For i = 0 To mCount - 1
        If StrComp(mBM(i).Comp, comp, vbTextCompare) = 0 Then
            If n > UBound(outLines) Then ReDim Preserve outLines(0 To n * 2)
            outLines(n) = mBM(i).LineNum
            n = n + 1
        End If
    Next
    BM_LinesForComp = n
End Function

' ---------------------------------------------------------------------

Public Sub BM_Toggle()
    On Error Resume Next
    BM_Poll
    Dim cp As VBIDE.CodePane
    Set cp = gVBE.ActiveCodePane
    If cp Is Nothing Then Exit Sub

    Dim sl As Long, sc As Long, el As Long, ec As Long, comp As String
    cp.GetSelection sl, sc, el, ec
    comp = cp.CodeModule.Parent.Name

    Dim i As Long, j As Long
    For i = 0 To mCount - 1
        If mBM(i).LineNum = sl And _
           StrComp(mBM(i).Comp, comp, vbTextCompare) = 0 Then
            For j = i To mCount - 2
                mBM(j) = mBM(j + 1)
            Next
            mCount = mCount - 1
            SaveToDisk
            Highlight_InvalidateAll
            Exit Sub
        End If
    Next

    If mCount = 0 Then
        ReDim mBM(0 To 8)
    ElseIf mCount > UBound(mBM) Then
        ReDim Preserve mBM(0 To mCount * 2)
    End If
    mBM(mCount).Comp = comp
    mBM(mCount).LineNum = sl
    mBM(mCount).LineText = cp.CodeModule.lines(sl, 1)
    mCount = mCount + 1
    SaveToDisk
    Highlight_InvalidateAll
End Sub

Public Sub BM_NextBookmark()
    On Error Resume Next
    BM_Poll
    If mCount = 0 Then Beep: Exit Sub

    Dim cp As VBIDE.CodePane, comp As String
    Dim sl As Long, sc As Long, el As Long, ec As Long
    Set cp = gVBE.ActiveCodePane
    If Not cp Is Nothing Then
        cp.GetSelection sl, sc, el, ec
        comp = cp.CodeModule.Parent.Name
    End If

    ' next in this module after the caret, else the first bookmark in
    ' another module, else wrap to the first in this module
    Dim i As Long, best As Long
    best = -1
    For i = 0 To mCount - 1
        If StrComp(mBM(i).Comp, comp, vbTextCompare) = 0 Then
            If mBM(i).LineNum > sl Then
                If best = -1 Then
                    best = i
                ElseIf mBM(i).LineNum < mBM(best).LineNum Then
                    best = i
                End If
            End If
        End If
    Next
    If best = -1 Then
        For i = 0 To mCount - 1
            If StrComp(mBM(i).Comp, comp, vbTextCompare) <> 0 Then
                best = i
                Exit For
            End If
        Next
    End If
    If best = -1 Then
        For i = 0 To mCount - 1
            If best = -1 Then
                best = i
            ElseIf mBM(i).LineNum < mBM(best).LineNum Then
                best = i
            End If
        Next
    End If
    If best = -1 Then Exit Sub

    JumpTo best
End Sub

Public Sub BM_ClearAll()
    On Error Resume Next
    BM_Poll
    mCount = 0
    SaveToDisk
    Highlight_InvalidateAll
End Sub

' ---------------------------------------------------------------------

Private Sub JumpTo(ByVal idx As Long)
    On Error Resume Next
    Dim proj As VBIDE.VBProject
    Set proj = gVBE.ActiveVBProject
    If proj Is Nothing Then Exit Sub

    Dim cm As VBIDE.CodeModule
    Set cm = proj.VBComponents(mBM(idx).Comp).CodeModule
    If cm Is Nothing Then Exit Sub
    Resync idx, cm

    Dim m As MatchInfo
    m.Proj = proj.Name
    m.Comp = mBM(idx).Comp
    m.LineNum = mBM(idx).LineNum
    m.Col = 1
    m.MatchLen = 0
    GoToMatch m
End Sub

' If the stored line no longer matches (code was edited above it),
' find the nearest line with the remembered text.
Private Sub Resync(ByVal idx As Long, cm As VBIDE.CodeModule)
    On Error Resume Next
    If Len(Trim$(mBM(idx).LineText)) = 0 Then Exit Sub
    Dim ln As Long, total As Long, d As Long
    ln = mBM(idx).LineNum
    total = cm.CountOfLines
    If ln >= 1 And ln <= total Then
        If cm.lines(ln, 1) = mBM(idx).LineText Then Exit Sub
    End If
    For d = 1 To total
        If ln - d >= 1 Then
            If cm.lines(ln - d, 1) = mBM(idx).LineText Then
                mBM(idx).LineNum = ln - d
                Exit Sub
            End If
        End If
        If ln + d <= total Then
            If cm.lines(ln + d, 1) = mBM(idx).LineText Then
                mBM(idx).LineNum = ln + d
                Exit Sub
            End If
        End If
        If ln - d < 1 And ln + d > total Then Exit For
    Next
End Sub

' ---------------------------------------------------------------------

Private Function ActiveProjFile() As String
    On Error Resume Next
    ActiveProjFile = gVBE.ActiveVBProject.FileName
End Function

Private Function StoreFile() As String
    If Len(mProjFile) > 0 Then StoreFile = mProjFile & ".bookmarks"
End Function

Private Sub SaveToDisk()
    On Error Resume Next
    Dim f As String, ff As Integer, i As Long
    f = StoreFile()
    If Len(f) = 0 Then Exit Sub
    If mCount = 0 Then
        If Len(Dir$(f)) > 0 Then Kill f
        Exit Sub
    End If
    ff = FreeFile
    Open f For Output As #ff
    For i = 0 To mCount - 1
        Print #ff, mBM(i).Comp & "|" & mBM(i).LineNum & "|" & mBM(i).LineText
    Next
    Close #ff
End Sub

Private Sub LoadFromDisk()
    On Error Resume Next
    Dim f As String, ff As Integer, s As String, p() As String
    mCount = 0
    ReDim mBM(0 To 8)
    f = StoreFile()
    If Len(f) = 0 Then Exit Sub
    If Len(Dir$(f)) = 0 Then Exit Sub
    ff = FreeFile
    Open f For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, s
        p = Split(s, "|", 3)
        If UBound(p) = 2 Then
            If mCount > UBound(mBM) Then ReDim Preserve mBM(0 To mCount * 2)
            mBM(mCount).Comp = p(0)
            mBM(mCount).LineNum = CLng(p(1))
            mBM(mCount).LineText = p(2)
            mCount = mCount + 1
        End If
    Loop
    Close #ff
End Sub
