Attribute VB_Name = "modGit"
Option Explicit

' =====================================================================
'  Git integration. Everything shells out to git.exe (must be on
'  PATH); if there is no repo or no git, every feature degrades to
'  silence. Long-running calls are asynchronous: one hidden process
'  at a time, output redirected to a temp file, completion polled by
'  the tab-bar timer via Git_Poll. Only user-invoked actions (blame,
'  commit) run synchronously with a bounded wait.
'
'  Caches exposed to the UI:
'    Git_Branch / Git_RepoDirty          -> tab bar label
'    Git_IsFileChanged / Git_IsCompChanged -> tab dots, changes window
'    Git_MarksForComp                    -> changed-line markers
' =====================================================================

Private Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwAccess As Long, ByVal bInherit As Long, ByVal pid As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProc As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal H As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Const STILL_ACTIVE As Long = 259
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const SYNCHRONIZE As Long = &H100000

' --- async job runner -------------------------------------------------

Private Const JOB_NONE As Long = 0
Private Const JOB_STATUS As Long = 1
Private Const JOB_DIFF As Long = 2

Private mJobKind As Long
Private mJobKey As String          ' for JOB_DIFF: "fullpath|compname"
Private mJobOut As String
Private mJobProc As Long

Private mQKind() As Long
Private mQKey() As String
Private mQCount As Long

Private mTick As Long
Private mTmpSeq As Long

' --- repo state --------------------------------------------------------

Private mProjFile As String        ' project the root was detected for
Private mRepoRoot As String        ' "" = no repo / not detected
Private mBranch As String
Private mRepoDirty As Boolean
Private mChanged As Collection     ' key = lcase full path, item = status
Private mChangedKeys As Collection ' full paths of changed files, in order
Private mGeneration As Long        ' bumps on every status parse
Private mDiffGen As Collection     ' key = lcase path, item = generation

' --- changed-line markers ----------------------------------------------

Public Const GITK_ADD As Long = 1
Public Const GITK_MOD As Long = 2
Public Const GITK_DEL As Long = 3

Private Type GitMark
    Comp As String
    LineNum As Long                ' module line
    Kind As Long
End Type
Private mMarks() As GitMark
Private mMarkCount As Long

' =====================================================================
'  Public state accessors
' =====================================================================

Public Function Git_Branch() As String
    Git_Branch = mBranch
End Function

Public Function Git_RepoDirty() As Boolean
    Git_RepoDirty = mRepoDirty
End Function

Public Function Git_HasRepo() As Boolean
    Git_HasRepo = (Len(mRepoRoot) > 0)
End Function

Public Function Git_RepoRoot() As String
    Git_RepoRoot = mRepoRoot
End Function

Public Function Git_IsFileChanged(ByVal fullPath As String) As Boolean
    On Error Resume Next
    If mChanged Is Nothing Then Exit Function
    Dim v As String
    v = mChanged(LCase$(fullPath))
    Git_IsFileChanged = (Err.Number = 0)
    Err.Clear
End Function

' Is any file of this component (active project) modified?
Public Function Git_IsCompChanged(ByVal compName As String) As Boolean
    On Error Resume Next
    Dim comp As VBIDE.VBComponent, j As Long
    Set comp = gVBE.ActiveVBProject.VBComponents(compName)
    If comp Is Nothing Then Exit Function
    For j = 1 To comp.FileCount
        If Git_IsFileChanged(comp.FileNames(j)) Then
            Git_IsCompChanged = True
            Exit Function
        End If
    Next
End Function

Public Function Git_MarkCount() As Long
    Git_MarkCount = mMarkCount
End Function

Public Function Git_MarksForComp(ByVal comp As String, _
        outLines() As Long, outKinds() As Long) As Long
    Dim i As Long, n As Long
    ReDim outLines(0 To 8)
    ReDim outKinds(0 To 8)
    For i = 0 To mMarkCount - 1
        If StrComp(mMarks(i).Comp, comp, vbTextCompare) = 0 Then
            If n > UBound(outLines) Then
                ReDim Preserve outLines(0 To n * 2)
                ReDim Preserve outKinds(0 To n * 2)
            End If
            outLines(n) = mMarks(i).LineNum
            outKinds(n) = mMarks(i).Kind
            n = n + 1
        End If
    Next
    Git_MarksForComp = n
End Function

' Changed files for the changes window: parallel arrays of status
' code and repo-relative path. Returns the count.
Public Function Git_ChangedList(outStatus() As String, _
        outPath() As String) As Long
    On Error Resume Next
    Dim n As Long, key As Variant
    ReDim outStatus(0 To 8)
    ReDim outPath(0 To 8)
    If mChanged Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To mChangedKeys.Count
        Dim fp As String
        fp = mChangedKeys(i)
        If n > UBound(outStatus) Then
            ReDim Preserve outStatus(0 To n * 2)
            ReDim Preserve outPath(0 To n * 2)
        End If
        outStatus(n) = mChanged(LCase$(fp))
        outPath(n) = fp
        n = n + 1
    Next
    Git_ChangedList = n
End Function

' =====================================================================
'  Poll driver (called from the tab-bar timer, ~400ms)
' =====================================================================

Public Sub Git_Poll()
    On Error Resume Next
    UpdateRepoRoot
    If Len(mRepoRoot) = 0 Then Exit Sub

    ' finish the running job, if any
    If mJobKind <> JOB_NONE Then
        If ProcessDone(mJobProc) Then FinishJob
    End If

    ' start the next queued job
    If mJobKind = JOB_NONE And mQCount > 0 Then StartNextJob

    ' periodic status refresh (~5 s)
    mTick = mTick + 1
    If mTick >= 12 Then
        mTick = 0
        Enqueue JOB_STATUS, ""
    End If
End Sub

Public Sub Git_RefreshNow()
    On Error Resume Next
    mTick = 999
End Sub

Private Sub UpdateRepoRoot()
    On Error Resume Next
    Dim f As String
    f = gVBE.ActiveVBProject.FileName
    If f = mProjFile Then Exit Sub
    mProjFile = f
    mRepoRoot = FindRepoRoot(f)
    mBranch = ""
    mRepoDirty = False
    Set mChanged = Nothing
    Set mChangedKeys = Nothing
    Set mDiffGen = Nothing
    mMarkCount = 0
    mTick = 999                       ' force a status run soon
End Sub

' Walk up from the project file's folder looking for a .git entry.
Private Function FindRepoRoot(ByVal projFile As String) As String
    On Error Resume Next
    If Len(projFile) = 0 Then Exit Function
    Dim d As String, p As Long
    p = InStrRev(projFile, "\")
    If p = 0 Then Exit Function
    d = Left$(projFile, p - 1)
    Do While Len(d) > 2
        If Len(Dir$(d & "\.git", vbDirectory Or vbHidden)) > 0 Then
            FindRepoRoot = d
            Exit Function
        End If
        p = InStrRev(d, "\")
        If p = 0 Then Exit Do
        d = Left$(d, p - 1)
    Loop
End Function

' =====================================================================
'  Job machinery
' =====================================================================

Private Sub Enqueue(ByVal kind As Long, ByVal key As String)
    Dim i As Long
    For i = 0 To mQCount - 1
        If mQKind(i) = kind And mQKey(i) = key Then Exit Sub  ' already queued
    Next
    If mQCount = 0 Then
        ReDim mQKind(0 To 8)
        ReDim mQKey(0 To 8)
    ElseIf mQCount > UBound(mQKind) Then
        ReDim Preserve mQKind(0 To mQCount * 2)
        ReDim Preserve mQKey(0 To mQCount * 2)
    End If
    mQKind(mQCount) = kind
    mQKey(mQCount) = key
    mQCount = mQCount + 1
End Sub

Private Sub StartNextJob()
    On Error Resume Next
    Dim kind As Long, key As String, i As Long
    kind = mQKind(0)
    key = mQKey(0)
    For i = 0 To mQCount - 2
        mQKind(i) = mQKind(i + 1)
        mQKey(i) = mQKey(i + 1)
    Next
    mQCount = mQCount - 1

    Dim args As String
    Select Case kind
    Case JOB_STATUS
        args = "status --porcelain=v1 -b"
    Case JOB_DIFF
        Dim fp As String
        fp = Split(key, "|")(0)
        args = "diff -U0 -- """ & RelPath(fp) & """"
    End Select

    mJobOut = TempFile()
    Dim pid As Long
    pid = Shell("cmd /c cd /d """ & mRepoRoot & """ && git " & args & _
                " > """ & mJobOut & """ 2>&1", vbHide)
    If pid = 0 Then Exit Sub
    mJobProc = OpenProcess(PROCESS_QUERY_INFORMATION Or SYNCHRONIZE, 0, pid)
    If mJobProc = 0 Then Exit Sub
    mJobKind = kind
    mJobKey = key
End Sub

Private Function ProcessDone(ByVal hProc As Long) As Boolean
    Dim code As Long
    If hProc = 0 Then ProcessDone = True: Exit Function
    If GetExitCodeProcess(hProc, code) = 0 Then ProcessDone = True: Exit Function
    ProcessDone = (code <> STILL_ACTIVE)
End Function

Private Sub FinishJob()
    On Error Resume Next
    Dim kind As Long, key As String, outText As String
    kind = mJobKind
    key = mJobKey
    CloseHandle mJobProc
    mJobProc = 0
    mJobKind = JOB_NONE

    outText = ReadAll(mJobOut)
    If Len(Dir$(mJobOut)) > 0 Then Kill mJobOut

    Select Case kind
    Case JOB_STATUS
        ParseStatus outText
        EnqueueDiffsForOpenPanes
        frmTabs.ForceRedraw
        frmChanges.NotifyGitChanged
    Case JOB_DIFF
        ParseDiff key, outText
        Highlight_InvalidateAll
    End Select
End Sub

' =====================================================================
'  Parsers
' =====================================================================

Private Sub ParseStatus(ByVal outText As String)
    On Error Resume Next
    Dim lines() As String, i As Long, s As String
    Set mChanged = New Collection
    Set mChangedKeys = New Collection
    mBranch = ""
    mRepoDirty = False

    lines = Split(outText, vbLf)
    For i = 0 To UBound(lines)
        s = Replace$(lines(i), vbCr, "")
        If Left$(s, 3) = "## " Then
            mBranch = Mid$(s, 4)
            Dim p As Long
            p = InStr(mBranch, "...")
            If p > 0 Then mBranch = Left$(mBranch, p - 1)
        ElseIf Len(s) > 3 Then
            Dim st As String, rel As String, fp As String
            st = Left$(s, 2)
            rel = Mid$(s, 4)
            p = InStr(rel, " -> ")            ' renames: take the new name
            If p > 0 Then rel = Mid$(rel, p + 4)
            If Left$(rel, 1) = """" Then rel = Mid$(rel, 2, Len(rel) - 2)
            fp = mRepoRoot & "\" & Replace$(rel, "/", "\")
            mChanged.Add Trim$(st), LCase$(fp)
            mChangedKeys.Add fp
            mRepoDirty = True
        End If
    Next
    mGeneration = mGeneration + 1
End Sub

' Queue a diff for every open, git-modified code pane whose diff is
' older than the latest status.
Private Sub EnqueueDiffsForOpenPanes()
    On Error Resume Next
    If mDiffGen Is Nothing Then Set mDiffGen = New Collection
    Dim cp As VBIDE.CodePane, comp As VBIDE.VBComponent
    For Each cp In gVBE.CodePanes
        Set comp = cp.CodeModule.Parent
        If comp.FileCount > 0 Then
            Dim fp As String
            fp = comp.FileNames(1)
            If Git_IsFileChanged(fp) Then
                Dim gen As Long
                gen = -1
                Err.Clear
                gen = mDiffGen(LCase$(fp))
                Err.Clear
                If gen < mGeneration Then
                    Enqueue JOB_DIFF, fp & "|" & comp.Name
                End If
            Else
                ' file no longer modified: clear its marks
                RemoveMarksFor comp.Name
            End If
        End If
    Next
End Sub

Private Sub ParseDiff(ByVal key As String, ByVal outText As String)
    On Error Resume Next
    Dim parts() As String
    parts = Split(key, "|")
    If UBound(parts) < 1 Then Exit Sub
    Dim fp As String, compName As String
    fp = parts(0)
    compName = parts(1)

    ' remember freshness
    If mDiffGen Is Nothing Then Set mDiffGen = New Collection
    mDiffGen.Remove LCase$(fp)
    Err.Clear
    mDiffGen.Add mGeneration, LCase$(fp)

    ' file line -> module line offset (header/designer block)
    Dim cm As VBIDE.CodeModule
    Set cm = gVBE.ActiveVBProject.VBComponents(compName).CodeModule
    If cm Is Nothing Then Exit Sub
    Dim offset As Long
    offset = CountFileLines(fp) - cm.CountOfLines

    RemoveMarksFor compName

    Dim lines() As String, i As Long, s As String
    lines = Split(outText, vbLf)
    For i = 0 To UBound(lines)
        s = Replace$(lines(i), vbCr, "")
        If Left$(s, 2) = "@@" Then
            Dim newStart As Long, newCount As Long, oldCount As Long
            ParseHunkHeader s, newStart, newCount, oldCount
            Dim kind As Long, l As Long, lo As Long, hi As Long
            If newCount = 0 Then
                kind = GITK_DEL
                lo = newStart + 1: hi = lo      ' mark line after deletion
            ElseIf oldCount = 0 Then
                kind = GITK_ADD
                lo = newStart: hi = newStart + newCount - 1
            Else
                kind = GITK_MOD
                lo = newStart: hi = newStart + newCount - 1
            End If
            For l = lo To hi
                Dim ml As Long
                ml = l - offset
                If ml >= 1 And ml <= cm.CountOfLines Then
                    AddMark compName, ml, kind
                End If
            Next
        End If
    Next
End Sub

' "@@ -12,3 +14,5 @@ ..." -> newStart=14, newCount=5, oldCount=3
Private Sub ParseHunkHeader(ByVal s As String, newStart As Long, _
        newCount As Long, oldCount As Long)
    On Error Resume Next
    Dim p As Long, q As Long, tok As String
    newStart = 0: newCount = 1: oldCount = 1

    p = InStr(s, "-")
    q = InStr(p, s, " ")
    tok = Mid$(s, p + 1, q - p - 1)
    If InStr(tok, ",") > 0 Then oldCount = CLng(Mid$(tok, InStr(tok, ",") + 1))

    p = InStr(s, "+")
    q = InStr(p, s, " ")
    tok = Mid$(s, p + 1, q - p - 1)
    If InStr(tok, ",") > 0 Then
        newCount = CLng(Mid$(tok, InStr(tok, ",") + 1))
        tok = Left$(tok, InStr(tok, ",") - 1)
    End If
    newStart = CLng(tok)
End Sub

Private Sub AddMark(ByVal comp As String, ByVal ln As Long, ByVal kind As Long)
    If mMarkCount = 0 Then
        ReDim mMarks(0 To 32)
    ElseIf mMarkCount > UBound(mMarks) Then
        ReDim Preserve mMarks(0 To mMarkCount * 2)
    End If
    mMarks(mMarkCount).Comp = comp
    mMarks(mMarkCount).LineNum = ln
    mMarks(mMarkCount).Kind = kind
    mMarkCount = mMarkCount + 1
End Sub

Private Sub RemoveMarksFor(ByVal comp As String)
    Dim i As Long, n As Long
    For i = 0 To mMarkCount - 1
        If StrComp(mMarks(i).Comp, comp, vbTextCompare) <> 0 Then
            mMarks(n) = mMarks(i)
            n = n + 1
        End If
    Next
    mMarkCount = n
End Sub

' =====================================================================
'  Synchronous helpers (user-invoked actions)
' =====================================================================

Public Function Git_RunSync(ByVal args As String, _
        ByVal timeoutMs As Long) As String
    On Error Resume Next
    If Len(mRepoRoot) = 0 Then Exit Function
    Dim out As String, pid As Long, hProc As Long, waited As Long
    out = TempFile()
    pid = Shell("cmd /c cd /d """ & mRepoRoot & """ && git " & args & _
                " > """ & out & """ 2>&1", vbHide)
    If pid = 0 Then Exit Function
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or SYNCHRONIZE, 0, pid)
    Do While Not ProcessDone(hProc)
        Sleep 40
        DoEvents
        waited = waited + 40
        If waited > timeoutMs Then Exit Do
    Loop
    CloseHandle hProc
    Git_RunSync = ReadAll(out)
    If Len(Dir$(out)) > 0 Then Kill out
End Function

' Blame the current line: author, date and summary in a message box.
Public Sub Git_BlameCurrentLine()
    On Error Resume Next
    If Len(mRepoRoot) = 0 Then Beep: Exit Sub
    Dim cp As VBIDE.CodePane
    Set cp = gVBE.ActiveCodePane
    If cp Is Nothing Then Beep: Exit Sub

    Dim comp As VBIDE.VBComponent
    Set comp = cp.CodeModule.Parent
    If comp.FileCount = 0 Then Beep: Exit Sub

    Dim fp As String, sl As Long, sc As Long, el As Long, ec As Long
    fp = comp.FileNames(1)
    cp.GetSelection sl, sc, el, ec

    Dim fileLine As Long
    fileLine = sl + (CountFileLines(fp) - cp.CodeModule.CountOfLines)
    If fileLine < 1 Then Beep: Exit Sub

    Dim res As String
    res = Git_RunSync("blame -L " & fileLine & "," & fileLine & _
                      " --porcelain -- """ & RelPath(fp) & """", 5000)
    If Len(res) = 0 Then Beep: Exit Sub

    Dim sha As String, author As String, when As String, summary As String
    Dim lines() As String, i As Long, s As String
    lines = Split(res, vbLf)
    sha = Left$(lines(0), 8)
    For i = 1 To UBound(lines)
        s = Replace$(lines(i), vbCr, "")
        If Left$(s, 7) = "author " Then
            author = Mid$(s, 8)
        ElseIf Left$(s, 12) = "author-time " Then
            when = Format$(DateAdd("s", CDbl(Mid$(s, 13)), #1/1/1970#), _
                           "yyyy-mm-dd hh:nn")
        ElseIf Left$(s, 8) = "summary " Then
            summary = Mid$(s, 9)
        End If
    Next

    If Len(author) = 0 Then
        MsgBox "No blame info (unsaved or uncommitted line?).", _
               vbInformation, "Git Blame"
    Else
        MsgBox "Line " & sl & " of " & comp.Name & vbCrLf & vbCrLf & _
               "Commit:  " & sha & vbCrLf & _
               "Author:  " & author & vbCrLf & _
               "Date:    " & when & " (UTC)" & vbCrLf & _
               "Summary: " & summary, vbInformation, "Git Blame"
    End If
End Sub

' Stage everything and commit with the given message. Returns git's
' output. The message goes via a temp file to dodge quoting issues.
Public Function Git_CommitAll(ByVal msg As String) As String
    On Error Resume Next
    If Len(mRepoRoot) = 0 Then Exit Function
    Dim mf As String, ff As Integer
    mf = TempFile()
    ff = FreeFile
    Open mf For Output As #ff
    Print #ff, msg
    Close #ff

    Git_RunSync "add -A", 15000
    Git_CommitAll = Git_RunSync("commit -F """ & mf & """", 15000)
    If Len(Dir$(mf)) > 0 Then Kill mf
    Git_RefreshNow
End Function

' =====================================================================
'  Small utilities
' =====================================================================

Private Function RelPath(ByVal fullPath As String) As String
    RelPath = fullPath
    If StrComp(Left$(fullPath, Len(mRepoRoot)), mRepoRoot, vbTextCompare) = 0 Then
        RelPath = Mid$(fullPath, Len(mRepoRoot) + 2)
    End If
    RelPath = Replace$(RelPath, "\", "/")
End Function

Private Function TempFile() As String
    mTmpSeq = mTmpSeq + 1
    TempFile = Environ$("TEMP") & "\vb6mrz_" & CLng(Timer * 100) & "_" & _
               mTmpSeq & ".txt"
End Function

Private Function ReadAll(ByVal path As String) As String
    On Error Resume Next
    Dim ff As Integer
    If Len(Dir$(path)) = 0 Then Exit Function
    ff = FreeFile
    Open path For Binary Access Read As #ff
    If LOF(ff) > 0 Then
        ReadAll = Space$(LOF(ff))
        Get #ff, 1, ReadAll
    End If
    Close #ff
End Function

Public Function CountFileLines(ByVal path As String) As Long
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
