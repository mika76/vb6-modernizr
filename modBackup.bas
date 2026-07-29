Attribute VB_Name = "modBackup"
Option Explicit

' =====================================================================
'  Auto-backup: zips the active project into "<projdir>\.backups" every
'  INTERVAL_MIN minutes, but only if some project file changed since the
'  last backup. What goes in is the project folder PLUS every component
'  file linked in from elsewhere - a .vbp routinely pulls in shared
'  modules from sibling folders, and zipping the folder alone left those
'  out of every backup while ChangedSince still triggered on them.
'  Entries are stored relative to the deepest folder common to all of
'  them, so the original layout survives a restore.
'  The zip runs in a hidden PowerShell (System.IO.Compression, ships
'  with Windows) so the IDE never waits; the file list is handed over in
'  a temp file rather than on the command line. Oldest archives are
'  rotated out beyond KEEP_COUNT.
'  Toggle + "Backup Now" live in the Modernizr menu; the enabled flag
'  persists via SaveSetting.
' =====================================================================

Private Const INTERVAL_MIN As Long = 10
Private Const KEEP_COUNT As Long = 20
Private Const TICKS_PER_CHECK As Long = 75    ' ~30 s at 400 ms

Private mEnabled As Boolean
Private mTicks As Long
Private mLastBackup As Date

Public Sub Backup_Init()
    On Error Resume Next
    mEnabled = (GetSetting("VB6Modernizr", "Backup", "Enabled", "1") = "1")
    mLastBackup = CDate(Val(GetSetting("VB6Modernizr", "Backup", "Last", "0")))
End Sub

Public Function Backup_Enabled() As Boolean
    Backup_Enabled = mEnabled
End Function

Public Sub Backup_Toggle()
    On Error Resume Next
    mEnabled = Not mEnabled
    SaveSetting "VB6Modernizr", "Backup", "Enabled", IIf(mEnabled, "1", "0")
    MsgBox "Auto-backup is now " & IIf(mEnabled, "ON", "OFF") & "." & _
           vbCrLf & vbCrLf & "Every " & INTERVAL_MIN & " minutes (when " & _
           "files changed), keeping the last " & KEEP_COUNT & _
           " zips in the project's .backups folder.", _
           vbInformation, "VB6 Modernizr"
End Sub

' called from the tab-bar timer (~400 ms)
Public Sub Backup_Poll()
    On Error Resume Next
    mTicks = mTicks + 1
    If mTicks < TICKS_PER_CHECK Then Exit Sub
    mTicks = 0
    If Not mEnabled Then Exit Sub
    If DateDiff("n", mLastBackup, Now) < INTERVAL_MIN Then Exit Sub
    If Not ChangedSince(mLastBackup) Then Exit Sub
    Backup_Now False
End Sub

Public Sub Backup_Now(ByVal interactive As Boolean)
    On Error Resume Next
    Dim projDir As String
    projDir = ActiveProjDir()
    If Len(projDir) = 0 Then
        If interactive Then MsgBox "Save the project first.", _
            vbExclamation, "VB6 Modernizr"
        Exit Sub
    End If

    Dim bdir As String
    bdir = projDir & "\.backups"
    If Len(Dir$(bdir, vbDirectory)) = 0 Then MkDir bdir
    Rotate bdir

    Dim files As New Collection
    CollectProjectFiles projDir, bdir, files
    If files.Count = 0 Then
        If interactive Then MsgBox "Nothing to back up.", _
            vbExclamation, "VB6 Modernizr"
        Exit Sub
    End If

    Dim listFile As String
    listFile = WriteListFile(files, CommonRoot(files))
    If Len(listFile) = 0 Then
        If interactive Then MsgBox "Could not write the backup file list.", _
            vbExclamation, "VB6 Modernizr"
        Exit Sub
    End If

    Dim zip As String
    zip = bdir & "\backup_" & Format$(Now, "yyyymmdd_hhnnss") & ".zip"
    Shell "powershell -NoProfile -ExecutionPolicy Bypass -Command """ & _
          ZipCommand(listFile, zip) & """", vbHide

    mLastBackup = Now
    SaveSetting "VB6Modernizr", "Backup", "Last", CStr(CDbl(Now))
    If interactive Then
        MsgBox "Backup started (runs in the background):" & vbCrLf & zip, _
               vbInformation, "VB6 Modernizr"
    End If
End Sub

' ---------------------------------------------------------------------

Private Function ActiveProjDir() As String
    On Error Resume Next
    Dim f As String, p As Long
    f = gVBE.ActiveVBProject.FileName
    p = InStrRev(f, "\")
    If p > 0 Then ActiveProjDir = Left$(f, p - 1)
End Function

' any project file (or the .vbp itself) newer than t?
Private Function ChangedSince(ByVal t As Date) As Boolean
    On Error Resume Next
    Dim proj As VBIDE.VBProject, comp As VBIDE.VBComponent, j As Long
    Set proj = gVBE.ActiveVBProject
    If proj Is Nothing Then Exit Function
    If FileDateTime(proj.FileName) > t Then ChangedSince = True: Exit Function
    For Each comp In proj.VBComponents
        For j = 1 To comp.FileCount
            If FileDateTime(comp.FileNames(j)) > t Then
                ChangedSince = True
                Exit Function
            End If
        Next
    Next
End Function

' ---------------------------------------------------------------------
'  What to back up
' ---------------------------------------------------------------------

' The project folder, plus every component file the .vbp links in from
' outside it. Order matters: the folder walk owns Dir$ until it is done,
' so nothing that calls Dir$ may run before it finishes.
Private Sub CollectProjectFiles(ByVal projDir As String, ByVal bdir As String, _
        files As Collection)
    On Error Resume Next
    Dim seen As New Collection
    AddTree projDir, bdir, files, seen

    Dim proj As VBIDE.VBProject, comp As VBIDE.VBComponent, j As Long
    Set proj = gVBE.ActiveVBProject
    If proj Is Nothing Then Exit Sub
    AddFile proj.FileName, files, seen

    For Each comp In proj.VBComponents
        For j = 1 To comp.FileCount
            AddFile comp.FileNames(j), files, seen
            AddFile Companion(comp.FileNames(j)), files, seen
        Next
    Next
End Sub

' Whole folder tree, minus the .backups folder itself. Dir$() is not
' reentrant, so each directory pass finishes before recursing.
Private Sub AddTree(ByVal path As String, ByVal skipDir As String, _
        files As Collection, seen As Collection)
    On Error Resume Next
    If StrComp(path, skipDir, vbTextCompare) = 0 Then Exit Sub

    Dim f As String, subdirs As New Collection, full As String
    f = Dir$(path & "\*", vbDirectory Or vbHidden Or vbSystem)
    Do While Len(f) > 0
        If f <> "." And f <> ".." Then
            full = path & "\" & f
            If (GetAttr(full) And vbDirectory) <> 0 Then
                subdirs.Add full
            Else
                AddKnown full, files, seen      ' no Dir$ in here
            End If
        End If
        f = Dir$()
    Loop

    Dim i As Long
    For i = 1 To subdirs.Count
        AddTree CStr(subdirs(i)), skipDir, files, seen
    Next
End Sub

' Existence-checked add. Uses Dir$, so never call it during a walk.
Private Sub AddFile(ByVal path As String, files As Collection, _
        seen As Collection)
    On Error Resume Next
    If Len(path) = 0 Then Exit Sub
    If Len(Dir$(path)) = 0 Then Exit Sub        ' never saved, or gone
    AddKnown path, files, seen
End Sub

Private Sub AddKnown(ByVal path As String, files As Collection, _
        seen As Collection)
    On Error Resume Next
    Err.Clear
    seen.Add True, LCase$(path)                 ' duplicate key -> had it
    If Err.Number = 0 Then files.Add path
    Err.Clear
End Sub

' Forms and their kin keep a binary half beside them. FileNames usually
' lists it, but it costs nothing to make sure.
Private Function Companion(ByVal path As String) As String
    Dim p As Long, base As String
    p = InStrRev(path, ".")
    If p = 0 Then Exit Function
    base = Left$(path, p)
    Select Case LCase$(Mid$(path, p + 1))
    Case "frm": Companion = base & "frx"
    Case "ctl": Companion = base & "ctx"
    Case "dsr": Companion = base & "dsx"
    Case "pag": Companion = base & "pgx"
    Case "dob": Companion = base & "dox"
    End Select
End Function

' ---------------------------------------------------------------------
'  Archive layout
' ---------------------------------------------------------------------

' Deepest folder every file sits under, so entries keep their real
' relative layout instead of colliding at the archive root. "" when the
' files span drives - RelOf falls back to a drive-named top folder.
Private Function CommonRoot(files As Collection) As String
    On Error Resume Next
    Dim i As Long, root As String, p As String
    root = ParentOf(CStr(files(1)))
    For i = 2 To files.Count
        p = ParentOf(CStr(files(i)))
        Do While Len(root) > 0
            If IsUnder(p, root) Then Exit Do
            root = ParentOf(root)               ' walk up one level
        Loop
        If Len(root) = 0 Then Exit Function
    Next
    CommonRoot = root
End Function

Private Function IsUnder(ByVal path As String, ByVal root As String) As Boolean
    If StrComp(path, root, vbTextCompare) = 0 Then IsUnder = True: Exit Function
    If Len(path) > Len(root) Then
        IsUnder = (StrComp(Left$(path, Len(root) + 1), root & "\", _
                           vbTextCompare) = 0)
    End If
End Function

Private Function ParentOf(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, "\")
    If p > 1 Then ParentOf = Left$(path, p - 1)
End Function

' Zip entry name: forward slashes, and a drive-named top folder for
' anything that shares no root with the rest.
Private Function RelOf(ByVal path As String, ByVal root As String) As String
    Dim s As String
    s = path
    If Len(root) > 0 Then
        If Len(path) > Len(root) Then
            If StrComp(Left$(path, Len(root) + 1), root & "\", _
                       vbTextCompare) = 0 Then s = Mid$(path, Len(root) + 2)
        End If
    End If
    RelOf = Replace$(Replace$(s, ":", ""), "\", "/")
End Function

' ---------------------------------------------------------------------
'  Handing the list to PowerShell
' ---------------------------------------------------------------------

' "<entry name><TAB><full path>" per line. A temp file keeps arbitrary
' paths off the command line entirely.
Private Function WriteListFile(files As Collection, ByVal root As String) As String
    On Error GoTo Fail
    Dim ff As Integer, i As Long, path As String, lf As String
    lf = Environ$("TEMP") & "\vb6mrz_bk" & Format$(Now, "yyyymmddhhnnss") & ".txt"
    ff = FreeFile
    Open lf For Output As #ff
    For i = 1 To files.Count
        path = CStr(files(i))
        Print #ff, RelOf(path, root) & vbTab & path
    Next
    Close #ff
    WriteListFile = lf
    Exit Function
Fail:
    Close #ff
End Function

' Compress-Archive cannot name its entries, so build the zip directly.
Private Function ZipCommand(ByVal listFile As String, ByVal zip As String) As String
    ZipCommand = _
        "Add-Type -AssemblyName System.IO.Compression.FileSystem; " & _
        "$l = '" & PsQuote(listFile) & "'; " & _
        "$d = '" & PsQuote(zip) & "'; " & _
        "if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Force }; " & _
        "try { " & _
          "$z = [IO.Compression.ZipFile]::Open($d, 'Create'); " & _
          "foreach ($e in Get-Content -LiteralPath $l -Encoding Default) { " & _
            "$p = $e.Split([char]9); " & _
            "if ($p.Length -eq 2) { try { " & _
              "[IO.Compression.ZipFileExtensions]::CreateEntryFromFile(" & _
                "$z, $p[1], $p[0]) " & _
            "} catch {} } " & _
          "}; " & _
          "$z.Dispose() " & _
        "} finally { Remove-Item -LiteralPath $l -Force }"
End Function

' apostrophes are the only escape hazard inside a single-quoted PS string
Private Function PsQuote(ByVal s As String) As String
    PsQuote = Replace$(s, "'", "''")
End Function

' ---------------------------------------------------------------------

' timestamped names sort chronologically; drop the oldest beyond quota
Private Sub Rotate(ByVal bdir As String)
    On Error Resume Next
    Dim f As String, names() As String, n As Long
    ReDim names(0 To 32)
    f = Dir$(bdir & "\backup_*.zip")
    Do While Len(f) > 0
        If n > UBound(names) Then ReDim Preserve names(0 To n * 2)
        names(n) = f
        n = n + 1
        f = Dir$()
    Loop
    If n < KEEP_COUNT Then Exit Sub

    Dim i As Long, j As Long, tmp As String
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If names(j) < names(i) Then
                tmp = names(i): names(i) = names(j): names(j) = tmp
            End If
        Next
    Next
    For i = 0 To n - KEEP_COUNT       ' leaves KEEP_COUNT-1 + the new one
        Kill bdir & "\" & names(i)
    Next
End Sub
